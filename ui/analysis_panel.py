"""
analysis_panel.py

Panel for analyzing fleet data and visualizing results in the
Fleet Electrification Analyzer.

Phase 18 redesign: top-to-bottom scrollable dashboard (findings first),
collapsible Parameters and Chart Browser sections.
"""

import os
import datetime
import logging
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, List, Any, Optional, Callable, Tuple

import io as _io

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

try:
    from PIL import Image as _PILImage, ImageTk as _ImageTk
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False

from settings import (
    PRIMARY_HEX_1,
    PRIMARY_HEX_2,
    PRIMARY_HEX_3,
    SECONDARY_HEX_1,
    DEFAULT_GAS_PRICE,
    DEFAULT_ELECTRICITY_PRICE,
    DEFAULT_EV_EFFICIENCY,
    DEFAULT_BATTERY_DEGRADATION,
    DEFAULT_RESIDUAL_VALUE_ICE_PCT,
    DEFAULT_RESIDUAL_VALUE_EV_PCT,
    CHART_TYPES
)
from utils import SimpleTooltip, ProgressDialog, ScrollableFrame
from data.models import FleetVehicle, Fleet
from ui.theme import Colors, Fonts, Spacing
from analysis.calculations import (
    analyze_fleet_electrification,
    create_emissions_inventory,
    analyze_charging_needs
)
from analysis.charts import ChartFactory
from analysis.reports import ReportGeneratorFactory, ExportCoordinator
from analysis.rate_database import get_rates_for_state, get_all_incentives, get_available_states
from analysis.scenarios import (
    compare_scenarios, PRESET_SCENARIOS, SCOPE_SCENARIO_KEYS,
    _get_vehicle_ev_cost, _get_vehicle_annual_co2, _get_vehicle_annual_savings,
)

logger = logging.getLogger(__name__)

try:
    from powerpoint_export import export_prelim_deck
    PPTX_EXPORT_AVAILABLE = True
except ImportError:
    PPTX_EXPORT_AVAILABLE = False
    logger.warning("PowerPoint export not available - powerpoint_export module not found")

CHARGING_POWER_LEVELS = {
    "LP": 7.2,
    "MP": 19.2,
    "HP": 50.0,
    "VHP": 150.0
}

# ACF category colours for donut / Gantt chart
ACF_COLORS = {
    "ZEV": "#48BB78",
    "A":   "#4299E1",
    "B":   "#ED8936",
    "C":   "#9F7AEA",
    "D":   "#F56565",
}

# Plain-English labels for ACF category codes (shared with timeline_panel)
ACF_LABELS = {
    "ZEV": "Already ZEV",
    "A":   "Light-Duty (Exempt)",
    "B":   "Mandate-Subject",
    "C":   "Body-Type Exempt",
    "D":   "Emergency Vehicle",
}


def _show_acf_ev_year_dialog(parent, old_acf: str, new_acf: str,
                              current_ev_year: str) -> str:
    """Modal dialog shown when the user changes a vehicle's ACF category.

    Warns that the ACF reclassification may affect the vehicle's EV year
    and offers three choices:
        'recalculate' — re-run the electrification year algorithm
        'keep'        — leave the current EV year unchanged
        'cancel'      — abort the ACF change entirely

    Args:
        parent:           Parent widget (for positioning).
        old_acf:          Original ACF code, e.g. "B".
        new_acf:          New ACF code, e.g. "D".
        current_ev_year:  Current "Proposed EV Year" value for display.

    Returns:
        One of 'recalculate', 'keep', or 'cancel'.
    """
    import tkinter as tk
    from tkinter import ttk

    old_label = ACF_LABELS.get(old_acf, old_acf)
    new_label = ACF_LABELS.get(new_acf, new_acf)

    result = tk.StringVar(value="cancel")
    dlg = tk.Toplevel(parent)
    dlg.title("ACF Category Changed")
    dlg.resizable(False, False)
    dlg.transient(parent)
    dlg.grab_set()

    try:
        dlg.geometry(
            f"+{parent.winfo_rootx() + 80}+{parent.winfo_rooty() + 80}"
        )
    except tk.TclError:
        pass

    msg_frame = ttk.Frame(dlg)
    msg_frame.pack(fill=tk.X, padx=20, pady=(16, 8))

    ttk.Label(
        msg_frame,
        text="ACF Reclassification",
        font=("TkDefaultFont", 11, "bold"),
    ).pack(anchor="w")

    ttk.Label(
        msg_frame,
        text=f"  From:  {old_acf} — {old_label}\n  To:      {new_acf} — {new_label}",
        justify="left",
    ).pack(anchor="w", pady=(4, 8))

    ttk.Label(
        msg_frame,
        text=f"Current proposed EV year: {current_ev_year}",
        foreground="#555",
    ).pack(anchor="w")

    ttk.Label(
        msg_frame,
        text="What should happen to the EV year assignment?",
        font=("TkDefaultFont", 10, "italic"),
    ).pack(anchor="w", pady=(8, 0))

    btn_frame = ttk.Frame(dlg)
    btn_frame.pack(fill=tk.X, padx=20, pady=(8, 16))

    def _pick(choice):
        result.set(choice)
        dlg.destroy()

    ttk.Button(
        btn_frame,
        text="Recalculate EV Year",
        style="Primary.TButton",
        command=lambda: _pick("recalculate"),
    ).pack(side=tk.LEFT, padx=(0, 6))

    keep_text = f"Keep {current_ev_year}" if current_ev_year not in ("N/A", "Exempt", "—", "") else "Keep As-Is"
    ttk.Button(
        btn_frame,
        text=keep_text,
        command=lambda: _pick("keep"),
    ).pack(side=tk.LEFT, padx=(0, 6))

    ttk.Button(
        btn_frame,
        text="Cancel",
        command=lambda: _pick("cancel"),
    ).pack(side=tk.LEFT)

    dlg.wait_window()
    return result.get()


def _compute_current_plan_result(vehicles: list) -> dict:
    """Build a scenario result from the fleet's current Proposed EV Year values.

    Reads custom_fields["Proposed EV Year"] directly, including any manual
    overrides, and returns a result dict in the same format as run_scenario().
    Safe to call from a background thread (pure computation).
    """
    current_year = datetime.datetime.now().year

    assignments = []
    for v in vehicles:
        yr_raw = v.custom_fields.get("Proposed EV Year", "")
        try:
            yr = int(yr_raw)
        except (ValueError, TypeError):
            continue
        if yr < current_year:
            continue
        assignments.append((yr, v))

    if not assignments:
        return {
            "name": "Current Plan",
            "description": "Fleet's current scheduled EV year assignments (includes overrides)",
            "end_year": current_year,
            "vehicle_filter": "all",
            "total_vehicles": 0,
            "vehicles_per_year": {}, "cumulative_vehicles": {},
            "cost_per_year": {}, "cumulative_cost": {},
            "co2_reduction_per_year": {}, "cumulative_co2_reduction": {},
            "savings_per_year": {}, "cumulative_savings": {},
            "total_investment": 0, "total_annual_savings": 0,
            "total_annual_co2_reduction": 0,
            "summary_text": "Current Plan: No vehicles with future scheduled EV years.",
        }

    all_years = sorted({yr for yr, _ in assignments})
    year_range = list(range(min(all_years), max(all_years) + 1))

    vpyr: dict = {}
    cpyr: dict = {}
    co2yr: dict = {}
    spyr: dict = {}
    for yr, v in assignments:
        vpyr[yr]  = vpyr.get(yr, 0) + 1
        cpyr[yr]  = cpyr.get(yr, 0.0) + _get_vehicle_ev_cost(v)
        co2yr[yr] = co2yr.get(yr, 0.0) + _get_vehicle_annual_co2(v)
        spyr[yr]  = spyr.get(yr, 0.0) + _get_vehicle_annual_savings(v)

    cumv = cumcost = cumco2 = cumsav = 0.0
    cv: dict = {}; cc: dict = {}; cco: dict = {}; cs: dict = {}
    for yr in year_range:
        cumv    += vpyr.get(yr, 0)
        cumcost += cpyr.get(yr, 0.0)
        cumco2  += co2yr.get(yr, 0.0)
        cumsav  += spyr.get(yr, 0.0)
        cv[yr]  = cumv;  cc[yr] = cumcost
        cco[yr] = cumco2; cs[yr] = cumsav

    return {
        "name": "Current Plan",
        "description": "Fleet's current scheduled EV year assignments (includes overrides)",
        "end_year": max(all_years),
        "vehicle_filter": "all",
        "total_vehicles": len(assignments),
        "vehicles_per_year": vpyr, "cumulative_vehicles": cv,
        "cost_per_year": cpyr, "cumulative_cost": cc,
        "co2_reduction_per_year": co2yr, "cumulative_co2_reduction": cco,
        "savings_per_year": spyr, "cumulative_savings": cs,
        "total_investment": cumcost,
        "total_annual_savings": cumsav,
        "total_annual_co2_reduction": cumco2,
        "summary_text": (
            f"Current Plan: {len(assignments)} vehicles scheduled through {max(all_years)}."
        ),
    }


GANTT_YEAR_MIN = 2026
GANTT_YEAR_MAX = 2040          # floor default; chart expands dynamically to fit fleet
_GANTT_CATEGORIES = ["B", "C", "D", "ZEV", "A"]


def _gantt_year_bounds(vehicles: list) -> tuple:
    """Return (year_min, year_max) for Gantt rendering.

    year_min is always GANTT_YEAR_MIN (2026).
    year_max is max(GANTT_YEAR_MAX, furthest scheduled EV year in the fleet),
    so overrides beyond 2040 remain visible.
    """
    max_yr = GANTT_YEAR_MAX
    for v in vehicles:
        yr_raw = v.custom_fields.get("Proposed EV Year", "")
        try:
            yr = int(yr_raw)
            if yr >= GANTT_YEAR_MIN:
                max_yr = max(max_yr, yr)
        except (ValueError, TypeError):
            continue
    return GANTT_YEAR_MIN, max_yr


def _draw_gantt_chart(ax, vehicles: list, view: str = "Grouped by ACF",
                      max_vehicles: int = 0,
                      scenario_results=None,
                      horizon_year: int = 2040) -> None:
    """Draw a Gantt-style timeline chart onto *ax* from a list of FleetVehicle.

    view: "Grouped by ACF"  — one row per ACF category, bubble sized by count
          "By Scenario"     — one row per scope scenario, bubble sized by count
          "Per Vehicle"     — one row per vehicle, dot at scheduled year
    max_vehicles: cap rows in Per Vehicle view (0 = show all)
    scenario_results: ScenarioResult list from compare_scenarios(); needed for By Scenario view
    horizon_year: upper X-axis bound used for By Scenario view (aligns with Scenario Comparison)
    """
    import matplotlib.ticker as mticker

    if view == "By Scenario":
        year_min = GANTT_YEAR_MIN
        year_max = max(horizon_year, GANTT_YEAR_MAX)
    else:
        year_min, year_max = _gantt_year_bounds(vehicles)

    years = list(range(year_min, year_max + 1))
    ax.set_xlim(year_min - 0.6, year_max + 0.6)
    ax.set_xticks(years)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: str(int(x))))
    ax.grid(axis="x", linestyle="--", alpha=0.4, zorder=0)
    ax.set_xlabel("Proposed EV Year", fontsize=8)
    ax.tick_params(axis="both", labelsize=7)

    if view == "Grouped by ACF":
        _gantt_grouped(ax, vehicles, years)
    elif view == "By Scenario":
        _gantt_by_scenario(ax, vehicles, years, scenario_results)
    else:
        _gantt_per_vehicle(ax, vehicles, years, max_vehicles=max_vehicles)


def _gantt_grouped(ax, vehicles: list, years: list) -> None:
    """Bubble chart: one row per ACF category, bubble size = vehicle count."""
    year_set = set(years)
    cats = [c for c in _GANTT_CATEGORIES
            if any(v.custom_fields.get("_acf_code") == c for v in vehicles)]
    if not cats:
        ax.text(0.5, 0.5, "No vehicles with ACF classifications.",
                ha="center", va="center", fontsize=9)
        ax.axis("off")
        return

    cat_pos = {c: i for i, c in enumerate(reversed(cats))}
    ax.set_yticks(list(cat_pos.values()))
    ax.set_yticklabels([ACF_LABELS.get(c, c) for c in reversed(cats)], fontsize=8)
    ax.set_ylim(-0.7, len(cats) - 0.3)

    for cat in cats:
        year_counts: dict = {}
        for v in vehicles:
            if v.custom_fields.get("_acf_code") != cat:
                continue
            yr_raw = v.custom_fields.get("Proposed EV Year", "")
            try:
                yr = int(yr_raw)
            except (ValueError, TypeError):
                continue
            if yr in year_set:
                year_counts[yr] = year_counts.get(yr, 0) + 1

        color = ACF_COLORS.get(cat, "#AAAAAA")
        for yr, cnt in year_counts.items():
            size = 200 + cnt * 80
            ax.scatter(yr, cat_pos[cat], s=size, color=color,
                       alpha=0.85, zorder=5, edgecolors="white", linewidths=0.8)
            ax.text(yr, cat_pos[cat], str(cnt),
                    ha="center", va="center", fontsize=7,
                    color="white", fontweight="bold", zorder=6)

    ax.set_title("Electrification Timeline — Grouped by ACF Category", fontsize=9, pad=4)


# Scenario colours for the By Scenario Gantt view (matches the Scenario Comparison charts)
_SCENARIO_COLORS = {
    "minimum_compliance":  "#37474F",   # dark slate
    "all_except_emergency": "#E64A19",  # orange
    "whole_fleet":          "#1B5E20",  # dark green
}
_SCENARIO_LABELS = {
    "minimum_compliance":   "Minimum Compliance",
    "all_except_emergency": "All Excl. Emergency",
    "whole_fleet":          "Whole Fleet",
}


def _gantt_by_scenario(ax, vehicles: list, years: list, scenario_results) -> None:
    """Bubble chart: one row per scope scenario, bubble size = vehicle count for that year.

    scenario_results: list of ScenarioResult namedtuples from compare_scenarios(), or None.
    When scenario_results is None (analysis not yet run), computes approximate counts from
    vehicle custom_fields directly.
    """
    from analysis.scenarios import get_scenario_year_assignments, PRESET_SCENARIOS

    year_set = set(years)
    scenario_keys = list(SCOPE_SCENARIO_KEYS)  # ["minimum_compliance", ...]

    # Build year→count dicts for each scenario
    scenario_year_counts: Dict[str, Dict[int, int]] = {}
    for key in scenario_keys:
        year_counts: Dict[int, int] = {}
        try:
            year_map = get_scenario_year_assignments(vehicles, key)
            for yr in year_map.values():
                if yr in year_set:
                    year_counts[yr] = year_counts.get(yr, 0) + 1
        except Exception:
            pass
        scenario_year_counts[key] = year_counts

    any_data = any(scenario_year_counts[k] for k in scenario_keys)
    if not any_data:
        ax.text(0.5, 0.5, "Run Full Analysis to generate scenario timelines.",
                ha="center", va="center", fontsize=9)
        ax.axis("off")
        return

    scenario_pos = {k: i for i, k in enumerate(reversed(scenario_keys))}
    ax.set_yticks(list(scenario_pos.values()))
    ax.set_yticklabels(
        [_SCENARIO_LABELS.get(k, k) for k in reversed(scenario_keys)],
        fontsize=8
    )
    ax.set_ylim(-0.7, len(scenario_keys) - 0.3)

    for key in scenario_keys:
        color = _SCENARIO_COLORS.get(key, "#888888")
        for yr, cnt in scenario_year_counts[key].items():
            size = 200 + cnt * 80
            ax.scatter(yr, scenario_pos[key], s=size, color=color,
                       alpha=0.80, zorder=5, edgecolors="white", linewidths=0.8)
            ax.text(yr, scenario_pos[key], str(cnt),
                    ha="center", va="center", fontsize=7,
                    color="white", fontweight="bold", zorder=6)

    ax.set_title("Electrification Timeline — By Scenario", fontsize=9, pad=4)


def _gantt_per_vehicle(ax, vehicles: list, years: list,
                       max_vehicles: int = 0) -> None:
    """One row per vehicle, dot at scheduled EV year, coloured by ACF.

    max_vehicles: if > 0, cap the number of rows shown (earliest-year first).
    """
    year_set = set(years)
    schedulable = []
    for v in vehicles:
        yr_raw = v.custom_fields.get("Proposed EV Year", "")
        try:
            yr = int(yr_raw)
        except (ValueError, TypeError):
            continue
        if yr in year_set:
            schedulable.append(v)

    if not schedulable:
        ax.text(0.5, 0.5, "No vehicles with scheduled EV years in range.",
                ha="center", va="center", fontsize=9)
        ax.axis("off")
        return

    # Sort by scheduled year then ACF
    schedulable.sort(key=lambda v: (
        int(v.custom_fields.get("Proposed EV Year", "9999")),
        v.custom_fields.get("_acf_code", ""),
    ))

    total = len(schedulable)
    capped = max_vehicles > 0 and total > max_vehicles
    if capped:
        schedulable = schedulable[:max_vehicles]

    for i, v in enumerate(schedulable):
        yr = int(v.custom_fields.get("Proposed EV Year"))
        cat = v.custom_fields.get("_acf_code", "")
        color = ACF_COLORS.get(cat, "#AAAAAA")
        is_override = v.custom_fields.get("EV Year Overridden") == "Yes"
        marker = "*" if is_override else "o"
        ax.scatter(yr, i, color=color, s=60, marker=marker, zorder=5,
                   edgecolors="black" if is_override else "none", linewidths=0.7)

    ax.set_yticks(range(len(schedulable)))
    label_fn = lambda v: (v.asset_id or v.vin[:6]) if v.asset_id or v.vin else "—"
    ax.set_yticklabels([label_fn(v) for v in schedulable], fontsize=6)
    ax.set_ylim(-0.8, len(schedulable) - 0.2)
    title = "Electrification Timeline — Per Vehicle  (★ = overridden)"
    if capped:
        title += f"  [showing {max_vehicles} of {total}]"
    ax.set_title(title, fontsize=9, pad=4)


class AnalysisPanel(ttk.Frame):
    """
    Dashboard panel for fleet electrification analysis.

    Layout (top to bottom, scrollable):
      Action bar → Fleet Snapshot KPIs → Top 5 / ACF donut →
      Scenario Comparison → TCO Summary →
      ▼ Chart Browser (collapsed) → ▼ Parameters (expanded) → Export bar
    """

    def __init__(self, parent, fleet=None, on_analysis_complete=None,
                 on_report_generation=None, sharing_data=None):
        super().__init__(parent)

        self.on_analysis_complete_callback = on_analysis_complete
        self.on_report_generation_callback = on_report_generation
        self.sharing_data = sharing_data

        self.fleet = fleet or Fleet(name="Empty Fleet")
        self.current_chart_type = tk.StringVar(value=CHART_TYPES[0] if CHART_TYPES else "")
        self.current_figure = None
        self.current_canvas = None

        # ── Analysis parameters ───────────────────────────────────────────────
        self.gas_price_var          = tk.DoubleVar(value=DEFAULT_GAS_PRICE)
        self.electricity_price_var  = tk.DoubleVar(value=DEFAULT_ELECTRICITY_PRICE)
        self.ev_efficiency_var      = tk.DoubleVar(value=DEFAULT_EV_EFFICIENCY)
        self.analysis_years_var     = tk.IntVar(value=10)
        self.discount_rate_var      = tk.DoubleVar(value=5.0)
        self.incentive_amount_var   = tk.DoubleVar(value=0.0)
        self.battery_degradation_var= tk.DoubleVar(value=DEFAULT_BATTERY_DEGRADATION)
        self.residual_ice_var       = tk.DoubleVar(value=DEFAULT_RESIDUAL_VALUE_ICE_PCT)
        self.residual_ev_var        = tk.DoubleVar(value=DEFAULT_RESIDUAL_VALUE_EV_PCT)
        self._incentive_check_vars: List[Tuple[tk.BooleanVar, float]] = []

        self.charging_pattern_var = tk.StringVar(value="standard")
        self.charging_start_var   = tk.IntVar(value=18)
        self.charging_end_var     = tk.IntVar(value=6)
        self.power_level_var      = tk.StringVar(value="LP")
        self.power_levels = {
            "LP":  tk.DoubleVar(value=7.2),
            "MP":  tk.DoubleVar(value=19.2),
            "HP":  tk.DoubleVar(value=50.0),
            "VHP": tk.DoubleVar(value=150.0),
        }

        # ── Fleet type (controls CARB ACF deadline table) ─────────────────────
        self._fleet_type_var = tk.StringVar(value="hpf")
        # ── Max annual replacements (0 = no cap, even-spread) ─────────────────
        self._max_per_year_var = tk.IntVar(value=0)

        # ── Analysis results ──────────────────────────────────────────────────
        self.electrification_analysis = None
        self.emissions_inventory      = None
        self.charging_analysis        = None
        self.scenario_results         = None

        # ── Scenario checkboxes (scope-based: Analysis tab only) ─────────────
        self._scenario_vars: Dict[str, tk.BooleanVar] = {
            key: tk.BooleanVar(value=True) for key in SCOPE_SCENARIO_KEYS
        }
        # Current Plan scenario: pre-checked, shows fleet's actual EV year assignments
        self._current_plan_var = tk.BooleanVar(value=True)
        # Horizon year for scope-based scenario comparison
        self._scenario_horizon_var = tk.IntVar(value=2040)

        self.export_coordinator = ExportCoordinator()

        # ── Chart style vars (used by chart browser) ──────────────────────────
        self.chart_style_var  = tk.StringVar(value="default")
        self.color_scheme_var = tk.StringVar(value="default")

        # ── Widget references populated during _create_ui ─────────────────────
        self.status_label: Optional[ttk.Label]     = None
        self._snapshot_kpi_labels: Dict[str, ttk.Label] = {}
        self._tco_kpi_labels:      Dict[str, ttk.Label] = {}
        self.kpi_labels:           Dict[str, ttk.Label] = {}  # compat alias
        self._top5_tree:           Optional[ttk.Treeview] = None
        self._top5_hscroll:        Optional[ttk.Scrollbar] = None
        self._top5_placeholder:    Optional[ttk.Label] = None
        self._acf_fig              = None
        self._acf_canvas           = None
        self._acf_count_labels:    Dict[str, ttk.Label] = {}
        self._scenario_fig         = None
        self._scenario_canvas      = None
        self._scenario_table_outer: Optional[ttk.LabelFrame] = None
        self._parameters_outer     = None
        self._chart_browser_outer  = None
        self.toolbar_frame         = None
        self.chart_frame           = None
        self._gantt_fig            = None
        self._gantt_canvas         = None
        self._gantt_outer          = None
        self._gantt_view_var: Optional[tk.StringVar] = None
        self._gantt_max_var_analysis: Optional[tk.StringVar] = None
        self._export_btn: Optional[ttk.Button]  = None
        self._present_btn: Optional[ttk.Button] = None

        # ── Cost trajectory chart (scenario section) ──────────────────────────
        self._scenario_cost_fig    = None
        self._scenario_cost_canvas = None

        # ── Stale analysis tracking ───────────────────────────────────────────
        self._analysis_is_stale = False

        self._create_ui()

        # Attach traces AFTER _create_ui so widgets are ready
        for _v in [self.gas_price_var, self.electricity_price_var,
                   self.ev_efficiency_var, self.analysis_years_var,
                   self.discount_rate_var, self.battery_degradation_var,
                   self.residual_ice_var, self.residual_ev_var,
                   self._fleet_type_var, self._max_per_year_var]:
            _v.trace_add("write", self._on_param_changed)

    # =========================================================================
    # UI Construction
    # =========================================================================

    def _create_ui(self):
        """Build the scrollable dashboard layout.

        Section order (top → bottom):
          1. Action bar (sticky, outside scroll)
          2. Parameters (configure before running)
          3. Fleet Snapshot KPIs
          4. Scenario Comparison (primary analytical view)
          5. Chart Gallery (replaces old Chart Browser)
          6. Electrification Timeline Gantt
          7. TCO Summary
          8. Top-5 Priority / ACF Donut (collapsible, starts collapsed)
          9. Export bar
        """
        # ── Sticky action bar (not inside scroll area) ────────────────────────
        self._create_action_bar(self)

        # ── Scrollable dashboard body ─────────────────────────────────────────
        self._dashboard_scroll = ScrollableFrame(self)
        self._dashboard_scroll.pack(fill=tk.BOTH, expand=True)
        dash = self._dashboard_scroll.scrollable_frame

        self._create_parameters_section_collapsible(dash)
        ttk.Separator(dash, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=Spacing.MD)

        self._create_fleet_snapshot_section(dash)
        ttk.Separator(dash, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=Spacing.MD)

        self._create_scenario_section_dashboard(dash)
        ttk.Separator(dash, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=Spacing.MD)

        self._create_chart_gallery_section(dash)
        ttk.Separator(dash, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=Spacing.MD)

        self._create_gantt_section(dash)
        ttk.Separator(dash, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=Spacing.MD)

        self._create_tco_summary_section(dash)
        ttk.Separator(dash, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=Spacing.MD)

        self._create_priority_acf_row_collapsible(dash)
        ttk.Separator(dash, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=Spacing.MD)

        self._create_export_bar(dash)

    # ── Action bar ────────────────────────────────────────────────────────────

    def _create_action_bar(self, parent):
        """Top bar with Run Full Analysis button and status label."""
        bar = ttk.Frame(parent)
        bar.pack(fill=tk.X, padx=Spacing.SM, pady=(Spacing.SM, Spacing.XS))

        run_btn = ttk.Button(
            bar,
            text="⚡  Run Full Analysis",
            command=self.run_full_analysis,
            style="Accent.TButton",
        )
        run_btn.pack(side=tk.LEFT, padx=(0, Spacing.MD))
        SimpleTooltip(run_btn,
            "Run electrification, emissions, and charging analyses.\n"
            "Populates the KPI chips, Top 5 table, and financial summary.")

        self.status_label = ttk.Label(
            bar,
            text="Ready — load a fleet and click Run Full Analysis",
            foreground=Colors.TEXT_TERTIARY,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL),
        )
        self.status_label.pack(side=tk.LEFT)

    # ── Collapsible section helper ────────────────────────────────────────────

    def _create_collapsible_section(self, parent, title: str,
                                    start_expanded: bool = True):
        """
        Create a collapsible section with a ▼/▶ toggle button.

        Returns:
            (outer_frame, content_frame) — pack children into content_frame.
            outer_frame exposes ._toggle() and ._collapsed bool.
        """
        outer = ttk.Frame(parent)
        outer.pack(fill=tk.X, padx=Spacing.SM, pady=(0, Spacing.XS))
        outer._collapsed = not start_expanded

        header = ttk.Frame(outer)
        header.pack(fill=tk.X, pady=(0, Spacing.XS))

        arrow = tk.StringVar(value="▼" if start_expanded else "▶")

        content = ttk.Frame(outer)
        if start_expanded:
            content.pack(fill=tk.X, padx=Spacing.SM)

        def toggle():
            if outer._collapsed:
                content.pack(fill=tk.X, padx=Spacing.SM)
                arrow.set("▼")
                outer._collapsed = False
            else:
                content.pack_forget()
                arrow.set("▶")
                outer._collapsed = True

        outer._toggle  = toggle
        outer._content = content

        ttk.Button(
            header, textvariable=arrow, command=toggle, width=2,
            style="Secondary.TButton",
        ).pack(side=tk.LEFT, padx=(0, Spacing.XS))

        ttk.Label(
            header, text=title,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_H3, Fonts.WEIGHT_BOLD),
            foreground=Colors.TEXT_PRIMARY,
        ).pack(side=tk.LEFT)

        return outer, content

    # ── Fleet Snapshot row ────────────────────────────────────────────────────

    def _create_fleet_snapshot_section(self, parent):
        """4 KPI chips: Fleet Size | ACF-B Count | Avg MPG | MPG Coverage."""
        section = ttk.Frame(parent)
        section.pack(fill=tk.X, padx=Spacing.SM, pady=(Spacing.SM, 0))

        ttk.Label(
            section, text="Fleet Snapshot",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_H3, Fonts.WEIGHT_BOLD),
            foreground=Colors.TEXT_PRIMARY,
        ).pack(anchor="w", pady=(0, Spacing.XS))

        chips_row = ttk.Frame(section)
        chips_row.pack(fill=tk.X)
        for col in range(4):
            chips_row.grid_columnconfigure(col, weight=1)

        snap_defs = [
            ("fleet_size",   "Fleet Size",    "—"),
            ("acf_b_count",  "ACF-B Count",   "—"),
            ("avg_mpg",      "Avg MPG",        "—"),
            ("mpg_coverage", "MPG Coverage",  "—"),
        ]
        self._snapshot_kpi_labels = {}
        for col, (key, title, default) in enumerate(snap_defs):
            card = ttk.Frame(chips_row, relief="solid", borderwidth=1)
            card.grid(row=0, column=col, padx=4, pady=4, sticky="nsew")
            chips_row.grid_rowconfigure(0, weight=1)
            ttk.Label(card, text=title,
                      font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL),
                      foreground=Colors.TEXT_TERTIARY).pack(anchor="w", padx=8, pady=(6, 0))
            val_lbl = ttk.Label(card, text=default,
                                font=(Fonts.FAMILY_SANS, 16, Fonts.WEIGHT_BOLD),
                                foreground=Colors.PRIMARY_GREEN)
            val_lbl.pack(anchor="w", padx=8, pady=(0, 6))
            self._snapshot_kpi_labels[key] = val_lbl

        self.kpi_labels = self._snapshot_kpi_labels

    # ── Two-column row: Top 5 | ACF donut ─────────────────────────────────────

    def _create_priority_acf_row_collapsible(self, parent):
        """Collapsible wrapper around the Top-5 + ACF donut row (starts collapsed)."""
        outer, content = self._create_collapsible_section(
            parent, "Fleet Detail — Priority Vehicles & ACF Breakdown",
            start_expanded=False
        )
        self._priority_acf_outer = outer
        self._create_priority_acf_row(content)

    def _create_priority_acf_row(self, parent):
        """Two-column: Top 5 Priority Vehicles (left) | ACF Compliance donut (right)."""
        row = ttk.Frame(parent)
        row.pack(fill=tk.BOTH, padx=Spacing.SM, pady=(0, Spacing.SM))
        row.grid_columnconfigure(0, weight=6)
        row.grid_columnconfigure(1, weight=4)
        row.grid_rowconfigure(0, weight=1)

        # ── Left: Top 5 table ─────────────────────────────────────────────────
        left_card = ttk.LabelFrame(row, text="Top 5 Priority Vehicles")
        left_card.grid(row=0, column=0, padx=(0, Spacing.SM), sticky="nsew")

        self._top5_placeholder = ttk.Label(
            left_card,
            text="Run Full Analysis to populate",
            foreground=Colors.TEXT_DISABLED,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL, "italic"),
        )
        self._top5_placeholder.pack(padx=12, pady=24)

        cols = [
            ("asset_id",  "Asset ID",   60),
            ("year",      "Year",        45),
            ("make",      "Make",        65),
            ("model",     "Model",       85),
            ("wt_class",  "Wt Class",    68),
            ("body_type", "Body Type",   80),
            ("acf_cat",   "ACF",         38),
            ("ev_year",   "EV Year",     52),
            ("priority",  "Priority $",  74),
            ("age",       "Age",         36),
            ("odometer",  "Odometer",    68),
        ]
        col_ids = [c[0] for c in cols]
        tree = ttk.Treeview(left_card, columns=col_ids, show="headings", height=5)
        for col_id, heading, width in cols:
            tree.heading(col_id, text=heading, anchor="w")
            tree.column(col_id, width=width, minwidth=width, anchor="w", stretch=False)
        h_scroll = ttk.Scrollbar(left_card, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=h_scroll.set)

        # Initially hidden — shown after analysis runs
        self._top5_tree    = tree
        self._top5_hscroll = h_scroll

        # Jump to vehicle on double-click
        tree.bind("<Double-Button-1>", self._on_top5_double_click)
        SimpleTooltip(tree, "Double-click a row to jump to that vehicle in the Results tab")

        # ── Right: ACF donut ──────────────────────────────────────────────────
        right_card = ttk.LabelFrame(row, text="ACF Compliance Summary")
        right_card.grid(row=0, column=1, sticky="nsew")

        self._acf_fig = Figure(figsize=(3.5, 2.6), dpi=80)
        self._acf_fig.patch.set_facecolor(Colors.SURFACE)
        acf_canvas = FigureCanvasTkAgg(self._acf_fig, master=right_card)
        acf_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=4, pady=(4, 0))
        self._acf_canvas = acf_canvas

        self._acf_text_frame = ttk.Frame(right_card)
        self._acf_text_frame.pack(fill=tk.X, padx=8, pady=(2, 6))
        for cat in ["ZEV", "A", "B", "C", "D"]:
            rf = ttk.Frame(self._acf_text_frame)
            rf.pack(anchor="w")
            tk.Label(rf, text="●", foreground=ACF_COLORS.get(cat, "#888"),
                     font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL)).pack(side=tk.LEFT)
            lbl = ttk.Label(rf, text=f"Category {cat}: —",
                            font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL))
            lbl.pack(side=tk.LEFT, padx=(2, 0))
            self._acf_count_labels[cat] = lbl

        # Populate donut immediately from fleet data if available
        self._update_acf_donut()

    # ── Scenario Comparison ───────────────────────────────────────────────────

    def _create_scenario_section_dashboard(self, parent):
        """Full-width Scenario Comparison section."""
        section = ttk.LabelFrame(parent, text="Scenario Comparison")
        section.pack(fill=tk.BOTH, padx=Spacing.SM, pady=(0, Spacing.SM))

        # Checkboxes + Compare button in one row
        cb_row = ttk.Frame(section)
        cb_row.pack(fill=tk.X, padx=Spacing.SM, pady=(Spacing.XS, 0))

        scope_labels = {
            "minimum_compliance":    "Minimum Compliance (Cat B only)",
            "all_except_emergency":  "All Excl. Emergency (A+B+C)",
            "whole_fleet":           "Whole Fleet (A+B+C+D)",
        }
        scope_tips = {
            "minimum_compliance":   "Only ACF mandate-subject vehicles (medium & heavy duty, Cat B).\nShows the bare minimum required by CARB.",
            "all_except_emergency": "All vehicles except emergency (Cat A light-duty, Cat B mandate-subject, Cat C body-type exempt).\nTypical planning scope for most clients.",
            "whole_fleet":          "Every vehicle including emergency vehicles (Cat A+B+C+D).\nShows the full fleet electrification cost and CO\u2082 benefit.",
        }
        for key, label in scope_labels.items():
            cb = ttk.Checkbutton(
                cb_row, text=label,
                variable=self._scenario_vars[key],
            )
            cb.pack(side=tk.LEFT, padx=(0, Spacing.MD))
            SimpleTooltip(cb, scope_tips[key])

        # Horizon year spinbox
        ttk.Separator(cb_row, orient=tk.VERTICAL).pack(
            side=tk.LEFT, fill=tk.Y, padx=(Spacing.XS, Spacing.SM), pady=2)
        ttk.Label(cb_row, text="Horizon:").pack(side=tk.LEFT, padx=(0, 4))
        horizon_spin = ttk.Spinbox(
            cb_row, from_=2026, to=2060,
            textvariable=self._scenario_horizon_var,
            width=5, command=lambda: None,
        )
        horizon_spin.pack(side=tk.LEFT, padx=(0, Spacing.MD))
        SimpleTooltip(horizon_spin,
            "End year for scope scenarios.\nAll three scope scenarios use this year as their planning horizon.")

        # Current Plan checkbox (4th, pre-checked)
        ttk.Separator(cb_row, orient=tk.VERTICAL).pack(
            side=tk.LEFT, fill=tk.Y, padx=(Spacing.XS, Spacing.SM), pady=2)
        cp_cb = ttk.Checkbutton(
            cb_row,
            text="Current Plan (Overrides)",
            variable=self._current_plan_var,
        )
        cp_cb.pack(side=tk.LEFT, padx=(0, Spacing.MD))
        SimpleTooltip(cp_cb,
            "Show the fleet's actual planned EV year assignments\n"
            "(includes any manual overrides from the Timeline tab)")

        compare_btn = ttk.Button(
            cb_row,
            text="Compare Scenarios",
            command=self._run_scenario_comparison,
            style="Primary.TButton",
        )
        compare_btn.pack(side=tk.LEFT, padx=(Spacing.SM, 0))
        SimpleTooltip(compare_btn,
            "Run selected scope scenarios and display:\n"
            "• Annual CO₂ burning-platform chart\n"
            "• Cumulative cost trajectory chart\n"
            "• Vehicles-per-year replacement schedule\n"
            "(Runs automatically after Run Full Analysis)")

        # ── CO₂ chart ─────────────────────────────────────────────────────────
        chart_outer = ttk.Frame(section)
        chart_outer.pack(fill=tk.X, padx=Spacing.SM, pady=(Spacing.XS, 0))

        self._scenario_fig = Figure(figsize=(9, 2.6), dpi=80)
        self._scenario_fig.patch.set_facecolor(Colors.SURFACE)
        scen_canvas = FigureCanvasTkAgg(self._scenario_fig, master=chart_outer)
        scen_canvas.get_tk_widget().pack(anchor="w")
        self._scenario_canvas = scen_canvas

        # Placeholder text
        ax = self._scenario_fig.add_subplot(111)
        ax.text(0.5, 0.5,
                "Run Full Analysis or click Compare Scenarios",
                ha="center", va="center", fontsize=9,
                color=Colors.TEXT_TERTIARY)
        ax.axis("off")
        self._scenario_canvas.draw()

        # ── Cost trajectory chart ─────────────────────────────────────────────
        cost_outer = ttk.Frame(section)
        cost_outer.pack(fill=tk.X, padx=Spacing.SM, pady=(Spacing.XS, 0))

        self._scenario_cost_fig = Figure(figsize=(9, 2.6), dpi=80)
        self._scenario_cost_fig.patch.set_facecolor(Colors.SURFACE)
        cost_canvas = FigureCanvasTkAgg(self._scenario_cost_fig, master=cost_outer)
        cost_canvas.get_tk_widget().pack(anchor="w")
        self._scenario_cost_canvas = cost_canvas

        cost_ax = self._scenario_cost_fig.add_subplot(111)
        cost_ax.text(0.5, 0.5, "",
                     ha="center", va="center", fontsize=9,
                     color=Colors.TEXT_TERTIARY)
        cost_ax.axis("off")
        self._scenario_cost_canvas.draw()

        # Vehicles-per-year table
        self._scenario_table_outer = ttk.LabelFrame(
            section, text="Scenario Results: Vehicles per Year"
        )
        self._scenario_table_outer.pack(
            fill=tk.X, padx=Spacing.SM, pady=(Spacing.XS, Spacing.SM)
        )
        ttk.Label(
            self._scenario_table_outer,
            text="Click 'Compare Scenarios' to see vehicles-per-year breakdown.",
            foreground=Colors.TEXT_DISABLED,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL, "italic"),
        ).pack(padx=10, pady=6)

    # ── TCO & Financial Summary ───────────────────────────────────────────────

    def _create_tco_summary_section(self, parent):
        """3 financial KPI cards: Annual Savings | Payback | Infrastructure."""
        section = ttk.LabelFrame(parent, text="TCO & Financial Summary")
        section.pack(fill=tk.X, padx=Spacing.SM, pady=(0, Spacing.SM))

        cards_row = ttk.Frame(section)
        cards_row.pack(fill=tk.X, padx=Spacing.SM, pady=Spacing.SM)
        for col in range(3):
            cards_row.grid_columnconfigure(col, weight=1)

        tco_defs = [
            ("annual_savings", "Annual Savings",     "—"),
            ("payback",        "Avg Payback Period", "—"),
            ("infra_cost",     "Infrastructure Est.", "—"),
        ]
        self._tco_kpi_labels = {}
        for col, (key, title, default) in enumerate(tco_defs):
            card = ttk.Frame(cards_row, relief="solid", borderwidth=1)
            card.grid(row=0, column=col, padx=4, pady=4, sticky="nsew")
            cards_row.grid_rowconfigure(0, weight=1)
            ttk.Label(card, text=title,
                      font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL),
                      foreground=Colors.TEXT_TERTIARY).pack(anchor="w", padx=8, pady=(6, 0))
            val_lbl = ttk.Label(card, text=default,
                                font=(Fonts.FAMILY_SANS, 14, Fonts.WEIGHT_BOLD),
                                foreground=Colors.PRIMARY_GREEN)
            val_lbl.pack(anchor="w", padx=8, pady=(0, 6))
            self._tco_kpi_labels[key] = val_lbl

        self.kpi_labels.update(self._tco_kpi_labels)

    # ── Electrification Gantt (collapsible, starts collapsed) ─────────────────

    def _create_gantt_section(self, parent):
        """Collapsible read-only Gantt chart (grouped by ACF, default).
        Full editing lives in the Timeline tab."""
        outer, content = self._create_collapsible_section(
            parent, "Electrification Timeline Gantt", start_expanded=True
        )
        self._gantt_outer = outer

        # Controls row: view toggle + open-in-timeline button
        ctrl = ttk.Frame(content)
        ctrl.pack(fill=tk.X, padx=Spacing.SM, pady=(0, Spacing.XS))

        ttk.Label(ctrl, text="View:").pack(side=tk.LEFT, padx=(0, 4))
        self._gantt_view_var = tk.StringVar(value="Grouped by ACF")
        view_combo = ttk.Combobox(
            ctrl,
            textvariable=self._gantt_view_var,
            values=["Grouped by ACF", "By Scenario", "Per Vehicle"],
            state="readonly",
            width=16,
        )
        view_combo.pack(side=tk.LEFT, padx=(0, Spacing.MD))
        view_combo.bind("<<ComboboxSelected>>", lambda _e: self._update_gantt_section())

        # Max rows control (only relevant for Per Vehicle view)
        self._gantt_max_var_analysis = tk.StringVar(value="50")
        ttk.Label(ctrl, text="Max rows:").pack(side=tk.LEFT, padx=(0, 3))
        max_combo = ttk.Combobox(
            ctrl,
            textvariable=self._gantt_max_var_analysis,
            values=["25", "50", "100", "All"],
            state="readonly",
            width=6,
        )
        max_combo.pack(side=tk.LEFT, padx=(0, Spacing.MD))
        max_combo.bind("<<ComboboxSelected>>", lambda _e: self._update_gantt_section())
        SimpleTooltip(max_combo, "Cap visible rows in Per Vehicle view")

        ttk.Button(
            ctrl,
            text="→ Edit in Timeline tab",
            command=self._navigate_to_timeline,
            style="Secondary.TButton",
        ).pack(side=tk.RIGHT)

        # Matplotlib canvas
        self._gantt_fig = Figure(figsize=(9, 3.0), dpi=80)
        self._gantt_fig.patch.set_facecolor(Colors.SURFACE)
        gantt_canvas = FigureCanvasTkAgg(self._gantt_fig, master=content)
        gantt_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True,
                                          padx=Spacing.SM, pady=(0, Spacing.XS))
        self._gantt_canvas = gantt_canvas

        # Placeholder text
        ax = self._gantt_fig.add_subplot(111)
        ax.text(0.5, 0.5,
                "Run Full Analysis to populate the timeline Gantt chart.",
                ha="center", va="center", fontsize=9,
                color=Colors.TEXT_TERTIARY)
        ax.axis("off")
        self._gantt_canvas.draw()

    # ── Chart Browser (collapsible, starts collapsed) ─────────────────────────

    def _create_chart_gallery_section(self, parent):
        """Collapsible Chart Gallery — big selected chart + scrollable thumbnail strip.

        UX:
          • Big canvas (main view) shows the currently selected chart.
          • Horizontally scrollable thumbnail strip below lets the user browse
            all chart types at a glance and click to load one.
          • Each thumbnail has an "Add to Presentation" checkbox that sets a flag
            in sharing_data so the Present tab can surface it.
          • Thumbnails are rendered lazily when the section is expanded after
            analysis has completed.
        """
        outer, content = self._create_collapsible_section(
            parent, "Chart Gallery", start_expanded=False
        )
        self._chart_browser_outer = outer  # keep attr name for compat

        # Style/color controls (no separate chart picker — thumbnails drive selection)
        ctrl = ttk.Frame(content)
        ctrl.pack(fill=tk.X, pady=(0, Spacing.XS))

        ttk.Label(ctrl, text="Style:").pack(side=tk.LEFT, padx=(0, 4))
        style_combo = ttk.Combobox(ctrl, textvariable=self.chart_style_var,
                                   values=["default", "minimal", "dark", "colorful"],
                                   state="readonly", width=10)
        style_combo.pack(side=tk.LEFT, padx=(0, Spacing.MD))
        style_combo.bind("<<ComboboxSelected>>",
                         lambda e: self._refresh_gallery_main_chart())

        ttk.Label(ctrl, text="Colors:").pack(side=tk.LEFT, padx=(0, 4))
        color_combo = ttk.Combobox(ctrl, textvariable=self.color_scheme_var,
                                   values=["default", "viridis", "magma", "plasma", "inferno"],
                                   state="readonly", width=10)
        color_combo.pack(side=tk.LEFT, padx=(0, Spacing.MD))
        color_combo.bind("<<ComboboxSelected>>",
                         lambda e: self._refresh_gallery_main_chart())

        # Selected chart label
        self._gallery_selected_label = ttk.Label(
            ctrl, text="", font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY, Fonts.WEIGHT_BOLD),
            foreground=Colors.TEXT_PRIMARY)
        self._gallery_selected_label.pack(side=tk.LEFT, padx=Spacing.SM)

        save_btn = ttk.Button(ctrl, text="Save Chart…", command=self._save_chart, width=12)
        save_btn.pack(side=tk.RIGHT, padx=2)
        copy_btn = ttk.Button(ctrl, text="Copy Chart", command=self._copy_chart, width=10)
        copy_btn.pack(side=tk.RIGHT, padx=2)

        # ── Main chart canvas ─────────────────────────────────────────────────
        self.toolbar_frame = ttk.Frame(content)
        self.toolbar_frame.pack(fill=tk.X)

        self.chart_frame = ttk.Frame(content, height=400)
        self.chart_frame.pack(fill=tk.BOTH, expand=True)
        self.chart_frame.pack_propagate(False)

        self._create_initial_chart()

        # ── Thumbnail strip ───────────────────────────────────────────────────
        thumb_wrapper = ttk.LabelFrame(content, text="All Charts — click to select")
        thumb_wrapper.pack(fill=tk.X, padx=0, pady=(Spacing.SM, 0))

        # Canvas + horizontal scrollbar for the strip
        self._gallery_strip_canvas = tk.Canvas(
            thumb_wrapper, height=130, highlightthickness=0,
            background=Colors.BACKGROUND if hasattr(Colors, "BACKGROUND") else "#F5F5F5",
        )
        strip_hbar = ttk.Scrollbar(thumb_wrapper, orient=tk.HORIZONTAL,
                                   command=self._gallery_strip_canvas.xview)
        self._gallery_strip_canvas.configure(xscrollcommand=strip_hbar.set)
        strip_hbar.pack(side=tk.BOTTOM, fill=tk.X)
        self._gallery_strip_canvas.pack(side=tk.TOP, fill=tk.BOTH)

        # Inner frame that holds thumbnail cards
        self._gallery_strip_inner = ttk.Frame(self._gallery_strip_canvas)
        self._gallery_strip_window = self._gallery_strip_canvas.create_window(
            (0, 0), window=self._gallery_strip_inner, anchor="nw"
        )
        self._gallery_strip_inner.bind(
            "<Configure>",
            lambda e: self._gallery_strip_canvas.configure(
                scrollregion=self._gallery_strip_canvas.bbox("all")
            )
        )

        # Storage for thumbnail state
        self._chart_thumbnails: Dict[str, object] = {}   # chart_type → PhotoImage
        self._gallery_pptx_vars: Dict[str, tk.BooleanVar] = {}
        self._gallery_thumb_frames: Dict[str, tk.Frame] = {}
        self._gallery_thumb_rendered = False

        # Populate thumbnail placeholders immediately (text cards, no renders yet)
        self._build_thumbnail_placeholders()

        # Expand → render thumbnails if analysis data is already available
        outer.bind("<Map>", lambda e: self._lazy_render_thumbnails(), add="+")

    def _build_thumbnail_placeholders(self):
        """Build one placeholder card per chart type in the thumbnail strip."""
        THUMB_W, THUMB_H = 148, 100

        for col, chart_type in enumerate(CHART_TYPES):
            var = tk.BooleanVar(value=False)
            self._gallery_pptx_vars[chart_type] = var

            card = tk.Frame(
                self._gallery_strip_inner, width=THUMB_W, height=THUMB_H,
                relief="raised", bd=1,
                bg="#E8EAF6",
            )
            card.pack(side=tk.LEFT, padx=4, pady=6)
            card.pack_propagate(False)
            self._gallery_thumb_frames[chart_type] = card

            # Placeholder label (replaced with image when rendered)
            name_lbl = tk.Label(
                card, text=chart_type, font=(Fonts.FAMILY_SANS, 7),
                wraplength=THUMB_W - 8, justify="center", bg="#E8EAF6",
                fg="#333",
            )
            name_lbl.place(relx=0.5, rely=0.45, anchor="center")

            # "Add to Presentation" checkbox at bottom
            cb_frame = tk.Frame(card, bg="#E8EAF6")
            cb_frame.place(relx=0.0, rely=1.0, anchor="sw", y=-2, x=2)
            cb = tk.Checkbutton(cb_frame, variable=var, bg="#E8EAF6",
                                command=lambda ct=chart_type: self._on_gallery_pptx_changed(ct))
            cb.pack(side=tk.LEFT)
            tk.Label(cb_frame, text="PPT", font=(Fonts.FAMILY_SANS, 6),
                     bg="#E8EAF6", fg="#555").pack(side=tk.LEFT)

            # Click anywhere on card to select that chart
            for w in (card, name_lbl):
                w.bind("<Button-1>",
                       lambda e, ct=chart_type: self._gallery_select_chart(ct))
                w.bind("<Enter>",
                       lambda e, fr=card: fr.config(relief="solid", bd=2))
                w.bind("<Leave>",
                       lambda e, fr=card: fr.config(relief="raised", bd=1))

    def _lazy_render_thumbnails(self):
        """Render actual chart thumbnails — called when gallery section is first opened."""
        if self._gallery_thumb_rendered:
            return
        if not self.fleet:
            return
        self._gallery_thumb_rendered = True
        # Render each chart in the background to avoid freezing the UI
        threading.Thread(target=self._render_all_thumbnails, daemon=True).start()

    def _render_all_thumbnails(self):
        """Background thread: render each chart type as a small thumbnail."""
        if not _PIL_AVAILABLE:
            return
        THUMB_W, THUMB_H = 136, 82
        for chart_type in CHART_TYPES:
            try:
                thumb_fig = Figure(figsize=(2.2, 1.35), dpi=62)
                # Build same extra_args logic as _update_chart
                chart_data = self.fleet
                extra_args = {}
                if chart_type == "Fleet Cash Flow" and self.electrification_analysis:
                    chart_data = self.electrification_analysis
                elif chart_type == "Electrification Potential" and self.electrification_analysis:
                    chart_data = self.electrification_analysis
                    extra_args = {
                        "gas_price": self.gas_price_var.get(),
                        "electricity_price": self.electricity_price_var.get(),
                        "ev_efficiency": self.ev_efficiency_var.get(),
                    }
                elif chart_type in ("Emissions Reduction", "ROI Analysis") \
                        and self.electrification_analysis:
                    chart_data = self.electrification_analysis
                elif chart_type in ("Annual Cost Comparison",):
                    extra_args = {
                        "gas_price": self.gas_price_var.get(),
                        "electricity_price": self.electricity_price_var.get(),
                        "ev_efficiency": self.ev_efficiency_var.get(),
                    }
                elif "Emission" in chart_type and self.emissions_inventory:
                    chart_data = self.emissions_inventory
                elif "Charging" in chart_type and self.charging_analysis:
                    chart_data = self.charging_analysis

                ChartFactory.create_chart(
                    chart_type=chart_type, data=chart_data, figure=thumb_fig,
                    chart_style="minimal", color_scheme=self.color_scheme_var.get(),
                    **extra_args,
                )
                thumb_fig.tight_layout(pad=0.1)

                buf = _io.BytesIO()
                thumb_fig.savefig(buf, format="png", dpi=62,
                                  bbox_inches="tight", facecolor="white")
                buf.seek(0)
                img = _PILImage.open(buf).resize((THUMB_W, THUMB_H), _PILImage.LANCZOS)
                photo = _ImageTk.PhotoImage(img)
                matplotlib.pyplot.close(thumb_fig)
                buf.close()

                self._chart_thumbnails[chart_type] = photo
                # Schedule UI update on main thread
                self.after(0, lambda ct=chart_type, ph=photo: self._apply_thumbnail(ct, ph))
                # Discard the small figure to free memory
                thumb_fig.clear()
                del thumb_fig
            except Exception as err:
                logger.debug(f"Thumbnail render failed for {chart_type!r}: {err}")

    def _apply_thumbnail(self, chart_type: str, photo):
        """Update a thumbnail card with a rendered image (called on main thread)."""
        card = self._gallery_thumb_frames.get(chart_type)
        if card is None:
            return
        # Replace the placeholder label with the actual image
        for w in card.winfo_children():
            if isinstance(w, tk.Label) and not isinstance(w, tk.Checkbutton):
                w.destroy()
        img_lbl = tk.Label(card, image=photo, cursor="hand2",
                           bg=card.cget("bg"))
        img_lbl.image = photo  # prevent GC
        img_lbl.place(relx=0.5, rely=0.45, anchor="center")
        img_lbl.bind("<Button-1>",
                     lambda e, ct=chart_type: self._gallery_select_chart(ct))
        img_lbl.bind("<Enter>",
                     lambda e, fr=card: fr.config(relief="solid", bd=2))
        img_lbl.bind("<Leave>",
                     lambda e, fr=card: fr.config(relief="raised", bd=1))
        # Add chart name as small overlay at top
        name_lbl = tk.Label(card, text=chart_type,
                            font=(Fonts.FAMILY_SANS, 6),
                            bg="#FFFFFF", fg="#333",
                            wraplength=card.winfo_width() - 4)
        name_lbl.place(relx=0.5, rely=0.06, anchor="n")
        name_lbl.bind("<Button-1>",
                      lambda e, ct=chart_type: self._gallery_select_chart(ct))

    def _gallery_select_chart(self, chart_type: str):
        """Select a chart type from the thumbnail strip and render it full-size."""
        self.current_chart_type.set(chart_type)
        if self._gallery_selected_label:
            self._gallery_selected_label.config(text=chart_type)
        # Highlight selected thumbnail
        for ct, frame in self._gallery_thumb_frames.items():
            frame.config(bd=3 if ct == chart_type else 1,
                         relief="solid" if ct == chart_type else "raised")
        self._update_chart()

    def _refresh_gallery_main_chart(self):
        """Re-render the current gallery chart with updated style/color."""
        self._update_chart()
        # Also flag thumbnails for re-render on next open
        self._gallery_thumb_rendered = False

    def _on_gallery_pptx_changed(self, chart_type: str):
        """Update sharing_data when a chart's 'Add to Presentation' checkbox changes."""
        selected = {ct for ct, var in self._gallery_pptx_vars.items() if var.get()}
        if self.sharing_data is not None:
            try:
                self.sharing_data.set("selected_chart_ids", selected)
            except AttributeError:
                self.sharing_data["selected_chart_ids"] = selected

    # ── Parameters (collapsible, starts expanded) ─────────────────────────────

    def _create_parameters_section_collapsible(self, parent):
        """Collapsible wrapper around the 3-tab parameter notebook."""
        outer, content = self._create_collapsible_section(
            parent, "Parameters", start_expanded=True
        )
        self._parameters_outer = outer
        self._create_parameters_section(content)

    # ── Export bar ────────────────────────────────────────────────────────────

    def _create_export_bar(self, parent):
        """Bottom row: Export Excel Report | Build Presentation | More Exports."""
        bar = ttk.Frame(parent)
        bar.pack(fill=tk.X, padx=Spacing.SM, pady=(0, Spacing.LG))

        self._export_btn = ttk.Button(bar, text="Export Excel Report",
                                      command=self.export_full_report,
                                      style="Primary.TButton",
                                      state="disabled")
        self._export_btn.pack(side=tk.LEFT, padx=(0, Spacing.SM))
        SimpleTooltip(self._export_btn,
            "Generate 8-tab Excel report: Vehicle Data, Summary,\n"
            "Electrification, Charging, Emissions, TCO, Schedule, Dashboard")

        self._present_btn = ttk.Button(bar, text="Build Presentation →",
                                       command=self._navigate_to_present,
                                       style="Accent.TButton",
                                       state="disabled")
        self._present_btn.pack(side=tk.LEFT)
        SimpleTooltip(self._present_btn,
            "Switch to the Present tab to configure and export a PowerPoint deck")

        ttk.Button(bar, text="More Exports ▾",
                   command=self._export_menu).pack(side=tk.RIGHT)

    # =========================================================================
    # Parameters section (intact logic; accepts parent arg for collapsible wrapping)
    # =========================================================================

    def _create_parameters_section(self, parent=None):
        """Create the 3-tab parameter notebook. parent defaults to self."""
        if parent is None:
            parent = self

        params_notebook = ttk.Notebook(parent)
        params_notebook.pack(fill=tk.X, pady=(0, Spacing.SM))

        # ── Costs tab ─────────────────────────────────────────────────────────
        cost_frame = ttk.LabelFrame(params_notebook, text="Cost Parameters")
        params_notebook.add(cost_frame, text="Costs")

        state_frame = ttk.Frame(cost_frame)
        state_frame.pack(fill=tk.X, padx=5, pady=(5, 2))
        state_label = ttk.Label(state_frame, text="State:")
        state_label.pack(side=tk.LEFT)
        self.state_var = tk.StringVar(value="")
        state_options = ["(National Avg)"] + get_available_states()
        state_combo = ttk.Combobox(state_frame, textvariable=self.state_var,
                                   values=state_options, state="readonly", width=12)
        state_combo.pack(side=tk.RIGHT)
        state_combo.bind("<<ComboboxSelected>>", self._on_state_selected)
        SimpleTooltip(state_label, "Select state to auto-populate gas and electricity prices")

        incentive_row = ttk.Frame(cost_frame)
        incentive_row.pack(fill=tk.X, padx=5, pady=(0, 0))
        ttk.Label(incentive_row, text="Incentives ($):").pack(side=tk.LEFT)
        self.incentive_total_label = ttk.Label(incentive_row, text="$0",
                                               foreground="#2a7a2a",
                                               font=("", 9, "bold"))
        self.incentive_total_label.pack(side=tk.RIGHT)
        SimpleTooltip(self.incentive_total_label,
                      "Total incentives applied to EV purchase price in TCO calculation")

        self.incentive_check_frame = ttk.Frame(cost_frame)
        self.incentive_check_frame.pack(fill=tk.X, padx=5, pady=(0, 4))

        for lbl_text, var, tip in [
            ("Gas Price ($/gal):",   self.gas_price_var,
             "Current price of gasoline per gallon"),
            ("Electricity ($/kWh):", self.electricity_price_var,
             "Current price of electricity per kWh"),
            ("Discount Rate (%):",   self.discount_rate_var,
             "Annual discount rate for future cost calculations"),
        ]:
            rf = ttk.Frame(cost_frame)
            rf.pack(fill=tk.X, padx=5, pady=2)
            lbl = ttk.Label(rf, text=lbl_text)
            lbl.pack(side=tk.LEFT)
            ttk.Entry(rf, textvariable=var, width=8).pack(side=tk.RIGHT)
            SimpleTooltip(lbl, tip)

        # ── Vehicle tab ───────────────────────────────────────────────────────
        vehicle_frame = ttk.LabelFrame(params_notebook, text="Vehicle Parameters")
        params_notebook.add(vehicle_frame, text="Vehicle")

        for lbl_text, var, tip in [
            ("EV Efficiency (kWh/mi):",      self.ev_efficiency_var,
             "Average electricity consumption per mile for electric vehicles"),
            ("Battery Degradation (%/yr):",   self.battery_degradation_var,
             "Annual EV battery capacity loss (%).\nTypical range: 1–3%/yr. Default: 2%/yr."),
            ("ICE Residual Value (%):",        self.residual_ice_var,
             "Estimated resale value of ICE vehicle at end of analysis period.\nDefault: 15%"),
            ("EV Residual Value (%):",         self.residual_ev_var,
             "Estimated resale value of EV at end of analysis period.\nDefault: 20%"),
        ]:
            rf = ttk.Frame(vehicle_frame)
            rf.pack(fill=tk.X, padx=5, pady=2)
            lbl = ttk.Label(rf, text=lbl_text)
            lbl.pack(side=tk.LEFT)
            ttk.Entry(rf, textvariable=var, width=8).pack(side=tk.RIGHT)
            SimpleTooltip(lbl, tip)

        years_frame = ttk.Frame(vehicle_frame)
        years_frame.pack(fill=tk.X, padx=5, pady=2)
        years_label = ttk.Label(years_frame, text="Analysis Period (yrs):")
        years_label.pack(side=tk.LEFT)
        ttk.Spinbox(years_frame, from_=1, to=20,
                    textvariable=self.analysis_years_var,
                    width=5).pack(side=tk.RIGHT)
        SimpleTooltip(years_label, "Number of years to consider in the analysis")

        # Residual values note
        ttk.Label(vehicle_frame, text="Residual Values (% of purchase price at end of analysis):",
                  font=("", 8), foreground="#555555").pack(anchor="w", padx=5, pady=(4, 0))

        # Note: Charging parameters have been moved to the dedicated Charging tab
        # (main notebook tab 5).  The tk.Vars (charging_pattern_var, charging_start_var,
        # charging_end_var, power_levels, power_level_var) are still initialised in
        # __init__ so that the analysis engine can read them; they are now surfaced
        # in the Charging tab UI instead of here.

        # ── Fleet tab ─────────────────────────────────────────────────────────
        fleet_frame = ttk.LabelFrame(params_notebook, text="Fleet Settings")
        params_notebook.add(fleet_frame, text="Fleet")

        ttk.Label(fleet_frame, text="CARB Fleet Type:").grid(
            row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 2))

        for i, (val, label) in enumerate([
            ("hpf",          "High-Priority Fleet (HPF)"),
            ("non_hpf",      "Non-HPF (State/Local Govt)"),
            ("state_agency", "State Agency"),
        ]):
            ttk.Radiobutton(
                fleet_frame, text=label,
                variable=self._fleet_type_var, value=val,
            ).grid(row=i + 1, column=0, columnspan=2, sticky="w", padx=18, pady=1)

        ttk.Label(
            fleet_frame,
            text="Affects CARB ACF mandate deadlines for Category B vehicles.\n"
                 "Verify non-HPF milestones against current CARB regulatory text.",
            foreground="#666666",
            justify="left",
            wraplength=220,
        ).grid(row=4, column=0, columnspan=2, sticky="w", padx=18, pady=(4, 6))

        ttk.Separator(fleet_frame, orient="horizontal").grid(
            row=5, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 6))

        ttk.Label(fleet_frame, text="Max Annual Replacements:").grid(
            row=6, column=0, sticky="w", padx=8, pady=(0, 2))
        max_spin = ttk.Spinbox(
            fleet_frame, from_=0, to=999, increment=1,
            textvariable=self._max_per_year_var, width=6,
        )
        max_spin.grid(row=6, column=1, sticky="w", padx=(0, 8), pady=(0, 2))
        SimpleTooltip(max_spin,
            "Maximum vehicles to schedule per year.\n"
            "0 = no cap (auto even-spread across horizon).\n"
            ">0 = greedy-fill left to right; earlier years fill first.")
        ttk.Label(
            fleet_frame,
            text="Set to 0 for automatic even distribution.\n"
                 "Use >0 to enforce a per-year procurement budget.",
            foreground="#666666",
            justify="left",
            wraplength=220,
        ).grid(row=7, column=0, columnspan=2, sticky="w", padx=18, pady=(0, 6))

        fleet_frame.columnconfigure(0, weight=1)

        # Reset to defaults
        def reset_defaults():
            self.gas_price_var.set(DEFAULT_GAS_PRICE)
            self.electricity_price_var.set(DEFAULT_ELECTRICITY_PRICE)
            self.ev_efficiency_var.set(DEFAULT_EV_EFFICIENCY)
            self.analysis_years_var.set(10)
            self.discount_rate_var.set(5.0)
            self.battery_degradation_var.set(DEFAULT_BATTERY_DEGRADATION)
            self.residual_ice_var.set(DEFAULT_RESIDUAL_VALUE_ICE_PCT)
            self.residual_ev_var.set(DEFAULT_RESIDUAL_VALUE_EV_PCT)
            self.charging_pattern_var.set("standard")
            self.charging_start_var.set(18)
            self.charging_end_var.set(6)

        reset_btn = ttk.Button(parent, text="↺ Reset to Defaults",
                               command=reset_defaults,
                               style="Secondary.TButton")
        reset_btn.pack(fill=tk.X, padx=Spacing.SM, pady=(0, Spacing.SM))
        SimpleTooltip(reset_btn,
            "Reset all parameters to default values\n"
            "Gas: $3.50/gal, Electricity: $0.13/kWh, EV Efficiency: 0.30 kWh/mi")

    # =========================================================================
    # Chart helpers (logic unchanged; targets Chart Browser canvas)
    # =========================================================================

    def _create_initial_chart(self):
        """Create the initial empty chart in the Chart Browser."""
        self.current_figure = Figure(figsize=(10, 4.5), dpi=90)
        self.current_figure.set_tight_layout(True)
        ax = self.current_figure.add_subplot(111)
        ax.text(0.5, 0.5,
                "No data available. Run analysis or select a chart type above.",
                ha="center", va="center", fontsize=11,
                color=Colors.TEXT_TERTIARY)
        ax.axis("off")

        self.current_canvas = FigureCanvasTkAgg(self.current_figure,
                                                master=self.chart_frame)
        self.current_canvas.draw()

        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        self.toolbar = NavigationToolbar2Tk(self.current_canvas, self.toolbar_frame)
        self.toolbar.update()

        canvas_widget = self.current_canvas.get_tk_widget()
        canvas_widget.pack(fill=tk.BOTH, expand=True)

        self.chart_menu = tk.Menu(self, tearoff=0)
        self.chart_menu.add_command(label="Copy Chart",  command=self._copy_chart)
        self.chart_menu.add_command(label="Save Chart...", command=self._save_chart)
        canvas_widget.bind("<Button-3>", self._show_chart_menu)

    def _show_chart_menu(self, event):
        try:
            self.chart_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.chart_menu.grab_release()

    def _copy_chart(self):
        """Copy the current Chart Browser figure to clipboard."""
        if not self.current_figure:
            return
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            self.current_figure.savefig(tmp.name, dpi=300, bbox_inches="tight",
                                        facecolor="white", edgecolor="none")
            import platform
            if platform.system() == "Darwin":
                subprocess.run([
                    "osascript", "-e",
                    f'set the clipboard to (read (POSIX file "{tmp.name}") as «class PNGf»)'
                ], check=False)
            elif platform.system() == "Windows":
                try:
                    from PIL import Image
                    import io as _io
                    image = Image.open(tmp.name)
                    output = _io.BytesIO()
                    image.convert("RGB").save(output, "BMP")
                    bmp_data = output.getvalue()[14:]
                    output.close()
                    import win32clipboard
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, bmp_data)
                    win32clipboard.CloseClipboard()
                except ImportError:
                    logger.warning("win32clipboard not available")
            else:
                subprocess.run(
                    ["xclip", "-selection", "clipboard", "-t", "image/png", "-i", tmp.name],
                    check=False)
        os.unlink(tmp.name)

    def _save_chart(self):
        """Save the Chart Browser figure to a file."""
        if not self.current_figure:
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf"),
                       ("SVG files", "*.svg"), ("All files", "*.*")]
        )
        if file_path:
            self.current_figure.savefig(file_path, dpi=300, bbox_inches="tight",
                                        facecolor="white", edgecolor="none")

    def _update_chart(self):
        """Update the Chart Browser chart."""
        chart_type = self.current_chart_type.get()
        if not chart_type:
            return

        if not self.fleet or not self.fleet.vehicles:
            self.current_figure.clear()
            ax = self.current_figure.add_subplot(111)
            ax.text(0.5, 0.5, "No data available. Run an analysis to view charts.",
                    ha="center", va="center", fontsize=12)
            ax.axis("off")
            self.current_canvas.draw()
            return

        chart_data = self.fleet
        extra_args = {}

        if chart_type == "Fleet Cash Flow" and self.electrification_analysis:
            chart_data = self.electrification_analysis
        elif chart_type in ("Replacement Priority", "Scenario Comparison"):
            chart_data = self.fleet
        elif chart_type in ("Annual Cost Comparison",):
            chart_data = self.fleet
            extra_args = {
                "gas_price":         self.gas_price_var.get(),
                "electricity_price": self.electricity_price_var.get(),
                "ev_efficiency":     self.ev_efficiency_var.get(),
            }
        elif chart_type == "Electrification Potential":
            if self.electrification_analysis:
                chart_data = self.electrification_analysis
                extra_args = {
                    "gas_price":         self.gas_price_var.get(),
                    "electricity_price": self.electricity_price_var.get(),
                    "ev_efficiency":     self.ev_efficiency_var.get(),
                }
        elif chart_type in ("Emissions Reduction", "ROI Analysis"):
            # These functions don't accept gas_price/electricity_price kwargs
            if self.electrification_analysis:
                chart_data = self.electrification_analysis
        elif "Emission" in chart_type and self.emissions_inventory:
            chart_data = self.emissions_inventory
        elif "Charging" in chart_type and self.charging_analysis:
            chart_data = self.charging_analysis

        try:
            self.current_figure.clear()
            ChartFactory.create_chart(
                chart_type=chart_type,
                data=chart_data,
                figure=self.current_figure,
                chart_style=self.chart_style_var.get(),
                color_scheme=self.color_scheme_var.get(),
                **extra_args,
            )
            self.current_canvas.draw()
        except Exception as e:
            logger.error(f"Error creating chart: {e}")
            self.current_figure.clear()
            ax = self.current_figure.add_subplot(111)
            ax.text(0.5, 0.5, f"Error creating chart:\n{str(e)}",
                    ha="center", va="center", fontsize=10)
            ax.axis("off")
            self.current_canvas.draw()

    def _previous_chart(self):
        if not CHART_TYPES:
            return
        try:
            idx = CHART_TYPES.index(self.current_chart_type.get())
        except ValueError:
            idx = 0
        self.current_chart_type.set(CHART_TYPES[(idx - 1) % len(CHART_TYPES)])
        self._update_chart()

    def _next_chart(self):
        if not CHART_TYPES:
            return
        try:
            idx = CHART_TYPES.index(self.current_chart_type.get())
        except ValueError:
            idx = 0
        self.current_chart_type.set(CHART_TYPES[(idx + 1) % len(CHART_TYPES)])
        self._update_chart()

    # =========================================================================
    # Analysis Methods  (logic unchanged from Phase 17)
    # =========================================================================

    def run_electrification_analysis(self):
        """Run electrification analysis on the current fleet."""
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Data", "No fleet data available for analysis.")
            return

        progress = ProgressDialog(self.master,
                                  "Running Electrification Analysis",
                                  "Analyzing electrification potential...")

        def analysis_task():
            try:
                progress.update(20, "Running analysis...")
                self.electrification_analysis = analyze_fleet_electrification(
                    fleet=self.fleet,
                    gas_price=self.gas_price_var.get(),
                    electricity_price=self.electricity_price_var.get(),
                    ev_efficiency=self.ev_efficiency_var.get(),
                    analysis_years=self.analysis_years_var.get(),
                    discount_rate=self.discount_rate_var.get(),
                    incentive_amount=self.incentive_amount_var.get(),
                    battery_degradation=self.battery_degradation_var.get(),
                    residual_value_ice_pct=self.residual_ice_var.get(),
                    residual_value_ev_pct=self.residual_ev_var.get(),
                )
                progress.update(80, "Updating display...")
                electrification_chart = next(
                    (c for c in CHART_TYPES if "Electrification" in c), CHART_TYPES[0])
                self.master.after(0, self._update_summary)
                self.master.after(0, lambda: self.current_chart_type.set(electrification_chart))
                self.master.after(100, self._update_chart)
                if self.on_analysis_complete_callback:
                    self.on_analysis_complete_callback(
                        "Electrification", self.electrification_analysis)
            except Exception as e:
                logger.error(f"Error in electrification analysis: {e}")
                self.master.after(100, lambda: messagebox.showerror(
                    "Analysis Error",
                    f"Error running electrification analysis:\n{str(e)}"))
            finally:
                self.master.after(100, progress.destroy)

        threading.Thread(target=analysis_task, daemon=True).start()

    def run_emissions_analysis(self):
        """Run emissions analysis on the current fleet."""
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Data", "No fleet data available for analysis.")
            return

        progress = ProgressDialog(self.master,
                                  "Running Emissions Analysis",
                                  "Creating emissions inventory...")

        def analysis_task():
            try:
                progress.update(20, "Calculating emissions...")
                self.emissions_inventory = create_emissions_inventory(self.fleet)
                progress.update(80, "Updating display...")
                emissions_chart = next(
                    (c for c in CHART_TYPES if "Emission" in c), CHART_TYPES[0])
                self.master.after(0, self._update_summary)
                self.master.after(0, lambda: self.current_chart_type.set(emissions_chart))
                self.master.after(100, self._update_chart)
                if self.on_analysis_complete_callback:
                    self.on_analysis_complete_callback("Emissions", self.emissions_inventory)
            except Exception as e:
                logger.error(f"Error in emissions analysis: {e}")
                self.master.after(100, lambda: messagebox.showerror(
                    "Analysis Error",
                    f"Error running emissions analysis:\n{str(e)}"))
            finally:
                self.master.after(100, progress.destroy)

        threading.Thread(target=analysis_task, daemon=True).start()

    def run_charging_analysis(self):
        """Run charging infrastructure analysis on the current fleet."""
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Data", "No fleet data available for analysis.")
            return

        progress = ProgressDialog(self.master,
                                  "Running Charging Analysis",
                                  "Analyzing charging infrastructure needs...")

        def analysis_task():
            try:
                power_level = self.power_level_var.get()
                charging_power_kw = self.power_levels[power_level].get()
                progress.update(20, "Calculating requirements...")
                if power_level in ("LP", "MP"):
                    l2_rate   = charging_power_kw
                    dcfc_rate = self.power_levels.get(
                        "HP", tk.DoubleVar(value=50.0)).get()
                else:
                    l2_rate   = self.power_levels.get(
                        "MP", tk.DoubleVar(value=19.2)).get()
                    dcfc_rate = charging_power_kw

                self.charging_analysis = analyze_charging_needs(
                    fleet=self.fleet,
                    daily_usage_pattern=self.charging_pattern_var.get(),
                    charging_window=(self.charging_start_var.get(),
                                     self.charging_end_var.get()),
                    level2_charging_rate=l2_rate,
                    dcfc_charging_rate=dcfc_rate,
                )
                progress.update(80, "Updating display...")
                charging_chart = next(
                    (c for c in CHART_TYPES if "Charging" in c), CHART_TYPES[0])
                self.master.after(0, self._update_summary)
                self.master.after(0, lambda: self.current_chart_type.set(charging_chart))
                self.master.after(100, self._update_chart)
                if self.on_analysis_complete_callback:
                    self.on_analysis_complete_callback("Charging", self.charging_analysis)
            except Exception as e:
                logger.error(f"Error in charging analysis: {e}")
                self.master.after(100, lambda: messagebox.showerror(
                    "Analysis Error",
                    f"Error running charging analysis:\n{str(e)}"))
            finally:
                self.master.after(100, progress.destroy)

        threading.Thread(target=analysis_task, daemon=True).start()

    def run_full_analysis(self):
        """Run all three analyses: electrification → emissions → charging."""
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Data", "No fleet data available for analysis.")
            return

        progress = ProgressDialog(
            self.master, "Running Full Analysis",
            "Running electrification, emissions, and charging analyses...")

        def full_task():
            try:
                # Re-run electrification timeline with the selected fleet type so
                # any deadlines that depend on HPF/Non-HPF are updated before analysis.
                progress.update(5, "Applying fleet type to electrification timeline...")
                from analysis.electrification_timeline import assign_electrification_years
                ft = self._fleet_type_var.get()
                self.fleet.fleet_type = ft
                mpy = self._max_per_year_var.get()
                self.fleet.max_vehicles_per_year = mpy
                # Only recalculate vehicles without manual year overrides
                non_overridden = [
                    v for v in self.fleet.vehicles
                    if v.custom_fields.get("EV Year Overridden") != "Yes"
                ]
                if non_overridden:
                    assign_electrification_years(
                        non_overridden, fleet_type=ft, max_per_year=mpy
                    )

                progress.update(10, "Running electrification analysis...")
                self.electrification_analysis = analyze_fleet_electrification(
                    fleet=self.fleet,
                    gas_price=self.gas_price_var.get(),
                    electricity_price=self.electricity_price_var.get(),
                    ev_efficiency=self.ev_efficiency_var.get(),
                    analysis_years=self.analysis_years_var.get(),
                    discount_rate=self.discount_rate_var.get(),
                    incentive_amount=self.incentive_amount_var.get(),
                    battery_degradation=self.battery_degradation_var.get(),
                    residual_value_ice_pct=self.residual_ice_var.get(),
                    residual_value_ev_pct=self.residual_ev_var.get(),
                )

                progress.update(45, "Creating emissions inventory...")
                self.emissions_inventory = create_emissions_inventory(self.fleet)

                progress.update(70, "Analyzing charging needs...")
                power_level = self.power_level_var.get()
                charging_power_kw = self.power_levels[power_level].get()
                if power_level in ("LP", "MP"):
                    l2_rate   = charging_power_kw
                    dcfc_rate = self.power_levels.get(
                        "HP", tk.DoubleVar(value=50.0)).get()
                else:
                    l2_rate   = self.power_levels.get(
                        "MP", tk.DoubleVar(value=19.2)).get()
                    dcfc_rate = charging_power_kw

                self.charging_analysis = analyze_charging_needs(
                    fleet=self.fleet,
                    daily_usage_pattern=self.charging_pattern_var.get(),
                    charging_window=(self.charging_start_var.get(),
                                     self.charging_end_var.get()),
                    level2_charging_rate=l2_rate,
                    dcfc_charging_rate=dcfc_rate,
                )

                progress.update(88, "Comparing scenarios...")
                # Auto-run scenario comparison in the same background thread
                try:
                    scenario_results = self._compute_all_scenarios_background()
                    self.scenario_results = scenario_results
                    # Push to sharing_data so Present panel can use them for PPTX export
                    if self.sharing_data is not None:
                        try:
                            self.sharing_data.set("scenario_results", scenario_results)
                        except AttributeError:
                            self.sharing_data["scenario_results"] = scenario_results
                except Exception as se:
                    logger.warning(f"Auto scenario comparison failed: {se}")

                progress.update(95, "Updating display...")
                self._analysis_is_stale = False
                self.master.after(0, self._update_summary)
                self.master.after(0, self._draw_scenario_emissions_chart)
                self.master.after(0, self._draw_scenario_cost_chart)
                self.master.after(0, self._update_scenario_table)
                self.master.after(100, lambda: self.current_chart_type.set("Fleet Cash Flow"))
                self.master.after(200, self._update_chart)

                if self.on_analysis_complete_callback:
                    self.on_analysis_complete_callback("Full", self.electrification_analysis)

            except Exception as e:
                logger.error(f"Error in full analysis: {e}")
                self.master.after(100, lambda: messagebox.showerror(
                    "Analysis Error",
                    f"Error running full analysis:\n{str(e)}"))
            finally:
                self.master.after(100, progress.destroy)

        threading.Thread(target=full_task, daemon=True).start()

    # =========================================================================
    # Navigation & incentive helpers (unchanged)
    # =========================================================================

    def _navigate_to_present(self):
        # Present is now tab index 4 (Timeline tab inserted at index 3)
        try:
            main_window = self.winfo_toplevel()
            if hasattr(main_window, "notebook"):
                main_window.notebook.select(4)
        except Exception as e:
            logger.warning(f"Could not navigate to Present tab: {e}")

    def _navigate_to_timeline(self):
        """Switch to the Electrification Timeline tab (index 3)."""
        try:
            main_window = self.winfo_toplevel()
            if hasattr(main_window, "notebook"):
                main_window.notebook.select(3)
        except Exception as e:
            logger.warning(f"Could not navigate to Timeline tab: {e}")

    def _on_incentive_toggled(self):
        total = sum(amt for var, amt in self._incentive_check_vars if var.get())
        self.incentive_amount_var.set(total)
        self.incentive_total_label.config(text=f"${total:,.0f}" if total > 0 else "$0")

    def _on_state_selected(self, event=None):
        selected = self.state_var.get()

        for widget in self.incentive_check_frame.winfo_children():
            widget.destroy()
        self._incentive_check_vars.clear()
        self.incentive_amount_var.set(0.0)
        self.incentive_total_label.config(text="$0")

        if not selected or selected == "(National Avg)":
            self.gas_price_var.set(3.50)
            self.electricity_price_var.set(0.13)
            return

        rates = get_rates_for_state(selected)
        self.gas_price_var.set(rates["gas_price"])
        self.electricity_price_var.set(rates["electricity_price"])

        incentives = get_all_incentives(selected)
        all_programs = (
            [("Federal", i) for i in incentives.get("federal_incentives", [])] +
            [("State",   i) for i in incentives.get("state_incentives",   [])]
        )

        if not all_programs:
            ttk.Label(self.incentive_check_frame, text="No incentives available",
                      font=("", 8), foreground="#888888").pack(anchor=tk.W)
            return

        for _source, prog in all_programs:
            amount = prog.get("max_amount", 0)
            if amount <= 0:
                continue
            var = tk.BooleanVar(value=False)
            self._incentive_check_vars.append((var, amount))
            label_text = f"{prog['name']} — up to ${amount:,}"
            cb = ttk.Checkbutton(self.incentive_check_frame, text=label_text,
                                 variable=var, command=self._on_incentive_toggled)
            cb.pack(anchor=tk.W)
            SimpleTooltip(cb, prog.get("description", ""))

    # =========================================================================
    # Export methods (unchanged)
    # =========================================================================

    def _export_menu(self):
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="PowerPoint Presentation",
                         command=self.export_preliminary_deck)
        menu.add_command(label="Excel / CSV Report",
                         command=self.export_full_report)
        menu.add_separator()
        menu.add_command(label="Current Chart (PNG/PDF/SVG)",
                         command=self.export_current_chart)
        try:
            menu.tk_popup(self.winfo_pointerx(), self.winfo_pointery())
        finally:
            menu.grab_release()

    def _show_timeline_options_dialog(self) -> Optional[dict]:
        """Modal dialog: 'Timelines to Include' at export time.

        Returns a dict {scenario_key: bool} or None if the user cancelled.
        """
        dlg = tk.Toplevel(self.master)
        dlg.title("Timelines to Include")
        dlg.resizable(False, False)
        dlg.transient(self.master)
        dlg.grab_set()

        # Centre over master
        dlg.geometry(
            f"+{self.master.winfo_rootx() + 80}"
            f"+{self.master.winfo_rooty() + 80}"
        )

        ttk.Label(
            dlg,
            text="Choose which timeline scenarios to include\n"
                 "as comparison columns in the Replacement Schedule sheet.",
            justify="left",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY),
        ).pack(padx=20, pady=(16, 8), anchor="w")

        scenario_labels = {
            "aggressive":     "Aggressive — electrify full fleet by 2030",
            "moderate":       "Moderate — electrify full fleet by 2035",
            "conservative":   "Conservative — electrify full fleet by 2040",
            "acf_compliance": "ACF Compliance Only — mandate-subject vehicles by 2035",
        }
        vars_: Dict[str, tk.BooleanVar] = {}
        for key, label in scenario_labels.items():
            var = tk.BooleanVar(value=True)
            vars_[key] = var
            ttk.Checkbutton(dlg, text=label, variable=var).pack(
                anchor="w", padx=28, pady=2
            )

        ttk.Label(
            dlg,
            text="Note: manual overrides are always reflected in the\n"
                 "current timeline (highlighted in amber).",
            foreground=Colors.TEXT_TERTIARY,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL),
            justify="left",
        ).pack(padx=20, pady=(10, 4), anchor="w")

        result: list = [None]

        def on_ok():
            result[0] = {k: v.get() for k, v in vars_.items()}
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        btn_row = ttk.Frame(dlg)
        btn_row.pack(fill=tk.X, padx=20, pady=12)
        ttk.Button(btn_row, text="Export", command=on_ok,
                   style="Accent.TButton").pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(btn_row, text="Cancel", command=on_cancel,
                   style="Secondary.TButton").pack(side=tk.RIGHT)

        dlg.wait_window()
        return result[0]

    def export_full_report(self):
        """Export a comprehensive report with all analyses."""
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Data", "No fleet data available for export.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Export Full Report",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("PDF Files", "*.pdf"),
                       ("All Files", "*.*")]
        )
        if not filepath:
            return

        # Show "Timelines to Include" dialog for Excel exports
        _, ext = os.path.splitext(filepath.lower())
        timeline_options: Optional[dict] = None
        if ext == ".xlsx":
            timeline_options = self._show_timeline_options_dialog()
            if timeline_options is None:
                return  # user cancelled

        progress = ProgressDialog(self.master, "Generating Report",
                                  "Please wait while the report is being generated...")

        def export_task():
            try:
                progress.update(25, "Preparing data...")

                if not self.electrification_analysis:
                    progress.update(40, "Running electrification analysis...")
                    self.electrification_analysis = analyze_fleet_electrification(
                        fleet=self.fleet,
                        gas_price=self.gas_price_var.get(),
                        electricity_price=self.electricity_price_var.get(),
                        ev_efficiency=self.ev_efficiency_var.get(),
                        analysis_years=self.analysis_years_var.get(),
                        discount_rate=self.discount_rate_var.get(),
                        incentive_amount=self.incentive_amount_var.get(),
                        battery_degradation=self.battery_degradation_var.get(),
                        residual_value_ice_pct=self.residual_ice_var.get(),
                        residual_value_ev_pct=self.residual_ev_var.get(),
                    )

                if not self.emissions_inventory:
                    progress.update(55, "Creating emissions inventory...")
                    self.emissions_inventory = create_emissions_inventory(self.fleet)

                if not self.charging_analysis:
                    progress.update(70, "Analyzing charging needs...")
                    power_level = self.power_level_var.get()
                    charging_power_kw = self.power_levels[power_level].get()
                    if power_level in ("LP", "MP"):
                        l2_rate   = charging_power_kw
                        dcfc_rate = self.power_levels["HP"].get()
                    else:
                        l2_rate   = self.power_levels["MP"].get()
                        dcfc_rate = charging_power_kw
                    self.charging_analysis = analyze_charging_needs(
                        fleet=self.fleet,
                        daily_usage_pattern=self.charging_pattern_var.get(),
                        charging_window=(self.charging_start_var.get(),
                                         self.charging_end_var.get()),
                        level2_charging_rate=l2_rate,
                        dcfc_charging_rate=dcfc_rate,
                    )

                progress.update(85, "Generating report...")
                generator = ReportGeneratorFactory.create_generator(filepath)
                if generator:
                    success = generator.generate(
                        fleet=self.fleet,
                        analysis=self.electrification_analysis,
                        charging=self.charging_analysis,
                        emissions=self.emissions_inventory,
                        timeline_options=timeline_options,
                    )
                    if success:
                        progress.update(100, "Report complete!")
                        if self.on_report_generation_callback:
                            self.on_report_generation_callback(filepath)
                        else:
                            self.master.after(500, lambda: messagebox.showinfo(
                                "Export Complete", f"Report exported to:\n{filepath}"))
                    else:
                        self.master.after(500, lambda: messagebox.showerror(
                            "Export Failed", "Failed to generate the report."))
                else:
                    self.master.after(500, lambda: messagebox.showerror(
                        "Export Failed",
                        "Unsupported file format or missing dependencies."))

            except Exception as e:
                logger.error(f"Error exporting report: {e}")
                self.master.after(500, lambda: messagebox.showerror(
                    "Export Failed", f"Error exporting report: {str(e)}"))
            finally:
                self.master.after(100, progress.destroy)

        threading.Thread(target=export_task, daemon=True).start()

    def export_current_chart(self):
        """Export the currently displayed Chart Browser chart."""
        if not self.current_figure:
            messagebox.showinfo("No Chart", "No chart available for export.")
            return
        filepath = filedialog.asksaveasfilename(
            title="Export Chart",
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png"), ("PDF File", "*.pdf"),
                       ("SVG Image", "*.svg"), ("All Files", "*.*")]
        )
        if not filepath:
            return
        try:
            self.current_figure.savefig(filepath, dpi=300, bbox_inches="tight")
            messagebox.showinfo("Export Complete", f"Chart exported to:\n{filepath}")
        except Exception as e:
            logger.error(f"Error exporting chart: {e}")
            messagebox.showerror("Export Failed", f"Error exporting chart: {str(e)}")

    def export_preliminary_deck(self):
        """Export preliminary PowerPoint deck."""
        if not PPTX_EXPORT_AVAILABLE:
            messagebox.showerror(
                "PowerPoint Export Unavailable",
                "PowerPoint export is not available. "
                "Please ensure python-pptx is installed.")
            return
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Data", "No fleet data available for PowerPoint export.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Export Preliminary Deck",
            defaultextension=".pptx",
            filetypes=[("PowerPoint Files", "*.pptx"), ("All Files", "*.*")]
        )
        if not filepath:
            return

        progress = ProgressDialog(self.master, "Generating PowerPoint",
                                  "Please wait while the presentation is being generated...")

        def export_task():
            try:
                progress.update(25, "Preparing fleet data...")
                export_data = {
                    "fleet":       self.fleet,
                    "client_name": getattr(self.fleet, "client_name", "Client"),
                    "stage":       "Preliminary",
                }
                progress.update(50, "Creating presentation...")
                result_path = export_prelim_deck(
                    data=export_data, template_path=None, out_path=filepath)
                progress.update(100, "Presentation complete!")

                def show_success():
                    d = tk.Toplevel(self.master)
                    d.title("Export Complete")
                    d.geometry("500x200")
                    d.transient(self.master)
                    d.grab_set()
                    d.geometry(f"+{self.master.winfo_rootx()+50}"
                               f"+{self.master.winfo_rooty()+50}")
                    ttk.Label(d,
                              text="PowerPoint presentation generated successfully!",
                              font=("", 12)).pack(pady=20)
                    pf = ttk.Frame(d)
                    pf.pack(fill=tk.X, padx=20, pady=10)
                    ttk.Label(pf, text="Saved to:").pack(anchor=tk.W)
                    pt = tk.Text(pf, height=2, wrap=tk.WORD)
                    pt.insert(tk.END, result_path)
                    pt.config(state=tk.DISABLED)
                    pt.pack(fill=tk.X, pady=5)
                    bf = ttk.Frame(d)
                    bf.pack(pady=20)
                    def copy_path():
                        self.master.clipboard_clear()
                        self.master.clipboard_append(result_path)
                        cb.config(text="Copied!")
                        self.master.after(1000, lambda: cb.config(text="Copy Path"))
                    cb = ttk.Button(bf, text="Copy Path", command=copy_path)
                    cb.pack(side=tk.LEFT, padx=5)
                    ttk.Button(bf, text="Close",
                               command=d.destroy).pack(side=tk.LEFT, padx=5)

                self.master.after(500, show_success)
            except Exception as e:
                logger.error(f"Error exporting PowerPoint: {e}")
                self.master.after(500, lambda: messagebox.showerror(
                    "Export Failed",
                    f"Error exporting PowerPoint presentation:\n{str(e)}"))
            finally:
                self.master.after(100, progress.destroy)

        threading.Thread(target=export_task, daemon=True).start()

    # =========================================================================
    # Scenario Comparison (logic unchanged; chart targets dedicated canvas)
    # =========================================================================

    def _compute_all_scenarios_background(self):
        """Compute scenario comparison + Current Plan (pure calculation, no UI).

        Safe to call from a background thread.  Returns the results dict
        that can be assigned to self.scenario_results.

        Uses the user-configured horizon year (self._scenario_horizon_var)
        to override the end_year of each scope-based preset scenario.
        """
        import dataclasses

        selected = [key for key, var in self._scenario_vars.items() if var.get()]
        horizon = self._scenario_horizon_var.get()

        # Build temporary scenario objects with the user-configured end year
        custom_scenarios = []
        for key in selected:
            if key in PRESET_SCENARIOS:
                base = PRESET_SCENARIOS[key]
                custom_scenarios.append(
                    dataclasses.replace(base, end_year=horizon)
                )

        results = compare_scenarios(
            vehicles=self.fleet.vehicles,
            custom_scenarios=custom_scenarios if custom_scenarios else None,
        )

        if self._current_plan_var.get() and self.fleet and self.fleet.vehicles:
            cp = _compute_current_plan_result(self.fleet.vehicles)
            results["scenarios"].insert(0, cp)
            for yr in cp.get("cumulative_vehicles", {}):
                results["all_years"] = sorted(
                    set(list(results["all_years"]) + [yr]))

        return results

    def _run_scenario_comparison(self):
        """Run selected preset scenarios + Current Plan; update charts + table."""
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Data", "No fleet data available for scenario comparison.")
            return

        selected = [key for key, var in self._scenario_vars.items() if var.get()]
        if not selected and not self._current_plan_var.get():
            messagebox.showinfo("No Scenarios Selected",
                                "Select at least one scenario or enable Current Plan.")
            return

        progress = ProgressDialog(self.master, "Comparing Scenarios",
                                  "Running scenario comparison...")

        def task():
            try:
                progress.update(20, "Computing scenarios...")
                self.scenario_results = self._compute_all_scenarios_background()
                if self.sharing_data is not None:
                    try:
                        self.sharing_data.set("scenario_results", self.scenario_results)
                    except AttributeError:
                        self.sharing_data["scenario_results"] = self.scenario_results
                progress.update(80, "Updating display...")
                self.master.after(0, self._draw_scenario_emissions_chart)
                self.master.after(0, self._draw_scenario_cost_chart)
                self.master.after(0, self._update_scenario_table)
            except Exception as exc:
                logger.error(f"Error in scenario comparison: {exc}")
                self.master.after(100, lambda: messagebox.showerror(
                    "Scenario Error",
                    f"Error running scenario comparison:\n{exc}"))
            finally:
                self.master.after(100, progress.destroy)

        threading.Thread(target=task, daemon=True).start()

    def _draw_scenario_emissions_chart(self):
        """Draw annual CO₂ remaining by scenario on the dedicated scenario canvas.

        Y: annual fleet CO₂ (MT CO₂e/yr). Baseline = dashed grey.
        """
        if not self.scenario_results or self._scenario_fig is None:
            return

        results = self.scenario_results.get("scenarios", [])
        if not results:
            return

        current_year = datetime.datetime.now().year

        # Baseline CO₂
        baseline_co2 = 0.0
        if self.emissions_inventory and hasattr(self.emissions_inventory, "total_emissions"):
            baseline_co2 = self.emissions_inventory.total_emissions
        elif self.fleet and self.fleet.vehicles:
            for v in self.fleet.vehicles:
                annual_miles = v.annual_mileage or 12000
                mpg = v.fuel_economy.combined_mpg or 0
                if mpg > 0:
                    baseline_co2 += (8900.0 / mpg * annual_miles) / 1_000_000

        all_years = sorted(self.scenario_results.get("all_years", []))
        if not all_years:
            return
        end_year   = max(all_years)
        plot_years = list(range(current_year, end_year + 2))

        self._scenario_fig.clear()
        ax = self._scenario_fig.add_subplot(111)

        ax.axhline(y=baseline_co2, color="#888888", linestyle="--", linewidth=1.5,
                   label="Baseline (No Electrification)", alpha=0.8)

        scenario_colors = [PRIMARY_HEX_1, "#e05c00", "#2a7a2a", "#7b52ab"]
        for i, result in enumerate(results):
            if result["total_vehicles"] == 0:
                continue
            color      = scenario_colors[i % len(scenario_colors)]
            cumul_co2  = result.get("cumulative_co2_reduction", {})
            cum        = 0.0
            emissions_line = []
            for yr in plot_years:
                if yr in cumul_co2:
                    cum = cumul_co2[yr]
                elif yr > end_year:
                    cum = cumul_co2.get(end_year, cum)
                emissions_line.append(max(0.0, baseline_co2 - cum))
            ax.plot(plot_years, emissions_line, color=color, linewidth=2,
                    marker="o", markersize=3, label=result["name"])

        ax.set_xlabel("Year")
        ax.set_ylabel("Annual Fleet CO₂ (MT CO₂e/yr)")
        ax.set_title("Annual Fleet Emissions by Scenario")
        ax.legend(loc="upper right", fontsize=8)
        ax.set_xlim(current_year - 0.5, end_year + 1.5)
        ax.set_ylim(bottom=0)
        ax.grid(axis="y", alpha=0.3)

        import matplotlib.ticker as mticker
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))

        self._scenario_fig.tight_layout()
        self._scenario_canvas.draw()

    def _draw_scenario_cost_chart(self):
        """Draw cumulative EV fleet investment cost by scenario on the cost canvas."""
        if self._scenario_cost_fig is None or self._scenario_cost_canvas is None:
            return

        self._scenario_cost_fig.clear()
        ax = self._scenario_cost_fig.add_subplot(111)

        if not self.scenario_results:
            ax.axis("off")
            self._scenario_cost_canvas.draw()
            return

        results = self.scenario_results.get("scenarios", [])
        all_years = sorted(self.scenario_results.get("all_years", []))
        if not results or not all_years:
            ax.axis("off")
            self._scenario_cost_canvas.draw()
            return

        current_year = datetime.datetime.now().year
        end_year = max(all_years)
        plot_years = list(range(current_year, end_year + 2))

        scenario_colors = [PRIMARY_HEX_1, "#e05c00", "#2a7a2a", "#7b52ab", "#0077b6"]
        for i, result in enumerate(results):
            if result["total_vehicles"] == 0:
                continue
            color = scenario_colors[i % len(scenario_colors)]
            cumul_cost = result.get("cumulative_cost", {})
            is_current_plan = result.get("name") == "Current Plan"

            cum = 0.0
            cost_line = []
            for yr in plot_years:
                if yr in cumul_cost:
                    cum = cumul_cost[yr]
                elif yr > end_year:
                    cum = cumul_cost.get(end_year, cum)
                cost_line.append(cum / 1_000_000)  # scale to $M

            linestyle = "--" if is_current_plan else "-"
            ax.plot(plot_years, cost_line, color=color, linewidth=2,
                    linestyle=linestyle, marker="o", markersize=3,
                    label=result["name"])

        import matplotlib.ticker as mticker
        ax.set_xlabel("Year")
        ax.set_ylabel("Cumulative EV Fleet Spend ($M)")
        ax.set_title("Cumulative Fleet Investment by Scenario")
        ax.legend(loc="upper left", fontsize=8)
        ax.set_xlim(current_year - 0.5, end_year + 1.5)
        ax.set_ylim(bottom=0)
        ax.grid(axis="y", alpha=0.3)
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda x, _: f"${x:,.1f}M"))

        self._scenario_cost_fig.tight_layout()
        self._scenario_cost_canvas.draw()

    def _update_scenario_table(self):
        """Populate the vehicles-per-year table in the Scenario section."""
        if not self.scenario_results or self._scenario_table_outer is None:
            return

        for widget in self._scenario_table_outer.winfo_children():
            widget.destroy()

        results   = self.scenario_results.get("scenarios", [])
        all_years = sorted(self.scenario_results.get("all_years", []))

        if not results or not all_years:
            ttk.Label(self._scenario_table_outer, text="No scenario data available.",
                      foreground=Colors.TEXT_DISABLED).pack(padx=10, pady=8)
            return

        display_years = all_years[:20]
        col_ids  = ["scenario"] + [str(y) for y in display_years] + ["total"]

        tree = ttk.Treeview(self._scenario_table_outer, columns=col_ids,
                            show="headings",
                            height=min(len(results) + 1, 5))
        tree.heading("scenario", text="Scenario", anchor="w")
        tree.column("scenario", width=140, anchor="w", stretch=False)
        for yr in display_years:
            tree.heading(str(yr), text=str(yr), anchor="center")
            tree.column(str(yr), width=44, anchor="center", stretch=False)
        tree.heading("total", text="Total", anchor="center")
        tree.column("total", width=50, anchor="center", stretch=False)

        for result in results:
            row   = [result["name"]]
            total = 0
            for yr in display_years:
                count = result.get("vehicles_per_year", {}).get(yr, 0)
                total += count
                row.append(str(count) if count > 0 else "—")
            row.append(str(total))
            tree.insert("", "end", values=row)

        h_scroll = ttk.Scrollbar(self._scenario_table_outer, orient="horizontal",
                                  command=tree.xview)
        tree.configure(xscrollcommand=h_scroll.set)
        tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=(5, 0))
        h_scroll.pack(fill=tk.X, padx=5, pady=(0, 5))

    # =========================================================================
    # Dashboard update methods
    # =========================================================================

    def _update_export_button_states(self):
        """Enable or disable export buttons based on whether a fleet is loaded."""
        state = "normal" if (self.fleet and self.fleet.vehicles) else "disabled"
        for btn in (self._export_btn, self._present_btn):
            if btn is not None:
                btn.configure(state=state)

    def _update_summary(self):
        """Coordinator: refresh all dashboard sections after analysis or fleet change."""
        self._update_fleet_snapshot_kpis()
        self._update_acf_donut()
        self._update_priority_vehicles()
        self._update_tco_kpis()
        self._update_status_label()
        self._update_gantt_section()
        self._update_export_button_states()

        # Flag gallery thumbnails for re-render with fresh analysis data
        if hasattr(self, "_gallery_thumb_rendered"):
            self._gallery_thumb_rendered = False

        # Auto-collapse Parameters once analysis has run
        if (self.electrification_analysis is not None
                and self._parameters_outer is not None
                and not self._parameters_outer._collapsed):
            self._parameters_outer._toggle()

    def _update_status_label(self):
        if self.status_label is None:
            return
        if not self.fleet or not self.fleet.vehicles:
            self.status_label.config(
                text="Ready — load a fleet and click Run Full Analysis",
                foreground=Colors.TEXT_TERTIARY)
        elif self.electrification_analysis is not None:
            n = len(self.fleet.vehicles)
            if getattr(self, "_analysis_is_stale", False):
                self.status_label.config(
                    text=f"Parameters changed — re-run analysis to update results",
                    foreground="#E65100")
            else:
                self.status_label.config(
                    text=f"Analysis complete — {n} vehicle{'s' if n != 1 else ''}",
                    foreground=Colors.SUCCESS)
        else:
            n = len(self.fleet.vehicles)
            self.status_label.config(
                text=f"{n} vehicle{'s' if n != 1 else ''} loaded — "
                     f"click Run Full Analysis",
                foreground=Colors.TEXT_SECONDARY)

    def _on_param_changed(self, *_args):
        """Mark results stale whenever a parameter var changes post-analysis."""
        if self.electrification_analysis is not None:
            self._analysis_is_stale = True
            self._update_status_label()

    def _update_fleet_snapshot_kpis(self):
        """Update the 4 Fleet Snapshot chips from fleet.vehicles (no analysis needed)."""
        if not self._snapshot_kpi_labels:
            return
        if not self.fleet or not self.fleet.vehicles:
            for lbl in self._snapshot_kpi_labels.values():
                lbl.config(text="—")
            return

        vehicles = self.fleet.vehicles
        n = len(vehicles)

        self._snapshot_kpi_labels["fleet_size"].config(text=str(n))

        acf_b = sum(1 for v in vehicles
                    if v.custom_fields.get("ACF Category", "") == "B")
        self._snapshot_kpi_labels["acf_b_count"].config(text=str(acf_b))

        avg_mpg = self.fleet.avg_mpg
        self._snapshot_kpi_labels["avg_mpg"].config(
            text=f"{avg_mpg:.1f}" if avg_mpg and avg_mpg > 0 else "—")

        with_mpg = sum(
            1 for v in vehicles
            if v.fuel_economy.combined_mpg and v.fuel_economy.combined_mpg > 0)
        coverage_pct = (with_mpg / n * 100) if n > 0 else 0
        self._snapshot_kpi_labels["mpg_coverage"].config(
            text=f"{coverage_pct:.0f}%" if n > 0 else "—")

    def _update_acf_donut(self):
        """Render ACF compliance donut from fleet.vehicles (no analysis needed)."""
        if self._acf_fig is None or self._acf_canvas is None:
            return

        self._acf_fig.clear()

        if not self.fleet or not self.fleet.vehicles:
            ax = self._acf_fig.add_subplot(111)
            ax.text(0.5, 0.5, "No data", ha="center", va="center",
                    fontsize=9, color=Colors.TEXT_DISABLED)
            ax.axis("off")
            self._acf_canvas.draw()
            for cat, lbl in self._acf_count_labels.items():
                lbl.config(text=f"Category {cat}: —")
            return

        counts: Dict[str, int] = {"ZEV": 0, "A": 0, "B": 0, "C": 0, "D": 0}
        for v in self.fleet.vehicles:
            cat = v.custom_fields.get("_acf_code", "")
            if cat in counts:
                counts[cat] += 1

        total       = sum(counts.values())
        active_cats = [(cat, cnt) for cat, cnt in counts.items() if cnt > 0]

        ax = self._acf_fig.add_subplot(111)
        if not active_cats or total == 0:
            ax.text(0.5, 0.5, "No ACF data\n(run processing first)",
                    ha="center", va="center", fontsize=8,
                    color=Colors.TEXT_DISABLED)
            ax.axis("off")
        else:
            sizes       = [cnt  for _, cnt in active_cats]
            colors_plot = [ACF_COLORS.get(cat, "#AAAAAA")
                           for cat, _ in active_cats]
            ax.pie(sizes, colors=colors_plot, startangle=90,
                   wedgeprops={"width": 0.55,
                               "edgecolor": Colors.SURFACE,
                               "linewidth": 1.5})
            ax.text(0, 0, f"{total}\nvehicles",
                    ha="center", va="center", fontsize=8,
                    fontweight="bold", color=Colors.TEXT_PRIMARY)
            ax.axis("equal")

        self._acf_fig.tight_layout(pad=0.2)
        self._acf_canvas.draw()

        for cat, lbl in self._acf_count_labels.items():
            lbl.config(text=f"Category {cat}: {counts.get(cat, 0)}")

    def _update_priority_vehicles(self):
        """Populate Top 5 Priority Vehicles table from electrification_analysis."""
        if self._top5_tree is None:
            return

        if self.electrification_analysis is None or not self.fleet or \
                not self.fleet.vehicles:
            # Hide table, show placeholder
            self._top5_tree.pack_forget()
            if self._top5_hscroll:
                self._top5_hscroll.pack_forget()
            if self._top5_placeholder:
                self._top5_placeholder.pack(padx=12, pady=24)
            return

        # Show table, hide placeholder
        if self._top5_placeholder:
            self._top5_placeholder.pack_forget()
        self._top5_tree.pack(fill=tk.BOTH, expand=True, padx=4, pady=(4, 0))
        if self._top5_hscroll:
            self._top5_hscroll.pack(fill=tk.X, padx=4, pady=(0, 4))

        # Clear existing rows
        for iid in self._top5_tree.get_children():
            self._top5_tree.delete(iid)

        vin_to_vehicle = {v.vin: v for v in self.fleet.vehicles}
        current_year   = datetime.datetime.now().year
        vehicle_results = self.electrification_analysis.vehicle_results
        top_vins        = self.electrification_analysis.prioritized_vehicles[:5]

        for vin in top_vins:
            v = vin_to_vehicle.get(vin)
            if v is None:
                continue
            vr = vehicle_results.get(vin, {})

            asset_id  = v.asset_id or vin[:8]
            year      = v.vehicle_id.year  or "—"
            make      = v.vehicle_id.make  or "—"
            model     = v.vehicle_id.model or "—"
            wt_class  = v.vehicle_id.vehicle_class or "—"
            body_type = v.vehicle_id.body_class    or "—"
            acf_cat   = v.custom_fields.get("ACF Category",   "—")
            ev_year   = v.custom_fields.get("Proposed EV Year", "—")

            npv = vr.get("total_npv_savings", 0)
            priority_str = f"${npv:,.0f}" if npv else "—"

            try:
                age_str = str(current_year - int(year))
            except (ValueError, TypeError):
                age_str = "—"

            odo = v.odometer
            odo_str = f"{int(odo):,}" if odo > 0 else "—"

            self._top5_tree.insert("", "end", values=(
                asset_id, year, make, model, wt_class, body_type,
                acf_cat, ev_year, priority_str, age_str, odo_str,
            ))

    def _update_tco_kpis(self):
        """Update TCO & Financial Summary cards."""
        if not self._tco_kpi_labels:
            return

        if self.electrification_analysis:
            savings = self.electrification_analysis.total_savings
            self._tco_kpi_labels["annual_savings"].config(
                text=f"${savings:,.0f}/yr" if savings > 0 else "—")
            payback = self.electrification_analysis.payback_period
            self._tco_kpi_labels["payback"].config(
                text=f"{payback:.1f} yr" if payback > 0 else "—")
        else:
            self._tco_kpi_labels["annual_savings"].config(text="—")
            self._tco_kpi_labels["payback"].config(text="—")

        if self.charging_analysis:
            cost = self.charging_analysis.estimated_installation_cost
            self._tco_kpi_labels["infra_cost"].config(
                text=f"${cost:,.0f}" if cost > 0 else "—")
        else:
            self._tco_kpi_labels["infra_cost"].config(text="—")

    def _update_gantt_section(self):
        """Redraw the Gantt chart in the Analysis tab dashboard."""
        if self._gantt_fig is None or self._gantt_canvas is None:
            return
        if not self.fleet or not self.fleet.vehicles:
            self._gantt_fig.clear()
            ax = self._gantt_fig.add_subplot(111)
            ax.text(0.5, 0.5, "No fleet data available.",
                    ha="center", va="center", fontsize=9,
                    color=Colors.TEXT_TERTIARY)
            ax.axis("off")
            self._gantt_canvas.draw()
            return

        self._gantt_fig.clear()
        ax = self._gantt_fig.add_subplot(111)
        view = self._gantt_view_var.get() if self._gantt_view_var else "Grouped by ACF"
        max_val = self._gantt_max_var_analysis.get() if self._gantt_max_var_analysis else "50"
        try:
            max_rows = 0 if max_val == "All" else int(max_val)
        except ValueError:
            max_rows = 50
        try:
            horizon = self._scenario_horizon_var.get()
        except Exception:
            horizon = 2040
        _draw_gantt_chart(
            ax, self.fleet.vehicles, view,
            max_vehicles=max_rows,
            scenario_results=self.scenario_results,
            horizon_year=horizon,
        )
        self._gantt_fig.tight_layout(pad=0.5)
        self._gantt_canvas.draw()

    # =========================================================================
    # Year-Override helpers (Phase 19)
    # =========================================================================

    @staticmethod
    def _apply_ev_year_override(vehicle, new_year_str: str) -> None:
        """Write an EV-year override into vehicle.custom_fields.

        Stores the original system-recommended year the first time an
        override is applied, sets the new year, and marks the vehicle
        as overridden so Excel export can highlight it.
        """
        # Preserve original only once so repeated edits don't clobber it
        if "System Recommended EV Year" not in vehicle.custom_fields:
            vehicle.custom_fields["System Recommended EV Year"] = (
                vehicle.custom_fields.get("Proposed EV Year", "N/A")
            )
        vehicle.custom_fields["Proposed EV Year"] = new_year_str
        vehicle.custom_fields["EV Year Overridden"] = "Yes"

    @staticmethod
    def _reset_ev_year_override(vehicle) -> None:
        """Restore the system-recommended EV year for a single vehicle."""
        original = vehicle.custom_fields.get("System Recommended EV Year")
        if original is not None:
            vehicle.custom_fields["Proposed EV Year"] = original
        vehicle.custom_fields.pop("EV Year Overridden", None)
        vehicle.custom_fields.pop("System Recommended EV Year", None)

    @staticmethod
    def _apply_acf_override(vehicle, new_acf_code: str) -> None:
        """Write an ACF category override into vehicle.custom_fields.

        Stores the original auto-classified ACF code the first time an
        override is applied, then sets the new code on both the display
        field (``ACF Category``) and the internal lookup field
        (``_acf_code``).
        """
        # Preserve original only once so repeated edits don't clobber it
        if "Original ACF Category" not in vehicle.custom_fields:
            vehicle.custom_fields["Original ACF Category"] = (
                vehicle.custom_fields.get("ACF Category", "")
            )
        vehicle.custom_fields["ACF Category"] = new_acf_code
        vehicle.custom_fields["_acf_code"] = new_acf_code
        vehicle.custom_fields["ACF Category Overridden"] = "Yes"

    @staticmethod
    def _reset_acf_override(vehicle) -> None:
        """Restore the original auto-classified ACF category for a single vehicle."""
        original = vehicle.custom_fields.get("Original ACF Category")
        if original is not None:
            vehicle.custom_fields["ACF Category"] = original
            vehicle.custom_fields["_acf_code"] = original
        vehicle.custom_fields.pop("ACF Category Overridden", None)
        vehicle.custom_fields.pop("Original ACF Category", None)

    def _reset_all_overrides(self) -> None:
        """Reset all manual EV-year overrides for the current fleet."""
        if not self.fleet:
            return
        for v in self.fleet.vehicles:
            if v.custom_fields.get("EV Year Overridden"):
                self._reset_ev_year_override(v)
        self._update_gantt_section()

    # =========================================================================
    # Jump to vehicle (Top 5 table → Results tab)
    # =========================================================================

    def _on_top5_double_click(self, event):
        """Double-click on Top 5 table: jump to that vehicle in Results tab."""
        if self._top5_tree is None:
            return
        region = self._top5_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        iid = self._top5_tree.identify_row(event.y)
        if not iid:
            return
        values = self._top5_tree.item(iid, "values")
        if not values:
            return
        asset_id = str(values[0])
        # Find the matching vehicle's VIN by asset_id
        if self.fleet:
            for v in self.fleet.vehicles:
                if (v.asset_id and v.asset_id == asset_id) or \
                        (v.vin and v.vin[:8] == asset_id):
                    self._jump_to_vehicle(v.vin)
                    return

    def _jump_to_vehicle(self, vin: str):
        """Switch to Results tab and select the row matching *vin*."""
        try:
            main_window = self.winfo_toplevel()
            if hasattr(main_window, "notebook"):
                main_window.notebook.select(1)
            if hasattr(main_window, "results_panel"):
                # Give the tab time to render before scrolling
                self.after(50, lambda: main_window.results_panel.select_by_vin(vin))
        except Exception as e:
            logger.warning(f"Could not jump to vehicle {vin}: {e}")

    # =========================================================================
    # Public API (unchanged signatures)
    # =========================================================================

    def set_fleet(self, fleet):
        """Set the fleet data for analysis."""
        self.fleet                    = fleet
        self.electrification_analysis = None
        self.emissions_inventory      = None
        self.charging_analysis        = None
        self.scenario_results         = None
        self._update_summary()
        self._update_chart()

    def update_parameters(self, **kwargs):
        """Batch-update analysis parameters by name."""
        if "gas_price"         in kwargs:
            self.gas_price_var.set(kwargs["gas_price"])
        if "electricity_price" in kwargs:
            self.electricity_price_var.set(kwargs["electricity_price"])
        if "ev_efficiency"     in kwargs:
            self.ev_efficiency_var.set(kwargs["ev_efficiency"])
        if "analysis_years"    in kwargs:
            self.analysis_years_var.set(kwargs["analysis_years"])
        if "discount_rate"     in kwargs:
            self.discount_rate_var.set(kwargs["discount_rate"])
        if "charging_pattern"  in kwargs:
            self.charging_pattern_var.set(kwargs["charging_pattern"])
        if "charging_window"   in kwargs and len(kwargs["charging_window"]) == 2:
            self.charging_start_var.set(kwargs["charging_window"][0])
            self.charging_end_var.set(kwargs["charging_window"][1])

    def copy_selection(self):
        """Copy Top 5 Priority Vehicles as tab-delimited text (Edit > Copy handler)."""
        if self._top5_tree is None:
            return
        headers = ["Asset ID", "Year", "Make", "Model", "Wt Class",
                   "Body Type", "ACF", "EV Year", "Priority $", "Age", "Odometer"]
        rows = ["\t".join(headers)]
        for iid in self._top5_tree.get_children():
            values = self._top5_tree.item(iid, "values")
            rows.append("\t".join(str(v) for v in values))
        if len(rows) > 1:
            try:
                self.clipboard_clear()
                self.clipboard_append("\n".join(rows))
            except tk.TclError:
                pass

    def get_charging_vars(self) -> dict:
        """Return charging parameter tk.Vars for sharing with ChargingPanel."""
        return {
            "power_level_var":      self.power_level_var,
            "power_levels":         self.power_levels,
            "charging_pattern_var": self.charging_pattern_var,
            "charging_start_var":   self.charging_start_var,
            "charging_end_var":     self.charging_end_var,
        }

    def refresh(self):
        """Refresh the display."""
        self._update_summary()
        self._update_chart()

    def on_resize(self):
        """Handle resize events."""
        if self.current_canvas:
            self.current_canvas.draw()
        if self._acf_canvas:
            self._acf_canvas.draw()
        if self._scenario_canvas:
            self._scenario_canvas.draw()
        if self._gantt_canvas:
            self._gantt_canvas.draw()
