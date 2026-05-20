"""
charging_panel.py

Charging tab (Tab 5) — charging infrastructure parameters and analysis results.

Surfaces the charging parameters previously buried in the Analysis tab's Parameters
notebook.  analysis_panel.py owns the authoritative tk.Vars; this panel receives them
via the `analysis_vars` dict passed at construction time so both panels stay in sync.

Phase 28 build-out:
  1 — Var wiring:  accepts analysis_panel's tk.Vars at construction time
  2 — Action bar:  Run Charging Analysis button + status label
  3 — Results UI:  KPI chips, load-profile chart, facility table, recommendations
  4 — Utility rates: state selector, $/kWh + demand charge, annual cost estimate
"""

import re
import tkinter as tk
from tkinter import ttk
from datetime import datetime
from typing import Optional, Dict, Any, Callable, List

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from settings import PRIMARY_HEX_1
from utils import ScrollableFrame
from ui.theme import Fonts
from analysis.rate_database import get_rates_for_state, get_available_states

import logging
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Module-level helpers
# ---------------------------------------------------------------------------

def _detect_state_from_fleet(fleet) -> Optional[str]:
    """
    Scan vehicle custom_fields for a 2-letter US state code.
    Returns the most-common code found, or None.
    """
    if fleet is None:
        return None

    _STATE_RE = re.compile(r'\b([A-Z]{2})\b')
    counts: Dict[str, int] = {}
    search_keys = ("State", "state", "Facility State", "Location", "location")

    for v in fleet.vehicles:
        cf = v.custom_fields
        for key in search_keys:
            val = str(cf.get(key, "") or "")
            for m in _STATE_RE.finditer(val):
                code = m.group(1)
                counts[code] = counts.get(code, 0) + 1

    if not counts:
        return None
    return max(counts, key=counts.get)


class ChargingPanel:
    """Charging Infrastructure tab panel."""

    def __init__(
        self,
        parent_frame: ttk.Frame,
        sharing_data: dict,
        analysis_vars: Optional[Dict[str, Any]] = None,
        on_run_analysis: Optional[Callable] = None,
    ):
        self.parent_frame = parent_frame
        self.sharing_data = sharing_data
        self._on_run_cb   = on_run_analysis
        self._fleet       = None
        self._last_result = None  # last ChargingAnalysis shown

        # ── Phase 1: accept shared vars from analysis_panel (or create own) ─
        if analysis_vars:
            self.power_level_var      = analysis_vars["power_level_var"]
            self.power_levels         = analysis_vars["power_levels"]
            self.charging_pattern_var = analysis_vars["charging_pattern_var"]
            self.charging_start_var   = analysis_vars["charging_start_var"]
            self.charging_end_var     = analysis_vars["charging_end_var"]
        else:
            self.charging_pattern_var = tk.StringVar(value="standard")
            self.charging_start_var   = tk.IntVar(value=18)
            self.charging_end_var     = tk.IntVar(value=6)
            self.power_level_var      = tk.StringVar(value="LP")
            self.power_levels: Dict[str, tk.DoubleVar] = {
                "LP":  tk.DoubleVar(value=7.2),
                "MP":  tk.DoubleVar(value=19.2),
                "HP":  tk.DoubleVar(value=50.0),
                "VHP": tk.DoubleVar(value=150.0),
            }

        # ── Utility rate state var ───────────────────────────────────────────
        self._state_var = tk.StringVar(value="")

        # ── Widget references ────────────────────────────────────────────────
        self._status_label: Optional[ttk.Label]  = None
        self._run_btn: Optional[ttk.Button]      = None

        # Results section
        self._results_body: Optional[tk.Frame]      = None
        self._results_hdr_var: Optional[tk.StringVar] = None
        self._results_expanded: bool                = False
        self._results_toggle: Optional[Callable]    = None

        self._kpi_vars: Dict[str, tk.StringVar]     = {}
        self._chart_frame: Optional[ttk.Frame]      = None
        self._chart_fig: Optional[Figure]           = None
        self._chart_ax                              = None
        self._chart_canvas: Optional[FigureCanvasTkAgg] = None
        self._facility_tree: Optional[ttk.Treeview] = None
        self._facility_frame: Optional[ttk.Frame]   = None
        self._reco_text: Optional[tk.Text]          = None

        # Rate section
        self._rate_kwh_var  = tk.StringVar(value="\u2014")
        self._rate_dem_var  = tk.StringVar(value="\u2014")
        self._rate_cost_var = tk.StringVar(value="\u2014")

        self._build_ui()

    # ─────────────────────────────────────────────────────────────────────────
    # UI Construction
    # ─────────────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self._sf = ScrollableFrame(self.parent_frame)
        self._sf.pack(fill=tk.BOTH, expand=True)
        body = self._sf.scrollable_frame

        # Header
        hdr_frame = ttk.Frame(body, padding=(12, 10))
        hdr_frame.pack(fill=tk.X)
        ttk.Label(
            hdr_frame,
            text="Charging Infrastructure Analysis",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_H1, Fonts.WEIGHT_BOLD),
        ).pack(side=tk.LEFT)
        ttk.Separator(body, orient="horizontal").pack(fill=tk.X, padx=8, pady=(0, 6))

        self._build_action_bar(body)
        self._build_params_section(body)
        self._build_results_section(body)

    def _build_action_bar(self, parent: tk.Frame) -> None:
        bar = ttk.Frame(parent, padding=(8, 4))
        bar.pack(fill=tk.X)

        self._run_btn = ttk.Button(
            bar,
            text="\u26a1 Run Charging Analysis",
            command=self._on_run_clicked,
        )
        self._run_btn.pack(side=tk.LEFT, padx=(0, 10))

        self._status_label = ttk.Label(
            bar,
            text="No analysis run yet",
            foreground="#757575",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL),
        )
        self._status_label.pack(side=tk.LEFT)

        ttk.Separator(parent, orient="horizontal").pack(fill=tk.X, padx=8, pady=(4, 0))

    def _build_params_section(self, parent: tk.Frame) -> None:
        _sec, body = self._collapsible_section(
            parent, "\u25bc  Charging Parameters", expanded=True)

        frm = ttk.Frame(body, padding=(12, 8))
        frm.pack(fill=tk.X)
        frm.columnconfigure(2, weight=1)

        row = 0

        # Default power level
        ttk.Label(frm, text="Default Power Level:").grid(
            row=row, column=0, sticky="w", padx=4, pady=3)
        ttk.Combobox(
            frm, textvariable=self.power_level_var,
            values=["LP", "MP", "HP", "VHP"],
            state="readonly", width=8,
        ).grid(row=row, column=1, sticky="w", padx=4, pady=3)
        ttk.Label(
            frm,
            text="LP=Level 1/2  MP=Mid-Power  HP=DC Fast  VHP=Ultra-Fast",
            foreground="gray",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL),
        ).grid(row=row, column=2, sticky="w", padx=4)
        row += 1

        # Individual power level kW values
        ttk.Label(
            frm, text="Power Levels (kW):",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY),
        ).grid(row=row, column=0, sticky="w", padx=4, pady=(8, 2))
        row += 1
        for code, label in [
            ("LP",  "Level 1/2 (LP)"),
            ("MP",  "Mid-Power (MP)"),
            ("HP",  "DC Fast (HP)"),
            ("VHP", "Ultra-Fast (VHP)"),
        ]:
            ttk.Label(frm, text=f"  {label}:").grid(
                row=row, column=0, sticky="w", padx=20, pady=2)
            ttk.Entry(frm, textvariable=self.power_levels[code], width=8).grid(
                row=row, column=1, sticky="w", padx=4, pady=2)
            ttk.Label(frm, text="kW", foreground="gray").grid(
                row=row, column=2, sticky="w", padx=2)
            row += 1

        ttk.Separator(frm, orient="horizontal").grid(
            row=row, column=0, columnspan=3, sticky="ew", pady=6)
        row += 1

        # Charging pattern
        ttk.Label(frm, text="Charging Pattern:").grid(
            row=row, column=0, sticky="w", padx=4, pady=3)
        ttk.Combobox(
            frm, textvariable=self.charging_pattern_var,
            values=["standard", "overnight", "opportunity", "managed"],
            state="readonly", width=14,
        ).grid(row=row, column=1, sticky="w", padx=4, pady=3)
        row += 1

        # Charging window
        ttk.Label(frm, text="Charging Window:").grid(
            row=row, column=0, sticky="w", padx=4, pady=3)
        win_frame = ttk.Frame(frm)
        win_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=4, pady=3)
        ttk.Label(win_frame, text="Start:").pack(side=tk.LEFT)
        ttk.Spinbox(win_frame, textvariable=self.charging_start_var,
                    from_=0, to=23, width=4).pack(side=tk.LEFT, padx=2)
        ttk.Label(win_frame, text=":00    End:").pack(side=tk.LEFT, padx=(4, 0))
        ttk.Spinbox(win_frame, textvariable=self.charging_end_var,
                    from_=0, to=23, width=4).pack(side=tk.LEFT, padx=2)
        ttk.Label(win_frame, text=":00").pack(side=tk.LEFT)

    def _build_results_section(self, parent: tk.Frame) -> None:
        """Build the collapsible results section (collapsed by default)."""
        hdr_var = tk.StringVar(value="\u25b6  Charging Demand Analysis")
        self._results_hdr_var = hdr_var

        container = ttk.Frame(parent, relief="flat")
        container.pack(fill=tk.X, padx=6, pady=4)

        hdr = tk.Frame(container, bg=PRIMARY_HEX_1, cursor="hand2")
        hdr.pack(fill=tk.X)
        hdr_lbl = tk.Label(
            hdr, textvariable=hdr_var, bg=PRIMARY_HEX_1, fg="white",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_H3, Fonts.WEIGHT_BOLD),
            anchor="w", padx=8, pady=4,
        )
        hdr_lbl.pack(fill=tk.X)

        body = ttk.Frame(container, relief="solid", borderwidth=1)
        self._results_body = body

        def _toggle(event=None):
            if body.winfo_ismapped():
                body.pack_forget()
                hdr_var.set("\u25b6  Charging Demand Analysis")
                self._results_expanded = False
            else:
                body.pack(fill=tk.X)
                hdr_var.set("\u25bc  Charging Demand Analysis")
                self._results_expanded = True

        hdr.bind("<Button-1>", _toggle)
        hdr_lbl.bind("<Button-1>", _toggle)
        self._results_toggle = _toggle

        # ── KPI chips row ────────────────────────────────────────────────────
        kpi_outer = ttk.Frame(body, padding=(8, 8))
        kpi_outer.pack(fill=tk.X)
        kpi_row = ttk.Frame(kpi_outer)
        kpi_row.pack(fill=tk.X)

        for key, label, color in [
            ("l2",   "L2 Chargers",        "#1565C0"),
            ("dcfc", "DCFC Chargers",      "#6A1B9A"),
            ("kw",   "Peak Demand (kW)",   "#E65100"),
            ("cost", "Infrastructure Est.", "#2E7D32"),
        ]:
            chip = ttk.Frame(kpi_row, relief="solid", borderwidth=1, padding=(12, 8))
            chip.pack(side=tk.LEFT, padx=6, pady=4, expand=True, fill=tk.X)
            var = tk.StringVar(value="\u2014")
            self._kpi_vars[key] = var
            ttk.Label(
                chip, text=label, foreground="#616161",
                font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL),
            ).pack()
            ttk.Label(
                chip, textvariable=var,
                font=(Fonts.FAMILY_SANS, Fonts.SIZE_H2, Fonts.WEIGHT_BOLD),
                foreground=color,
            ).pack()

        ttk.Separator(body, orient="horizontal").pack(fill=tk.X, padx=8, pady=2)

        # ── Utility Rates sub-section ────────────────────────────────────────
        self._build_rate_section(body)

        ttk.Separator(body, orient="horizontal").pack(fill=tk.X, padx=8, pady=2)

        # ── Load profile chart ───────────────────────────────────────────────
        self._chart_frame = ttk.Frame(body, padding=(8, 4))
        self._chart_frame.pack(fill=tk.X)
        ttk.Label(
            self._chart_frame,
            text="Hourly Charging Load Profile",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY, Fonts.WEIGHT_BOLD),
        ).pack(anchor="w")

        fig = Figure(figsize=(9, 2.4), dpi=80)
        self._chart_fig = fig
        self._chart_ax  = fig.add_subplot(111)
        self._chart_canvas = FigureCanvasTkAgg(fig, master=self._chart_frame)
        self._chart_canvas.get_tk_widget().pack(fill=tk.X)

        ttk.Separator(body, orient="horizontal").pack(fill=tk.X, padx=8, pady=2)

        # ── Facility breakdown table ─────────────────────────────────────────
        self._facility_frame = ttk.Frame(body, padding=(8, 4))
        # (packed or hidden dynamically in show_results)

        fac_hdr = ttk.Frame(self._facility_frame)
        fac_hdr.pack(fill=tk.X)
        ttk.Label(
            fac_hdr,
            text="By-Facility Breakdown",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY, Fonts.WEIGHT_BOLD),
        ).pack(side=tk.LEFT)

        cols = ("Facility", "Vehicles", "L2 Chargers", "DCFC Chargers", "Est. Cost")
        self._facility_tree = ttk.Treeview(
            self._facility_frame, columns=cols, show="headings", height=5)
        widths = (160, 70, 90, 100, 110)
        for col, w in zip(cols, widths):
            self._facility_tree.heading(col, text=col)
            self._facility_tree.column(
                col, width=w, minwidth=60,
                anchor="w" if col == "Facility" else "center",
            )
        sb_fac = ttk.Scrollbar(
            self._facility_frame, orient="vertical",
            command=self._facility_tree.yview)
        self._facility_tree.configure(yscrollcommand=sb_fac.set)
        self._facility_tree.pack(side=tk.LEFT, fill=tk.X, expand=True)
        sb_fac.pack(side=tk.LEFT, fill=tk.Y)

        ttk.Separator(body, orient="horizontal").pack(fill=tk.X, padx=8, pady=2)

        # ── Recommendations text ─────────────────────────────────────────────
        rec_frame = ttk.Frame(body, padding=(8, 4))
        rec_frame.pack(fill=tk.X, pady=(0, 8))
        self._rec_outer = rec_frame  # kept for reference
        ttk.Label(
            rec_frame,
            text="Recommendations",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY, Fonts.WEIGHT_BOLD),
        ).pack(anchor="w")
        self._reco_text = tk.Text(
            rec_frame, height=7, wrap=tk.WORD, state=tk.DISABLED,
            relief="flat", background="#F8F8F8",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY),
        )
        self._reco_text.pack(fill=tk.X)

    def _build_rate_section(self, parent: tk.Frame) -> None:
        """Utility rates sub-section inside the results body."""
        rate_hdr_var = tk.StringVar(value="\u25bc  Utility Rates")

        container = ttk.Frame(parent)
        container.pack(fill=tk.X, padx=6, pady=2)

        hdr = tk.Frame(container, bg="#5B7553", cursor="hand2")
        hdr.pack(fill=tk.X)
        hdr_lbl = tk.Label(
            hdr, textvariable=rate_hdr_var, bg="#5B7553", fg="white",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL, Fonts.WEIGHT_BOLD),
            anchor="w", padx=8, pady=3,
        )
        hdr_lbl.pack(fill=tk.X)

        rate_body = ttk.Frame(container, relief="solid", borderwidth=1)
        rate_body.pack(fill=tk.X)

        def _toggle(event=None):
            if rate_body.winfo_ismapped():
                rate_body.pack_forget()
                rate_hdr_var.set("\u25b6  Utility Rates")
            else:
                rate_body.pack(fill=tk.X)
                rate_hdr_var.set("\u25bc  Utility Rates")

        hdr.bind("<Button-1>", _toggle)
        hdr_lbl.bind("<Button-1>", _toggle)

        frm = ttk.Frame(rate_body, padding=(12, 6))
        frm.pack(fill=tk.X)
        frm.columnconfigure(1, weight=1)

        # State selector
        ttk.Label(frm, text="State:").grid(row=0, column=0, sticky="w", padx=4, pady=3)
        states = get_available_states()
        state_cb = ttk.Combobox(
            frm, textvariable=self._state_var, values=states,
            state="readonly", width=6,
        )
        state_cb.grid(row=0, column=1, sticky="w", padx=4, pady=3)
        state_cb.bind("<<ComboboxSelected>>", self._on_state_changed)

        ttk.Label(frm, text="Commercial Rate:").grid(
            row=1, column=0, sticky="w", padx=4, pady=2)
        ttk.Label(
            frm, textvariable=self._rate_kwh_var,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY, Fonts.WEIGHT_BOLD),
        ).grid(row=1, column=1, sticky="w", padx=4, pady=2)

        ttk.Label(frm, text="Demand Charge:").grid(
            row=2, column=0, sticky="w", padx=4, pady=2)
        ttk.Label(
            frm, textvariable=self._rate_dem_var,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY, Fonts.WEIGHT_BOLD),
        ).grid(row=2, column=1, sticky="w", padx=4, pady=2)

        ttk.Separator(frm, orient="horizontal").grid(
            row=3, column=0, columnspan=2, sticky="ew", pady=4)

        ttk.Label(frm, text="Est. Annual Charging Cost:").grid(
            row=4, column=0, sticky="w", padx=4, pady=2)
        ttk.Label(
            frm, textvariable=self._rate_cost_var,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY, Fonts.WEIGHT_BOLD),
            foreground="#2E7D32",
        ).grid(row=4, column=1, sticky="w", padx=4, pady=2)

    # ─────────────────────────────────────────────────────────────────────────
    # Event handlers
    # ─────────────────────────────────────────────────────────────────────────

    def _on_run_clicked(self) -> None:
        if self._on_run_cb is None:
            return
        self._set_status("Running\u2026", color="#E65100")
        try:
            self._on_run_cb()
        except Exception as exc:
            logger.error(f"ChargingPanel run error: {exc}")
            self._set_status(f"Error: {exc}", color="red")

    def _on_state_changed(self, event=None) -> None:
        """Re-compute annual cost when the state dropdown changes."""
        state = self._state_var.get()
        if not state:
            return
        rates = get_rates_for_state(state)
        elec  = rates.get("electricity_price", 0.0)
        dem   = rates.get("demand_charge", 0.0)
        self._rate_kwh_var.set(f"${elec:.3f} / kWh")
        self._rate_dem_var.set(f"${dem:.2f} / kW-month")
        if self._last_result is not None:
            ca = self._last_result
            annual_cost = (
                ca.daily_energy_kwh * 365.0 * elec
                + ca.max_power_required * dem * 12.0
            )
            self._rate_cost_var.set(f"${annual_cost:,.0f} / yr")
        else:
            self._rate_cost_var.set("\u2014")

    # ─────────────────────────────────────────────────────────────────────────
    # Results population
    # ─────────────────────────────────────────────────────────────────────────

    def show_results(self, ca) -> None:
        """
        Expand the results section and populate all widgets.

        Args:
            ca: ChargingAnalysis dataclass returned by analyze_charging_needs()
        """
        self._last_result = ca

        # Expand results section if collapsed
        if not self._results_expanded and self._results_toggle:
            self._results_toggle()

        # KPI chips
        self._kpi_vars["l2"].set(str(ca.level2_chargers_needed))
        self._kpi_vars["dcfc"].set(str(ca.dcfc_chargers_needed))
        self._kpi_vars["kw"].set(f"{ca.max_power_required:.0f}")
        self._kpi_vars["cost"].set(f"${ca.estimated_installation_cost:,.0f}")

        # Status label
        ts = datetime.now().strftime("%I:%M %p")
        self._set_status(f"Analysis complete \u2014 {ts}", color="#2E7D32")

        # Auto-detect state and populate utility rates
        state = _detect_state_from_fleet(self._fleet)
        if state and state in get_available_states():
            self._state_var.set(state)
            self._on_state_changed()

        # Load profile chart
        self._draw_load_chart(ca)

        # Facility breakdown table
        self._populate_facility_table(ca)

        # Recommendations text
        self._generate_recommendations(ca)

    def _draw_load_chart(self, ca) -> None:
        if self._chart_ax is None:
            return
        ax = self._chart_ax
        ax.clear()

        hours = list(range(24))
        load  = ca.hourly_load_kw if ca.hourly_load_kw else [0.0] * 24

        s, e = ca.charging_window
        colors = []
        for h in hours:
            if s < e:
                in_window = s <= h < e
            else:
                in_window = h >= s or h < e
            colors.append(PRIMARY_HEX_1 if in_window else "#BDBDBD")

        ax.bar(hours, load, color=colors, width=0.8,
               edgecolor="white", linewidth=0.4)
        ax.set_xlabel("Hour of Day", fontsize=8)
        ax.set_ylabel("kW Demand", fontsize=8)
        ax.set_xticks(hours)
        ax.set_xticklabels([str(h) for h in hours], fontsize=7)
        ax.tick_params(axis="y", labelsize=7)
        ax.grid(axis="y", alpha=0.3, linewidth=0.5)
        ax.set_xlim(-0.5, 23.5)
        self._chart_fig.tight_layout(pad=0.6)
        self._chart_canvas.draw()

    def _populate_facility_table(self, ca) -> None:
        if not ca.facility_breakdown:
            self._facility_frame.pack_forget()
            return

        # Show the facility frame above the separator before recommendations
        self._facility_frame.pack(fill=tk.X, before=self._rec_outer)

        tree = self._facility_tree
        for iid in tree.get_children():
            tree.delete(iid)
        for row in ca.facility_breakdown:
            tree.insert("", tk.END, values=(
                row["name"],
                row["vehicle_count"],
                row["l2"],
                row["dcfc"],
                f"${row['cost']:,.0f}",
            ))

    def _generate_recommendations(self, ca) -> None:
        bullets: List[str] = []
        bullets.append(
            f"\u2022  Install {ca.level2_chargers_needed} Level 2 charger"
            f"{'s' if ca.level2_chargers_needed != 1 else ''} and "
            f"{ca.dcfc_chargers_needed} DC fast charger"
            f"{'s' if ca.dcfc_chargers_needed != 1 else ''}."
        )

        s, e = ca.charging_window
        bullets.append(
            f"\u2022  Charging window: {s:02d}:00\u2013{e:02d}:00 "
            f"({ca.charging_hours:.0f} available hours / day)."
        )
        bullets.append(
            f"\u2022  Peak electrical demand: {ca.max_power_required:.0f} kW "
            f"\u2014 verify utility service capacity before procurement."
        )

        phasing = ca.recommended_layout.get("phasing", [])
        if len(phasing) >= 2:
            p1 = phasing[0]
            bullets.append(
                f"\u2022  Suggested 2-phase deployment: Phase 1 \u2014 "
                f"{p1['level2_chargers']} L2 + {p1['dcfc_chargers']} DCFC "
                f"(est. ${p1['estimated_cost']:,.0f}); "
                f"Phase 2 \u2014 remainder."
            )

        if ca.facility_breakdown:
            n = len(ca.facility_breakdown)
            bullets.append(
                f"\u2022  {n} facilit{'ies' if n != 1 else 'y'} identified "
                f"\u2014 see breakdown table above for per-facility estimates."
            )

        state = self._state_var.get()
        if state:
            bullets.append(
                f"\u2022  Utility rates loaded for {state}. "
                "Check with your local utility for commercial EV rate programs "
                "that may reduce demand charges."
            )

        text_content = "\n".join(bullets)
        self._reco_text.configure(state=tk.NORMAL)
        self._reco_text.delete("1.0", tk.END)
        self._reco_text.insert(tk.END, text_content)
        self._reco_text.configure(state=tk.DISABLED)

    # ─────────────────────────────────────────────────────────────────────────
    # Collapsible section helper (shared pattern across tabs)
    # ─────────────────────────────────────────────────────────────────────────

    def _collapsible_section(self, parent: tk.Frame, label: str,
                              expanded: bool = True) -> tuple:
        container = ttk.Frame(parent, relief="flat")
        container.pack(fill=tk.X, padx=6, pady=4)

        hdr = tk.Frame(container, bg=PRIMARY_HEX_1, cursor="hand2")
        hdr.pack(fill=tk.X)
        lbl_var = tk.StringVar(value=label)
        lbl_widget = tk.Label(
            hdr, textvariable=lbl_var, bg=PRIMARY_HEX_1, fg="white",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_H3, Fonts.WEIGHT_BOLD),
            anchor="w", padx=8, pady=4,
        )
        lbl_widget.pack(fill=tk.X)

        body = ttk.Frame(container, relief="solid", borderwidth=1)
        if expanded:
            body.pack(fill=tk.X)

        def _toggle(event=None):
            if body.winfo_ismapped():
                body.pack_forget()
                lbl_var.set(label.replace("\u25bc", "\u25b6", 1))
            else:
                body.pack(fill=tk.X)
                lbl_var.set(label.replace("\u25b6", "\u25bc", 1))

        hdr.bind("<Button-1>", _toggle)
        lbl_widget.bind("<Button-1>", _toggle)

        return hdr, body

    # ─────────────────────────────────────────────────────────────────────────
    # Helpers
    # ─────────────────────────────────────────────────────────────────────────

    def _set_status(self, text: str, color: str = "#757575") -> None:
        if self._status_label:
            self._status_label.configure(text=text, foreground=color)

    # ─────────────────────────────────────────────────────────────────────────
    # Public API
    # ─────────────────────────────────────────────────────────────────────────

    def refresh_data(self) -> None:
        """Called by MainWindow when a fleet is loaded or tab is activated."""
        self._fleet = self.sharing_data.get("fleet")

    def get_panel_frame(self) -> ttk.Frame:
        return self.parent_frame
