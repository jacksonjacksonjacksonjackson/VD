"""
timeline_panel.py

Electrification Timeline editor tab (Phase 19).

Provides:
  • Full-fleet Treeview with all processed vehicles
  • Double-click inline editing of the "Proposed EV Year" cell
  • Manual overrides written back to FleetVehicle.custom_fields
  • "Reset Selected" / "Reset All" buttons
  • Live Gantt chart (Grouped by ACF default; toggle to Per Vehicle)
"""

import logging
import warnings
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable, Dict, List, Optional

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from data.models import FleetVehicle, Fleet
from ui.theme import Colors, Fonts, Spacing
from utils import SimpleTooltip, ScrollableFrame
from ui.analysis_panel import (
    ACF_COLORS, ACF_LABELS, GANTT_YEAR_MIN, GANTT_YEAR_MAX,
    _draw_gantt_chart, _show_acf_ev_year_dialog,
    AnalysisPanel,
)

logger = logging.getLogger(__name__)

# Columns shown in the timeline editor table
_COLS = [
    ("asset_id",  "Asset ID",            70),
    ("yr",        "Year",                45),
    ("make",      "Make",                70),
    ("model",     "Model",               90),
    ("acf",       "ACF",                 38),
    ("wt_class",  "Wt Class",            70),
    ("body_type", "Body Type",           85),
    ("ev_year",   "Proposed EV Year ✎", 115),   # ✎ signals double-click to edit
    ("sys_rec",   "System Rec.",          80),
    ("override",  "Override?",            68),
]

# Lookup table of base heading text (without sort arrows added at runtime)
_HEADING_BASE: dict = {col_id: heading for col_id, heading, _ in _COLS}

_EV_YEAR_COL = "ev_year"   # column id for the editable cell
_EV_YEAR_COL_NUM = 8       # 1-based column number ("#8") in the Treeview

_ACF_COL = "acf"           # column id for the ACF category cell
_ACF_COL_NUM = 5           # 1-based column number ("#5") in the Treeview

# Combobox options for ACF category inline edit: "CODE — Plain Label"
# Dropdown options for the inline ACF editor — match ACF Category stored values exactly
_ACF_OPTIONS = [
    "Zero-Emission Vehicle",
    "Exempt — Light-Duty",
    "Subject to ACF",
    "Exempt — Body Type",
    "Emergency Vehicle",
]
_ACF_OPTION_TO_CODE = {
    "Zero-Emission Vehicle": "ZEV",
    "Exempt — Light-Duty":   "A",
    "Subject to ACF":        "B",
    "Exempt — Body Type":    "C",
    "Emergency Vehicle":     "D",
}

# Short labels used in the filter bar checkbuttons
_ACF_FILTER_LABELS = {
    "ZEV": "Already ZEV",
    "A":   "Light-Duty (Exempt)",
    "B":   "Mandate-Subject",
    "C":   "Body-Type Exempt",
    "D":   "Emergency Vehicle",
}


class TimelinePanel(ttk.Frame):
    """
    Electrification Timeline editor — full-fleet table + live Gantt chart.

    Public API
    ----------
    set_fleet(fleet, on_year_changed=None)
        Called by MainWindow whenever the fleet is loaded/updated.
    notify_year_changed()
        Called externally to force a Gantt redraw (e.g., from AnalysisPanel).
    copy_selection()
        Copy selected rows as tab-delimited text (Edit > Copy handler).
    """

    def __init__(self, parent, fleet: Optional[Fleet] = None,
                 on_year_changed: Optional[Callable] = None,
                 on_acf_changed: Optional[Callable] = None):
        super().__init__(parent)
        self.fleet = fleet
        self.on_year_changed: Optional[Callable] = on_year_changed
        self.on_acf_changed: Optional[Callable] = on_acf_changed

        # {iid: FleetVehicle} — populated when the table is filled
        self._iid_to_vehicle: Dict[str, FleetVehicle] = {}
        # Insertion-ordered list of all iids (used by filter to detach/reattach)
        self._all_iids: List[str] = []
        self._edit_entry: Optional[tk.Entry] = None  # inline edit widget

        self._tree: Optional[ttk.Treeview] = None
        self._override_count_label: Optional[ttk.Label] = None
        self._reset_sel_btn: Optional[ttk.Button] = None
        self._gantt_fig: Optional[Figure] = None
        self._gantt_canvas: Optional[FigureCanvasTkAgg] = None
        self._gantt_view_var: Optional[tk.StringVar] = None
        self._gantt_max_var: Optional[tk.StringVar] = None
        self._sort_state: Dict[str, bool] = {}  # col_id -> ascending

        # ── Filter state vars (populated in _create_filter_bar) ───────────────
        self._filter_search_var: Optional[tk.StringVar] = None
        self._filter_acf_vars: Dict[str, tk.BooleanVar] = {}
        self._filter_year_from_var: Optional[tk.StringVar] = None
        self._filter_year_to_var: Optional[tk.StringVar] = None

        self._create_ui()
        if self.fleet:
            self._populate_table()

    # =========================================================================
    # UI construction
    # =========================================================================

    def _create_ui(self):
        """Build: toolbar → filter bar → table → Gantt section."""
        self._create_toolbar()
        self._create_filter_bar()
        self._create_table()
        self._create_gantt_section()

    def _create_toolbar(self):
        bar = ttk.Frame(self)
        bar.pack(fill=tk.X, padx=Spacing.SM, pady=(Spacing.SM, Spacing.XS))

        ttk.Label(
            bar,
            text="Electrification Timeline Editor",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_H3, Fonts.WEIGHT_BOLD),
            foreground=Colors.TEXT_PRIMARY,
        ).pack(side=tk.LEFT, padx=(0, Spacing.LG))

        self._reset_sel_btn = ttk.Button(
            bar, text="Reset Selected",
            command=self._reset_selected,
            style="Secondary.TButton",
            state="disabled",  # enabled when rows are selected
        )
        self._reset_sel_btn.pack(side=tk.LEFT, padx=(0, Spacing.XS))
        SimpleTooltip(self._reset_sel_btn,
                      "Restore the system-recommended EV year for selected vehicles.")

        reset_all_btn = ttk.Button(
            bar, text="Reset All",
            command=self._reset_all,
            style="Secondary.TButton",
        )
        reset_all_btn.pack(side=tk.LEFT, padx=(0, Spacing.MD))
        SimpleTooltip(reset_all_btn, "Remove ALL manual EV-year overrides.")

        self._override_count_label = ttk.Label(
            bar, text="",
            foreground=Colors.TEXT_SECONDARY,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL),
        )
        self._override_count_label.pack(side=tk.LEFT)

        ttk.Label(
            bar,
            text="Double-click a vehicle's 'Proposed EV Year' to edit.",
            foreground=Colors.TEXT_TERTIARY,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL, "italic"),
        ).pack(side=tk.RIGHT)

    def _create_filter_bar(self):
        """Filter bar: text search | ACF checkboxes | EV year range | Clear."""
        bar = ttk.Frame(self, relief="groove", borderwidth=1)
        bar.pack(fill=tk.X, padx=Spacing.SM, pady=(0, Spacing.XS))

        inner = ttk.Frame(bar)
        inner.pack(fill=tk.X, padx=Spacing.SM, pady=Spacing.XS)

        # ── Text search ───────────────────────────────────────────────────────
        ttk.Label(inner, text="Search:").pack(side=tk.LEFT, padx=(0, 3))
        self._filter_search_var = tk.StringVar()
        search_entry = ttk.Entry(inner, textvariable=self._filter_search_var, width=20)
        search_entry.pack(side=tk.LEFT, padx=(0, Spacing.MD))
        SimpleTooltip(search_entry, "Filter by asset ID, make, or model (updates live)")
        self._filter_search_var.trace_add("write", lambda *_: self._apply_filter())

        # ── ACF category checkboxes ───────────────────────────────────────────
        ttk.Label(inner, text="ACF:").pack(side=tk.LEFT, padx=(0, 3))
        for cat in ("ZEV", "A", "B", "C", "D"):
            var = tk.BooleanVar(value=True)
            self._filter_acf_vars[cat] = var
            cb = ttk.Checkbutton(
                inner, text=_ACF_FILTER_LABELS[cat], variable=var,
                command=self._apply_filter,
            )
            cb.pack(side=tk.LEFT, padx=(0, 6))
            SimpleTooltip(cb, f"Show/hide {_ACF_FILTER_LABELS[cat]} vehicles")

        # ── EV year range ─────────────────────────────────────────────────────
        ttk.Label(inner, text="  EV Year:").pack(side=tk.LEFT, padx=(Spacing.SM, 3))
        self._filter_year_from_var = tk.StringVar(value="")
        ttk.Spinbox(
            inner, from_=2026, to=2060,
            textvariable=self._filter_year_from_var,
            width=5, command=self._apply_filter,
        ).pack(side=tk.LEFT, padx=(0, 2))
        self._filter_year_from_var.trace_add("write", lambda *_: self._apply_filter())

        ttk.Label(inner, text="–").pack(side=tk.LEFT, padx=2)
        self._filter_year_to_var = tk.StringVar(value="")
        ttk.Spinbox(
            inner, from_=2026, to=2060,
            textvariable=self._filter_year_to_var,
            width=5, command=self._apply_filter,
        ).pack(side=tk.LEFT, padx=(0, Spacing.MD))
        self._filter_year_to_var.trace_add("write", lambda *_: self._apply_filter())

        # ── Clear button ──────────────────────────────────────────────────────
        ttk.Button(
            inner, text="Clear Filters",
            command=self._clear_filter,
            style="Secondary.TButton",
        ).pack(side=tk.LEFT)

    def _apply_filter(self, *_args):
        """Show/hide rows based on current filter state (detach / reattach)."""
        if self._tree is None or not self._all_iids:
            return

        search = (self._filter_search_var.get().strip().lower()
                  if self._filter_search_var else "")
        acf_shown = {cat for cat, var in self._filter_acf_vars.items() if var.get()}

        try:
            year_from = int(self._filter_year_from_var.get())
        except (ValueError, TypeError, AttributeError):
            year_from = 0
        try:
            year_to = int(self._filter_year_to_var.get())
        except (ValueError, TypeError, AttributeError):
            year_to = 9999

        has_year_filter = (year_from > 0 or year_to < 9999)

        # Detach all first, then reattach matching rows in insertion order
        for iid in self._all_iids:
            try:
                self._tree.detach(iid)
            except tk.TclError:
                pass

        for iid in self._all_iids:
            v = self._iid_to_vehicle.get(iid)
            if v is None:
                continue

            # ACF filter — use the letter-code field (_acf_code) so it matches
            # the checkbox keys {"ZEV","A","B","C","D"} regardless of how the
            # full label string was written by the processor.
            acf = v.custom_fields.get("_acf_code", "")
            if acf not in acf_shown:
                continue

            # EV year filter
            if has_year_filter:
                yr_raw = v.custom_fields.get("Proposed EV Year", "")
                try:
                    yr = int(yr_raw)
                    if not (year_from <= yr <= year_to):
                        continue
                except (ValueError, TypeError):
                    continue  # skip unscheduled when year filter is active

            # Text search (asset ID / make / model)
            if search:
                searchable = " ".join([
                    str(v.asset_id or ""),
                    str(v.vehicle_id.make or ""),
                    str(v.vehicle_id.model or ""),
                ]).lower()
                if search not in searchable:
                    continue

            try:
                self._tree.reattach(iid, "", "end")
            except tk.TclError:
                pass

    def _clear_filter(self):
        """Reset all filter controls and show all rows."""
        if self._filter_search_var:
            self._filter_search_var.set("")
        for var in self._filter_acf_vars.values():
            var.set(True)
        if self._filter_year_from_var:
            self._filter_year_from_var.set("")
        if self._filter_year_to_var:
            self._filter_year_to_var.set("")
        # Reattach everything
        for iid in self._all_iids:
            try:
                self._tree.reattach(iid, "", "end")
            except tk.TclError:
                pass

    def _create_table(self):
        table_frame = ttk.LabelFrame(self, text="Fleet Vehicles")
        table_frame.pack(fill=tk.BOTH, expand=True,
                         padx=Spacing.SM, pady=(0, Spacing.XS))

        # Columns that should expand to fill available horizontal space
        _STRETCH_COLS = {"make", "model", "body_type"}

        col_ids = [c[0] for c in _COLS]
        self._tree = ttk.Treeview(
            table_frame, columns=col_ids, show="headings", height=14
        )
        for col_id, heading, width in _COLS:
            self._tree.heading(
                col_id, text=heading, anchor="w",
                command=lambda c=col_id: self._sort_by_column(c),
            )
            self._tree.column(col_id, width=width, minwidth=width,
                              anchor="w", stretch=(col_id in _STRETCH_COLS))

        vsb = ttk.Scrollbar(table_frame, orient="vertical",
                             command=self._tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal",
                             command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # EV year override → amber tint
        self._tree.tag_configure("overridden", background="#FFF3E0")
        # ACF category override → light blue tint
        self._tree.tag_configure("acf_overridden", background="#E3F2FD")
        # Both overrides active → light purple tint
        self._tree.tag_configure("both_overridden", background="#F3E8FD")
        # Placeholder tag (no visual change, just marks non-editable rows)
        self._tree.tag_configure("normal")

        self._tree.bind("<Double-Button-1>", self._on_double_click)
        self._tree.bind("<Button-1>", self._cancel_edit)
        self._tree.bind("<<TreeviewSelect>>", self._on_selection_changed)

    def _create_gantt_section(self):
        gantt_frame = ttk.LabelFrame(self, text="Timeline Gantt Chart")
        gantt_frame.pack(fill=tk.BOTH, padx=Spacing.SM, pady=(0, Spacing.SM))

        ctrl = ttk.Frame(gantt_frame)
        ctrl.pack(fill=tk.X, padx=Spacing.SM, pady=(Spacing.XS, 0))

        ttk.Label(ctrl, text="View:").pack(side=tk.LEFT, padx=(0, 4))
        self._gantt_view_var = tk.StringVar(value="Grouped by ACF")
        view_combo = ttk.Combobox(
            ctrl,
            textvariable=self._gantt_view_var,
            values=["Grouped by ACF", "Per Vehicle"],
            state="readonly",
            width=16,
        )
        view_combo.pack(side=tk.LEFT)
        view_combo.bind("<<ComboboxSelected>>", lambda _e: self._update_gantt())

        ttk.Label(ctrl, text="  Max rows:").pack(side=tk.LEFT, padx=(Spacing.MD, 4))
        self._gantt_max_var = tk.StringVar(value="50")
        max_combo = ttk.Combobox(
            ctrl,
            textvariable=self._gantt_max_var,
            values=["25", "50", "100", "All"],
            state="readonly",
            width=5,
        )
        max_combo.pack(side=tk.LEFT)
        max_combo.bind("<<ComboboxSelected>>", lambda _e: self._update_gantt())
        SimpleTooltip(max_combo,
                      "Maximum rows shown in Per Vehicle view.\n"
                      "Vehicles are sorted earliest EV year first.")

        self._gantt_fig = Figure(figsize=(10, 3.2), dpi=80)
        self._gantt_fig.patch.set_facecolor(Colors.SURFACE)
        canvas = FigureCanvasTkAgg(self._gantt_fig, master=gantt_frame)
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True,
                                    padx=Spacing.SM, pady=(Spacing.XS, Spacing.SM))
        self._gantt_canvas = canvas
        self._update_gantt()

    # =========================================================================
    # Table population
    # =========================================================================

    def _populate_table(self):
        """Clear and refill the Treeview from self.fleet.vehicles."""
        if self._tree is None:
            return
        # Delete all rows and reset tracking lists
        for iid in list(self._iid_to_vehicle.keys()):
            try:
                self._tree.delete(iid)
            except tk.TclError:
                pass
        self._iid_to_vehicle.clear()
        self._all_iids.clear()

        if not self.fleet or not self.fleet.vehicles:
            self._update_override_label()
            return

        for v in self.fleet.vehicles:
            asset_id   = v.asset_id or v.vin[:8]
            yr         = v.vehicle_id.year or "—"
            make       = v.vehicle_id.make  or "—"
            model      = v.vehicle_id.model or "—"
            acf        = v.custom_fields.get("ACF Category", "—")
            wt_class   = v.vehicle_id.vehicle_class or "—"
            body_type  = v.vehicle_id.body_class    or "—"
            ev_year    = v.custom_fields.get("Proposed EV Year", "—")
            sys_rec    = v.custom_fields.get("System Recommended EV Year", ev_year)
            is_ev_ovr  = v.custom_fields.get("EV Year Overridden", "") == "Yes"
            is_acf_ovr = v.custom_fields.get("ACF Category Overridden", "") == "Yes"
            ovr_parts  = []
            if is_ev_ovr:
                ovr_parts.append("EV Year")
            if is_acf_ovr:
                ovr_parts.append("ACF Cat.")
            ovr_str = " & ".join(ovr_parts) if ovr_parts else ""

            if is_ev_ovr and is_acf_ovr:
                tag = "both_overridden"
            elif is_ev_ovr:
                tag = "overridden"
            elif is_acf_ovr:
                tag = "acf_overridden"
            else:
                tag = "normal"

            iid = self._tree.insert("", "end", tags=(tag,), values=(
                asset_id, yr, make, model, acf, wt_class, body_type,
                ev_year, sys_rec, ovr_str,
            ))
            self._iid_to_vehicle[iid] = v
            self._all_iids.append(iid)

        self._update_override_label()
        # Re-apply any active filter so newly loaded data is filtered correctly
        self._apply_filter()

    def _update_row(self, iid: str, vehicle: FleetVehicle):
        """Refresh a single Treeview row after any override change."""
        acf        = vehicle.custom_fields.get("ACF Category", "—")
        ev_year    = vehicle.custom_fields.get("Proposed EV Year", "—")
        sys_rec    = vehicle.custom_fields.get("System Recommended EV Year", ev_year)
        is_ev_ovr  = vehicle.custom_fields.get("EV Year Overridden", "") == "Yes"
        is_acf_ovr = vehicle.custom_fields.get("ACF Category Overridden", "") == "Yes"
        ovr_parts  = []
        if is_ev_ovr:
            ovr_parts.append("EV Year")
        if is_acf_ovr:
            ovr_parts.append("ACF Cat.")
        ovr_str = " & ".join(ovr_parts) if ovr_parts else ""

        if is_ev_ovr and is_acf_ovr:
            tag = "both_overridden"
        elif is_ev_ovr:
            tag = "overridden"
        elif is_acf_ovr:
            tag = "acf_overridden"
        else:
            tag = "normal"

        self._tree.item(iid, tags=(tag,))
        self._tree.set(iid, _ACF_COL,   acf)
        self._tree.set(iid, "ev_year",  ev_year)
        self._tree.set(iid, "sys_rec",  sys_rec)
        self._tree.set(iid, "override", ovr_str)
        self._update_override_label()

    def _update_override_label(self):
        if self._override_count_label is None or not self.fleet:
            return

        ev_count = earlier = later = 0
        acf_count = 0
        for v in self.fleet.vehicles:
            if v.custom_fields.get("EV Year Overridden") == "Yes":
                ev_count += 1
                try:
                    proposed = int(v.custom_fields.get("Proposed EV Year", 0))
                    system   = int(v.custom_fields.get("System Recommended EV Year", 0))
                    if proposed < system:
                        earlier += 1
                    elif proposed > system:
                        later += 1
                except (ValueError, TypeError):
                    pass
            if v.custom_fields.get("ACF Category Overridden") == "Yes":
                acf_count += 1

        parts = []
        if ev_count:
            ev_detail = ""
            if earlier or later:
                sub = []
                if earlier:
                    sub.append(f"{earlier} earlier")
                if later:
                    sub.append(f"{later} later")
                ev_detail = " (" + ", ".join(sub) + ")"
            parts.append(
                f"{ev_count} EV year override{'s' if ev_count != 1 else ''}{ev_detail}"
            )
        if acf_count:
            parts.append(
                f"{acf_count} ACF category override{'s' if acf_count != 1 else ''}"
            )

        text = "  |  ".join(parts) if parts else ""
        self._override_count_label.config(text=text)

    # =========================================================================
    # Inline edit (double-click on Proposed EV Year cell)
    # =========================================================================

    def _on_double_click(self, event: tk.Event):
        """Spawn an editor widget over the clicked cell.

        Handles two editable columns:
          • EV Year (col #8) → plain Entry (free-text numeric year)
          • ACF Category (col #5) → read-only Combobox (dropdown)
        """
        if self._tree is None:
            return
        region = self._tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self._tree.identify_column(event.x)
        row_iid = self._tree.identify_row(event.y)
        if not row_iid:
            return
        vehicle = self._iid_to_vehicle.get(row_iid)
        if vehicle is None:
            return

        if col == f"#{_EV_YEAR_COL_NUM}":
            self._open_year_editor(row_iid, vehicle, col)
        elif col == f"#{_ACF_COL_NUM}":
            self._open_acf_editor(row_iid, vehicle, col)

    def _open_year_editor(self, row_iid: str, vehicle: FleetVehicle, col: str):
        """Open an inline Entry for the EV Year column."""
        bbox = self._tree.bbox(row_iid, col)
        if not bbox:
            return
        x, y, w, h = bbox

        self._cancel_edit()

        current = self._tree.set(row_iid, _EV_YEAR_COL)
        entry = ttk.Entry(self._tree, width=8)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, current)
        entry.select_range(0, tk.END)
        entry.focus_set()
        self._edit_entry = entry

        def commit(_e=None):
            raw = entry.get().strip()
            self._cancel_edit()
            self._commit_year(row_iid, vehicle, raw)

        def cancel(_e=None):
            self._cancel_edit()

        entry.bind("<Return>",   commit)
        entry.bind("<KP_Enter>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>",   cancel)

    def _open_acf_editor(self, row_iid: str, vehicle: FleetVehicle, col: str):
        """Open an inline Combobox for the ACF Category column."""
        bbox = self._tree.bbox(row_iid, col)
        if not bbox:
            return
        x, y, w, h = bbox

        self._cancel_edit()

        current_label = self._tree.set(row_iid, _ACF_COL)
        current_option = current_label if current_label in _ACF_OPTIONS else _ACF_OPTIONS[0]

        combo = ttk.Combobox(
            self._tree,
            values=_ACF_OPTIONS,
            state="readonly",
            width=22,
        )
        combo.place(x=x, y=y, width=max(w, 190), height=h)
        combo.set(current_option)
        combo.focus_set()
        combo.event_generate("<Button-1>")  # open dropdown immediately
        self._edit_entry = combo  # reuse the cancel mechanism

        def commit(_e=None):
            selected = combo.get()
            self._cancel_edit()
            if selected:
                self._commit_acf(row_iid, vehicle, selected)

        def cancel(_e=None):
            self._cancel_edit()

        combo.bind("<<ComboboxSelected>>", commit)
        combo.bind("<Return>",   commit)
        combo.bind("<KP_Enter>", commit)
        combo.bind("<Escape>",   cancel)
        combo.bind("<FocusOut>", lambda e: self.after(50, lambda: self._cancel_edit()))

    def _cancel_edit(self, _e=None):
        if self._edit_entry is not None:
            try:
                self._edit_entry.destroy()
            except tk.TclError:
                pass
            self._edit_entry = None

    def _commit_year(self, row_iid: str, vehicle: FleetVehicle, raw: str):
        """Validate the entered year and apply the override."""
        _MAX_YEAR = 2060
        try:
            year = int(raw)
        except (ValueError, TypeError):
            if raw.strip():  # warn only if the user actually typed something
                messagebox.showwarning(
                    "Invalid Year",
                    f"'{raw}' is not a valid year. Please enter a whole number "
                    f"between {GANTT_YEAR_MIN} and {_MAX_YEAR} (e.g. 2032).",
                )
            return
        if not (GANTT_YEAR_MIN <= year <= _MAX_YEAR):
            messagebox.showwarning(
                "Invalid Year",
                f"Please enter a year between {GANTT_YEAR_MIN} and {_MAX_YEAR}.\n"
                f"Years beyond the Gantt window ({GANTT_YEAR_MAX}) will extend "
                f"the chart automatically.",
            )
            return
        AnalysisPanel._apply_ev_year_override(vehicle, str(year))
        self._update_row(row_iid, vehicle)
        self._update_gantt()
        if self.on_year_changed:
            self.on_year_changed()

    def _commit_acf(self, row_iid: str, vehicle: FleetVehicle, selected: str):
        """Validate and apply an ACF category override after dropdown selection.

        ``selected`` is a full label string like "Subject to ACF".
        """
        new_code = _ACF_OPTION_TO_CODE.get(selected)
        if not new_code:
            logger.warning(f"Unrecognised ACF option selected: {selected!r}")
            return

        old_code = vehicle.custom_fields.get("_acf_code", "")
        if new_code == old_code:
            return  # No change — nothing to do

        current_ev_year = vehicle.custom_fields.get("Proposed EV Year", "N/A")
        choice = _show_acf_ev_year_dialog(self, old_code, new_code, current_ev_year)

        if choice == "cancel":
            return

        # Apply the ACF override
        AnalysisPanel._apply_acf_override(vehicle, new_code)

        if choice == "recalculate":
            # Clear any existing EV year override markers and re-run assignment
            vehicle.custom_fields.pop("EV Year Overridden", None)
            vehicle.custom_fields.pop("System Recommended EV Year", None)
            from analysis.electrification_timeline import assign_electrification_years
            ft = self.fleet.fleet_type if self.fleet else "hpf"
            assign_electrification_years([vehicle], fleet_type=ft)

        self._update_row(row_iid, vehicle)
        self._update_gantt()
        if self.on_year_changed:
            self.on_year_changed()
        if self.on_acf_changed:
            self.on_acf_changed()

    # =========================================================================
    # Reset helpers
    # =========================================================================

    def _reset_selected(self):
        selected = self._tree.selection() if self._tree else []
        if not selected:
            return  # button is disabled when nothing selected; guard is defensive
        acf_changed = False
        for iid in selected:
            v = self._iid_to_vehicle.get(iid)
            if v is None:
                continue
            if v.custom_fields.get("EV Year Overridden") == "Yes":
                AnalysisPanel._reset_ev_year_override(v)
            if v.custom_fields.get("ACF Category Overridden") == "Yes":
                AnalysisPanel._reset_acf_override(v)
                acf_changed = True
            self._update_row(iid, v)
        self._update_gantt()
        if self.on_year_changed:
            self.on_year_changed()
        if acf_changed and self.on_acf_changed:
            self.on_acf_changed()

    def _reset_all(self):
        if not self.fleet or not self.fleet.vehicles:
            return
        acf_changed = False
        for iid, v in self._iid_to_vehicle.items():
            if v.custom_fields.get("EV Year Overridden") == "Yes":
                AnalysisPanel._reset_ev_year_override(v)
            if v.custom_fields.get("ACF Category Overridden") == "Yes":
                AnalysisPanel._reset_acf_override(v)
                acf_changed = True
            self._update_row(iid, v)
        self._update_gantt()
        if self.on_year_changed:
            self.on_year_changed()
        if acf_changed and self.on_acf_changed:
            self.on_acf_changed()

    # =========================================================================
    # Gantt chart
    # =========================================================================

    def _update_gantt(self):
        """Redraw the Gantt chart from current fleet data."""
        if self._gantt_fig is None or self._gantt_canvas is None:
            return
        self._gantt_fig.clear()
        ax = self._gantt_fig.add_subplot(111)
        vehicles = self.fleet.vehicles if self.fleet else []
        view = self._gantt_view_var.get() if self._gantt_view_var else "Grouped by ACF"

        max_v = 0
        if self._gantt_max_var:
            max_str = self._gantt_max_var.get()
            if max_str != "All":
                try:
                    max_v = int(max_str)
                except ValueError:
                    max_v = 0

        _draw_gantt_chart(ax, vehicles, view, max_vehicles=max_v)
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            self._gantt_fig.tight_layout(pad=0.5)
        self._gantt_canvas.draw()

    # =========================================================================
    # Sort & selection helpers
    # =========================================================================

    def _sort_by_column(self, col_id: str):
        """Sort the Treeview by *col_id*, toggling ascending/descending."""
        if self._tree is None:
            return
        ascending = not self._sort_state.get(col_id, True)
        self._sort_state[col_id] = ascending

        items = [(self._tree.set(iid, col_id), iid)
                 for iid in self._tree.get_children("")]

        def sort_key(item):
            val = item[0]
            try:
                return (0, float(str(val).replace(",", "").replace("$", "")))
            except (ValueError, AttributeError):
                return (1, str(val).lower())

        items.sort(key=sort_key, reverse=not ascending)

        for idx, (_, iid) in enumerate(items):
            self._tree.move(iid, "", idx)

        # Update heading arrows; preserve base text (including ✎ indicator)
        arrow = " ▲" if ascending else " ▼"
        for cid, _, _ in _COLS:
            base = _HEADING_BASE[cid]
            self._tree.heading(cid, text=(base + arrow) if cid == col_id else base)

    def _on_selection_changed(self, event=None):
        """Enable/disable Reset Selected based on whether rows are selected."""
        if self._reset_sel_btn is None or self._tree is None:
            return
        selected = self._tree.selection()
        self._reset_sel_btn.configure(state="normal" if selected else "disabled")

    # =========================================================================
    # Public API
    # =========================================================================

    def set_fleet(self, fleet: Optional[Fleet],
                  on_year_changed: Optional[Callable] = None,
                  on_acf_changed: Optional[Callable] = None):
        """Update the panel with a new fleet object."""
        self.fleet = fleet
        if on_year_changed is not None:
            self.on_year_changed = on_year_changed
        if on_acf_changed is not None:
            self.on_acf_changed = on_acf_changed
        self._populate_table()
        self._update_gantt()

    def notify_year_changed(self):
        """Called by AnalysisPanel if it triggers a year change externally."""
        self._populate_table()
        self._update_gantt()

    def copy_selection(self):
        """Copy selected Timeline rows as tab-delimited text (Edit > Copy handler)."""
        if self._tree is None:
            return
        headers = [h for _, h, _ in _COLS]
        rows = ["\t".join(headers)]
        selected = self._tree.selection()
        target_iids = selected if selected else self._tree.get_children()
        for iid in target_iids:
            values = self._tree.item(iid, "values")
            rows.append("\t".join(str(v) for v in values))
        if len(rows) > 1:
            try:
                self.clipboard_clear()
                self.clipboard_append("\n".join(rows))
            except tk.TclError:
                pass
