"""
ui/database_panel.py

Vehicle reference database management panel for the Fleet Electrification Analyzer.

Provides a GUI for browsing, searching, adding, editing, and deleting vehicle MPG
reference entries stored in the SQLite vehicle database (data/vehicle_database.db).

The panel has two sub-tabs:
  - ICE Vehicles: full CRUD for analyst-sourced MPG entries
  - EV Vehicles: read-only placeholder (active in a future phase)

Phase 16 of the Fleet Electrification Analyzer improvement track.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import logging
from typing import Optional, Dict, Any, List

from ui.theme import Colors, Fonts, Spacing
from ui.widgets import ScrollableFrame, SimpleTooltip

logger = logging.getLogger(__name__)

# Fuel type options shown in dropdowns
_FUEL_TYPES = ["", "gasoline", "diesel", "flex", "hybrid", "cng", "propane", "other"]

# Source options shown in dropdowns
_SOURCES = [
    "analyst",
    "manufacturer_spec",
    "fuelly",
    "epa_label",
    "fleet_record",
    "other",
]

# Treeview columns: (internal_key, display_name, width, anchor)
_ICE_COLUMNS = [
    ("year",         "Year",       55,  "center"),
    ("make",         "Make",      100,  "w"),
    ("model",        "Model",     130,  "w"),
    ("fuel_type",    "Fuel",       85,  "w"),
    ("body_class",   "Body",      100,  "w"),
    ("mpg_combined", "MPG Comb",   75,  "e"),
    ("mpg_city",     "City",       65,  "e"),
    ("mpg_highway",  "Hwy",        65,  "e"),
    ("source",       "Source",    130,  "w"),
    ("notes",        "Notes",     200,  "w"),
]


###############################################################################
# DatabasePanel
###############################################################################

class DatabasePanel(ttk.Frame):
    """
    Fifth tab of the main notebook.  Provides CRUD access to the SQLite vehicle
    reference database so analysts can build a team MPG lookup table over time.
    """

    def __init__(self, parent, db_manager, root, status_bar):
        super().__init__(parent)
        self.db_manager  = db_manager
        self.root        = root
        self.status_bar  = status_bar

        # State
        self._selected_id:   Optional[int]       = None   # DB row id of selected row
        self._all_rows:      List[Dict[str, Any]] = []     # Full unfiltered data
        self._id_map:        Dict[str, int]       = {}     # treeview iid -> DB id
        self._sort_col:      str                  = "make"
        self._sort_reverse:  bool                 = False

        # Search / filter vars
        self._search_var    = tk.StringVar()
        self._make_var      = tk.StringVar(value="All Makes")
        self._fuel_var      = tk.StringVar(value="All Fuel Types")

        self._search_var.trace_add("write", lambda *_: self._apply_search())
        self._make_var.trace_add("write",   lambda *_: self._apply_search())
        self._fuel_var.trace_add("write",   lambda *_: self._apply_search())

        self._create_layout()
        self.refresh()

    # -------------------------------------------------------------------------
    # Layout
    # -------------------------------------------------------------------------

    def _create_layout(self) -> None:
        """Build the sub-notebook with ICE and EV tabs."""
        sub_nb = ttk.Notebook(self)
        sub_nb.pack(fill=tk.BOTH, expand=True, padx=Spacing.SM, pady=Spacing.SM)

        # -- ICE Vehicles tab --
        ice_frame = ttk.Frame(sub_nb)
        sub_nb.add(ice_frame, text="ICE Vehicles")
        self._create_ice_tab(ice_frame)

        # -- EV Vehicles tab (placeholder) --
        ev_frame = ttk.Frame(sub_nb)
        sub_nb.add(ev_frame, text="EV Vehicles (Future)")
        self._create_ev_tab(ev_frame)

    def _create_ice_tab(self, parent: ttk.Frame) -> None:
        """Build the ICE vehicles tab: toolbar + treeview + edit pane."""
        # Toolbar
        self._create_toolbar(parent)

        # Main area: treeview left, edit pane right (hidden by default)
        content = ttk.Frame(parent)
        content.pack(fill=tk.BOTH, expand=True)

        # Treeview (left, fills remaining space)
        self._create_treeview(content)

        # Edit pane (right, fixed width, hidden initially)
        self._edit_pane = ttk.LabelFrame(
            content, text="Entry Details",
            padding=(Spacing.SM, Spacing.SM)
        )
        # Do NOT pack yet — revealed by _on_row_select / _show_add_form

        self._create_edit_form(self._edit_pane)

    def _create_toolbar(self, parent: ttk.Frame) -> None:
        """Create search, filter, and action controls."""
        bar = ttk.Frame(parent)
        bar.pack(fill=tk.X, padx=Spacing.SM, pady=(Spacing.SM, Spacing.XS))

        # Search
        ttk.Label(bar, text="Search:").pack(side=tk.LEFT, padx=(0, Spacing.XS))
        search_entry = ttk.Entry(bar, textvariable=self._search_var, width=22)
        search_entry.pack(side=tk.LEFT, padx=(0, Spacing.SM))
        SimpleTooltip(search_entry, "Filter by make, model, or any text field")

        # Make filter
        self._make_combo = ttk.Combobox(
            bar, textvariable=self._make_var,
            values=["All Makes"], width=14, state="readonly"
        )
        self._make_combo.pack(side=tk.LEFT, padx=(0, Spacing.XS))

        # Fuel filter
        ttk.Combobox(
            bar, textvariable=self._fuel_var,
            values=["All Fuel Types"] + _FUEL_TYPES[1:],
            width=13, state="readonly"
        ).pack(side=tk.LEFT, padx=(0, Spacing.MD))

        # Actions (right-aligned)
        ttk.Button(
            bar, text="Import CSV",
            command=self._import_csv
        ).pack(side=tk.RIGHT, padx=(Spacing.XS, 0))

        ttk.Button(
            bar, text="+ Add Entry",
            command=self._show_add_form,
            style="Accent.TButton"
        ).pack(side=tk.RIGHT, padx=(0, Spacing.XS))

        # Row count label
        self._count_label = ttk.Label(
            bar, text="0 entries",
            foreground=Colors.TEXT_SECONDARY
        )
        self._count_label.pack(side=tk.RIGHT, padx=(0, Spacing.MD))

    def _create_treeview(self, parent: ttk.Frame) -> None:
        """Create the sortable, scrollable treeview."""
        tv_frame = ttk.Frame(parent)
        tv_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True,
                      padx=(Spacing.SM, 0), pady=(0, Spacing.SM))

        col_ids = [c[0] for c in _ICE_COLUMNS]
        self._tree = ttk.Treeview(tv_frame, columns=col_ids, show="headings",
                                  selectmode="browse")

        for col_id, display, width, anchor in _ICE_COLUMNS:
            self._tree.heading(
                col_id, text=display,
                command=lambda c=col_id: self._sort_by_column(c)
            )
            self._tree.column(col_id, width=width, anchor=anchor, minwidth=40)

        # Scrollbars
        vsb = ttk.Scrollbar(tv_frame, orient=tk.VERTICAL,   command=self._tree.yview)
        hsb = ttk.Scrollbar(tv_frame, orient=tk.HORIZONTAL, command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT,  fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self._tree.pack(fill=tk.BOTH, expand=True)

        self._tree.bind("<<TreeviewSelect>>", self._on_row_select)
        self._tree.bind("<Double-Button-1>",  self._on_row_select)

    def _create_edit_form(self, parent: ttk.Frame) -> None:
        """Build the entry edit form inside the edit pane."""
        self._form_vars: Dict[str, tk.StringVar] = {}

        def lbl_entry(row: int, label: str, key: str, width: int = 18) -> ttk.Entry:
            ttk.Label(parent, text=label, anchor="w").grid(
                row=row, column=0, sticky="w", pady=2, padx=(0, Spacing.XS)
            )
            var = tk.StringVar()
            self._form_vars[key] = var
            e = ttk.Entry(parent, textvariable=var, width=width)
            e.grid(row=row, column=1, sticky="ew", pady=2)
            return e

        def lbl_combo(row: int, label: str, key: str, values: list,
                      width: int = 16) -> ttk.Combobox:
            ttk.Label(parent, text=label, anchor="w").grid(
                row=row, column=0, sticky="w", pady=2, padx=(0, Spacing.XS)
            )
            var = tk.StringVar()
            self._form_vars[key] = var
            cb = ttk.Combobox(parent, textvariable=var, values=values, width=width)
            cb.grid(row=row, column=1, sticky="ew", pady=2)
            return cb

        parent.columnconfigure(1, weight=1)

        r = 0
        lbl_entry(r, "Year:",         "year",      8);  r += 1
        lbl_entry(r, "Make *:",       "make",      18); r += 1
        lbl_entry(r, "Model *:",      "model",     18); r += 1
        lbl_combo(r, "Fuel Type:",    "fuel_type", _FUEL_TYPES); r += 1
        lbl_entry(r, "Body Class:",   "body_class", 18); r += 1
        lbl_entry(r, "GVWR Min lbs:", "gvwr_lbs_min", 10); r += 1
        lbl_entry(r, "GVWR Max lbs:", "gvwr_lbs_max", 10); r += 1

        # MPG row with three entries side by side
        ttk.Label(parent, text="MPG Comb *:", anchor="w").grid(
            row=r, column=0, sticky="w", pady=2
        )
        mpg_frame = ttk.Frame(parent)
        mpg_frame.grid(row=r, column=1, sticky="ew", pady=2)
        for key, hint in [("mpg_combined", "Comb"), ("mpg_city", "City"), ("mpg_highway", "Hwy")]:
            var = tk.StringVar()
            self._form_vars[key] = var
            ttk.Entry(mpg_frame, textvariable=var, width=6).pack(side=tk.LEFT, padx=(0, 2))
            ttk.Label(mpg_frame, text=hint, foreground=Colors.TEXT_TERTIARY,
                      font=(Fonts.FAMILY_SANS, Fonts.SIZE_TINY)).pack(side=tk.LEFT, padx=(0, 4))
        r += 1

        lbl_combo(r, "Source:", "source", _SOURCES, 16); r += 1

        # Notes (multi-line)
        ttk.Label(parent, text="Notes:", anchor="nw").grid(
            row=r, column=0, sticky="nw", pady=2
        )
        self._notes_text = tk.Text(parent, height=3, width=22, wrap=tk.WORD)
        self._notes_text.grid(row=r, column=1, sticky="ew", pady=2)
        r += 1

        # Button row
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=r, column=0, columnspan=2, pady=(Spacing.MD, 0), sticky="ew")

        self._delete_btn = ttk.Button(
            btn_frame, text="Delete",
            command=self._delete_entry,
            style="Danger.TButton"
        )
        self._delete_btn.pack(side=tk.LEFT)

        ttk.Button(
            btn_frame, text="Cancel",
            command=self._cancel_edit
        ).pack(side=tk.RIGHT, padx=(Spacing.XS, 0))

        ttk.Button(
            btn_frame, text="Save",
            command=self._save_entry,
            style="Primary.TButton"
        ).pack(side=tk.RIGHT)

    def _create_ev_tab(self, parent: ttk.Frame) -> None:
        """EV vehicles tab — placeholder for a future phase."""
        frame = ttk.Frame(parent)
        frame.pack(expand=True)

        ttk.Label(
            frame,
            text="EV Replacement Candidate Database",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_H2, "bold"),
            foreground=Colors.TEXT_PRIMARY,
        ).pack(pady=(40, Spacing.MD))

        ttk.Label(
            frame,
            text=(
                "The EV vehicles table has been created and is ready for data.\n"
                "EV candidate management will be added in a future phase.\n\n"
                "The existing EV matching engine (analysis/ev_database.py)\n"
                "continues to operate in the background."
            ),
            foreground=Colors.TEXT_SECONDARY,
            justify=tk.CENTER,
        ).pack()

    # -------------------------------------------------------------------------
    # Data loading and filtering
    # -------------------------------------------------------------------------

    def refresh(self) -> None:
        """Reload all data from the database and repopulate the treeview."""
        if self.db_manager is None:
            self._all_rows = []
            self._update_make_filter([])
            self._populate_treeview([])
            return

        try:
            self._all_rows = self.db_manager.get_all_ice_vehicles()
        except Exception as e:
            logger.error(f"Failed to load database entries: {e}")
            self._all_rows = []

        self._update_make_filter(self._all_rows)
        self._apply_search()

    def _update_make_filter(self, rows: List[Dict]) -> None:
        """Rebuild the Make filter combobox from current data."""
        makes = sorted({r["make"].title() for r in rows if r.get("make")})
        self._make_combo["values"] = ["All Makes"] + makes
        if self._make_var.get() not in (["All Makes"] + makes):
            self._make_var.set("All Makes")

    def _apply_search(self) -> None:
        """Filter _all_rows client-side and repopulate the treeview."""
        search  = self._search_var.get().strip().lower()
        make_f  = self._make_var.get()
        fuel_f  = self._fuel_var.get()

        def matches(row: Dict) -> bool:
            if make_f != "All Makes":
                if (row.get("make") or "").title() != make_f:
                    return False
            if fuel_f != "All Fuel Types":
                if (row.get("fuel_type") or "").lower() != fuel_f.lower():
                    return False
            if search:
                haystack = " ".join(str(v) for v in row.values() if v).lower()
                if search not in haystack:
                    return False
            return True

        filtered = [r for r in self._all_rows if matches(r)]
        self._populate_treeview(filtered)

    def _populate_treeview(self, rows: List[Dict]) -> None:
        """Clear and refill the treeview from a list of row dicts."""
        self._tree.delete(*self._tree.get_children())
        self._id_map.clear()

        for row in rows:
            values = []
            for col_id, *_ in _ICE_COLUMNS:
                val = row.get(col_id, "")
                if val is None:
                    val = ""
                # Tidy up floats: show as int if whole number
                if col_id in ("mpg_combined", "mpg_city", "mpg_highway") and val:
                    try:
                        fval = float(val)
                        val = int(fval) if fval == int(fval) else round(fval, 1)
                    except (ValueError, TypeError):
                        pass
                values.append(val)

            iid = self._tree.insert("", tk.END, values=values)
            self._id_map[iid] = row["id"]

        total = len(self._all_rows)
        shown = len(rows)
        if total == shown:
            self._count_label.config(text=f"{total} entr{'y' if total == 1 else 'ies'}")
        else:
            self._count_label.config(text=f"{shown} of {total} entries")

    # -------------------------------------------------------------------------
    # Edit pane visibility
    # -------------------------------------------------------------------------

    def _show_edit_pane(self) -> None:
        if not self._edit_pane.winfo_ismapped():
            self._edit_pane.pack(
                side=tk.RIGHT, fill=tk.Y,
                padx=(Spacing.XS, Spacing.SM),
                pady=(0, Spacing.SM)
            )

    def _hide_edit_pane(self) -> None:
        if self._edit_pane.winfo_ismapped():
            self._edit_pane.pack_forget()

    # -------------------------------------------------------------------------
    # Row selection
    # -------------------------------------------------------------------------

    def _on_row_select(self, event=None) -> None:
        """Populate the edit form when a row is selected."""
        sel = self._tree.selection()
        if not sel:
            return

        iid = sel[0]
        db_id = self._id_map.get(iid)
        if db_id is None:
            return

        # Find the row in _all_rows by id
        row = next((r for r in self._all_rows if r["id"] == db_id), None)
        if row is None:
            return

        self._selected_id = db_id
        self._populate_form(row)
        self._delete_btn.config(state="normal")
        self._show_edit_pane()

    def _populate_form(self, row: Dict) -> None:
        """Fill form vars from a database row dict."""
        for key in self._form_vars:
            val = row.get(key)
            self._form_vars[key].set("" if val is None else str(val))

        self._notes_text.delete("1.0", tk.END)
        self._notes_text.insert("1.0", row.get("notes") or "")

    def _clear_form(self) -> None:
        """Reset all form fields to empty."""
        for var in self._form_vars.values():
            var.set("")
        self._form_vars["source"].set("analyst")
        self._notes_text.delete("1.0", tk.END)

    # -------------------------------------------------------------------------
    # CRUD actions
    # -------------------------------------------------------------------------

    def _show_add_form(self) -> None:
        """Switch the edit pane to Add mode (blank form)."""
        self._tree.selection_remove(self._tree.selection())
        self._selected_id = None
        self._clear_form()
        self._delete_btn.config(state="disabled")
        self._show_edit_pane()
        # Focus Year field
        for widget in self._edit_pane.winfo_children():
            if isinstance(widget, ttk.Entry):
                widget.focus_set()
                break

    def _save_entry(self) -> None:
        """Validate and save the current form as an add or update."""
        make  = self._form_vars["make"].get().strip()
        model = self._form_vars["model"].get().strip()
        mpg_c = self._form_vars["mpg_combined"].get().strip()

        if not make:
            messagebox.showerror("Validation Error", "Make is required.", parent=self)
            return
        if not model:
            messagebox.showerror("Validation Error", "Model is required.", parent=self)
            return
        if not mpg_c:
            messagebox.showerror(
                "Validation Error", "MPG Combined is required.", parent=self
            )
            return

        try:
            mpg_combined = float(mpg_c)
            if mpg_combined <= 0:
                raise ValueError("must be > 0")
        except ValueError:
            messagebox.showerror(
                "Validation Error",
                "MPG Combined must be a number greater than zero.",
                parent=self
            )
            return

        # Parse optional numeric fields
        def _opt_int(key: str) -> Optional[int]:
            v = self._form_vars[key].get().strip()
            return int(v) if v else None

        def _opt_float(key: str) -> float:
            v = self._form_vars[key].get().strip()
            try:
                return float(v) if v else 0.0
            except ValueError:
                return 0.0

        year_str = self._form_vars["year"].get().strip()
        try:
            year = int(year_str) if year_str else None
        except ValueError:
            messagebox.showerror("Validation Error", "Year must be a number.", parent=self)
            return

        kwargs = dict(
            year         = year,
            make         = make,
            model        = model,
            fuel_type    = self._form_vars["fuel_type"].get().strip() or None,
            body_class   = self._form_vars["body_class"].get().strip() or None,
            mpg_combined = mpg_combined,
            mpg_city     = _opt_float("mpg_city"),
            mpg_highway  = _opt_float("mpg_highway"),
            notes        = self._notes_text.get("1.0", tk.END).strip(),
            source       = self._form_vars["source"].get().strip() or "analyst",
            gvwr_lbs_min = _opt_int("gvwr_lbs_min"),
            gvwr_lbs_max = _opt_int("gvwr_lbs_max"),
        )

        try:
            if self._selected_id is None:
                # Add new entry
                self.db_manager.add_ice_vehicle(**kwargs)
                msg = f"Added: {make} {model}"
            else:
                # Update existing entry — update_ice_vehicle takes id + kwargs
                update_kwargs = {k: v for k, v in kwargs.items()}
                self.db_manager.update_ice_vehicle(self._selected_id, **update_kwargs)
                msg = f"Updated: {make} {model}"

            self.refresh()
            self._hide_edit_pane()
            self._selected_id = None
            if self.status_bar:
                self.status_bar.set(msg)

        except Exception as e:
            logger.error(f"Failed to save database entry: {e}")
            messagebox.showerror("Save Failed", f"An error occurred:\n{e}", parent=self)

    def _delete_entry(self) -> None:
        """Confirm and delete the selected entry."""
        if self._selected_id is None:
            return

        # Find display name for the confirm dialog
        row = next((r for r in self._all_rows if r["id"] == self._selected_id), None)
        label = f"{row.get('make','')} {row.get('model','')}" if row else "this entry"

        confirmed = messagebox.askyesno(
            "Delete Entry",
            f"Delete the entry for '{label}' from the database?\n\n"
            "This cannot be undone.",
            parent=self,
        )
        if not confirmed:
            return

        try:
            self.db_manager.delete_ice_vehicle(self._selected_id)
            self._selected_id = None
            self._hide_edit_pane()
            self.refresh()
            if self.status_bar:
                self.status_bar.set(f"Deleted: {label}")
        except Exception as e:
            logger.error(f"Failed to delete database entry {self._selected_id}: {e}")
            messagebox.showerror("Delete Failed", f"An error occurred:\n{e}", parent=self)

    def _cancel_edit(self) -> None:
        """Hide the edit pane without saving."""
        self._hide_edit_pane()
        self._selected_id = None
        self._tree.selection_remove(self._tree.selection())

    # -------------------------------------------------------------------------
    # Sorting
    # -------------------------------------------------------------------------

    def _sort_by_column(self, col: str) -> None:
        """Sort the treeview by the clicked column header."""
        if self._sort_col == col:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_col    = col
            self._sort_reverse = False

        # Collect (value, iid) pairs
        items = [
            (self._tree.set(iid, col), iid)
            for iid in self._tree.get_children("")
        ]

        # Numeric sort for MPG / year columns
        numeric_cols = {"year", "mpg_combined", "mpg_city", "mpg_highway",
                        "gvwr_lbs_min", "gvwr_lbs_max"}

        def sort_key(pair):
            val = pair[0]
            if col in numeric_cols:
                try:
                    return float(val) if val else -1.0
                except ValueError:
                    return -1.0
            return val.lower()

        items.sort(key=sort_key, reverse=self._sort_reverse)

        for index, (_, iid) in enumerate(items):
            self._tree.move(iid, "", index)

        # Update heading arrows
        for c_id, *_ in _ICE_COLUMNS:
            arrow = ""
            if c_id == col:
                arrow = " ▲" if not self._sort_reverse else " ▼"
            # Find display name
            display = next(d for cid, d, *_ in _ICE_COLUMNS if cid == c_id)
            self._tree.heading(c_id, text=display + arrow)

    # -------------------------------------------------------------------------
    # Stubs
    # -------------------------------------------------------------------------

    def _import_csv(self) -> None:
        """Bulk-import vehicle MPG entries from a CSV file into the database.

        Accepts flexible column headers via alias mapping.  Required columns:
        make, model, and mpg_combined (or mpg / combined_mpg / combined).
        All other columns are optional.  Rows with blank make/model or a
        non-positive MPG value are silently skipped; any database errors are
        collected and shown in a summary dialog after the import completes.
        """
        if self.db_manager is None:
            messagebox.showerror("Import Error", "Database is not available.", parent=self)
            return

        path = filedialog.askopenfilename(
            title="Import Vehicle MPG CSV",
            filetypes=[
                ("CSV files", "*.csv"),
                ("Text files", "*.txt"),
                ("All files", "*.*"),
            ],
            parent=self,
        )
        if not path:
            return

        import csv as _csv

        # 1. Read the file
        try:
            with open(path, newline="", encoding="utf-8-sig") as f:
                reader = _csv.DictReader(f)
                raw_rows = list(reader)
        except Exception as exc:
            messagebox.showerror(
                "Import Error", f"Could not read file:\n{exc}", parent=self
            )
            return

        if not raw_rows:
            messagebox.showinfo("Import", "File contains no data rows.", parent=self)
            return

        # 2. Build a normalised column map  (canonical field → actual header)
        ALIASES = {
            "make":         ["make"],
            "model":        ["model"],
            "mpg_combined": ["mpg_combined", "mpg", "combined_mpg", "combined"],
            "year":         ["year", "model_year"],
            "fuel_type":    ["fuel_type", "fuel"],
            "body_class":   ["body_class", "body", "body_type"],
            "mpg_city":     ["mpg_city", "city_mpg", "city"],
            "mpg_highway":  ["mpg_highway", "hwy_mpg", "hwy", "highway"],
            "gvwr_lbs_min": ["gvwr_lbs_min", "gvwr_min"],
            "gvwr_lbs_max": ["gvwr_lbs_max", "gvwr_max"],
            "notes":        ["notes"],
            "source":       ["source"],
        }
        header_lower = {h.strip().lower(): h for h in raw_rows[0].keys()}
        col_map: dict = {}
        for field, aliases in ALIASES.items():
            for alias in aliases:
                if alias in header_lower:
                    col_map[field] = header_lower[alias]
                    break

        # Require the three essential columns
        if not all(f in col_map for f in ("make", "model", "mpg_combined")):
            messagebox.showerror(
                "Import Error",
                "CSV must contain columns: make, model, and mpg_combined "
                "(also accepted: mpg / combined_mpg / combined).\n\n"
                f"Headers found: {', '.join(raw_rows[0].keys())}",
                parent=self,
            )
            return

        # 3. Process each row
        added = skipped = 0
        errors: list = []

        def _get(row: dict, field: str, default: str = "") -> str:
            col = col_map.get(field)
            return row.get(col, default).strip() if col else default

        def _opt_float(row: dict, field: str) -> float:
            v = _get(row, field)
            try:
                return float(v) if v else 0.0
            except ValueError:
                return 0.0

        def _opt_int(row: dict, field: str):
            v = _get(row, field)
            try:
                return int(v) if v else None
            except ValueError:
                return None

        for row_num, row in enumerate(raw_rows, start=2):
            make = _get(row, "make")
            model = _get(row, "model")
            mpg_str = _get(row, "mpg_combined")

            if not make or not model:
                skipped += 1
                continue

            try:
                mpg = float(mpg_str)
                if mpg <= 0.0:
                    raise ValueError("MPG must be positive")
            except ValueError:
                skipped += 1
                continue

            year_str = _get(row, "year")
            year = int(year_str) if year_str.isdigit() else None

            try:
                self.db_manager.add_ice_vehicle(
                    make=make,
                    model=model,
                    mpg_combined=mpg,
                    year=year,
                    fuel_type=_get(row, "fuel_type") or None,
                    body_class=_get(row, "body_class") or None,
                    mpg_city=_opt_float(row, "mpg_city"),
                    mpg_highway=_opt_float(row, "mpg_highway"),
                    notes=_get(row, "notes"),
                    source=_get(row, "source") or "analyst",
                    gvwr_lbs_min=_opt_int(row, "gvwr_lbs_min"),
                    gvwr_lbs_max=_opt_int(row, "gvwr_lbs_max"),
                )
                added += 1
            except Exception as exc:
                errors.append(f"Row {row_num}: {exc}")

        # 4. Refresh the table and show a summary
        self.refresh()

        summary = f"Import complete.\n\nAdded: {added}\nSkipped (blank/invalid): {skipped}"
        if errors:
            summary += f"\nErrors: {len(errors)}\n" + "\n".join(errors[:5])
            if len(errors) > 5:
                summary += f"\n… and {len(errors) - 5} more"
        messagebox.showinfo("Import Complete", summary, parent=self)
