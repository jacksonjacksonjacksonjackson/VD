"""
present_panel.py

Present tab (Tab 4) — client-facing PowerPoint export.
Provides:
  • Client/consultant profile form (saved as per-fleet sidecar JSON)
  • Draggable, toggleable slide list (core + optional slides)
  • Consulting content editor (agenda, data needs, next steps)
  • Template browser
  • Build Presentation button  →  export_presentation()
  • Export as PDF button       →  export_pdf()
"""

import os
import platform
import subprocess
import logging
import threading
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, List, Any, Optional

try:
    from PIL import Image as _PILImage, ImageTk as _ImageTk
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False

from settings import (
    PRIMARY_HEX_1, PRIMARY_HEX_2, PRIMARY_HEX_3,
    SECONDARY_HEX_1, SECONDARY_HEX_2,
    TEMPLATE_SLIDE_IDS, DEFAULT_SLIDE_IDS, DEFAULT_TEMPLATE_PATH,
    EXPORT_DIR, ASSETS_DIR,
)
from utils import SimpleTooltip, ScrollableFrame
from data.models import PresentationProfile
from powerpoint_export import export_presentation, export_pdf
from ui.theme import Colors, Fonts, Spacing

logger = logging.getLogger(__name__)

###############################################################################
# Slide metadata (display names + types for all supported slides)
###############################################################################

_TEMPLATE_SLIDE_META = {
    "cover":           {"name": "Cover Slide",                    "type": "Token"},
    "agenda":          {"name": "Agenda",                         "type": "Static"},
    "carb_overview":   {"name": "CARB / ACF Overview",            "type": "Static"},
    "acf_scenarios":   {"name": "ACF Compliance Scenarios",       "type": "Static"},
    "acf_exemptions":  {"name": "ACF Exemptions",                 "type": "Static"},
    "key_findings":    {"name": "Key Findings",                   "type": "Data"},
    "timeline_chart":  {"name": "Electrification Timeline Chart", "type": "Chart"},
    "emissions_chart": {"name": "GHG Emissions Reduction",        "type": "Chart"},
    "incentives":      {"name": "Incentives & Other Support",     "type": "Static"},
    "data_needs":      {"name": "Data Needs",                     "type": "Editable"},
    "next_steps":      {"name": "Next Steps",                     "type": "Editable"},
    "contact":         {"name": "Contact Information",            "type": "Token"},
    "appendix":        {"name": "Appendix (section break)",       "type": "Static"},
    "infra_costs_chart": {"name": "Charging Infrastructure Costs", "type": "Chart"},
    "tco_chart":       {"name": "Annual Marginal EV TCO",         "type": "Chart"},
}

_OPTIONAL_SLIDE_META = {
    # ── Phase 24: new slides — pre-checked by default ─────────────────────────
    "acf_composition":    {"name": "Fleet Composition by ACF Category",       "type": "Chart",    "default": True},
    "timeline_moderate":  {"name": "Electrification Timeline — Moderate 2035","type": "Chart",    "default": True},
    "timeline_current_plan": {"name": "Electrification Timeline — Current Plan", "type": "Chart", "default": True},
    "invalid_vin":        {"name": "Vehicle Data Assumptions (if any)",        "type": "Data",     "default": True},
    # ── Phase 24: new slides — optional ───────────────────────────────────────
    "timeline_aggressive":   {"name": "Electrification Timeline — Aggressive 2030",   "type": "Chart",    "default": False},
    "timeline_conservative": {"name": "Electrification Timeline — Conservative 2040", "type": "Chart",    "default": False},
    "department_summary":    {"name": "Department Summary (if dept data in CSV)",      "type": "Chart",    "default": False},
    "facility_summary":      {"name": "Domicile Facility Summary (if location in CSV)","type": "Chart",    "default": False},
    # ── Phase 27: scenario comparison slides ──────────────────────────────────
    "scenario_investment": {"name": "Cumulative Fleet Investment by Scenario", "type": "Chart",    "default": True},
    # scenario_co2 is now the template emissions_chart slide (always present); redundant here
    "scenario_co2":        {"name": "Annual Fleet Emissions by Scenario (duplicate)", "type": "Optional", "default": False},
    # ── Phase 28: Milestone Option timeline ───────────────────────────────────
    "timeline_milestone":  {"name": "Electrification Timeline — ZEV Milestone Option", "type": "Chart", "default": False},
    # ── Existing optional slides ───────────────────────────────────────────────
    "fleet_composition":       {"name": "Fleet Composition by Body Type",          "type": "Optional", "default": False},
    "age_analysis":            {"name": "Fleet Age Distribution",                   "type": "Optional", "default": False},
    "scenario_comparison":     {"name": "Scenario Comparison (line charts)",        "type": "Optional", "default": False},
    "replacement_table":       {"name": "Priority Replacement Schedule",            "type": "Optional", "default": False},
    "data_quality":            {"name": "Data Quality & Completeness",              "type": "Optional", "default": False},
}

_TYPE_COLOR = {
    "Token":    "#E8F5E9",  # pale green
    "Chart":    "#E3F2FD",  # pale blue
    "Data":     "#FFF3E0",  # pale amber
    "Editable": "#F3E5F5",  # pale purple
    "Static":   "#FAFAFA",  # near-white
    "Optional": "#FCE4EC",  # pale rose
}

class PresentPanel:
    """
    Present tab panel — exports client-facing PowerPoint presentations using
    a per-fleet profile and the bundled visual template.
    """

    def __init__(self, parent_frame: ttk.Frame, sharing_data: dict):
        self.parent_frame = parent_frame
        self.sharing_data = sharing_data
        self.root = parent_frame.winfo_toplevel()

        # State
        self._fleet = None
        self._fleet_path: Optional[str] = None
        self._profile: Optional[PresentationProfile] = None
        self._profile_dirty = False
        self._last_pptx_path: Optional[str] = None
        self._building = False

        # Slide order state: list of slide_id strings (template + optional)
        self._all_slide_rows: List[dict] = []  # {id, name, type, included}
        self._gallery_check_vars: Dict[str, tk.BooleanVar] = {}
        self._gallery_thumb_images: Dict[str, Any] = {}  # keep PIL refs alive

        self._build_ui()

    # ─────────────────────────────────────────────────────────────────────────
    # UI construction
    # ─────────────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        # Sticky action bar at the very top of parent_frame (not scrollable)
        self._action_bar = ttk.Frame(self.parent_frame, padding=(8, 6))
        self._action_bar.pack(fill=tk.X, side=tk.TOP)
        self._build_action_bar(self._action_bar)

        # Everything else scrolls
        self._sf = ScrollableFrame(self.parent_frame)
        self._sf.pack(fill=tk.BOTH, expand=True)
        body = self._sf.scrollable_frame

        self._build_profile_section(body)
        self._build_slide_section(body)
        self._build_content_section(body)
        self._build_template_section(body)
        self._build_output_section(body)

    def _build_action_bar(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="Build Presentation", font=(Fonts.FAMILY_SANS, Fonts.SIZE_H2, Fonts.WEIGHT_BOLD)).pack(side=tk.LEFT, padx=(0, 12))
        self._build_btn = ttk.Button(parent, text="⚡ Build Presentation",
                                     command=self._on_build, state="disabled")
        self._build_btn.pack(side=tk.LEFT, padx=4)
        self._pdf_btn = ttk.Button(parent, text="📄 Export as PDF",
                                   command=self._on_export_pdf, state="disabled")
        self._pdf_btn.pack(side=tk.LEFT, padx=4)
        self._status_lbl = ttk.Label(parent, text="Load a fleet to enable export",
                                     foreground="gray")
        self._status_lbl.pack(side=tk.LEFT, padx=12)

    # ── Client Profile ────────────────────────────────────────────────────────

    def _build_profile_section(self, parent: tk.Frame) -> None:
        sec, body = self._collapsible_section(parent, "▼  Client Profile", expanded=True)
        frm = ttk.Frame(body)
        frm.pack(fill=tk.X, padx=8, pady=6)
        frm.columnconfigure(1, weight=1)
        frm.columnconfigure(3, weight=1)

        self._pv: Dict[str, tk.StringVar] = {k: tk.StringVar() for k in [
            "client_name", "meeting_date", "presentation_type",
            "presenter_name", "presenter_title", "presenter_company",
            "partner_1_name", "partner_1_title", "partner_1_org", "partner_1_email",
            "partner_2_name", "partner_2_title", "partner_2_org", "partner_2_email",
        ]}
        # Default meeting date
        self._pv["meeting_date"].set(datetime.datetime.now().strftime("%B %-d, %Y"))
        self._pv["presentation_type"].set("Kickoff")

        # Trace changes → mark dirty
        for var in self._pv.values():
            var.trace_add("write", lambda *_: self._mark_dirty())

        row = 0
        def _lbl(text, r, c, span=1):
            ttk.Label(frm, text=text, font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY)).grid(
                row=r, column=c, sticky="w", padx=4, pady=2, columnspan=span)

        def _entry(var_key, r, c, width=28, colspan=1):
            e = ttk.Entry(frm, textvariable=self._pv[var_key], width=width)
            e.grid(row=r, column=c, sticky="ew", padx=4, pady=2, columnspan=colspan)
            return e

        # Row 0: client name + type
        _lbl("Client / City Name:", row, 0)
        _entry("client_name", row, 1)
        _lbl("Type:", row, 2)
        cb = ttk.Combobox(frm, textvariable=self._pv["presentation_type"], width=15,
                          values=["Kickoff", "Update 1", "Update 2", "Update 3", "Final"],
                          state="readonly")
        cb.grid(row=row, column=3, sticky="w", padx=4, pady=2)
        row += 1

        # Row 1: date
        _lbl("Meeting Date:", row, 0)
        _entry("meeting_date", row, 1)
        row += 1

        # Separator
        ttk.Separator(frm, orient="horizontal").grid(
            row=row, column=0, columnspan=4, sticky="ew", pady=4)
        row += 1

        # Presenter rows
        _lbl("Presenter Name:", row, 0)
        _entry("presenter_name", row, 1)
        _lbl("Title:", row, 2)
        _entry("presenter_title", row, 3)
        row += 1

        _lbl("Company:", row, 0)
        _entry("presenter_company", row, 1)
        row += 1

        ttk.Separator(frm, orient="horizontal").grid(
            row=row, column=0, columnspan=4, sticky="ew", pady=4)
        row += 1

        _lbl("Partner 1 Name:", row, 0)
        _entry("partner_1_name", row, 1)
        _lbl("Title / Org:", row, 2)
        _entry("partner_1_title", row, 3)
        row += 1

        _lbl("Partner 1 Email:", row, 0)
        _entry("partner_1_email", row, 1)
        _lbl("Organization:", row, 2)
        _entry("partner_1_org", row, 3)
        row += 1

        ttk.Separator(frm, orient="horizontal").grid(
            row=row, column=0, columnspan=4, sticky="ew", pady=4)
        row += 1

        _lbl("Partner 2 Name:", row, 0)
        _entry("partner_2_name", row, 1)
        _lbl("Title / Org:", row, 2)
        _entry("partner_2_title", row, 3)
        row += 1

        _lbl("Partner 2 Email:", row, 0)
        _entry("partner_2_email", row, 1)
        _lbl("Organization:", row, 2)
        _entry("partner_2_org", row, 3)
        row += 1

        save_btn = ttk.Button(frm, text="💾  Save Profile", command=self._on_save_profile)
        save_btn.grid(row=row, column=0, columnspan=2, sticky="w", padx=4, pady=6)
        self._profile_status_lbl = ttk.Label(frm, text="", foreground="gray")
        self._profile_status_lbl.grid(row=row, column=2, columnspan=2, sticky="w")

    # ── Slide Selection (card gallery) ────────────────────────────────────────

    def _build_slide_section(self, parent: tk.Frame) -> None:
        sec, body = self._collapsible_section(parent, "▼  Slide Selection", expanded=True)

        hint = ttk.Label(body,
                         text="Check slides to include in your presentation.  "
                              "Template slides appear first; optional slides follow.",
                         foreground="gray", font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL))
        hint.pack(anchor="w", padx=8, pady=(4, 2))

        # Toolbar
        btn_row = ttk.Frame(body)
        btn_row.pack(fill=tk.X, padx=8, pady=(0, 4))
        ttk.Button(btn_row, text="Check All",   command=self._check_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="Uncheck All", command=self._uncheck_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="Reset Order", command=self._reset_slide_order).pack(side=tk.RIGHT, padx=2)

        # Scrollable canvas for card grid
        gallery_outer = ttk.Frame(body)
        gallery_outer.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self._gallery_canvas = tk.Canvas(gallery_outer, bg="#F5F5F5",
                                         highlightthickness=0, height=500)
        vsb = ttk.Scrollbar(gallery_outer, orient="vertical",
                            command=self._gallery_canvas.yview)
        self._gallery_canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._gallery_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._gallery_inner = ttk.Frame(self._gallery_canvas)
        self._gallery_cwin = self._gallery_canvas.create_window(
            (0, 0), window=self._gallery_inner, anchor="nw")

        self._gallery_inner.bind(
            "<Configure>",
            lambda e: self._gallery_canvas.configure(
                scrollregion=self._gallery_canvas.bbox("all")))
        self._gallery_canvas.bind(
            "<Configure>",
            lambda e: self._gallery_canvas.itemconfig(
                self._gallery_cwin, width=e.width))
        # Mousewheel scroll
        self._gallery_canvas.bind(
            "<MouseWheel>",
            lambda e: self._gallery_canvas.yview_scroll(
                -1 * (e.delta // 120), "units"))

        self._populate_slide_tree()

    def _populate_slide_tree(self) -> None:
        """Build _all_slide_rows (template first, then optional) and refresh gallery."""
        self._all_slide_rows = []

        for sid in TEMPLATE_SLIDE_IDS:
            meta = _TEMPLATE_SLIDE_META.get(sid, {"name": sid, "type": "Static"})
            row = {"id": sid, "name": meta["name"], "type": meta["type"], "included": True}
            self._all_slide_rows.append(row)

        # Separator row
        self._all_slide_rows.append({"id": "_sep", "name": "── Optional Add-On Slides ──",
                                      "type": "Optional", "included": False})

        for sid, meta in _OPTIONAL_SLIDE_META.items():
            default_included = meta.get("default", False)
            row = {"id": sid, "name": meta["name"], "type": meta["type"], "included": default_included}
            self._all_slide_rows.append(row)

        self._refresh_gallery()

    def _refresh_gallery(self) -> None:
        """Rebuild the card grid from self._all_slide_rows."""
        if not hasattr(self, "_gallery_inner"):
            return
        # Destroy existing card widgets
        for w in self._gallery_inner.winfo_children():
            w.destroy()
        self._gallery_check_vars.clear()
        self._gallery_thumb_images.clear()

        self._gallery_inner.columnconfigure(0, weight=1, minsize=180)
        self._gallery_inner.columnconfigure(1, weight=1, minsize=180)

        card_col = 0
        card_row = 0
        for row in self._all_slide_rows:
            sid = row["id"]
            if sid == "_sep":
                # Full-width separator label between template and optional slides
                sep_frame = ttk.Frame(self._gallery_inner)
                sep_frame.grid(row=card_row, column=0, columnspan=2,
                               sticky="ew", padx=8, pady=(14, 4))
                ttk.Separator(sep_frame, orient="horizontal").pack(
                    fill=tk.X, side=tk.TOP, pady=2)
                ttk.Label(sep_frame, text="Optional Add-On Slides",
                          foreground="gray",
                          font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL, "italic"),
                          ).pack(anchor="w")
                card_row += 1
                card_col = 0
                continue

            self._create_slide_card(self._gallery_inner, row, card_row, card_col)
            card_col += 1
            if card_col >= 2:
                card_col = 0
                card_row += 1

        self._gallery_canvas.yview_moveto(0)

    def _create_slide_card(self, parent: ttk.Frame,
                            row: dict, grid_row: int, grid_col: int) -> None:
        """Create a single slide card and grid it into `parent`."""
        sid       = row["id"]
        type_name = row["type"]
        bg_color  = _TYPE_COLOR.get(type_name, "#FAFAFA")

        card = tk.Frame(parent, bg=bg_color, relief="solid", borderwidth=1,
                        cursor="hand2", padx=4, pady=4)
        card.grid(row=grid_row, column=grid_col, padx=6, pady=6, sticky="nsew")

        # Thumbnail area (16:9 proportions)
        THUMB_W, THUMB_H = 160, 90
        thumb_canvas = tk.Canvas(card, width=THUMB_W, height=THUMB_H,
                                 bg=bg_color, highlightthickness=1,
                                 highlightbackground="#BDBDBD")
        thumb_canvas.pack()

        img = self._load_thumb_image(sid, THUMB_W, THUMB_H)
        if img is not None:
            self._gallery_thumb_images[sid] = img
            thumb_canvas.create_image(THUMB_W // 2, THUMB_H // 2, image=img, anchor="center")
        else:
            # Colored placeholder with icon
            thumb_canvas.create_rectangle(2, 2, THUMB_W - 2, THUMB_H - 2,
                                          fill=bg_color, outline="#BDBDBD")
            icon = {"Chart": "CH", "Token": "TK", "Static": "ST",
                    "Data": "DT", "Editable": "ED", "Optional": "OPT"}.get(type_name, "SL")
            thumb_canvas.create_text(THUMB_W // 2, THUMB_H // 2 - 8,
                                     text=icon, font=("Calibri", 22, "bold"),
                                     fill="#BDBDBD", anchor="center")
            thumb_canvas.create_text(THUMB_W // 2, THUMB_H // 2 + 20,
                                     text=row["name"][:28],
                                     font=("Calibri", 7), fill="#9E9E9E", anchor="center")

        # Checkbox + title row
        var = tk.BooleanVar(value=row["included"])
        self._gallery_check_vars[sid] = var

        def _on_toggle(s=sid, v=var, r=row):
            r["included"] = v.get()

        chk_frame = tk.Frame(card, bg=bg_color)
        chk_frame.pack(fill=tk.X, pady=(4, 0))
        chk = tk.Checkbutton(chk_frame, variable=var, command=_on_toggle,
                              bg=bg_color, activebackground=bg_color, cursor="hand2")
        chk.pack(side=tk.LEFT)
        name_lbl = tk.Label(chk_frame, text=row["name"], bg=bg_color,
                             font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL, "bold"),
                             wraplength=130, justify="left", anchor="w")
        name_lbl.pack(side=tk.LEFT, fill=tk.X)

        # Type badge
        badge = tk.Label(card, text=type_name, bg=bg_color,
                          font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL),
                          foreground="#757575", anchor="e")
        badge.pack(fill=tk.X)

        # Clicking anywhere on the card toggles inclusion
        def _card_click(event, v=var, r=row):
            v.set(not v.get())
            r["included"] = v.get()

        for widget in (card, thumb_canvas, chk_frame, name_lbl, badge):
            widget.bind("<Button-1>", _card_click)

    def _load_thumb_image(self, sid: str, width: int, height: int):
        """Load a pre-rendered thumbnail PNG for the slide, or return None."""
        png_path = os.path.join(str(ASSETS_DIR), "slide_thumbnails", f"{sid}.png")
        if not os.path.isfile(png_path) or not _PIL_AVAILABLE:
            return None
        try:
            img = _PILImage.open(png_path).resize((width, height), _PILImage.LANCZOS)
            return _ImageTk.PhotoImage(img)
        except Exception:
            return None

    # ── Consulting content ────────────────────────────────────────────────────

    def _build_content_section(self, parent: tk.Frame) -> None:
        sec, body = self._collapsible_section(parent, "▶  Consulting Content", expanded=False)

        for label, key in [
            ("Agenda Items (one per line):", "agenda"),
            ("Data Needs (one per line):", "data_needs"),
            ("Next Steps (one per line):", "next_steps"),
        ]:
            ttk.Label(body, text=label, font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY)).pack(anchor="w", padx=8, pady=(6, 0))
            frm = ttk.Frame(body)
            frm.pack(fill=tk.X, padx=8, pady=(0, 4))
            txt = tk.Text(frm, height=4, wrap="word",
                          font=("Calibri", 10), relief="solid", borderwidth=1)
            txt.pack(side=tk.LEFT, fill=tk.X, expand=True)
            sb = ttk.Scrollbar(frm, orient="vertical", command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            sb.pack(side=tk.RIGHT, fill=tk.Y)
            setattr(self, f"_txt_{key}", txt)

    # ── Template ──────────────────────────────────────────────────────────────

    def _build_template_section(self, parent: tk.Frame) -> None:
        sec, body = self._collapsible_section(parent, "▶  Template", expanded=False)
        frm = ttk.Frame(body)
        frm.pack(fill=tk.X, padx=8, pady=6)

        self._tpl_var = tk.StringVar(value=DEFAULT_TEMPLATE_PATH)
        ttk.Label(frm, text="Template file:").pack(side=tk.LEFT, padx=4)
        self._tpl_lbl = ttk.Label(frm, textvariable=self._tpl_var,
                                   foreground="gray", wraplength=380)
        self._tpl_lbl.pack(side=tk.LEFT, padx=4, fill=tk.X, expand=True)
        ttk.Button(frm, text="Browse…", command=self._browse_template).pack(side=tk.LEFT, padx=4)
        ttk.Button(frm, text="Reset",   command=self._reset_template).pack(side=tk.LEFT, padx=2)

    # ── Output ────────────────────────────────────────────────────────────────

    def _build_output_section(self, parent: tk.Frame) -> None:
        sec, body = self._collapsible_section(parent, "▶  Output", expanded=False)
        frm = ttk.Frame(body)
        frm.pack(fill=tk.X, padx=8, pady=6)

        ttk.Label(frm, text="Save to:").pack(side=tk.LEFT, padx=4)
        self._out_var = tk.StringVar(value=str(EXPORT_DIR))
        ttk.Entry(frm, textvariable=self._out_var, width=40).pack(side=tk.LEFT, padx=4)
        ttk.Button(frm, text="Browse…",
                   command=self._browse_output).pack(side=tk.LEFT, padx=2)

        ttk.Label(frm, text="Filename:").pack(side=tk.LEFT, padx=(12, 4))
        self._fname_var = tk.StringVar(value="presentation.pptx")
        ttk.Entry(frm, textvariable=self._fname_var, width=28).pack(side=tk.LEFT, padx=4)

    # ─────────────────────────────────────────────────────────────────────────
    # Collapsible section helper
    # ─────────────────────────────────────────────────────────────────────────

    def _collapsible_section(self, parent: tk.Frame, label: str,
                              expanded: bool = True) -> tuple:
        """Create a collapsible labeled section. Returns (header_frame, body_frame)."""
        container = ttk.Frame(parent, relief="flat")
        container.pack(fill=tk.X, padx=6, pady=4)

        hdr = tk.Frame(container, bg=PRIMARY_HEX_1, cursor="hand2")
        hdr.pack(fill=tk.X)
        lbl_var = tk.StringVar(value=label)
        lbl_widget = tk.Label(hdr, textvariable=lbl_var, bg=PRIMARY_HEX_1,
                              fg="white", font=(Fonts.FAMILY_SANS, Fonts.SIZE_H3, Fonts.WEIGHT_BOLD),
                              anchor="w", padx=8, pady=4)
        lbl_widget.pack(fill=tk.X)

        body = ttk.Frame(container, relief="solid", borderwidth=1)
        if expanded:
            body.pack(fill=tk.X)

        def _toggle(event=None):
            if body.winfo_ismapped():
                body.pack_forget()
                txt = label.replace("▼", "▶", 1)
            else:
                body.pack(fill=tk.X)
                txt = label.replace("▶", "▼", 1)
            lbl_var.set(txt)

        hdr.bind("<Button-1>", _toggle)
        lbl_widget.bind("<Button-1>", _toggle)

        return hdr, body

    # ─────────────────────────────────────────────────────────────────────────
    # Slide gallery interactions
    # ─────────────────────────────────────────────────────────────────────────

    def _check_all(self) -> None:
        for row in self._all_slide_rows:
            if row["id"] != "_sep":
                row["included"] = True
        self._refresh_gallery()

    def _uncheck_all(self) -> None:
        for row in self._all_slide_rows:
            if row["id"] not in ("_sep", "cover"):
                row["included"] = False
        self._refresh_gallery()

    def _reset_slide_order(self) -> None:
        """Restore default slide order (template first, then optional)."""
        state = {r["id"]: r["included"] for r in self._all_slide_rows}
        self._populate_slide_tree()
        for row in self._all_slide_rows:
            if row["id"] in state:
                row["included"] = state[row["id"]]
        self._refresh_gallery()

    # ─────────────────────────────────────────────────────────────────────────
    # Profile helpers
    # ─────────────────────────────────────────────────────────────────────────

    def _mark_dirty(self) -> None:
        self._profile_dirty = True

    def load_profile(self, profile: PresentationProfile) -> None:
        """Populate form fields from a PresentationProfile."""
        self._profile = profile
        mapping = {
            "client_name":       profile.client_name,
            "meeting_date":      profile.meeting_date,
            "presentation_type": profile.presentation_type,
            "presenter_name":    profile.presenter_name,
            "presenter_title":   profile.presenter_title,
            "presenter_company": profile.presenter_company,
            "partner_1_name":    profile.partner_1_name,
            "partner_1_title":   profile.partner_1_title,
            "partner_1_org":     profile.partner_1_org,
            "partner_1_email":   profile.partner_1_email,
            "partner_2_name":    profile.partner_2_name,
            "partner_2_title":   profile.partner_2_title,
            "partner_2_org":     profile.partner_2_org,
            "partner_2_email":   profile.partner_2_email,
        }
        for key, val in mapping.items():
            if val and key in self._pv:
                self._pv[key].set(val or "")

        # Restore consulting content text areas
        for attr, items in [
            ("_txt_agenda",     profile.agenda_items),
            ("_txt_data_needs", profile.data_needs_items),
            ("_txt_next_steps", profile.next_steps_items),
        ]:
            txt = getattr(self, attr, None)
            if txt and items:
                txt.delete("1.0", tk.END)
                txt.insert("1.0", "\n".join(items))

        # Restore slide inclusion / order
        if profile.included_slides:
            self._apply_profile_slides(profile)

        # Template override
        if profile.template_path:
            self._tpl_var.set(profile.template_path)

        self._profile_dirty = False
        self._profile_status_lbl.configure(text="Profile loaded", foreground="gray")

    def _apply_profile_slides(self, profile: PresentationProfile) -> None:
        """Update slide inclusion and order from profile.included_slides."""
        included_set = set(profile.included_slides)
        optional_set = set(profile.optional_slides or [])

        state_map: Dict[str, bool] = {}
        for sid in TEMPLATE_SLIDE_IDS:
            state_map[sid] = sid in included_set
        for sid in _OPTIONAL_SLIDE_META:
            state_map[sid] = sid in optional_set

        for row in self._all_slide_rows:
            if row["id"] in state_map:
                row["included"] = state_map[row["id"]]

        # Reorder template slides to match profile order
        template_order = [sid for sid in profile.included_slides
                          if sid in TEMPLATE_SLIDE_IDS]
        optional_rows = [r for r in self._all_slide_rows
                         if r["id"] not in TEMPLATE_SLIDE_IDS and r["id"] != "_sep"]
        template_rows_map = {r["id"]: r for r in self._all_slide_rows
                             if r["id"] in TEMPLATE_SLIDE_IDS}
        new_template_rows = [template_rows_map[sid] for sid in template_order
                             if sid in template_rows_map]
        remaining = [r for r in self._all_slide_rows
                     if r["id"] in TEMPLATE_SLIDE_IDS and r["id"] not in template_order]
        sep = {"id": "_sep", "name": "── Optional Add-On Slides ──",
               "type": "Optional", "included": False}
        self._all_slide_rows = new_template_rows + remaining + [sep] + optional_rows
        self._refresh_gallery()

    def _read_profile(self) -> PresentationProfile:
        """Build a PresentationProfile from the current form state."""
        p = PresentationProfile()
        for key in self._pv:
            setattr(p, key, self._pv[key].get().strip())

        def _lines(attr) -> List[str]:
            txt = getattr(self, attr, None)
            if not txt:
                return []
            content = txt.get("1.0", tk.END).strip()
            return [l.strip() for l in content.splitlines() if l.strip()]

        p.agenda_items     = _lines("_txt_agenda")
        p.data_needs_items = _lines("_txt_data_needs")
        p.next_steps_items = _lines("_txt_next_steps")

        # Slide selection + order from Treeview
        p.included_slides = [r["id"] for r in self._all_slide_rows
                              if r["included"] and r["id"] in TEMPLATE_SLIDE_IDS]
        p.optional_slides  = [r["id"] for r in self._all_slide_rows
                               if r["included"] and r["id"] in _OPTIONAL_SLIDE_META]

        tpl = self._tpl_var.get().strip()
        p.template_path = tpl if tpl != DEFAULT_TEMPLATE_PATH else None

        return p

    def _on_save_profile(self) -> None:
        """Save current profile to sidecar file."""
        if not self._fleet_path:
            messagebox.showinfo("Save Profile",
                "Load a fleet file first. The profile is saved next to the fleet CSV.")
            return
        from data.processor import save_presentation_profile
        profile = self._read_profile()
        ok = save_presentation_profile(self._fleet_path, profile)
        if ok:
            self._profile_dirty = False
            self._profile_status_lbl.configure(text="✓ Saved", foreground="green")
            self.root.after(3000, lambda: self._profile_status_lbl.configure(
                text="", foreground="gray"))
        else:
            messagebox.showerror("Save Failed", "Could not save profile. Check file permissions.")

    # ─────────────────────────────────────────────────────────────────────────
    # Template / output browsing
    # ─────────────────────────────────────────────────────────────────────────

    def _browse_template(self) -> None:
        path = filedialog.askopenfilename(
            title="Select PowerPoint Template",
            filetypes=[("PowerPoint", "*.pptx *.potx"), ("All files", "*.*")],
        )
        if path:
            self._tpl_var.set(path)

    def _reset_template(self) -> None:
        self._tpl_var.set(DEFAULT_TEMPLATE_PATH)

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self._out_var.set(path)

    def _get_out_path(self) -> str:
        folder = self._out_var.get().strip() or str(EXPORT_DIR)
        fname  = self._fname_var.get().strip() or "presentation.pptx"
        if not fname.endswith(".pptx"):
            fname += ".pptx"
        return os.path.join(folder, fname)

    # ─────────────────────────────────────────────────────────────────────────
    # Build / PDF export
    # ─────────────────────────────────────────────────────────────────────────

    def _on_build(self) -> None:
        if self._building:
            return
        fleet = self.sharing_data.get("fleet")
        if not fleet or not getattr(fleet, "vehicles", None):
            messagebox.showwarning("No Fleet", "Load a fleet before building the presentation.")
            return

        profile  = self._read_profile()
        out_path = self._get_out_path()
        tpl      = self._tpl_var.get().strip()

        # Auto-update filename from client + type
        if profile.client_name:
            slug = profile.client_name.replace(" ", "_")[:25]
            ptype = (profile.presentation_type or "Kickoff").replace(" ", "_")
            ts = datetime.datetime.now().strftime("%Y%m%d")
            self._fname_var.set(f"{slug}_{ptype}_{ts}.pptx")
            out_path = self._get_out_path()

        self._building = True
        self._build_btn.configure(state="disabled", text="Building…")
        self._set_status("Building presentation…", color="blue")

        scenario_results = self.sharing_data.get("scenario_results")

        def _worker():
            try:
                result = export_presentation(
                    fleet_data=fleet,
                    profile=profile,
                    out_path=out_path,
                    template_path=tpl if os.path.isfile(tpl) else None,
                    scenario_results=scenario_results,
                )
                # Support both old (str) and new (dict) return formats
                if isinstance(result, dict):
                    self.root.after(0, lambda: self._on_build_done(
                        result["path"], result))
                else:
                    self.root.after(0, lambda: self._on_build_done(result))
            except Exception as exc:
                logger.exception("Presentation export failed")
                msg = str(exc)
                self.root.after(0, lambda m=msg: self._on_build_error(m))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_build_done(self, path: str, stats: dict = None) -> None:
        self._building = False
        self._last_pptx_path = path
        self._build_btn.configure(state="normal", text="⚡ Build Presentation")
        self._pdf_btn.configure(state="normal")
        self._set_status(f"✓ Saved: {os.path.basename(path)}", color="green")

        # Build summary message
        summary = f"Presentation saved to:\n{path}"
        if stats:
            total = stats.get("total_slides", "?")
            charts_ok = stats.get("charts_succeeded", "?")
            charts_total = stats.get("charts_attempted", "?")
            optional = stats.get("optional_slides_added", 0)
            summary += f"\n\n{total} slides generated"
            summary += f"\n{charts_ok}/{charts_total} charts rendered"
            if optional:
                summary += f"\n{optional} optional slides added"
            if charts_ok != charts_total:
                summary += "\n\nNote: Some charts could not be rendered (check logs for details)"

        if messagebox.askyesno("Export Complete", f"{summary}\n\nOpen now?"):
            self._open_file(path)

    def _on_build_error(self, msg: str) -> None:
        self._building = False
        self._build_btn.configure(state="normal", text="⚡ Build Presentation")
        self._set_status("✗ Export failed", color="red")
        messagebox.showerror("Export Failed", f"Could not build presentation:\n\n{msg}")

    def _on_export_pdf(self) -> None:
        if not self._last_pptx_path:
            messagebox.showinfo("No PPTX", "Build a presentation first, then export as PDF.")
            return
        self._set_status("Converting to PDF…", color="blue")

        def _worker():
            pdf_path = export_pdf(self._last_pptx_path)
            self.root.after(0, lambda: self._on_pdf_done(pdf_path))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_pdf_done(self, pdf_path: Optional[str]) -> None:
        if pdf_path and os.path.isfile(pdf_path):
            self._set_status(f"✓ PDF: {os.path.basename(pdf_path)}", color="green")
            if messagebox.askyesno("PDF Ready", f"PDF saved to:\n{pdf_path}\n\nOpen now?"):
                self._open_file(pdf_path)
        else:
            self._set_status("PDF export unavailable", color="gray")
            messagebox.showinfo(
                "PDF Export",
                "Automatic PDF conversion requires LibreOffice or the pptx2pdf package.\n\n"
                "To install: pip install pptx2pdf\n\n"
                "Or open the .pptx in PowerPoint and use File > Save as PDF.",
            )

    def _set_status(self, msg: str, color: str = "gray") -> None:
        self._status_lbl.configure(text=msg, foreground=color)

    @staticmethod
    def _open_file(path: str) -> None:
        try:
            if platform.system() == "Darwin":
                subprocess.Popen(["open", path])
            elif platform.system() == "Windows":
                os.startfile(path)
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception:
            pass

    # ─────────────────────────────────────────────────────────────────────────
    # Public API (called by main_window.py)
    # ─────────────────────────────────────────────────────────────────────────

    def refresh_data(self) -> None:
        """Called by MainWindow after a fleet loads. Updates status and pre-fills fields."""
        fleet = self.sharing_data.get("fleet")
        if fleet and getattr(fleet, "vehicles", None):
            n = len(fleet.vehicles)
            self._fleet = fleet
            self._build_btn.configure(state="normal")
            self._set_status(f"Ready — {n} vehicles loaded", color="gray")

            # Pre-fill client name from fleet name if field is blank
            if not self._pv["client_name"].get().strip() and fleet.name:
                # Clean up auto-generated fleet names
                name = fleet.name.replace("Fleet Analysis - ", "").strip()
                self._pv["client_name"].set(name)

            # Auto-update output filename suggestion
            client = self._pv["client_name"].get().strip()
            ptype  = self._pv["presentation_type"].get().strip() or "Kickoff"
            ts = datetime.datetime.now().strftime("%Y%m%d")
            slug = (client or "Fleet").replace(" ", "_")[:25]
            self._fname_var.set(f"{slug}_{ptype}_{ts}.pptx")
        else:
            self._build_btn.configure(state="disabled")
            self._set_status("Load a fleet to enable export", color="gray")

    def set_fleet_path(self, path: Optional[str]) -> None:
        """Called by MainWindow when a fleet CSV is opened. Used for sidecar profile save."""
        self._fleet_path = path

    def get_panel_frame(self) -> ttk.Frame:
        """Return the parent frame (required by MainWindow)."""
        return self.parent_frame


# ─────────────────────────────────────────────────────────────────────────────
# Module-level factory (backward compat)
# ─────────────────────────────────────────────────────────────────────────────

def create_present_panel(parent_frame: ttk.Frame, sharing_data: dict) -> PresentPanel:
    return PresentPanel(parent_frame, sharing_data)
