# CLAUDE.md — Fleet Electrification Analyzer

## Project Overview

Desktop Python/Tkinter app for fleet electrification consulting. Decodes vehicle VINs, retrieves MPG data from government APIs, classifies vehicles for CARB ACF compliance, assigns electrification timelines, and generates analysis + client deliverables.

**Stack:** Python 3.8+, Tkinter, pandas, requests, python-pptx, matplotlib, openpyxl, BeautifulSoup/Selenium
**Entry point:** `python app.py` (GUI) or `python app.py -b -i in.csv -o out.csv` (batch)
**Current version:** 3.0.8 on branch `uiux/v3_0_3`

---

## Build & Run

```bash
pip install -r requirements.txt
python app.py              # GUI mode
python app.py --help       # CLI options
python app.py -b -i in.csv -o out.csv  # Batch mode
```

### Testing

```bash
python -m pytest tests/ -v           # Run all 194 tests
python -m pytest tests/ -v -k acf    # ACF tests only
python -m pytest tests/ -v --tb=short
```

---

## Architecture

```
app.py                  → Entry point, CLI args, window setup
settings.py             → All configuration constants
utils.py                → Shared utilities, VIN validation, Cache

data/
  models.py             → Dataclasses (VehicleIdentification, FuelEconomyData, FleetVehicle, Fleet)
  providers.py          → API clients (NHTSA VIN decoder, FuelEconomy.gov)
  processor.py          → CSV I/O, ProcessingPipeline, BatchProcessor
  vehicle_database.py   → SQLite vehicle MPG reference database (analyst-maintained)

ui/
  main_window.py        → Main window with tabbed interface
  theme.py              → Centralized styling
  widgets.py            → Reusable UI widgets (tooltips, status bar, progress dialog)
  process_panel.py      → Tab 1: VIN processing
  results_panel.py      → Tab 2: Results table
  analysis_panel.py     → Tab 3: Analysis tools + KPI cards
  timeline_panel.py     → Tab 4: Electrification Timeline editor (Phase 19)
  present_panel.py      → Tab 5: PowerPoint generation
  database_panel.py     → Tab 6: Vehicle MPG reference database CRUD

analysis/
  calculations.py       → TCO, ROI, emissions, year-by-year cash flow model
  charts.py             → Matplotlib chart generation (22 chart types via ChartFactory)
  reports.py            → Excel multi-tab report generation (8 worksheets)
  acf_compliance.py     → CARB ACF classification per vehicle (ZEV/A/B/C/D)
  electrification_timeline.py → ACF-deadline-first year assignment + score-based queue
  ev_database.py        → EV equivalent matching (~25 entries, reliability limited for heavy-duty)
  rate_database.py      → State energy rates + federal/state incentive programs
  scenarios.py          → ElectrificationScenario dataclass + 4 presets + compare_scenarios()

commercial_vehicle_scraper.py  → Web scraping for commercial vehicle specs
powerpoint_export.py           → PPTX generation engine (14 slide builders, all implemented)
powerpoint_charts.py           → Native PPTX chart creation (9 chart types)
powerpoint_customizer.py       → PPTX customization helpers

tests/
  conftest.py                      → Factory helpers & pytest fixtures
  test_vin_validation.py
  test_models.py
  test_normalize.py
  test_calculations.py
  test_acf_compliance.py
  test_electrification_timeline.py  ← 3 regression tests added (Phase 17)
  test_csv_mapping.py
  test_match_confidence.py
  test_timeline_override.py         ← 14 regression tests added (Phase 19)
```

---

## Data Pipeline

```
CSV input
  → VIN validation (format check)
  → NHTSA API       (decode VIN: make, model, year, body class, GVWR, engine, fuel type)
  → FuelEconomy.gov (MPG — light-duty focused; poor coverage for Class 4-8)
  → Commercial scraper (Fuelly, EPA SmartWay, Ford — 3 live / 10 stubs)
  → SQLite DB lookup (analyst-maintained MPG reference — Tier 3b)
  → EPA class-average fallback (GVWR-bucket estimate — Tier 4, tagged as estimate)
  → Quality scoring (0-100)
  → ACF classification (ZEV/A/B/C/D per vehicle)
  → Electrification year assignment (ACF deadline-first for Cat B; score-based for C/D)
  → EV equivalent matching (heuristic, ~25 entries, hidden from default display)
  → FleetVehicle objects → UI table / CSV / Excel report / PowerPoint
```

---

## Key Feature State

### ACF Classification
- **ZEV** — Already zero-emission
- **A** — Exempt, GVWR ≤ 8,500 lbs (light-duty)
- **B** — Mandate-subject, medium/heavy-duty
- **C** — Exempt by body type (dump truck, crane, concrete mixer, etc.)
- **D** — Emergency vehicle (PPV/SSV trim, ambulance, fire apparatus)

### Electrification Timeline
- **Category B:** Assigned CARB deadline year by GVWR class (`ACF_DEADLINE_TABLE`): Class 2b-4→2035, Class 5-8a→2039, Class 8b→2042. High-urgency (score ≥ 0.55) → earliest purchase milestone checkpoint.
- **Categories C/D and unknown-GVWR B:** Score-based priority queue: `0.55*age + 0.25*odometer + 0.10*annual_mileage + acf_boost`. Budget smoothing distributes evenly across available years.
  - **N ≤ num_years:** Vehicles are linearly spaced across the full horizon using `idx = round(k * (num_years-1) / (n-1))`. Highest-priority → earliest year, lowest-priority → last year.
  - **N > num_years:** Floor-based distribution (`base = n // num_years`, `extra = n % num_years`). All years receive vehicles; earlier years get one extra if N is not evenly divisible.
  - The same algorithm is used in both `electrification_timeline.py` and `scenarios.py`.
- **N/A reason codes:** "Already ZEV," "Processing failed," "ACF classification unavailable"

### Excel Export (Two Paths — Known UX Issue)
- **Results tab "Export"** → single-sheet "Fleet Analysis" (raw vehicle data)
- **Analysis tab "Export" or File > Export Report** → 8-tab report: Vehicle Data, Summary, Electrification Analysis, Charging Infrastructure, Emissions Inventory, TCO Model (live Excel formulas), Replacement Schedule, Summary Dashboard. The extra sheets only generate if analysis has been run first.

### PowerPoint (15 template slides + optional slides)
Template slides: cover, agenda, carb_overview, acf_scenarios, acf_exemptions, key_findings, timeline_chart, emissions_chart, incentives, data_needs, next_steps, contact, appendix, infra_costs_chart, tco_chart.
Optional slides: acf_composition (pie), timeline_moderate/aggressive/conservative/current_plan (scenario timelines), invalid_vin (data assumptions), department_summary, facility_summary, age_analysis, scenario_comparison, replacement_table, data_quality.
- **ACF timeline chart:** Always adds all 4 ACF series (consistent legend across scenario slides); data labels enabled (zeros suppressed via numFmt).
- **GHG chart:** 3-series line (Baseline flat / M·H Duty Only / Whole Fleet); axis title "Metric Tons CO₂e".
- **TCO chart:** Clustered columns (ICE vs EV side-by-side); axis title "Annual Cost ($)".
- **Key Findings:** Fleet count → dept/facility context → ZEV → EV-replaceable % → exemptions → earliest mandate year → achievability.
- **Template support:** Loads `.pptx`/`.potx` via Present tab. Gold-standard deck: `Atascadero 1st Fleet Update Presentation.pptx`.

### Scenario Engine
- `analysis/scenarios.py` — 7 presets total: **4 time-based** (Aggressive 2030, Moderate 2035, Conservative 2040, ACF Compliance Only 2035) used by Present tab + PowerPoint; **3 scope-based** (Minimum Compliance, All Excl. Emergency, Whole Fleet) used by Analysis tab Scenario Comparison. `SCOPE_SCENARIO_KEYS = ("minimum_compliance", "all_except_emergency", "whole_fleet")` exported for the UI.
- `ElectrificationScenario` dataclass now has `include_light_duty: bool` field — controls whether Cat A vehicles are scheduled in the scenario.
- `compare_scenarios()` returns per-year vehicles/cost/co2/savings for each scenario
- Wired into Present panel (scenario checkboxes + custom scenario builder — still uses the 4 time-based presets)
- **Analysis tab Scenario Comparison (Phase 22):** 3 scope-based checkboxes (Minimum Compliance / All Excl. Emergency / Whole Fleet) + **"Current Plan (Overrides)"** + Horizon spinbox (default 2040, configurable 2026–2060). Scope presets run to the user-configured horizon year via `dataclasses.replace()`. Auto-runs after Run Full Analysis.

### Notebook Tab Order (as of Phase 19)

| Index | Tab | Notes |
|-------|-----|-------|
| 0 | Process | VIN processing |
| 1 | Results | Vehicle table |
| 2 | Analysis | Dashboard + KPIs |
| 3 | Timeline | Year editor (Phase 19) |
| 4 | Present | PowerPoint export |
| 5 | Database | MPG reference DB |

`_navigate_to_present()` in `analysis_panel.py` uses `notebook.select(4)`.
`_navigate_to_timeline()` uses `notebook.select(3)`.
`on_tab_changed()` in `main_window.py` handles index 3 (Timeline) and 5 (Database).

### Analysis Tab — Current State (as of Phase 22)
- **Layout:** top-to-bottom `ScrollableFrame` dashboard (findings-first design)
- **Sticky action bar:** `⚡ Run Full Analysis` button + status label
- **Fleet Snapshot:** 4 KPI chips (Fleet Size, ACF-B Count, Avg MPG, MPG Coverage %)
- **Two-column row:** Top 5 Priority Vehicles `Treeview` | ACF Compliance donut
- **Scenario Comparison:** 3 scope-based checkboxes (Minimum Compliance / All Excl. Emergency / Whole Fleet) + Horizon spinbox (2040 default) + separator + **"Current Plan (Overrides)"** + Compare button. Auto-runs after Run Full Analysis. Dual charts: CO₂ trajectory + cumulative cost ($M). Vehicles-per-year Treeview.
- **Stale indicator:** status label turns amber when parameters change after analysis has run: *"Parameters changed — re-run analysis to update results."*
- **TCO Summary:** 3 KPI cards (Annual Savings, Payback Period, Infrastructure Est.)
- **Electrification Timeline Gantt** *(Phase 20, collapsible — starts expanded)*: read-only Gantt (grouped-by-ACF default, per-vehicle toggle); **Max rows combobox** (25/50/100/All, default 50, applies to Per Vehicle view); "→ Edit in Timeline tab" button; refreshed by `_update_summary()` and on Analysis tab activation. **Double-click any Top 5 row → switches to Results tab and selects that vehicle.**
- **Chart Browser:** collapsible (▶ starts collapsed)
- **Parameters:** collapsible (▼ starts expanded; auto-collapses post-analysis)
- **Export bar:** Export Excel Report (triggers "Timelines to Include" dialog for .xlsx, all 4 scenarios checked by default) + Build Presentation buttons; both buttons disabled until fleet is loaded
- **Override helpers (static):** `_apply_ev_year_override(vehicle, year_str)`, `_reset_ev_year_override(vehicle)`, `_apply_acf_override(vehicle, code)`, `_reset_acf_override(vehicle)`, `_reset_all_overrides()` — used by Timeline tab, Results tab, and Tools menu
- **Module-level helpers:** `ACF_LABELS` dict (code → plain-English), `_show_acf_ev_year_dialog(parent, old, new, year)` → returns 'recalculate'/'keep'/'cancel'
- **Module-level Gantt functions:** `_draw_gantt_chart(ax, vehicles, view, max_vehicles=0)`, `_gantt_grouped()`, `_gantt_per_vehicle()`, `_gantt_year_bounds()`, `GANTT_YEAR_MIN=2026`, `GANTT_YEAR_MAX=2040` (floor; chart expands dynamically to show years beyond 2040)

### Timeline Tab — Current State (as of Phase 22)
- **File:** `ui/timeline_panel.py`
- **Full-fleet `Treeview`:** all processed vehicles; columns: Asset ID, Year, Make, Model, ACF, Wt Class, Body Type, Proposed EV Year ✎ (editable), System Rec., Override?
- **Click-to-sort:** all column headers are sortable (ascending/descending, toggling); sort arrow shown in header
- **Column stretch:** Make, Model, Body Type columns expand to fill available width
- **Inline editing:** two editable columns:
  - **Proposed EV Year ✎** (col #8): double-click → `ttk.Entry`; validates year 2026–2060; `<Return>`/`<FocusOut>` commits; `<Escape>` cancels
  - **ACF** (col #5): double-click → `ttk.Combobox` dropdown with 5 options; on selection → warns user and offers "Recalculate EV Year" / "Keep Current Year" / "Cancel"
- **Override write-back:**
  - EV year edits: `_apply_ev_year_override()` → `custom_fields["Proposed EV Year"]`, stores original in `["System Recommended EV Year"]`, sets `["EV Year Overridden"] = "Yes"`
  - ACF edits: `_apply_acf_override()` → `custom_fields["ACF Category"]` + `["_acf_code"]`, stores original in `["Original ACF Category"]`, sets `["ACF Category Overridden"] = "Yes"`; if Recalculate chosen, clears EV year override markers and calls `assign_electrification_years([vehicle])`
- **Row tints:** amber #FFF3E0 (EV year override), light blue #E3F2FD (ACF override only), light purple #F3E8FD (both overrides)
- **Override count label:** shows e.g. `"2 EV year overrides (1 earlier, 1 later)  |  1 ACF category override"`
- **Toolbar:** Reset Selected (clears both EV year AND ACF overrides for selection), Reset All, override count label
- **Live Gantt chart:** grouped-by-ACF (default) or per-vehicle (toggle); X axis extends dynamically to cover all scheduled EV years (floor 2026–2040); Per Vehicle view has configurable Max rows (25/50/100/All, default 50); refreshes on every edit
- **Filter bar** (between toolbar and table): live search box (make/model/asset); ACF category checkboxes with **full plain-English labels** (Already ZEV / Light-Duty (Exempt) / Mandate-Subject / Body-Type Exempt / Emergency Vehicle); EV year range spinboxes (From–To); Clear Filters button. Uses detach/reattach to preserve IID stability.
- **`_all_iids: list`**: insertion-ordered IID list maintained in `_populate_table()`, used by `_apply_filter()`.
- **`copy_selection()`:** public method for Edit > Copy — copies selected rows (or all if none selected) as tab-delimited text.
- **`set_fleet(fleet, on_year_changed=None, on_acf_changed=None)`:** public API called by MainWindow
- **`on_year_changed` callback:** notifies `main_window._on_timeline_year_changed()` which refreshes the Analysis tab's Gantt
- **`on_acf_changed` callback:** notifies `main_window._on_acf_override()` which refreshes Timeline + Analysis Gantt

### Excel Export — Timelines to Include Dialog (Phase 19/20)
- Triggered when user clicks "Export Excel Report" or File > Export Report (`.xlsx` only)
- Checkboxes: Aggressive (2030) | Moderate (2035) | Conservative (2040) | ACF Compliance (2035) — **all 4 checked by default**
- Selected scenarios → additional columns in the **Replacement Schedule** sheet (same sheet, side-by-side, no extra sheets)
- Override flagging in Replacement Schedule: amber cell fill for overridden rows + "Overridden?" boolean column ("Yes" / "")
- `get_scenario_year_assignments(vehicles, scenario_name)` in `scenarios.py` computes per-vehicle year for a scenario without mutating originals

### Results Tab — Right-Click Context Menu (as of Phase 22)
Right-click any row to access:
- **View Details** — show vehicle detail popup
- **Copy Selected** — copy to clipboard
- **Analyze Selected** — summary stats for selection
- **Select All / Deselect All**
- **Save MPG to Database** — save vehicle MPG to the SQLite reference DB
- **Override ACF Category...** — opens a category picker dialog; on confirm, shows `_show_acf_ev_year_dialog()` to handle EV year; applies via `AnalysisPanel._apply_acf_override()`; fires `on_acf_override_callback` so Timeline + Analysis Gantt refresh

### Main Window — Tools Menu (as of Phase 20)
- **Reset All EV Year Overrides:** resets all manual year overrides fleet-wide; confirms count before acting; disabled if no fleet loaded

---

## Active Backlog

### High Priority

*(Previously: "Consolidate the two Excel export paths" — resolved in Phase 20. File > Export Report now routes through `analysis_panel.export_full_report()`, ensuring the Timelines dialog, override flagging, and 8-tab structure are always applied.)*

### Medium Priority

**3. ACF HPF vs. non-HPF distinction**
Timeline applies CARB High-Priority Fleet deadlines to all Cat B vehicles. Non-HPF fleets have different compliance schedules. Add a "Fleet Type" toggle in Analysis tab (HPF / Non-HPF / State Agency) to adjust deadline lookup.

**4. Underscore-prefixed fields leak into CSV**
`models.py:740-742` exports all `custom_fields` entries without filtering. `_acf_code`, `_ev_purchase_price`, `_ice_purchase_price` appear as columns in exported CSVs. Fix: add `if not key.startswith("_")` guard.

**5. Present tab UX — template-centric redesign**
When a client template `.pptx` is provided, it should be the primary workflow entry point (choose template → select slides → export). Currently template selection is a small buried section. Prepare the tab so the template-based workflow is natural before the sample deck is ready.

### Low Priority

**6. EV matching overhead** — `match_fleet_ev_equivalents()` runs on every processing job even though EV columns are hidden from defaults (Phase 15). Consider making this opt-in.

**7. Vehicle database CSV bulk import** — `database_panel.py` "Import CSV" shows "Coming Soon."

**8. Vehicle database tests** — `data/vehicle_database.py` has no test coverage. Add `tests/test_vehicle_database.py` covering 4-tier lookup fallback, NULL-year UNIQUE handling, and `update_ice_vehicle` allow-list filtering.

**9. Module-level side effects** — `settings.py` creates directories on import; `utils.py` calls `setup_logging()` on import. Move to explicit init functions in `app.py` for cleaner testing.

---

## Phase History (Summary)

| Phase | Key Changes |
|-------|-------------|
| 1 | Core UX fixes: button states, debug logging removal, thread safety, dead code cleanup |
| 2 | Results table polish: number formatting, column widths, summary bar, persistent API cache |
| 3 | ACF compliance classification + electrification timeline (score-based v1) |
| 3.5 | Chart crash fixes, ACF false-positive prevention, filtered summary stats |
| 4 | Process panel rewrite, results toolbar cleanup, diesel MPG handling, match confidence display |
| 5 | Calculation fixes (ICE TCO, emissions overcount), analysis panel repairs, dead code cleanup |
| 6 | Crash fixes (Python 3.8 compat, CSV export, logger NameError), scraper MPG validation, UX polish |
| 7 | 177-test pytest suite (VIN, models, normalize, calculations, ACF, timeline, CSV mapping) |
| 8 | utils.py cleanup, widget extraction to ui/widgets.py |
| 9A-9J | Multi-year TCO model, EV matching, real data in PPTX, financial/exec slides, scenario engine, presentation-quality charts, unified analysis workflow, incentive/rate database, analysis-ready Excel export, present panel customization |
| 10 | Double-validation fix, configurable matching weights, ScrollableFrame consolidation |
| 11 | Live Excel formulas in TCO sheet, exposed battery degradation/residual value params, scenario selector in Present panel |
| 12 | Custom scenario UI in Present panel, figure size enforcement (FIG_SIZE_HALF/SLIDE), live TCO anchor cells |
| 13 | Thread-safety fixes in analysis methods, export parameter wiring, OS-specific file open, dead PDF code removal |
| 14 | GVWR column added to defaults, Customize Columns dialog redesign, Fuelly 403 hardening + fuzzy year fallback |
| 15 | MPG source tracking, EPA class-average fallback, MPG coverage badge, electrification timeline redesign (ACF-deadline-first for Cat B), ACF Relevance column, EV matching removed from default display |
| 16 | SQLite vehicle MPG database (4-tier lookup, CRUD panel, pipeline integration, right-click save from Results) |
| 17 | Budget smoothing bug fixed (N≤Y linear spacing + N>Y floor distribution) in electrification_timeline.py and scenarios.py; 3 regression tests added (180 total); scenario comparison wired into Analysis tab (checkboxes + emissions chart + vehicles-per-year table via PanedWindow); full dashboard redesign spec locked in |
| 18 | Analysis tab full dashboard redesign: replaced PanedWindow layout with top-to-bottom ScrollableFrame; sticky action bar; 4-chip Fleet Snapshot; two-column Top 5 Priority Vehicles + ACF Compliance donut row; Scenario Comparison with dedicated chart canvas; TCO KPI cards; collapsible Chart Browser (collapsed by default) + collapsible Parameters (expanded by default, auto-collapses post-analysis); updated copy_selection() to copy Top 5 table; removed summary_text widget and individual analysis UI buttons (methods kept for Tools menu) |
| 19 | Manual EV-year override + live Gantt timeline: new Timeline tab (`ui/timeline_panel.py`) with full-fleet table, double-click inline editing, amber row highlighting, Reset Selected/All; collapsible Gantt section in Analysis tab (grouped-by-ACF default, per-vehicle toggle, X=2026–2040); override write-back to `custom_fields` via `_apply_ev_year_override()`/`_reset_ev_year_override()`; "Timelines to Include" dialog at Excel export (scenario year columns + amber override flagging in Replacement Schedule sheet); `get_scenario_year_assignments()` helper in scenarios.py; notebook tab order updated (Timeline at index 3, Present at 4, Database at 5); 194 tests passing (+14 new) |
| 20 | Audit + polish pass: fixed File > Export Report routing through `analysis_panel.export_full_report()` (was broken duplicate path); Gantt X-axis now dynamic (extends beyond 2040 to include any override years); Gantt starts expanded in Analysis tab; "Timelines to Include" dialog defaults all 4 scenarios checked; Timeline Treeview click-to-sort on all columns + stretch on Make/Model/Body Type columns; "Proposed EV Year ✎" header signals editability; non-integer year input shows warning instead of silent discard; year validation ceiling raised to 2060 with helpful message; Reset Selected disabled when nothing selected; Export Excel/Build Presentation buttons disabled until fleet loaded; Analysis Gantt auto-refreshes on Analysis tab activation; configurable Max rows (25/50/100/All) for Per Vehicle Gantt view; Tools > Reset All EV Year Overrides menu item; status bar message for Present tab (tab 4); 194 tests still passing |
| 21 | Full audit + workflow integration: **Bugs fixed** — B1 `new_fleet()` now updates Present panel (`sharing_data` + `refresh_data()`); B2 `refresh_view()` handles tabs 3/4/5; B3 `copy_selection()` delegates to Timeline tab; B4 Documentation dialog updated to 6-tab workflow; B5 copyright year corrected to 2026. **Scenario improvements** — "Current Plan (Overrides)" as 5th scenario checkbox (pre-checked); scenarios auto-run at end of Run Full Analysis (no separate button press needed); CO₂ chart retained + new cumulative cost ($M) trajectory chart added; `_compute_current_plan_result()` module-level helper reads `Proposed EV Year` custom_fields directly. **Stale analysis indicator** — parameter changes after analysis turn the status label amber with warning text; `_on_param_changed()` trace method. **Timeline filter bar** — live search (make/model/asset), ACF category checkboxes, EV year range spinboxes, Clear Filters button; uses detach/reattach to preserve IID stability; `_all_iids` list. **Override impact summary** — expanded label shows "X overrides active — Y moved earlier, Z moved later". **Timeline copy** — `copy_selection()` public method added. **Results tab** — `select_by_vin(vin)` public method. **Jump to vehicle** — double-click Top 5 row → switch to Results tab + select that vehicle. **Analysis Gantt** — Max rows combobox added to Per Vehicle view. **Tools menu** — "Run Full Analysis" primary entry; three individual analyses moved to "Individual Analyses ▸" submenu. 194 tests still passing |
| 22 | ACF category usability + scope-based scenario comparison: **ACF filter labels** — Timeline tab filter bar checkboxes now show full plain-English names (Already ZEV / Light-Duty (Exempt) / Mandate-Subject / Body-Type Exempt / Emergency Vehicle) instead of raw letter codes. **ACF category override** — double-click ACF column in Timeline tab → Combobox dropdown; right-click any row in Results tab → "Override ACF Category..."; both paths show `_show_acf_ev_year_dialog()` offering Recalculate / Keep / Cancel for the EV year; `_apply_acf_override()` + `_reset_acf_override()` static methods on AnalysisPanel; overridden rows tinted light blue (#E3F2FD) / light purple (#F3E8FD) for both types; Reset Selected + Reset All clear ACF overrides too; `on_acf_changed` callback wired through to `main_window._on_acf_override()`. **Scope-based scenario comparison** — Analysis tab Scenario Comparison section replaces the 4 time-based presets with 3 scope-based presets: Minimum Compliance (Cat B only) / All Excl. Emergency (Cat A+B+C) / Whole Fleet (Cat A+B+C+D); Horizon spinbox (2040 default, 2026–2060) controls the end year for all scope scenarios; `SCOPE_SCENARIO_KEYS` exported from scenarios.py; `ElectrificationScenario` gains `include_light_duty: bool` field; old 4 time-based presets retained for Present tab + PowerPoint + Excel export. 194 tests still passing |
| 23 | PowerPoint template-driven rebuild: **Architecture** — `export_presentation()` in `powerpoint_export.py` replaces all 14 old slide builders; loads example PPTX as base template (inheriting backgrounds, layouts, fonts), modifies in-place, deletes unchecked slides in reverse order, then saves. **New infrastructure** — `assets/template_default.pptx` (template copy); `PresentationProfile` dataclass (`data/models.py`) with client/presenter/partner/content/slide fields; per-fleet sidecar JSON (`{stem}_profile.json`) via `load_presentation_profile()` / `save_presentation_profile()` in `data/processor.py`; `settings.py` gains `ASSETS_DIR`, `DEFAULT_TEMPLATE_PATH`, `PROFILE_SIDECAR_SUFFIX`, `TEMPLATE_SLIDE_IDS`, `DEFAULT_SLIDE_IDS`. **Export engine** — token replacement on cover/contact slides; auto-generated Key Findings bullets (fleet size, EV-replaceable %, ACF exemptions, timeline achievability, Cat B urgency); new ACF Electrification Timeline stacked column chart (ACF cat × year); new GHG Emissions line chart (yearly remaining ICE emissions per scenario); slide deletion via `prs.slides._sldIdLst` XML; `export_pdf()` with LibreOffice → pptx2pdf → macOS fallback. **Present tab redesign** (`ui/present_panel.py`) — Client Profile form (all PresentationProfile fields, auto-dirty tracking); draggable Treeview slide list with `☑`/`☐` checkboxes + Move Up/Down buttons; Consulting Content text areas (agenda, data needs, next steps); Template browser; Build + Export PDF buttons; profile auto-loaded from sidecar when fleet CSV opens. **main_window.py** — `on_fleet_loaded` calls `load_presentation_profile()` → pushes to `sharing_data["presentation_profile"]` → `present_panel.load_profile()`. 194 tests still passing |
| 24 | PowerPoint Quality Overhaul: **Chart fixes** — `add_tco_comparison_chart()` changed from COLUMN_STACKED → COLUMN_CLUSTERED (ICE vs EV side-by-side); all 4 ACF series now rendered in timeline chart with plain-English names matching user-built deck ("Light Duty (Excluded)", "Medium or Heavy Duty", etc.); year cap at 2040; GHG emissions chart rewritten to 3 series (Baseline flat, M/H Duty Only, Whole Fleet). **New chart builders** — `add_acf_category_composition_chart()` (pie by ACF category, replaces body-type pie); `add_department_summary_chart()` (horizontal stacked bar, conditional on dept data); `add_facility_summary_chart()` (horizontal stacked bar, conditional on location data). **Smart Key Findings** — `_generate_key_findings()` rewrote with narrative conditional bullets: fleet size, department/facility context (if data present), ZEV vehicles (if any), ACF exemption breakdown, mandate urgency with earliest year, achievability statement. **New conditional slide builders** — `_create_acf_composition_slide()`, `_create_scenario_timeline_slide()` (one slide per scenario: ACF Compliance, Moderate 2035, Aggressive 2030, Conservative 2040, Current Plan), `_create_invalid_vin_slide()` (auto-included when VINs failed decoding), `_create_department_summary_slide()`, `_create_facility_summary_slide()`. **Smart positioning** — `_move_slide_to_index()` utility + `_INSERT_BEFORE` dict places optional slides mid-deck rather than always appending at end. **Present tab** — expanded `_OPTIONAL_SLIDE_META` with 9 new entries (4 pre-checked by default: acf_composition, timeline_moderate, timeline_current_plan, invalid_vin); fixed `_apply_profile_slides()` separator row preservation bug. 194 tests still passing |
| 25 | PowerPoint audit + bug fixes: **Bug A** — `_reorder_slides()` positional mapping corrected; was mapping TEMPLATE_SLIDE_IDS[i]→items[i] after deletions (wrong whenever any template slides excluded), now maps against `remaining_ids` (TEMPLATE_SLIDE_IDS filtered to included set). **Bug B** — removed `if any(v>0)` guard from `_add_acf_electrification_chart()`; all 4 ACF series always added for consistent legend across scenario slides. **Bug C** — `_create_scenario_timeline_slide()` now propagates the chart function's return value instead of always returning True. **Bug D** — `APP_VERSION` bumped to 3.0.8. **Quality** — GHG b_values fallback restructured for clarity; TCO docstring fixed ("clustered column" not "stacked bar"); unused imports (`Union`, `Tuple`, `ChartData`, `datetime`, `timedelta`) removed from powerpoint_charts.py; Key Findings gains EV-replaceable % bullet ("X of Y vehicles (Z%) mandate-subject") and urgency bullet reworded to earliest-year + procurement call-to-action; GHG chart gains "Metric Tons CO₂e" value-axis title; TCO chart gains "Annual Cost ($)" value-axis title; ACF timeline stacked column chart enables per-series data labels with zero suppression (`#,##0;;` numFmt). 194 tests still passing |
