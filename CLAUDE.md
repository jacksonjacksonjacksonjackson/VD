# CLAUDE.md — Fleet Electrification Analyzer

## Project Overview

Desktop Python/Tkinter app for fleet electrification consulting. Decodes vehicle VINs, retrieves MPG data from government APIs, classifies vehicles for CARB ACF compliance, assigns electrification timelines, and generates analysis + client deliverables.

**Stack:** Python 3.8+, Tkinter, pandas, requests, python-pptx, matplotlib, openpyxl, BeautifulSoup/Selenium
**Entry point:** `python3 app.py` (GUI) or `python3 app.py -b -i in.csv -o out.csv` (batch)
**Current version:** 3.0.11 on branch `uiux/v3_0_3`

---

## Feature Status

| Feature | Status | Notes |
|---------|--------|-------|
| VIN pipeline (Process tab) | ✓ Working | Core workflow, 279 tests passing |
| Results table | ✓ Working | Right-click menu, ACF override, MPG DB save |
| Analysis dashboard | ✓ Working | Charts, KPIs, scenario comparison, Gantt |
| ACF classification | ✓ Working | ZEV/A/B/C/D; HPF/Non-HPF/State Agency support |
| Electrification timeline | ✓ Working | Phase 28 even-spread w/ GVWR-tiered boost |
| Excel export (cover + 6) | ✓ Working | v2 (Phase 29): Cover-linked assumptions, incentives, capital plan, scenarios, action items |
| PowerPoint export | ⚠️ Known issues | Charts occasionally wrong/incomplete; template quality variable |
| Timeline tab | ✓ Working (low use) | Inline editing works; rarely needed in practice |
| Charging tab | 🚧 Stub | UI scaffolding present; no analysis engine yet |
| Database tab | ✓ Working | CRUD + CSV bulk import (Phase 26) |

---

## Build & Run

```bash
pip install -r requirements.txt
python3 app.py              # GUI mode
python3 app.py --help       # CLI options
python3 app.py -b -i in.csv -o out.csv  # Batch mode
```

### Testing

```bash
python3 -m pytest tests/ -v           # Run all 279 tests
python3 -m pytest tests/ -v -k acf    # ACF tests only
python3 -m pytest tests/ -v --tb=short
```

---

## Architecture

```
app.py                  → Entry point, CLI args, window setup
settings.py             → All configuration constants (ASSETS_DIR, TEMPLATE_SLIDE_IDS, etc.)
utils.py                → Shared utilities, VIN validation, Cache

data/
  models.py             → Dataclasses (VehicleIdentification, FuelEconomyData, FleetVehicle,
                          Fleet, PresentationProfile)
  providers.py          → API clients (NHTSA VIN decoder, FuelEconomy.gov)
  processor.py          → CSV I/O, ProcessingPipeline, BatchProcessor,
                          load/save_presentation_profile()
  vehicle_database.py   → SQLite vehicle MPG reference database (analyst-maintained)

ui/
  main_window.py        → Main window, 7-tab notebook, sharing_data SafeDict, all callbacks
  theme.py              → Centralized styling
  widgets.py            → Reusable UI widgets (tooltips, status bar, progress dialog)
  process_panel.py      → Tab 0: VIN processing, batch input
  results_panel.py      → Tab 1: Results table, column customization, right-click menu
  analysis_panel.py     → Tab 2: Analysis dashboard (KPIs, scenarios, Chart Gallery, Gantt)
  timeline_panel.py     → Tab 3: Full-fleet Treeview with inline EV year + ACF editing
  present_panel.py      → Tab 4: PowerPoint export, client profile form, card gallery
  charging_panel.py     → Tab 5: Charging params UI (stub — no analysis engine)
  database_panel.py     → Tab 6: Vehicle MPG reference DB CRUD + CSV bulk import

analysis/
  calculations.py       → TCO, ROI, emissions, year-by-year cash flow model
  charts.py             → Matplotlib chart generation (22 chart types via ChartFactory)
  reports.py            → Excel report generation (v2: cover + 6 worksheets)
  acf_compliance.py     → CARB ACF classification per vehicle (ZEV/A/B/C/D)
  electrification_timeline.py → Score-based even-spread year assignment + GVWR-tiered boost
  ev_database.py        → EV equivalent matching (~25 entries, reliability limited for heavy-duty)
  rate_database.py      → State energy rates + federal/state incentive programs
  scenarios.py          → ElectrificationScenario dataclass + 7 presets + compare_scenarios()

assets/
  template_default.pptx      → Default PPTX template (Avenir LT Std Book body font)
  slide_thumbnails/           → Optional PNG thumbnails for Present tab card gallery (may be empty)

powerpoint_export.py           → PPTX generation engine (template-driven, 15 slide builders)
powerpoint_charts.py           → Native PPTX chart creation (9 chart types)
powerpoint_customizer.py       → PPTX customization helpers
commercial_vehicle_scraper.py  → Web scraping for commercial vehicle specs
scripts/
  update_template_font.py     → Patches theme minorFont to "Avenir LT Std Book" in template

tests/
  conftest.py                      → Factory helpers & pytest fixtures
  test_vin_validation.py
  test_models.py
  test_normalize.py
  test_calculations.py
  test_acf_compliance.py
  test_electrification_timeline.py
  test_csv_mapping.py
  test_match_confidence.py
  test_timeline_override.py
  test_pptx_smoke.py               → 13 headless PPTX assertions (Phase 26)
  test_excel_report.py             → Excel report v2 structural assertions (Phase 29)
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
  → Electrification year assignment (score-based even-spread for all schedulable cats)
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

ACF codes: `custom_fields["_acf_code"]` = letter code (use for logic); `custom_fields["ACF Category"]` = human-readable label (use for display only).

### Electrification Timeline (Phase 28)
All schedulable vehicles (B, C, D) use a unified score-based even-spread queue:
- **Score formula:** `0.55*age + 0.25*odometer + 0.10*annual_mileage + acf_boost`
- **Cat B GVWR-tiered boost:** Class 2b-4 → +0.35, Class 5-8a → +0.25, Class 8b → +0.15, Unknown → +0.30
- **Cat C/D boost:** `ACF_BOOST["C"] = 0.10`, `ACF_BOOST["D"] = 0.05`
- **Distribution:** N ≤ num_years → linear spacing (`idx = round(k*(num_years-1)/(n-1))`); N > num_years → floor-based (`base = n//num_years`, `extra = n%num_years`)
- **CARB deadline (reference only):** `custom_fields["ACF Deadline Year"]` — HPF/non-HPF milestone year, used by Milestone chart + compliance warnings; does NOT drive `Proposed EV Year`
- **Budget cap:** `Fleet.max_vehicles_per_year > 0` → greedy-fill left-to-right cap per year
- **Fleet type:** `Fleet.fleet_type` ∈ `("hpf", "non_hpf", "state_agency")` — controls deadline table for `ACF Deadline Year` reference field
- **N/A reason codes:** "Already ZEV," "Processing failed," "ACF classification unavailable"

### Notebook Tab Order

| Index | Tab | Key File |
|-------|-----|----------|
| 0 | Process | `ui/process_panel.py` |
| 1 | Results | `ui/results_panel.py` |
| 2 | Analysis | `ui/analysis_panel.py` |
| 3 | Timeline | `ui/timeline_panel.py` |
| 4 | Present | `ui/present_panel.py` |
| 5 | Charging | `ui/charging_panel.py` (stub) |
| 6 | Database | `ui/database_panel.py` |

Navigation helpers: `_navigate_to_present()` → `notebook.select(4)`; `_navigate_to_timeline()` → `notebook.select(3)`. `on_tab_changed()` handles indices 3, 5, 6.

### Analysis Tab Layout (Phase 27/28)
Top-to-bottom `ScrollableFrame` with sticky action bar outside scroll:
1. **Sticky bar:** `⚡ Run Full Analysis` + status label (amber when stale)
2. **Parameters** (collapsible, starts expanded; auto-collapses post-analysis) — Fleet type radios, Max Annual Replacements spinbox, financial params
3. **Fleet Snapshot** — 4 KPI chips (Fleet Size, ACF-B Count, Avg MPG, MPG Coverage %)
4. **Scenario Comparison** — 3 scope checkboxes (Min Compliance / All Excl. Emergency / Whole Fleet) + "Current Plan (Overrides)" + Horizon spinbox (2040 default); dual charts: CO₂ trajectory + cumulative cost ($M); vehicles-per-year Treeview; auto-runs after analysis
5. **Chart Gallery** (collapsible, starts collapsed) — main canvas + thumbnail strip; "Include in PPTX" checkbox wired to `sharing_data["selected_chart_ids"]`
6. **Electrification Timeline Gantt** (collapsible, starts expanded) — 3 views: Grouped by ACF / Per Vehicle / By Scenario; Max rows combobox; auto-refreshes on tab activation
7. **TCO Summary** — 3 KPI cards (Annual Savings, Payback Period, Infrastructure Est.)
8. **Top 5 Priority / ACF Donut** (collapsible, starts collapsed)
9. **Export bar** — Export Excel Report (triggers Timelines dialog) + Build Presentation; disabled until fleet loaded

Override helpers (static on `AnalysisPanel`): `_apply_ev_year_override()`, `_reset_ev_year_override()`, `_apply_acf_override()`, `_reset_acf_override()`, `_reset_all_overrides()`.
Module-level: `ACF_LABELS` dict, `_show_acf_ev_year_dialog()`, `_draw_gantt_chart()`, `_gantt_grouped()`, `_gantt_per_vehicle()`, `_gantt_by_scenario()`.

### PowerPoint Engine (Phase 27/28)

**Template slides (15):** cover, agenda, carb_overview, acf_scenarios, acf_exemptions, key_findings, timeline_chart, emissions_chart, incentives, data_needs, next_steps, contact, appendix, infra_costs_chart, tco_chart.

**Optional slides:** acf_composition (pie), timeline_moderate/aggressive/conservative/current_plan (scenario timelines), timeline_milestone, invalid_vin, department_summary, facility_summary, age_analysis, scenario_comparison, replacement_table, data_quality, scenario_co2, scenario_investment.

**Key chart builders:**
- `_add_purchase_schedule_chart()` — ZEV Purchase Schedule: 4 series (A=blue/B=orange/C=yellow/D=grey), fixed 2026–2040 x-axis; used by `timeline_chart` slide
- `_add_milestone_option_chart()` — Cat B split by GVWR class; optional `timeline_milestone` slide
- `add_co2_trajectory_chart()` — LINE_MARKERS: baseline + 3 scope scenarios; used by `emissions_chart` slide (fallback: 3-series when no scenario data)
- `add_tco_comparison_chart()` — COLUMN_CLUSTERED: ICE vs EV side-by-side; axis title "Annual Cost ($)"
- `add_acf_category_composition_chart()` — pie by ACF category
- `_create_scenario_timeline_slide()` — one slide per time-based scenario

**Data flow:** `scenario_results` pushed to `sharing_data` after `_run_scenario_comparison()`; `present_panel._on_build()` reads and passes to `export_presentation()`.

**Slide profile:** `PresentationProfile` dataclass in `data/models.py`; saved/loaded as `{fleet_stem}_profile.json` sidecar via `data/processor.py`.

### PowerPoint Template Guide

Template file: `assets/template_default.pptx`. If building a new/improved template:
- **Title placeholder** (type 15 or 13) required on every slide — engine uses `_find_placeholder()` to locate and replace text
- **Body placeholder** (type 14 or 1) for bullet content — cleared by `_clear_body_placeholders()` to remove "Click to add text"
- **Do NOT embed charts** — leave chart areas as empty placeholders or blank space; all charts are generated programmatically from data
- **Token replacement:** `{{CLIENT_NAME}}` (cover), `{{PRESENTER}}` and `{{PARTNER}}` and `{{DATE}}` (contact slide)
- **Body font:** Avenir LT Std Book (applied via `scripts/update_template_font.py`; run once after building new template)
- **Backgrounds/branding:** safe to use any background image or brand colors — engine only modifies text frames and injects chart shapes

### Scenario Engine

`analysis/scenarios.py` — 7 presets total:
- **4 time-based** (Aggressive 2030, Moderate 2035, Conservative 2040, ACF Compliance Only 2035) — used by Present tab + PowerPoint + Excel export
- **3 scope-based** (Minimum Compliance, All Excl. Emergency, Whole Fleet) — used by Analysis tab Scenario Comparison; exported as `SCOPE_SCENARIO_KEYS`

`ElectrificationScenario` dataclass has `include_light_duty: bool` field. `compare_scenarios()` returns per-year vehicles/cost/co2/savings. Scope presets run to user-configured horizon year via `dataclasses.replace()`.

### Excel Export (v2 — Phase 29: cover + 6 sheets)

- **Results tab "Export"** → single-sheet "Fleet Analysis" (raw vehicle data only)
- **Analysis tab "Export" or File > Export Report** → **cover + 6-sheet** report:
  0. **Cover & Methodology** — client header (from `PresentationProfile`), **canonical single-source-of-truth assumptions block** (`'Cover & Methodology'!$B$9:$B$18`, keyed by `_ASSUMPTION_ROW` / `cover_aref()`), data-quality KPIs, methodology + data-source note, ACF glossary, disclaimer
  1. **Vehicle Data** — flat per-VIN table with quality-flag highlighting (MPG estimate amber, low quality/failed red, low match confidence)
  2. **Fleet Overview** — merged old Summary + Summary Dashboard: KPI band + ACF/fuel/make composition tables with native pie/column charts
  3. **TCO & Financials** — live cash-flow model (assumption cells B4–B12/B14 mirror the Cover via formula), Funding & Incentives block (`get_all_incentives()`, light/medium_heavy buckets), per-vehicle savings (real `_payback_years`, no $15k hack)
  4. **Replacement & Capital Plan** — Gantt schedule + per-fiscal-year capital plan (vehicle+infra capex, net-of-incentive, cumulative, over-cap flag) + native spend chart
  5. **Emissions & Scenarios** — emissions inventory + 4-scenario comparison (`compare_scenarios`) + native CO₂ trajectory chart; projected emissions flagged when `is_synthetic`
  6. **Infrastructure & Action Items** — charging summary + per-year buildout + Data Gaps punch list
- `generate()` takes `client_profile` + `state_code` (default `"CA"`); wired from `analysis_panel.export_full_report()` via `sharing_data["presentation_profile"]`.
- **Timelines to Include dialog** (triggered at `.xlsx` export): 4 scenario checkboxes (all checked by default) → additional year columns in Replacement schedule + amber override flagging
- Structural tests in `tests/test_excel_report.py`.

---

## Active Backlog

### High Priority

**1. PowerPoint chart data quality**
Charts occasionally show wrong or incomplete data (wrong series, bad numbers). No systematic regression tests for optional slides or custom-template workflows. Needs investigation per-slide with a real fleet dataset.

**2. PowerPoint visual quality / template**
Visual output quality is variable depending on template. User wants to build a better template — see Template Guide above. After new template is built, audit all slide builders for layout consistency and data accuracy.

### Medium Priority

**3. No slide thumbnail assets**
`assets/slide_thumbnails/` is empty. The Present tab card gallery falls back to colored placeholder boxes. Generating real PNG thumbnails would significantly improve the slide selection UX.

**4. Present tab UX — template-centric workflow**
Template selection is still secondary in the UI. When a client `.pptx` template is loaded, it should immediately be the primary workflow context (slide list updates from template, client profile visible first).

**5. EV matching overhead**
`match_fleet_ev_equivalents()` runs on every processing job even though EV columns are hidden from default display. Consider making this opt-in via a checkbox.

### Low Priority

**6. Vehicle database tests**
`data/vehicle_database.py` has no test coverage. Add `tests/test_vehicle_database.py` covering 4-tier lookup fallback, NULL-year UNIQUE handling, and `update_ice_vehicle` allow-list filtering.

**7. Module-level side effects**
`settings.py` creates directories on import; `utils.py` calls `setup_logging()` on import. Move to explicit init functions in `app.py` for cleaner testing.

**8. Charging tab implementation**
`ui/charging_panel.py` has UI scaffolding (power levels, pattern, window params) but no analysis engine. Implement charging demand calculation: vehicles scheduled per year × charger type → infrastructure cost estimate.

---

## Phase History

**Phases 1–25 (2024–early 2025):** Built the core VIN processing pipeline (NHTSA + FuelEconomy.gov + commercial scraper + SQLite MPG DB + EPA fallback), ACF compliance classification (ZEV/A/B/C/D), multi-year TCO/ROI/emissions model, score-based electrification timeline with budget smoothing, multi-sheet Excel export with live formulas, CARB scenario engine (7 presets), SQLite vehicle MPG reference DB with CRUD panel (Phase 16), Analysis tab dashboard redesign with ScrollableFrame + KPI chips (Phase 18), Timeline tab with inline EV year override editing (Phase 19), Excel "Timelines to Include" dialog with override flagging (Phase 20), ACF category override system with right-click + Timeline inline editing (Phase 22), PowerPoint template-driven architecture rebuild with `export_presentation()` engine (Phase 23), chart quality overhaul + conditional optional slides + smart Key Findings (Phase 24), PPTX bug fixes + `test_pptx_smoke.py` + HPF/Non-HPF toggle + DB CSV import (Phases 25–26).

| Phase | Key Changes |
|-------|-------------|
| 26 | PPTX smoke test (13 assertions). Underscore-field CSV guard. HPF/Non-HPF/State Agency fleet type toggle (`ACF_DEADLINE_TABLE_NON_HPF`, `fleet_type` param, 3 radio buttons in Analysis tab). DB CSV bulk import implemented (~115 lines, flexible header aliasing). 220 tests passing. |
| 27 | ACF display bug fixes (donut, Gantt, Timeline filter all now use `_acf_code` not label). Chart kwargs crash fix. Analysis tab layout reorder (Parameters first). Chart Gallery replaces Chart Browser (main canvas + thumbnail strip, PIL, "Include in PPTX" checkbox). Multi-scenario Gantt view (`_gantt_by_scenario()`). PPTX: `add_co2_trajectory_chart()` + `add_cumulative_investment_chart()`; `scenario_results` wired through `sharing_data`. Present tab: Treeview replaced with 2-column card gallery. Charging tab added (Tab 5, stub). Database shifts to Tab 6. 220 tests passing. |
| 28 | Cat B even-spread: all Cat B now use score queue with GVWR-tiered boost constants (0.35/0.25/0.15/0.30); CARB deadline stored in `custom_fields["ACF Deadline Year"]` for reference only. New PPTX chart builders: `_add_purchase_schedule_chart()` (4-series A/B/C/D) + `_add_milestone_option_chart()` (Cat B split by GVWR); new optional `timeline_milestone` slide. GHG fix: `_scenario_baseline_co2` now computed via `_calculate_baseline_emissions()` (was always 0.0). `_clear_body_placeholders()` on optional slides. Font patched to Avenir LT Std Book via `scripts/update_template_font.py`. `Fleet.max_vehicles_per_year` + "Max Annual Replacements" spinbox. 279 tests passing (+6). |
| 29 | **Excel report v2 rebuild** (`analysis/reports.py`, branch `feature/excel-report-v2`): 8 loosely-coupled sheets → **cover + 6**. New **Cover & Methodology** sheet holds the canonical assumptions block (`cover_aref()` / `_ASSUMPTION_ROW`); TCO cells B4–B12/B14 mirror it by formula (single source of truth). Merged Summary+Dashboard → **Fleet Overview** (native charts); merged Electrification+TCO → **TCO & Financials** (+ Funding & Incentives via `get_all_incentives()`; **fixed the hardcoded $15k EV-premium payback bug** → real `_payback_years`); **Replacement & Capital Plan** (per-year capex, net-of-incentive, over-cap flag); **Emissions & Scenarios** (`compare_scenarios` table + CO₂ trajectory chart, `is_synthetic` flag); **Infrastructure & Action Items** (per-year buildout + data-gap punch list). Vehicle Data gained quality-flag highlighting. `generate()` gained `client_profile` + `state_code`. Removed dead `_create_summary_sheet` / `_create_summary_dashboard_sheet` / `_create_electrification_sheet`. `tests/test_excel_report.py` (+9). 279 tests passing. |

---

## Next-Session Prompt

Copy and paste this to start a new development session:

```
I'm continuing development of my Fleet Electrification Analyzer app.
Branch: uiux/v3_0_3 | Version: see settings.py APP_VERSION
Review CLAUDE.md fully before starting — it is the single source of truth.
Run: python3 -m pytest tests/ -v (expect 279 passing)
Focus this session: [DESCRIBE TASK]
```
