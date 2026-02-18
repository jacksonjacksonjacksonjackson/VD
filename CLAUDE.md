# CLAUDE.md - Fleet Electrification Analyzer

## Project Overview

A desktop Python application (Tkinter) that decodes vehicle VINs, retrieves fuel economy data from government APIs, and generates fleet electrification analysis reports. The core value is automating MPG lookups that previously required hours of manual searching.

**Stack:** Python 3.8+, Tkinter, pandas, requests, python-pptx, matplotlib, BeautifulSoup/Selenium

**Entry point:** `python app.py` (GUI mode) or `python app.py --batch -i input.csv -o output.csv`

**Current version:** 3.0.3 on branch `uiux/v3_0_3`

---

## Build & Run

```bash
pip install -r requirements.txt
python app.py              # GUI mode
python app.py --help       # Show CLI options
python app.py -b -i in.csv -o out.csv  # Batch mode
```

### Testing

```bash
python -m pytest tests/ -v           # Run all 151 tests
python -m pytest tests/ -v -k acf    # Run only ACF compliance tests
python -m pytest tests/ -v --tb=short  # Shorter tracebacks
```

Test suite covers pure logic functions (no API mocking needed). See [Phase 7](#phase-7-automated-test-suite) for details.

---

## Architecture

```
app.py                  -> Entry point, CLI args, window setup
settings.py             -> All configuration constants
utils.py                -> Shared utilities, UI widgets, VIN validation

data/
  models.py             -> Dataclasses (VehicleIdentification, FuelEconomyData, FleetVehicle, Fleet)
  providers.py           -> API clients (NHTSA VIN decoder, FuelEconomy.gov)
  processor.py           -> CSV I/O, ProcessingPipeline, BatchProcessor

ui/
  main_window.py         -> Main window with tabbed interface
  theme.py               -> Centralized styling
  widgets.py             -> Reusable UI widgets (tooltips, status bar, progress dialog, error dialogs)
  process_panel.py       -> VIN processing tab
  results_panel.py       -> Results display tab
  analysis_panel.py      -> Analysis tools tab
  present_panel.py       -> PowerPoint generation tab

analysis/
  calculations.py        -> TCO, ROI, emissions calculations
  charts.py              -> Matplotlib chart generation
  reports.py             -> Report formatting
  acf_compliance.py      -> CARB ACF compliance classification per vehicle
  electrification_timeline.py -> Fleet-wide electrification year assignment

commercial_vehicle_scraper.py  -> Web scraping for commercial vehicle specs
powerpoint_export.py           -> PPTX generation engine
powerpoint_customizer.py       -> PPTX customization
powerpoint_charts.py           -> PPTX chart creation

tests/
  conftest.py              -> Factory helpers & pytest fixtures
  test_vin_validation.py   -> VIN format validation tests
  test_models.py           -> GVWR, commercial, diesel, quality tests
  test_normalize.py        -> Vehicle model normalization tests
  test_calculations.py     -> TCO, emissions, fuel cost tests
  test_acf_compliance.py   -> ACF classification tests
  test_electrification_timeline.py -> Timeline scoring & year assignment tests
  test_csv_mapping.py      -> Column mapping, VIN detection, export tests
```

**Data flow:** CSV input -> VIN validation -> NHTSA API (decode VIN) -> FuelEconomy.gov API (match MPG) -> quality scoring -> ACF classification -> electrification timeline assignment -> FleetVehicle objects -> UI/export

---

## Recommendations for Improvement

The recommendations below are ordered roughly by impact and risk. Items near the top are high-value and low-risk; items at the bottom are larger refactors.

### ~~1. Remove Excessive Diagnostic Logging (High Priority, Low Risk)~~ — DONE (Phase 1, Fix 2)

Removed ~250 lines of `DIAGNOSTIC:` and `DEBUG:` logging from `providers.py`, `processor.py`, `models.py`, `main_window.py`, `process_panel.py`. Deleted `diagnostic_logger` stdout handler from `providers.py`.

### ~~2. Eliminate Duplicated Logic (High Priority, Low Risk)~~ — PARTIALLY DONE (Phase 1, Fix 8)

Dead `utils.py` duplicates removed (`extract_gvwr_pounds`, `detect_commercial_vehicle`, `detect_diesel_engine`, `classify_commercial_category`, `get_commercial_summary`, `extract_engine_power`). Remaining minor duplication:
- **Column mapping** (`_map_additional_columns`) still in both `CsvFileValidator` and `ProcessingPipeline`
- **VIN column detection** deprecated `CsvReader._find_vin_column` wrapper still exists in `processor.py:486`

### 3. Fix the Double-Validation Problem (Medium Priority, Low Risk)

`CsvReader.read_vins()` calls `self.validate_file()` which runs the full `CsvFileValidator.validate_and_preview()` pipeline. But the UI also calls validation separately before processing starts. This means every CSV is validated twice — once for preview and once when processing begins. The validation should happen once, and the result should be passed to the reader.

### ~~4. Add a Proper Test Suite (High Priority, Medium Risk)~~ — DONE (Phase 7)

151 pytest tests covering VIN validation, GVWR parsing, model normalization, quality scoring, CSV column mapping, TCO/ROI calculations, ACF compliance, electrification timeline, and emissions inventory. `diagnostic_test.py` still exists but is superseded.

### 5. Fix the Batch Mode Wait Loop (Medium Priority, Low Risk)

In `app.py:117-121`, batch mode uses a busy-wait polling loop:

```python
while processor.current_pipeline and processor.current_pipeline.processing_thread and processor.current_pipeline.processing_thread.is_alive():
    time.sleep(0.1)
```

This is fragile. `BatchProcessor` doesn't expose a `processing_thread` attribute — it starts a daemon thread in `process_file()` but stores no reference. This code likely doesn't work as intended. Replace with `threading.Event` that's set when processing completes.

### ~~6. Clean Up `utils.py` Responsibilities (Medium Priority, Low Risk)~~ — DONE (Phase 8)

UI widgets (`SimpleTooltip`, `StatusBar`, `ProgressDialog`, `ScrollableFrame`) and error dialog classes (`ErrorCommunicator`, `ContextHelp`) extracted to `ui/widgets.py`. `utils.py` re-exports them for backward compatibility. Commercial vehicle detection functions were already removed in Phase 1 (Fix 8). `utils.py` reduced from ~1,040 to ~630 lines.

### ~~7. Handle Diesel Vehicle MPG Better (Medium Priority, Medium Risk)~~ — DONE (Phase 4, Fix 29)

Added `is_diesel` parameter to `FuelEconomyClient.find_vehicle_matches()` with fuel-type filtering, gasoline fallback with warning, fuel type mismatch detection, and 15-point confidence penalty. Mismatch surfaced in results as "Gas proxy (diesel data unavailable)".

### ~~8. Make the Cache Persistent (Medium Priority, Medium Risk)~~ — DONE (Phase 2, Fix 15)

`Cache` class extended with JSON disk persistence at `data/cache/api_cache.json`. Type-aware serialization round-trips dataclass objects. Loads on startup, saves after each processing run.

### 9. Remove Dead/Incomplete Features (Low Priority, Low Risk) — MOSTLY DONE

PDF/JSON/HTML removed from `EXPORT_FORMATS` (Phase 5, Fix 40). `CACHE_CONFIG` deleted from `settings.py` (Phase 8 — unused; actual Cache class uses `CACHE_EXPIRY` and `DEFAULT_CACHE_FILE`). `CsvReader._find_vin_column` deprecated wrapper deleted from `processor.py` (Phase 8). `diagnostic_test.py` deleted (Phase 8 — superseded by pytest suite).

**Intentionally kept:**
- `MATCHING_WEIGHTS` and `MIN_MATCH_CONFIDENCE` in `settings.py` — not currently wired into the matching code, but represent a valuable future improvement (see Recommendation #14 below). The matching logic in `providers.py` currently uses hardcoded point values; these config entries define a more configurable design that should be implemented.
- `ScrollableFrame` in `ui/widgets.py` — not currently instantiated, but the exact same scrollable-canvas pattern is reimplemented ad-hoc in 4+ places (`main_window.py` column dialog, `analysis_panel.py` left panel, `present_panel.py` slide options, `theme.py` helper). Worth consolidating these callers to use `ScrollableFrame` instead.

### ~~10. Improve Error Handling in `commercial_vehicle_scraper.py` (Low Priority, Low Risk)~~ — DONE (Phase 1, Fix 5)

Selenium imports guarded with try/except and `SELENIUM_AVAILABLE` flag, matching the existing `pdfplumber` pattern.

### ~~11. Address Threading Safety in UI Callbacks (Medium Priority, High Risk)~~ — DONE (Phase 1, Fix 3)

`log_callback` and `progress_callback` wrapped in `root.after()` in `main_window.py`. `done_callback` already marshalled correctly.

### 12. Reduce Module-Level Side Effects (Low Priority, Low Risk)

- `settings.py` creates directories on import (lines 34-35). If `settings.py` is imported for any reason (testing, tooling), directories are created as a side effect.
- `utils.py` calls `setup_logging()` at module level (line 66), creating a log file on import.

Move these to explicit initialization functions called from `app.py`.

### 13. Consider Packaging for Distribution (Low Priority, High Risk)

The app uses `sys.path.insert(0, ...)` in `app.py` to fix imports. A proper `pyproject.toml` or `setup.py` with package structure would make imports clean and enable `pip install -e .` for development.

### 14. Wire Up Configurable Matching Weights (Medium Priority, Medium Risk)

`MATCHING_WEIGHTS` and `MIN_MATCH_CONFIDENCE` exist in `settings.py` but the matching code in `providers.py` (`FuelEconomyClient.find_vehicle_matches()`, around line 860) uses hardcoded point values instead. The config defines a richer weighted scheme (exact_vin: 100, year_make_model: 80, engine_match: 20, displacement: 15, cylinders: 10, fuel_type: 10, drive: 5, transmission: 5) that would make matching tunable without code changes. Implementation would involve:

- Refactoring the confidence calculation in `providers.py` to read weights from `MATCHING_WEIGHTS`
- Applying `MIN_MATCH_CONFIDENCE` as a threshold to reject low-quality matches
- Exposing these settings in the Preferences UI for advanced users

### 15. Consolidate Scrollable Frame Implementations (Low Priority, Low Risk)

`ui/widgets.py` has a reusable `ScrollableFrame` class, but the codebase has 4+ ad-hoc reimplementations of the same canvas-scrollbar-inner-frame pattern:
- `main_window.py` `customize_columns()` (column selection dialog)
- `analysis_panel.py` `_create_scrollable_left_panel()`
- `present_panel.py` `_create_scrollable_container()` and slide options area
- `theme.py` `create_scrollable_frame()` helper

These should be migrated to use `ScrollableFrame` for consistency and to reduce duplicated scrolling logic.

---

## Known Limitations

- **Diesel MPG:** FuelEconomy.gov coverage for diesel variants is inconsistent. Fix 29 added diesel-specific filtering and fallback detection, but some diesel vehicles still match to gasoline trims (now flagged with "Fuel Type Mismatch" column)
- ~~**No automated tests:**~~ Fixed in Phase 7 — 151 pytest tests covering core logic
- ~~**In-memory cache only:**~~ Fixed in Phase 2 (Fix 15) — persistent JSON cache at `data/cache/api_cache.json`
- ~~**Selenium dependency:**~~ Fixed in Phase 1 (Fix 5) — guarded with try/except and `SELENIUM_AVAILABLE` flag
- ~~**Thread safety:**~~ Fixed in Phase 1 (Fix 3) — log/progress callbacks marshalled via `root.after()`
- **Large fleets:** No pagination or streaming — all vehicles are held in memory

---

## File Naming Conventions

- Python files use `snake_case.py`
- Classes use `PascalCase`
- Constants use `UPPER_SNAKE_CASE`
- Config keys in JSON use `snake_case`

---

## Broken Features & UX Improvements

Issues below are grouped by severity. "Crash" means the feature will raise an exception when invoked. "Broken" means it runs but produces wrong results. "Dead" means code exists but is never reachable. "UX" means it works but the user experience is poor.

### Crashes (features that error when used)

1. ~~**Charging Analysis button crashes**~~ — **FIXED in Phase 5 (Fix 31).** Added `power_level_var`, radio button selector, and corrected `analyze_charging_needs()` call signature.

2. ~~**Electrification Potential chart crashes**~~ — **FIXED in Phase 3.5 (Fix 19).** Rewrote to aggregate `vehicle_results` by make/model, showing total NPV savings per vehicle type.

3. ~~**Two PowerPoint slides crash on generation**~~ — **FIXED in Phase 5 (Fix 40).** Removed all 6 orphaned slide builders (including these two) — they were unreachable dead code never registered in `slide_builders`.

4. ~~**Selenium import crashes the app if not installed**~~ — **FIXED in Phase 1 (Fix 5).** Selenium imports guarded with try/except and `SELENIUM_AVAILABLE` flag.

5. ~~**Python 3.10+ syntax in `powerpoint_export.py`**~~ — **FIXED in Phase 6 (Fix 41).** Replaced `str | None` with `Optional[str]` for Python 3.8+ compatibility.

6. ~~**CSV export through ExportCoordinator crashes**~~ — **FIXED in Phase 6 (Fix 42).** Added `**kwargs` to `CsvReportGenerator.generate()` for API compatibility.

7. ~~**`analysis_panel.py` logger used before definition**~~ — **FIXED in Phase 6 (Fix 43).** Moved `logger = logging.getLogger(__name__)` above the `try/except` import block.

### Broken Logic (runs but produces wrong results)

8. ~~**ICE Total Cost of Ownership calculation is wrong**~~ — **FIXED in Phase 5 (Fix 32).** Rewrote to compute ICE and EV operating costs (fuel + maintenance) separately and add to respective purchase prices.

9. ~~**Emissions overcount by ~10% in PowerPoint charts**~~ — **FIXED in Phase 5 (Fix 33).** Removed spurious `* 1.1023` factor from both occurrences.

10. ~~**Body class chart "Other" category is always zero**~~ — **FIXED in Phase 3.5 (Fix 18).** Reordered truncation so `other_count` is calculated before slicing to top 10.

11. ~~**Historical emissions data is fabricated**~~ — **FIXED in Phase 5 (Fix 34).** Added `is_synthetic` flag to `EmissionsInventory` and disclaimer in chart titles/footnotes.

12. ~~**Copy-to-clipboard for charts never works**~~ — **FIXED in Phase 5 (Fix 35).** Replaced with `platform.system()` and proper platform-specific clipboard commands.

13. ~~**Chart style and color scheme dropdowns have no effect**~~ — **FIXED in Phase 5 (Fix 36).** Wired through to `ChartFactory.create_chart()` via `plt.style.context()` and colormap-derived prop_cycle.

14. ~~**Slide ordering is randomized**~~ — **FIXED in Phase 5 (Fix 39).** Replaced `set` conversion with ordered dedup loop.

15. ~~**PowerPoint template path is silently ignored**~~ — **FIXED in Phase 6 (Fix 44).** Now loads user-provided `.pptx`/`.potx` templates, with fallback to blank + warning log on failure.

16. ~~**PowerPoint branding box is off-screen**~~ — **FIXED in Phase 6 (Fix 45).** Moved from y=8" to y=6.9" (within the 7.5" slide height).

17. ~~**Commercial vehicle MPG validation rejects valid data**~~ — **FIXED in Phase 6 (Fix 46).** Raised caps to 50/45/55 for combined/city/highway to accommodate hybrid fleet vehicles.

18. ~~**Scraper indentation bug corrupts MPG values**~~ — **FIXED in Phase 6 (Fix 47).** Corrected indentation so `match`, `if match`, and value assignment are all inside the `if key not in mpg_data` guard.

### Dead Code / Stub Features

19. ~~**`analysis/charging.py` is entirely dead**~~ — **FIXED in Phase 5 (Fix 40).** File deleted.

20. ~~**6 PowerPoint slide builders are orphaned**~~ — **FIXED in Phase 5 (Fix 40).** All 6 functions removed (~320 lines).

21. **10 of 13 web scrapers are stubs** — In `commercial_vehicle_scraper.py`, only 3 scrapers are implemented (Ford, EPA SmartWay, Fuelly). The other 10 (Freightliner, Peterbilt, Kenworth, Mack, Volvo, International, Isuzu, CARB, Truck Trader, Truck Paper) return `"not yet implemented"`.

22. ~~**HTML export is listed but doesn't exist**~~ — **FIXED in Phase 5 (Fix 40).** Removed PDF/JSON/HTML from `EXPORT_FORMATS`; only CSV and Excel remain.

23. ~~**`processing_complete()` and `processing_stopped()` are never called**~~ — **FIXED in Phase 1 (Fix 1).** Wired `processing_complete()` into `done_callback` and `processing_stopped()` into `stop_processing()` in `main_window.py`.

24. ~~**Present panel format dropdown has one option**~~ — **FIXED in Phase 5 (Fix 40).** Removed the vestigial dropdown.

### UX Improvements

25. ~~**Start/Stop buttons don't reset after processing**~~ — **FIXED in Phase 1 (Fix 1).** See #23.

26. ~~**No save dialog for PowerPoint from Present tab**~~ — **FIXED in Phase 5 (Fix 37).** Added `filedialog.asksaveasfilename()` before export.

27. ~~**Analysis results return before computation finishes**~~ — **FIXED in Phase 5 (Fix 38).** Removed premature return statements; UI updates happen via `root.after()` inside threads.

28. ~~**Results table is populated twice on every load**~~ — **FIXED in Phase 1 (Fix 4).** Removed redundant `_apply_filter()` call from `populate_data()`.

29. ~~**Debug messages shown in the user-facing processing log**~~ — **FIXED in Phase 1 (Fix 2).** Removed all DEBUG: prefixed messages from process panel.

30. ~~**Empty "Appearance" tab in Preferences**~~ — **FIXED in Phase 6 (Fix 48).** Added placeholder text explaining future options; no longer a blank panel.

31. **Validation warnings in Present panel are discarded** — `ui/present_panel.py` `_validate_current_selection()` builds a `warning_text` string but never displays it. A comment acknowledges: "You could show this in a separate label or tooltip."

32. ~~**Generate PowerPoint button loses its icon after first export**~~ — **FIXED in Phase 6 (Fix 49).** Reset text now includes the `📊` emoji prefix.

33. ~~**`log_callback` and `progress_callback` are not thread-safe**~~ — **FIXED in Phase 1 (Fix 3).** Both callbacks wrapped in `root.after()` in `main_window.py`.

34. ~~**PDF export listed in File > Export but not implemented**~~ — **Partially addressed in Phase 5 (Fix 40).** PDF removed from `EXPORT_FORMATS` in `settings.py`. Note: File > Export menu may still reference PDF if it reads format types independently — check `main_window.py` if the issue persists.

35. ~~**"Reprocess Last File" button is unstyled**~~ — **FIXED in Phase 2 (Fix 14).** Added `Accent.TButton` style to `theme.py`.

36. ~~**Status filter and color coding use different success criteria**~~ — **FIXED in Phase 1 (Fix 7).** Both now use `vehicle.processing_success`.

---

## Active Improvement Progress

Tracking ongoing work on core UX fixes (upload → process → results flow).

- [x] **Fix 1: Start/Stop button states** — Wired `processing_complete()` into `done_callback`'s `update_ui` and `processing_stopped()` into `stop_processing()` in `main_window.py`. Buttons now properly re-enable after processing finishes or is stopped.
- [x] **Fix 2: Remove debug/diagnostic messages** — Removed ~250 lines of `DIAGNOSTIC:` and `DEBUG:` logging from `providers.py`, `processor.py`, `models.py`, `main_window.py`, `process_panel.py`. Deleted `diagnostic_logger` stdout handler from `providers.py`. Kept meaningful error/warning logs at appropriate levels.
- [x] **Fix 3: Thread-safe log/progress callbacks** — Wrapped `log_callback` and `progress_callback` in `root.after()` in `main_window.py` so Tkinter widget updates are always on the main thread.
- [x] **Fix 4: Double-population in results table** — Removed `_apply_filter()` call from end of `populate_data()` in `results_panel.py`. Data now inserts once instead of insert→clear→re-insert. Also cleaned ~40 lines of debug logging from `set_data()` and `populate_data()`.
- [x] **Fix 5: Selenium import guard** — Wrapped Selenium imports in try/except in `commercial_vehicle_scraper.py` with `SELENIUM_AVAILABLE` flag, matching the existing `pdfplumber` pattern. App no longer crashes if Selenium/Chrome isn't installed.
- [x] **Fix 6: Double CSV validation** — Added optional `cached_validation` parameter to `CsvReader.read_vins()` so callers can skip re-validation. Plumbing from `ProcessPanel → BatchProcessor → Pipeline → CsvReader` not yet wired (would touch 4 files); the parameter is ready for future use.
- [x] **Fix 7: Status filter vs color coding inconsistency** — Changed `_apply_filter()` status check from `vehicle.vehicle_id.year and make and model` to `vehicle.processing_success`, matching the color coding logic.
- [x] **Fix 8: Dead code cleanup in utils.py** — Removed ~245 lines of dead code: `extract_gvwr_pounds`, `detect_commercial_vehicle`, `detect_diesel_engine`, `classify_commercial_category`, `get_commercial_summary`, `extract_engine_power`. All duplicated by `models.py` `VehicleIdentification.__post_init__` and never called.

**Phase 1 Status:** All 8 core UX fixes complete. Files modified: `ui/main_window.py`, `ui/process_panel.py`, `ui/results_panel.py`, `data/providers.py`, `data/models.py`, `data/processor.py`, `commercial_vehicle_scraper.py`, `utils.py`. Total lines removed: ~600+ (debug logging, dead code, redundant operations).

Issues addressed from "Broken Features & UX Improvements" list: #4 (Selenium crash), #23 (processing_complete never called), #25 (Start/Stop buttons don't reset), #28 (results table populated twice), #29 (debug messages in user log), #33 (thread-unsafe callbacks), #36 (status filter inconsistency). Also addressed Recommendations #1 (diagnostic logging) and #2 (duplicated logic in utils.py).

### Phase 2: Results Table Polish & Persistent Cache

- [x] **Fix 9: Regression — stale `diagnostic_logger` references** — Removed 2 leftover `diagnostic_logger.info()` calls in `providers.py` `find_vehicle_matches()` that would crash every VIN lookup after Fix 2 deleted the logger instance.
- [x] **Fix 10: Number formatting in results table** — Added `_fmt_mpg()` and `_fmt_number()` helpers to `FleetVehicle` in `models.py`. MPG now displays clean values (e.g. `"24.5"` or `"17"` instead of `"17.0"`). Odometer, annual mileage, and GVWR show thousand separators (e.g. `"145,000"` instead of `"145000.0"`). Empty/zero values show as blank instead of `"0.0"`.
- [x] **Fix 11: Default visible columns** — Replaced always-empty commercial scraper fields (`payload_capacity_lbs`, `towing_capacity_lbs`, `duty_cycle`, `electrification_suitability`) with the columns users actually care about: `FuelTypePrimary`, `BodyClass`, `MPG City`, `MPG Highway`, `CO2 emissions`, `Department`. Fuel economy data — the core value of the app — is now front-and-center.
- [x] **Fix 12: Content-aware column widths** — Added `_get_column_width_and_anchor()` to `results_panel.py` with a width preset map for ~30 column types. VIN gets 170px, Year gets 70px, MPG columns get 80-100px, etc. Numeric columns (MPG, CO2, odometer) are now right-aligned. Previously all columns used `max(100, len(header) * 10)` which ignored content.
- [x] **Fix 13: Summary bar improvements** — Rewrote `_update_summary()` to use `processing_success` (matching Fix 7), show fleet stats (unique makes, avg MPG), and display an MPG coverage indicator on the right side. Added `count_label` to the summary frame. Fixed `_update_summary_filtered()` to update both labels.
- [x] **Fix 14: Accent.TButton style** — Added `Accent.TButton` style to `theme.py` (blue, bold, with hover/active states). The "Reprocess Last File" button in `process_panel.py` now renders with proper styling instead of falling back to default gray.
- [x] **Fix 15: Persistent API cache** — Extended the `Cache` class in `utils.py` to support JSON disk persistence with type-aware serialization. Dataclass objects (`VinDecoderResponse`, `FuelEconomyResponse`) are round-tripped through `to_dict()` / `from_dict()` with a `_cache_type` tag. `VehicleDataProvider` now creates a single shared persistent cache at `data/cache/api_cache.json` used by both NHTSA and FuelEconomy clients. Cache loads on startup and saves after each processing run. Re-processing the same fleet skips all API calls.

**Phase 2 Status:** Complete. Files modified: `data/providers.py` (shared cache, type registration), `data/models.py` (formatting helpers), `data/processor.py` (cache save on completion), `ui/results_panel.py` (column widths, summary bar, anchors), `ui/theme.py` (Accent.TButton), `utils.py` (disk-persistent Cache), `settings.py` (default columns).

Issues addressed: #12 (copy-to-clipboard platform check — partial, summary bar now uses consistent check), #35 (Reprocess button unstyled). Also addressed Recommendations #8 (persistent cache) and #9 (dead default columns pointing to unbuilt features).

### Phase 3: ACF Compliance & Electrification Timeline

- [x] **Feature 16: ACF Compliance Category** — Created `analysis/acf_compliance.py` with `classify_acf_vehicle()` function. Classifies each vehicle into one of five CARB ACF categories:
  - **ZEV** — Already electric, no action needed
  - **A (Exempt — Light-Duty)** — GVWR ≤ 8,500 lbs; not subject to ACF
  - **B (Subject to ACF)** — Medium/heavy-duty; covered by regulation
  - **C (Exempt — Body Type)** — Body type on CARB ZEV Purchase Exemption List (dump truck, crane, concrete mixer, etc.)
  - **D (Emergency Vehicle)** — Detected via body class, model/trim, make, or department name. Uses word-boundary regex to prevent false positives. Model keywords split into strong (PPV, SSV — unambiguous alone) and weak (police, interceptor, patrol — require corroborating emergency department name).

  Classification uses GVWR, body class, make/model/trim/series, fuel type, and department field. Runs per-vehicle during processing (no fleet context needed). Stores results in `custom_fields` for display and timeline use.

- [x] **Feature 17: Proposed Electrification Year** — Created `analysis/electrification_timeline.py` with `assign_electrification_years()` function. Assigns a target replacement year to each vehicle with:
  - **Priority scoring:** Age (35%), odometer (25%), annual mileage (15%), plus ACF boost (+0.30 for ACF-subject, +0.10 for exempt categories). Data-completeness penalty halves the ACF boost when all three metrics are missing.
  - **Budget smoothing:** Spreads replacements roughly evenly across years from now through configurable end year (default 2040)
  - **ACF dependency:** Requires ACF classification first; reads `_acf_code` from custom_fields. Light-duty exempt vehicles show "Exempt". ZEVs and failed vehicles show "N/A".
  - Runs as fleet-wide post-processing step after all VINs are processed.

- [x] **Integration:** ACF classification added to `_process_single_vin()` in `processor.py` (after quality scoring). Timeline assignment added to `process()` (after all VINs, before cache save). Three new columns (`ACF Category`, `ACF Detail`, `Proposed EV Year`) added to `to_row_dict()` in `models.py`, registered in `settings.py` column maps and default visible columns, with width presets in `results_panel.py`.

**Phase 3 Status:** Complete. New files: `analysis/acf_compliance.py`, `analysis/electrification_timeline.py`. Modified files: `data/processor.py`, `data/models.py`, `settings.py`, `ui/results_panel.py`.

### Phase 3.5: Chart Fixes, ACF/Timeline Polish, Results Improvements

- [x] **Fix 18: Body class chart "Other" always zero** — `analysis/charts.py` truncated the list *before* summing the overflow, so "Other" was always 0. Swapped the two lines so `other_count` is calculated before slicing to top 10. (Issue #10)
- [x] **Fix 19: Electrification Potential chart crash** — `analysis/charts.py` `electrification_potential()` accessed non-existent fields (`potential_by_vehicle_type`, `fleet_average_potential`, `vehicles`) on `ElectrificationAnalysis`. Rewrote the method to aggregate `vehicle_results` by make/model, showing total NPV savings per vehicle type as a horizontal bar chart with a summary panel (fleet totals, CO₂ reduction, payback period). Now handles empty analysis gracefully. (Issue #2)
- [x] **Fix 20: Timeline scoring for missing data** — `analysis/electrification_timeline.py` `_score_vehicle()` gave vehicles with zero age, odometer, and annual mileage (all missing) the full ACF boost (+0.30), pushing them ahead of vehicles with real data. Now tracks how many of the 3 metrics are present; if none are, the ACF boost is halved so these vehicles sort to the back of their tier.
- [x] **Fix 21: Category A "Proposed EV Year" blank → "Exempt"** — `analysis/electrification_timeline.py` gave light-duty exempt vehicles `""` (blank) for Proposed EV Year, which looked like missing data. Changed to `"Exempt"`. Also changed the `end_year <= current_year` fallback from empty string to `"N/A"`.
- [x] **Fix 22: Emergency vehicle detection false positives** — `analysis/acf_compliance.py` used simple substring matching, causing false positives ("fire" matched "Crossfire", "patrol" matched Nissan Patrol, "hme" matched "scheme"). Rewrote with:
  - Word-boundary regex (`\b`) for all keyword matching via new `_word_match()` helper
  - Split model keywords into **strong** (ppv, ssv, special service, responder — unambiguous alone) and **weak** (police, interceptor, pursuit, patrol — require corroborating emergency department name)
  - Tightened department keywords (e.g. "fire dept"/"fire department" instead of bare "fire", `\bems\b` to avoid matching "systems")
  - Applied word-boundary matching to exempt body type detection (`_detect_exempt_body`) too
- [x] **Fix 23: ACF Detail in default visible columns** — Added `"ACF Detail"` to `DEFAULT_VISIBLE_COLUMNS` in `settings.py` between ACF Category and Proposed EV Year, so users see the classification reasoning without manually adding the column.
- [x] **Fix 24: Column map only sampling first vehicle** — `results_panel.py` `_build_column_map()` only looked at `self.data[0].to_row_dict()` to discover custom field columns. If the first vehicle's processing failed, ACF columns were missing for the entire table. Now unions keys across up to 50 vehicles.
- [x] **Fix 25: Filtered summary shows fleet-wide stats** — `results_panel.py` `_update_summary_filtered()` showed fleet-wide averages (avg MPG, unique makes) even when a filter was active. Rewrote to compute stats from `self.data_map.values()` (the filtered subset) including success/fail counts, avg MPG, and unique makes.

**Phase 3.5 Status:** Complete. Files modified: `analysis/charts.py` (Fixes 18, 19), `analysis/electrification_timeline.py` (Fixes 20, 21), `analysis/acf_compliance.py` (Fix 22), `settings.py` (Fix 23), `ui/results_panel.py` (Fixes 24, 25).

Issues addressed from "Broken Features & UX Improvements" list: #2 (Electrification Potential chart crash), #10 (body class "Other" always zero).

### Phase 4: Process & Results Panel Overhaul, Data Acquisition Improvements

Focus: Simplify the two most-used tabs (Process and Results), improve data acquisition quality.

- [x] **Fix 26: Simplify Process Panel flow** — Rewrote `process_panel.py` from scratch with a step-based layout: **Step 1** (Upload) is a compact drop zone with browse + sample CSV buttons and a "Recent" dropdown for reprocessing history. **Step 2** (Review) appears inline after file selection showing VIN count, validity, encoding, column mappings, and a 4-row data preview. **Step 3** (Process) has the Start/Stop buttons with threading and output options collapsed behind an "Advanced Options" toggle. Output path auto-generated silently. Reduced from 1,223 lines to ~855 lines. All public methods preserved (`add_log`, `update_progress`, `processing_complete`, `processing_stopped`, `set_input_file`, `set_max_threads`, `refresh`).
- [x] **Fix 27: Consolidate Results export buttons** — Replaced two export buttons ("Export to Excel" + "Export Options...") with a single "Export" button. Opens a standard save dialog offering CSV and Excel (.xlsx) formats. CSV uses UTF-8 BOM for Excel compatibility. Excel export uses `openpyxl` with header styling and auto-width columns. Removed dead PDF/JSON/HTML format options. Added `_get_export_columns()`, `_write_csv()`, `_write_xlsx()` helper methods.
- [x] **Fix 28: Clean up Results toolbar** — Replaced the 8-control toolbar (search + field selector + LabelFrame with 2 filters + 2 export buttons + Clear Filters + Refresh) with a streamlined single-row layout: search entry + field selector | vertical separator | Status dropdown + Quality dropdown + Reset button | Export button. Removed redundant Refresh button. Shortened quality filter labels ("high (80%+)" instead of "high quality (80%+)"). Filter matching now uses `startswith()` for robustness.
- [x] **Fix 29: Improve diesel vehicle handling** — Added `is_diesel` parameter to `FuelEconomyClient.find_vehicle_matches()`. When detected: (a) filters options by diesel/biodiesel keywords via new `_filter_options_by_fuel_type()` static method, (b) if no diesel options found, falls back to gasoline with a log warning. In `get_vehicle_by_vin()`: detects fuel type mismatch when diesel VIN matched to non-diesel option text, sets `fuel_type_mismatch` flag in return data, penalizes match confidence by 15 points. Mismatch propagated through `processor.py` into `custom_fields["Fuel Type Mismatch"]`.
- [x] **Fix 30: Surface match confidence in results** — Added `Match Confidence` to `DEFAULT_VISIBLE_COLUMNS` in `settings.py` and `COLUMN_NAME_MAP`. Added column width preset (110px, right-aligned) in `results_panel.py`. Formatted as percentage (e.g. "85%") in `to_row_dict()` in `models.py`. Added `Fuel Type Mismatch` column for diesel vehicles matched to gasoline data (shows "Gas proxy (diesel data unavailable)").

**Phase 4 Status:** Complete. Files modified: `ui/process_panel.py` (full rewrite — Fix 26), `ui/results_panel.py` (toolbar + export — Fixes 27, 28, 30), `data/providers.py` (diesel handling — Fix 29), `data/processor.py` (mismatch propagation — Fix 29), `data/models.py` (match confidence + mismatch columns — Fix 30), `settings.py` (default columns, column name map — Fix 30).

Issues addressed from "Broken Features & UX Improvements" list: partial #13 (redundant export buttons consolidated).

### Phase 5: Calculation Fixes, Analysis Panel Repairs, Dead Code Cleanup

Focus: Fix all remaining crashes and broken logic in calculations/analysis, repair Analysis panel UX issues, clean up dead code.

- [x] **Fix 31: Charging Analysis crash** — Added `self.power_level_var` StringVar (default `"LP"`), added radio buttons in the Charging Parameters tab for selecting the active power level, and rewrote the `run_charging_analysis()` function call to map the selected power level to `level2_charging_rate`/`dcfc_charging_rate` parameters matching the actual `analyze_charging_needs()` signature. (Issue #1)
- [x] **Fix 32: ICE TCO calculation** — Rewrote TCO computation in `calculate_ev_roi()` to properly compute both sides: ICE TCO = purchase price + (annual fuel cost + annual maintenance) × years. EV TCO = purchase price + (annual electricity cost + annual EV maintenance) × years. Previously added EV savings to ICE price (nonsensical). (Issue #8)
- [x] **Fix 33: Emissions overcount** — Removed spurious `* 1.1023` short-ton conversion factor from both emissions calculation sites in `powerpoint_charts.py`. The `/ 1000000` already produces metric tons correctly. Both occurrences fixed via `replace_all`. (Issue #9)
- [x] **Fix 34: Fabricated historical data warning** — Added `is_synthetic: bool = False` field to `EmissionsInventory` dataclass in `models.py`. Set to `True` in `create_emissions_inventory()` when generating illustrative historical/projected data. Added subtitle disclaimer and footnote annotation to the Emissions Trends chart in `charts.py` when `is_synthetic` is set. (Issue #11)
- [x] **Fix 35: Copy-to-clipboard** — Replaced `os.system == 'Darwin'` (function-to-string comparison, always False) with `platform.system()` in `_copy_chart()`. Added `subprocess` import. macOS uses `osascript` with PNGf class, Windows uses PIL+win32clipboard with ImportError guard, Linux uses `xclip`. (Issue #12)
- [x] **Fix 36: Chart style/color dropdowns** — Added `STYLE_MAP` and `COLOR_SCHEME_MAP` dicts to `charts.py`. `ChartFactory.create_chart()` now accepts `chart_style` and `color_scheme` kwargs, applies them via `plt.style.context()` with colormap-derived prop_cycle overrides. `_update_chart()` in `analysis_panel.py` passes `self.chart_style_var.get()` and `self.color_scheme_var.get()` through to the factory. (Issue #13)
- [x] **Fix 37: PowerPoint save dialog** — Added `filedialog.asksaveasfilename()` to `_export_presentation()` in `present_panel.py` before export begins. User chooses save location and filename; cancelled dialog aborts export. User-chosen path passed as `out_path` to `export_prelim_deck()`. (Issue #26)
- [x] **Fix 38: Analysis threading** — Removed premature `return self.*_analysis` statements from `run_electrification_analysis()`, `run_emissions_analysis()`, and `run_charging_analysis()`. These returned stale/None values before the thread finished. All UI updates already happen via `root.after()` inside the threads. (Issue #27)
- [x] **Fix 39: Slide ordering** — Replaced `list(set(slide_ids + required_slides))` in `set_selected_slides()` with an ordered dedup loop using a `seen` set, preserving the user's chosen slide order. (Issue #14)
- [x] **Fix 40: Dead code cleanup** — Deleted `analysis/charging.py` (never imported, incompatible API — Issue #19). Removed 6 orphaned slide builders (~320 lines) from `powerpoint_export.py`: `_add_duty_age_slide`, `_add_screening_rules_slide`, `_add_candidates_slide`, `_add_example_replacements_slide`, `_add_energy_charging_slide`, `_add_cost_emissions_slide` (Issue #20). Removed PDF/JSON/HTML from `EXPORT_FORMATS` in `settings.py` (Issue #22). Removed single-option format dropdown from `present_panel.py` (Issue #24).

**Phase 5 Status:** Complete. Files modified: `ui/analysis_panel.py` (Fixes 31, 35, 36, 38), `analysis/calculations.py` (Fixes 32, 34), `powerpoint_charts.py` (Fixes 33, 39), `analysis/charts.py` (Fixes 34, 36), `data/models.py` (Fix 34), `ui/present_panel.py` (Fixes 37, 40), `powerpoint_export.py` (Fix 40), `settings.py` (Fix 40). Files deleted: `analysis/charging.py` (Fix 40). Total lines removed: ~400+ (dead code, orphaned functions, unused config).

Issues addressed from "Broken Features & UX Improvements" list: #1 (Charging Analysis crash), #8 (ICE TCO calculation), #9 (emissions overcount), #11 (fabricated historical data), #12 (copy-to-clipboard), #13 (chart dropdowns), #14 (slide ordering), #19 (dead charging.py), #20 (orphaned slide builders), #22 (dead export formats), #24 (single-option dropdown), #26 (no PowerPoint save dialog), #27 (analysis threading).

### Phase 6: Remaining Crashes, Data Correctness, UX Polish

Focus: Fix all remaining crash bugs, correct data output issues in scrapers and PowerPoint, polish remaining UX papercuts.

- [x] **Fix 41: Python 3.10+ syntax** — Replaced `str | None` and `SlideConfiguration | None` union type hints in `powerpoint_export.py` with `Optional[str]` and `Optional[SlideConfiguration]` for Python 3.8+ compatibility. (Issue #5)
- [x] **Fix 42: CSV export through ExportCoordinator** — Added `**kwargs` to `CsvReportGenerator.generate()` so it accepts (and ignores) the `analysis=`, `charging=`, `emissions=` kwargs passed by `ExportCoordinator.export_to_format()`. (Issue #6)
- [x] **Fix 43: Logger before definition** — Moved `logger = logging.getLogger(__name__)` in `analysis_panel.py` above the `try/except` block that imports `powerpoint_export`, so the `logger.warning()` in the `except` branch no longer raises `NameError`. (Issue #7)
- [x] **Fix 44: PowerPoint template path** — `_get_template_presentation()` now actually loads user-provided `.pptx`/`.potx` templates via `Presentation(str(path))`. Falls back to blank presentation with a warning log if the file doesn't exist or fails to load. (Issue #15)
- [x] **Fix 45: PowerPoint branding box off-screen** — Moved the "Powered by Optony" text box from y=8" (off a 7.5" slide) to y=6.9" so it's visible at the bottom of the cover slide. (Issue #16)
- [x] **Fix 46: MPG validation rejects valid data** — Raised `_is_valid_mpg()` caps from 25/20/30 to 50/45/55 for combined/city/highway MPG. Light-duty fleet vehicles (hybrid vans, Transit Connect) with >25 MPG are no longer rejected. (Issue #17)
- [x] **Fix 47: Scraper indentation bug** — Fixed broken indentation in `commercial_vehicle_scraper.py` regex pattern loop. The `match = re.search(...)`, `if match:`, and value assignment are now all correctly nested inside the `if key not in mpg_data:` guard, preventing stale matches from previous iterations from corrupting data. (Issue #18)
- [x] **Fix 48: Empty Appearance tab** — Added placeholder text to the Preferences > Appearance tab in `main_window.py` explaining that appearance settings are not yet configurable. No longer a blank, confusing panel. (Issue #30)
- [x] **Fix 49: Generate PowerPoint button loses emoji** — `_reset_export_ui()` in `present_panel.py` now restores the button text to `"📊 Generate PowerPoint"` (with emoji) instead of `"Generate PowerPoint"`. (Issue #32)

**Phase 6 Status:** Complete. Files modified: `ui/analysis_panel.py` (Fix 43), `powerpoint_export.py` (Fixes 41, 44, 45), `analysis/reports.py` (Fix 42), `commercial_vehicle_scraper.py` (Fixes 46, 47), `ui/main_window.py` (Fix 48), `ui/present_panel.py` (Fix 49).

Issues addressed from "Broken Features & UX Improvements" list: #5 (Python 3.10+ syntax), #6 (ExportCoordinator crash), #7 (logger NameError), #15 (template path ignored), #16 (branding off-screen), #17 (MPG validation), #18 (scraper indentation), #30 (empty Appearance tab), #32 (button emoji lost).

### Phase 7: Automated Test Suite

Focus: Address Recommendation #4 — build a proper pytest test suite covering all pure-logic functions that can be tested without API mocking.

**Infrastructure:**
- `tests/__init__.py` — Package marker
- `tests/conftest.py` — Factory helpers (`make_vehicle_id`, `make_fuel_economy`, `make_fleet_vehicle`) and pytest fixtures (`sample_vehicle`, `diesel_vehicle`, `electric_vehicle`, `emergency_vehicle`, `sample_fleet`). Factory functions accept `vid_overrides` and `fuel_overrides` dicts for targeted field customization.

**Test modules (151 tests total):**

- [x] `tests/test_vin_validation.py` (15 tests) — `validate_vin_detailed()`: valid VINs, empty/None/whitespace, wrong length, invalid characters (I/O/Q), placeholder detection (all-zeros, all-ones, excessive zeros).
- [x] `tests/test_models.py` (26 tests) — `VehicleIdentification.__post_init__`: GVWR parsing from NHTSA class format, simple pounds, commas, raw numbers, empty strings (11 tests). Commercial detection via body class, model name, GVWR threshold (4 tests). Diesel detection from fuel_type, fuel_type_secondary, biodiesel (5 tests). Quality scoring: full data, minimal data, score cap at 100, breakdown keys, fuel economy points (6 tests).
- [x] `tests/test_normalize.py` (16 tests) — `normalize_vehicle_model()`: Ford F-series (F150/F-150/F 150 → F-150), E-series, Chevy Silverado/Sierra numbering, Ram numbering, Express/Savana, case handling, whitespace stripping, special character removal.
- [x] `tests/test_calculations.py` (19 tests) — `calculate_annual_fuel_cost()`, `calculate_annual_ev_cost()`, `calculate_annual_co2_emissions()`, `calculate_emissions_reduction()`, `calculate_ev_roi()` (regression tests for Fix 32 — ICE/EV TCO both include operating costs, savings is the difference), `create_emissions_inventory()` (regression test for Fix 34 — `is_synthetic` flag set).
- [x] `tests/test_acf_compliance.py` (16 tests) — `classify_acf_vehicle()`: ZEV detection (BEV, fuel cell), Category A light-duty (sedan, SUV, boundary 8500 lb), Category B medium/heavy-duty, Category C exempt body types (dump truck, concrete mixer, crane), Category D emergency vehicles (PPV trim, ambulance body class, fire apparatus make). False positive prevention: Crossfire ≠ fire, Nissan Patrol ≠ patrol. Weak keyword + emergency department triggers (with GVWR >8500 to avoid light-duty preemption).
- [x] `tests/test_electrification_timeline.py` (22 tests) — `_score_vehicle()`: age/mileage/annual-usage component contributions, ACF boost levels (B > C), unknown code gives zero boost, data-completeness penalty (Fix 20 regression — halved boost when all three metrics missing), score bounds. `assign_electrification_years()`: ZEV→N/A, failed→N/A, missing ACF→N/A, Category A→Exempt, B/C/D→year, higher-score→earlier-year ordering, budget smoothing distribution, past end_year→N/A, mixed fleet triage, empty fleet.
- [x] `tests/test_csv_mapping.py` (37 tests) — `CsvFileValidator._find_vin_column()`: exact match, case-insensitive, partial match, no VIN column. `_map_additional_columns()`: standard field mapping (department, odometer, location, asset_id), variant names (dept→department, mileage→odometer), case-insensitive mapping, unmapped columns preserved, VIN column skipped, fleet_management_fields populated. `FleetVehicle.to_row_dict()`: key presence (core, fuel economy, commercial, ACF), MPG formatting (clean numbers, zero→blank), match confidence as percentage, processing status, diesel flag, odometer formatting, custom field passthrough. `validate_and_preview()` integration: valid CSV, empty CSV, no VIN column, additional column detection, invalid VIN counting, nonexistent file.

**Phase 7 Status:** Complete. New files: `tests/__init__.py`, `tests/conftest.py`, `tests/test_vin_validation.py`, `tests/test_models.py`, `tests/test_normalize.py`, `tests/test_calculations.py`, `tests/test_acf_compliance.py`, `tests/test_electrification_timeline.py`, `tests/test_csv_mapping.py`. All 151 tests passing.

Recommendation addressed: #4 (Add a Proper Test Suite). All five originally recommended areas covered (VIN validation, GVWR parsing, model normalization, quality scoring, CSV column mapping) plus four additional areas (TCO/ROI calculations, ACF compliance, electrification timeline, emissions inventory).

### Phase 9: Decision-Support & Client-Ready Output

Focus: Transform the tool from a data processing step into a complete decision-support platform. Every improvement targets reducing manual post-processing work — the spreadsheets, slide edits, and financial models that consultants currently build after exporting from this app.

**Guiding principle:** After Phase 9, a consultant should be able to process a fleet CSV, run analysis, and hand the PowerPoint + Excel export directly to a client with minimal editing.

#### Phase 9A: Multi-Year Cash Flow TCO Model — DONE

The current TCO in `calculate_ev_roi()` produced a single lump sum. Replaced with a transparent year-by-year model.

**Changes to `analysis/calculations.py`:**
- [x] `calculate_yearly_cash_flows()` — New function (~170 lines) returning `{"yearly_flows": [...], "summary": {...}}` with year-0 purchase + years 1..N operating costs, fuel escalation, battery degradation, incentive deduction, infrastructure amortization, residual values, payback interpolation, NPV calculation.
- [x] Added `fuel_escalation_rate` (3%/yr), `incentive_amount`, `infrastructure_cost_per_vehicle`, `residual_value_ice_pct` (15%), `residual_value_ev_pct` (20%) parameters — all configurable with defaults in `settings.py`.
- [x] Refactored `calculate_ev_roi()` to delegate to `calculate_yearly_cash_flows()` — backward compatible, returns `yearly_cash_flows` key.
- [x] Updated `analyze_fleet_electrification()` — accepts new parameters, computes per-vehicle cash flows when EV pricing available, builds fleet-wide aggregate `fleet_cash_flows`.
- [x] Added `fleet_cash_flows` field to `ElectrificationAnalysis` dataclass in `models.py`.

**Phase 9A Status:** Complete. Files modified: `analysis/calculations.py`, `data/models.py`, `settings.py`. 5 new constants in settings. All 151 tests passing.

#### Phase 9B: EV Equivalent Mapping & Replacement Recommendations — DONE

**New file `analysis/ev_database.py` (~400 lines):**
- [x] `EVEquivalent` dataclass and `EV_DATABASE` list with ~25 entries covering sedans, SUVs, pickups (LD/MD), cargo vans, passenger vans, medium-duty trucks, heavy-duty trucks, buses, specialty vehicles.
- [x] `find_ev_equivalent(vehicle)` — Scoring engine (body class 30pts, GVWR range 25pts, keyword 20-30pts, make bonus 10pts, capability 5pts).
- [x] `generate_replacement_recommendation()` — Structured dict with current vehicle, proposed EV, payback, rationale string.
- [x] `get_priority_replacements(fleet, n=15)` — Top N sorted by NPV savings.
- [x] `match_fleet_ev_equivalents(fleet)` — Fleet-wide matching, populates `_ev_purchase_price`, `_ice_purchase_price`, `EV Equivalent`, `EV MSRP Range`, `EV EPA Range`, `EV Fit Score` in `custom_fields`.

**Integration:** Step 4b in `processor.py` after timeline assignment. EV columns added to `settings.py` COLUMN_NAME_MAP and DEFAULT_VISIBLE_COLUMNS. Column width presets in `results_panel.py`.

**Phase 9B Status:** Complete. New file: `analysis/ev_database.py`. Modified: `data/processor.py`, `settings.py`, `ui/results_panel.py`. All 151 tests passing.

#### Phase 9C: Wire Real Data into PowerPoint Charts — DONE

Replaced all hardcoded percentage schedules and fabricated curves with real fleet data.

**Changes to `powerpoint_charts.py`:**
- [x] Rewrote `add_emissions_timeline_chart()` — reads `Proposed EV Year` from vehicle `custom_fields`, computes real baseline emissions and projects cumulative CO₂ reduction year by year.
- [x] Rewrote `add_electrification_timeline_by_weight_chart()` — groups vehicles by actual assigned year × weight class (Light/Medium/Heavy Duty). No more hardcoded arrays.
- [x] Rewrote `add_electrification_timeline_by_body_type_chart()` — groups by actual year × body type (top 5 + Other).
- [x] Added `add_tco_comparison_chart()` — Stacked bar chart: ICE vs EV fleet TCO with purchase/fuel/maintenance components.
- [x] Added `add_payback_timeline_chart()` — Dual line chart: cumulative ICE vs EV cost over time showing crossover point.
- [x] Updated `SlideConfiguration.__init__()` — Added `financial_summary`, `executive_recommendations`, `replacement_schedule` slides; `stacked_bar_tco`, `line_payback` chart types; updated default selected_slides.

**Phase 9C Status:** Complete. File modified: `powerpoint_charts.py`. All 151 tests passing.

#### Phase 9D: Financial Summary & Executive Summary Slides — DONE

Added three new slide builders and rewrote Next Steps with data-driven content.

**Changes to `powerpoint_export.py`:**
- [x] New `_add_financial_summary_slide()` — 5 KPI boxes (EV Investment, Investment Premium, Annual Savings, Simple Payback, Lifetime CO₂ Saved) + TCO comparison chart (left) + payback timeline chart (right) + assumptions footnote. Computes fleet-wide metrics from real vehicle data.
- [x] New `_add_executive_recommendations_slide()` — Data-driven narrative paragraph ("Based on analysis of N vehicles..."), top 5 priority replacement table with formatted headers and alternating row shading, risk/opportunity callouts (no-EV-equivalent count, ACF-mandated count, immediate-replacement count).
- [x] New `_add_replacement_schedule_slide()` — Table of top 12 priority vehicles: Current Vehicle, Department, Proposed EV, Annual Savings, Payback, Target Year. Sorted by proposed year then NPV savings. Includes combined annual savings footer.
- [x] Rewrote `_add_next_steps_slide()` — No more generic boilerplate. Now uses fleet data to generate specific phase counts ("Begin procurement for N priority vehicles"), ACF compliance actions ("Develop strategy for N regulated vehicles"), and data-driven fiscal year references.
- [x] Registered all 3 new builders in `slide_builders` dict. Added `add_tco_comparison_chart` and `add_payback_timeline_chart` to imports.

**Phase 9D Status:** Complete. File modified: `powerpoint_export.py`. All 151 tests passing.

#### Phase 9E: Scenario Comparison for Electrification Timelines — DONE

**New file `analysis/scenarios.py` (~280 lines):**
- [x] `ElectrificationScenario` dataclass — name, end_year, budget_per_year (optional), vehicle_filter ("all"/"acf_only"/"medium_heavy_only"), custom_weights, description.
- [x] `PRESET_SCENARIOS` — 4 presets: "Aggressive" (2030, all), "Moderate" (2035, all), "Conservative" (2040, all), "ACF Compliance Only" (2035, acf_only).
- [x] `run_scenario(vehicles, scenario)` → dict with vehicles_per_year, cumulative_cost/co2/savings per year, total_investment, total_annual_savings, total_annual_co2_reduction, summary_text. Uses same `_score_vehicle()` from electrification_timeline for consistency. Supports budget_per_year constraint.
- [x] `compare_scenarios(vehicles, scenario_names, custom_scenarios)` → dict with per-scenario results, comparison_table, all_years, best_roi/lowest_cost/fastest identifiers.

**PowerPoint integration:**
- [x] New `add_scenario_comparison_chart()` in `powerpoint_charts.py` — Multi-line chart with configurable metric (vehicles/cost/co2/savings). Shows all scenarios as color-coded lines.
- [x] New `_add_scenario_comparison_slide()` in `powerpoint_export.py` — Runs all 4 preset scenarios, shows vehicles chart (left) + investment chart (right), adds comparison metrics table with formatted headers and alternating row shading, plus best-in-class callouts at bottom.
- [x] `scenario_comparison` slide and chart types registered in `SlideConfiguration`.

**Phase 9E Status:** Complete. New file: `analysis/scenarios.py`. Modified: `powerpoint_charts.py`, `powerpoint_export.py`. All 151 tests passing. UI integration (analysis_panel scenario selector) deferred to Phase 9G (unified workflow).

#### Phase 9F: Presentation-Quality Chart Improvements — DONE

Upgraded all matplotlib charts from data-exploration style to presentation-ready client deliverables.

**Changes to `analysis/charts.py`:**
- [x] New `apply_presentation_style(fig, ax, title, subtitle, footnote)` helper — white background, minimal gridlines, no top/right spines, standard title/subtitle/footnote placement.
- [x] Constants `FIG_SIZE_SLIDE = (13.33, 7.5)` and `FIG_SIZE_HALF = (10, 5.6)` for 16:9 slide-native rendering.
- [x] Added executive insight subtitles to all chart methods — computed from data (median MPG, high-emitter counts, top department %, savings %, etc.).
- [x] Added data callouts and annotations — worst emitter on scatter plots, total values on bar charts, percentage labels on pie/donut charts, median lines on distributions.
- [x] New `DecisionCharts` class with 3 methods:
  - `fleet_cashflow_chart()` — Dual cumulative cost curves (ICE vs EV) with payback crossover annotation and fill_between savings area.
  - `replacement_priority_chart()` — Horizontal bars by NPV savings with EV equivalent labels and payback annotations.
  - `scenario_comparison_chart()` — Overlaid area charts for multiple scenarios showing vehicles electrified per year.
- [x] All 3 registered in `ChartFactory.create_chart()` under "Fleet Cash Flow", "Replacement Priority", "Scenario Comparison".
- [x] Converted fuel type distribution to donut chart with center total.

**Phase 9F Status:** Complete. File modified: `analysis/charts.py` (~1,910 lines). All 151 tests passing.

**Known remaining polish:** `FIG_SIZE_SLIDE`/`FIG_SIZE_HALF` constants are used by `DecisionCharts` but not enforced in older chart methods (which use the figure passed by the caller). This is cosmetic — all charts render correctly at any size.

#### Phase 9G: Unified Analysis Workflow — DONE

Consolidated the Analysis panel from 3 separate buttons + text summary into a single-click workflow with KPI cards.

**Changes to `ui/analysis_panel.py`:**
- [x] New `run_full_analysis()` method — runs electrification → emissions → charging sequentially in a background thread, then updates KPI cards and switches chart to "Fleet Cash Flow".
- [x] Replaced text-only "Results Summary" with 6 KPI cards (`self.kpi_labels` dict): fleet_size, avg_mpg, annual_savings, co2_reduction, payback, infra_cost. Grid layout 2×3.
- [x] Added "Build Presentation" button (`_navigate_to_present()`) that switches to the Present tab (index 3).
- [x] Consolidated 3 export buttons into single "Export" button with popup menu (`_export_menu()`) offering PowerPoint, Excel/CSV, and chart image.

**Changes to `settings.py`:**
- [x] `CHART_TYPES` expanded from 12 to 22 entries, adding: "Emissions Inventory", "Emissions Trends", "Emissions by Department", "Emissions by Vehicle Type", "Fleet Cash Flow", "Replacement Priority", "Scenario Comparison" (plus pre-existing ones that were missing from the list).

**Phase 9G Status:** Complete. Files modified: `ui/analysis_panel.py`, `settings.py`. All 151 tests passing.

#### Phase 9H: Incentive & Rate Database — DONE

Added state-level energy rates and federal/state incentive programs so TCO calculations can use localized data.

**New file `analysis/rate_database.py` (~280 lines):**
- [x] `STATE_ENERGY_RATES` dict — 50 states + DC with `gas_price`, `electricity_price`, `demand_charge`.
- [x] `FEDERAL_INCENTIVES` list — IRA 45W ($40K commercial), 30D ($7.5K consumer), 30C ($100K infrastructure).
- [x] `STATE_INCENTIVES` dict — 9 states covered: CA (HVIP $120K, CVRP $7.5K), NY (NYTVIP $185K, Drive Clean $2K), CO ($5K), NJ ($4K), MA ($3.5K), OR ($7.5K), TX (TERP $60K), WA (tax exempt), IL ($4K).
- [x] `get_rates_for_state(state_code)` → dict with gas_price, electricity_price, demand_charge, source.
- [x] `get_federal_incentives(vehicle_class)`, `get_state_incentives(state_code, vehicle_class)`, `get_all_incentives(state_code, vehicle_class)` → combined with max_federal, max_state, max_total.
- [x] `get_available_states()` → sorted state code list.

**UI integration in `ui/analysis_panel.py`:**
- [x] State selector combobox in Cost Parameters tab with `_on_state_selected()` handler that auto-populates gas/electricity prices.
- [x] Incentive display label showing available programs and max total when state is selected.

**Phase 9H Status:** Complete. New file: `analysis/rate_database.py`. Modified: `ui/analysis_panel.py`. All 151 tests passing.

**Known remaining polish:** Incentives are displayed as information text, not as user-selectable checkboxes. The incentive total is not yet automatically wired into the TCO `incentive_amount` parameter — user must manually enter it. Future enhancement: add checkboxes and auto-sum into the TCO input.

#### Phase 9I: Analysis-Ready Excel Export — DONE

Added 3 new Excel worksheets that transform the export from a data dump into a structured financial deliverable.

**Changes to `analysis/reports.py`:**
- [x] New `_create_tco_model_sheet()` — Year-by-year columns: Year, ICE Annual, EV Annual, Annual Savings, ICE Cumulative, EV Cumulative, Cumulative Savings, Notes (marks PAYBACK YEAR). Plus assumptions section.
- [x] New `_create_replacement_schedule_sheet()` — Gantt-style table: vehicle rows × year columns with "X" markers for replacement year. Summary row with vehicles per year and total estimated cost.
- [x] New `_create_summary_dashboard_sheet()` — KPI cards (Fleet Size, Avg MPG, Annual Savings, CO₂ Reduction, Payback, Infrastructure Cost), emissions by department table, ACF classification breakdown, replacement year distribution.
- [x] All 3 called from `generate()` method conditionally (TCO sheet requires `analysis.fleet_cash_flows`).

**Phase 9I Status:** Complete. File modified: `analysis/reports.py`. All 151 tests passing.

**Known remaining polish:** TCO Model sheet writes static computed values, not live Excel `=` formulas. Users cannot yet change assumptions in Excel and see results update automatically. Future enhancement: replace static values with `openpyxl` formula cells referencing an assumptions range.

#### Phase 9J: Present Panel Editable Content & Vehicle Filtering — DONE

Added client-facing customization controls so users can personalize presentations and filter which vehicles are included.

**Changes to `ui/present_panel.py`:**
- [x] New `_create_details_section()` — Editable fields: Client Name (`client_name_var`), Subtitle (`subtitle_var`, default "Fleet Electrification Analysis"), Stage (`stage_var`, combobox with Preliminary/Final/Draft/Revised). These appear on the cover slide and headers.
- [x] New `_create_vehicle_filter_section()` — Department dropdown, ACF Category dropdown, Max Payback dropdown (No Limit / <3 / <5 / <7 / <10 years). Dropdowns populated from fleet data via `_populate_filter_dropdowns()`.
- [x] New `_get_filtered_vehicles()` — Returns subset of fleet vehicles matching all active filters.
- [x] New `_build_slide_context(vehicles)` — Generates per-slide context strings from vehicle data (e.g., "142 vehicles, 8 departments, avg 18.3 MPG") for data-aware preview.
- [x] Enhanced `_update_preview()` — Shows client name, subtitle, stage, template, vehicle count (with filter indicator), and per-slide descriptions with context and chart counts.
- [x] Vehicle filter applied in `_export_presentation()` — filtered vehicles passed as `export_data['vehicles']` to `export_prelim_deck()`. Empty filter result shows error.
- [x] `refresh_data()` updated to populate filter dropdowns and pre-fill client name from fleet name.

**Changes to `powerpoint_export.py`:**
- [x] Cover slide now uses `data.get('subtitle', ...)` to support custom subtitles from the UI.

**Phase 9J Status:** Complete. Files modified: `ui/present_panel.py` (~1,015 lines, up from 755), `powerpoint_export.py` (1 line). All 151 tests passing.

**Phase 9 Status:** All 10 sub-phases (9A–9J) complete. See "Known Remaining Polish" notes on 9F, 9H, and 9I for future enhancement opportunities.

---

## Known Remaining Polish (Post-Phase 9)

These are minor gaps identified during reconciliation audits. None are bugs — the features work correctly. They represent future enhancement opportunities.

1. **9F — Figure size enforcement:** `FIG_SIZE_SLIDE`/`FIG_SIZE_HALF` constants defined but only used by `DecisionCharts`. Older chart methods use caller-provided figure sizes. Cosmetic only.
2. **9H — Incentive auto-wiring:** Incentives display as informational text. They are not yet user-selectable checkboxes that auto-sum into the TCO `incentive_amount` parameter.
3. **9I — Excel formula cells:** TCO Model sheet writes static Python-computed values, not live `=` formulas. Users cannot change assumptions in Excel and see live updates.
4. **9E — Scenario UI selector:** The scenario comparison slide always runs all 4 preset scenarios. No UI exists to let users select which scenarios to include or define custom ones.
5. **9A — Parameter exposure:** `battery_degradation`, `residual_value_ice_pct`, and `residual_value_ev_pct` are not exposed as parameters on `analyze_fleet_electrification()` — they use defaults internally. Callers cannot customize without modifying `calculate_yearly_cash_flows()` directly.
