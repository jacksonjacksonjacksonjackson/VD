"""
reports.py

Report generation and exports for the Fleet Electrification Analyzer.
Provides functions to create various report formats (CSV, Excel).
"""

import os
import csv
import json
import logging
import datetime
from typing import Dict, List, Any, Optional, Union, Tuple, Set
from pathlib import Path
from io import BytesIO

from matplotlib import pyplot as plt

from settings import (
    DEFAULT_GAS_PRICE,
    DEFAULT_ELECTRICITY_PRICE,
    DEFAULT_EV_EFFICIENCY,
    DEFAULT_ANNUAL_MILEAGE,
    EXPORT_FORMATS,
    COLUMN_NAME_MAP
)
from data.models import FleetVehicle, Fleet, ElectrificationAnalysis, ChargingAnalysis, EmissionsInventory
from analysis.charts import ChartFactory

# Set up module logger
logger = logging.getLogger(__name__)

# Check for optional dependencies
try:
    import xlsxwriter
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False
    logger.warning("XlsxWriter not available. Excel export will be disabled.")


###############################################################################
# Canonical assumptions — single source of truth on the Cover sheet
###############################################################################
# Every financial sheet references these cells via cover_aref(); the Cover sheet
# is the only place a user edits them. Percentages are stored as whole numbers
# (e.g. 3.0 = 3%/yr) to match the TCO model's existing formula convention.

COVER_SHEET_NAME = "Cover & Methodology"

# First assumption value lives in column B of this 1-indexed Excel row.
_ASSUMPTION_FIRST_ROW = 9

# Ordered list of (key, label, kind). kind ∈ {money2, money0, num2, pct, int}.
_ASSUMPTION_DEFS = [
    ("gas_price",           "Gas price ($/gal)",                  "money2"),
    ("electricity_price",   "Electricity price ($/kWh)",          "money2"),
    ("ev_efficiency",       "EV efficiency (kWh/mile)",           "num2"),
    ("ice_maintenance",     "ICE maintenance ($/mile)",           "money2"),
    ("ev_maintenance",      "EV maintenance ($/mile)",            "money2"),
    ("fuel_escalation",     "Fuel escalation (%/yr)",             "pct"),
    ("discount_rate",       "Discount rate (%/yr)",               "pct"),
    ("battery_degradation", "Battery degradation (%/yr)",         "pct"),
    ("infra_cost_per_veh",  "Infrastructure cost / vehicle ($)",  "money0"),
    ("analysis_period",     "Analysis period (years)",            "int"),
]

# key -> 1-indexed Excel row of its value cell in column B of the Cover sheet.
_ASSUMPTION_ROW = {
    key: _ASSUMPTION_FIRST_ROW + i for i, (key, _label, _kind) in enumerate(_ASSUMPTION_DEFS)
}


def cover_aref(key: str) -> str:
    """Absolute cross-sheet reference to a canonical assumption cell (no leading '=').

    e.g. cover_aref("gas_price") -> "'Cover & Methodology'!$B$9"
    Wrap in "=" + cover_aref(...) to write a mirror formula, or embed in a larger one.
    """
    return f"'{COVER_SHEET_NAME}'!$B${_ASSUMPTION_ROW[key]}"


###############################################################################
# Report Generator Base Class
###############################################################################

class ReportGenerator:
    """Base class for report generators."""
    
    def __init__(self, output_path: str):
        """
        Initialize the report generator.
        
        Args:
            output_path: Path to the output file
        """
        self.output_path = output_path
    
    def generate(self, **kwargs) -> bool:
        """
        Generate the report.
        
        Args:
            **kwargs: Additional arguments for the specific report type
            
        Returns:
            True if successful, False otherwise
        """
        raise NotImplementedError("Subclasses must implement this method")


###############################################################################
# CSV Export
###############################################################################

class CsvReportGenerator(ReportGenerator):
    """Generate CSV reports from fleet data."""
    
    def generate(self, fleet: Union[Fleet, List[FleetVehicle]], fields: Optional[List[str]] = None,
                 **kwargs) -> bool:
        """
        Generate a CSV report from fleet data.

        Args:
            fleet: Fleet object or list of vehicles
            fields: List of fields to include (None for default fields)
            **kwargs: Accepts (and ignores) analysis, charging, emissions for API
                      compatibility with ExportCoordinator

        Returns:
            True if successful, False otherwise
        """
        try:
            # Extract vehicles from fleet if needed
            vehicles = fleet.vehicles if isinstance(fleet, Fleet) else fleet
            
            if not vehicles:
                logger.warning("No vehicles to export")
                return False
            
            # Determine fields to include
            if fields is None:
                # Use all fields from the first vehicle as reference
                sample = vehicles[0].to_row_dict()
                fields = list(sample.keys())
            
            # Create user-friendly headers
            headers = [COLUMN_NAME_MAP.get(field, field) for field in fields]
            
            with open(self.output_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                
                # Write headers
                writer.writerow(headers)
                
                # Write data
                for vehicle in vehicles:
                    data = vehicle.to_row_dict()
                    writer.writerow([data.get(field, "") for field in fields])
            
            logger.info(f"CSV report generated successfully: {self.output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error generating CSV report: {e}")
            return False


###############################################################################
# Excel Export
###############################################################################

class ExcelReportGenerator(ReportGenerator):
    """Generate Excel reports from fleet data with charts and analysis."""
    
    def generate(self, fleet: Union[Fleet, List[FleetVehicle]],
               analysis: Optional[ElectrificationAnalysis] = None,
               charging: Optional[ChargingAnalysis] = None,
               emissions: Optional[EmissionsInventory] = None,
               fields: Optional[List[str]] = None,
               timeline_options: Optional[dict] = None,
               client_profile: Optional[Any] = None,
               state_code: str = "CA") -> bool:
        """
        Generate an Excel report with data, charts, and analysis.

        Args:
            fleet: Fleet object or list of vehicles
            analysis: Optional electrification analysis
            charging: Optional charging analysis
            emissions: Optional emissions inventory
            fields: List of fields to include (None for default fields)
            timeline_options: scenario year columns to append to Replacement sheet
            client_profile: Optional PresentationProfile for cover-page client/date/presenter
            state_code: Two-letter state code for incentive lookups (default "CA")

        Returns:
            True if successful, False otherwise
        """
        if not EXCEL_AVAILABLE:
            logger.error("Excel export is not available (xlsxwriter not installed)")
            return False
        
        try:
            # Extract vehicles from fleet if needed
            vehicles = fleet.vehicles if isinstance(fleet, Fleet) else fleet
            
            if not vehicles:
                logger.warning("No vehicles to export")
                return False
            
            # Determine fields to include
            if fields is None:
                # Use all fields from the first vehicle as reference
                sample = vehicles[0].to_row_dict()
                fields = list(sample.keys())
            
            # Create Excel workbook
            workbook = xlsxwriter.Workbook(self.output_path)
            
            # Defined styles
            title_format = workbook.add_format({
                'bold': True,
                'font_size': 14,
                'align': 'center',
                'valign': 'vcenter'
            })
            
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#3C465A',  # PRIMARY_HEX_1
                'font_color': 'white',
                'border': 1
            })
            
            cell_format = workbook.add_format({
                'border': 1
            })
            
            number_format = workbook.add_format({
                'border': 1,
                'num_format': '#,##0.0'
            })

            # Stash report-wide context for sheet builders
            self._client_profile = client_profile
            self._state_code = state_code or "CA"

            # Create Cover & Methodology sheet FIRST so cross-sheet assumption
            # references ('Cover & Methodology'!$B$N) resolve for later sheets.
            self._create_cover_sheet(workbook, vehicles, analysis, charging, emissions,
                                     client_profile, title_format, header_format,
                                     cell_format, number_format)

            # Create Vehicle Data sheet
            self._create_vehicle_data_sheet(workbook, vehicles, fields, title_format, header_format, cell_format, number_format)
            
            # Create Fleet Overview sheet (merged Summary + Summary Dashboard)
            self._create_fleet_overview_sheet(workbook, vehicles, analysis, charging, emissions,
                                              title_format, header_format, cell_format, number_format)

            # Create Electrification Analysis sheet if available
            if analysis:
                self._create_electrification_sheet(workbook, analysis, title_format, header_format, cell_format, number_format)
            
            # Create Charging Infrastructure sheet if available
            if charging:
                self._create_charging_sheet(workbook, charging, title_format, header_format, cell_format, number_format)
            
            # Create Emissions Inventory sheet if available
            if emissions:
                self._create_emissions_sheet(workbook, emissions, title_format, header_format, cell_format, number_format)

            # Phase 9I: Analysis-ready sheets
            if analysis and hasattr(analysis, 'fleet_cash_flows') and analysis.fleet_cash_flows:
                self._create_tco_model_sheet(workbook, analysis, vehicles, title_format, header_format, cell_format, number_format)

            self._create_replacement_schedule_sheet(
                workbook, vehicles,
                title_format, header_format, cell_format, number_format,
                timeline_options=timeline_options,
            )

            # Close workbook to save changes
            workbook.close()
            
            logger.info(f"Excel report generated successfully: {self.output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error generating Excel report: {e}")
            return False
    
    def _create_cover_sheet(self, workbook, vehicles, analysis, charging, emissions,
                            client_profile, title_format, header_format,
                            cell_format, number_format):
        """Create the Cover & Methodology sheet.

        Holds the canonical (single-source-of-truth) assumptions block that every
        financial sheet references, plus client header, data-quality KPIs, a
        methodology / data-source note, and the ACF glossary.
        """
        from settings import (
            DEFAULT_GAS_PRICE, DEFAULT_ELECTRICITY_PRICE, DEFAULT_EV_EFFICIENCY,
            DEFAULT_ICE_MAINTENANCE, DEFAULT_EV_MAINTENANCE,
            DEFAULT_FUEL_ESCALATION_RATE, DEFAULT_BATTERY_DEGRADATION,
            DEFAULT_INFRASTRUCTURE_COST_PER_VEHICLE,
        )

        ws = workbook.add_worksheet(COVER_SHEET_NAME)
        ws.set_column('A:A', 34)
        ws.set_column('B:B', 20)
        ws.set_column('C:F', 16)
        ws.hide_gridlines(2)

        # ── Formats ──────────────────────────────────────────────────────────────
        big_title_fmt = workbook.add_format({
            'bold': True, 'font_size': 22, 'font_color': '#3C465A', 'valign': 'vcenter'})
        client_fmt = workbook.add_format({
            'bold': True, 'font_size': 16, 'font_color': '#C45911', 'valign': 'vcenter'})
        meta_fmt = workbook.add_format({'font_size': 11, 'font_color': '#444444'})
        section_fmt = workbook.add_format({
            'bold': True, 'font_size': 12, 'bg_color': '#3C465A', 'font_color': 'white',
            'border': 1, 'align': 'left', 'valign': 'vcenter'})
        label_fmt = workbook.add_format({'border': 1, 'align': 'left'})
        note_fmt = workbook.add_format({'italic': True, 'font_size': 9, 'font_color': '#7F8C8D'})
        body_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top', 'font_size': 10})
        kpi_label_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#EAF2FB', 'border': 1, 'align': 'center', 'font_size': 9})
        kpi_value_fmt = workbook.add_format({
            'bold': True, 'font_size': 16, 'align': 'center', 'border': 1})
        # Editable canonical assumption cells (yellow) — the only place users edit.
        money2_in = workbook.add_format({'border': 2, 'bg_color': '#FFF9C4', 'num_format': '$#,##0.00', 'align': 'right'})
        money0_in = workbook.add_format({'border': 2, 'bg_color': '#FFF9C4', 'num_format': '$#,##0', 'align': 'right'})
        num2_in   = workbook.add_format({'border': 2, 'bg_color': '#FFF9C4', 'num_format': '#,##0.00', 'align': 'right'})
        pct_in    = workbook.add_format({'border': 2, 'bg_color': '#FFF9C4', 'num_format': '0.00"%"', 'align': 'right'})
        int_in    = workbook.add_format({'border': 2, 'bg_color': '#FFF9C4', 'num_format': '#,##0', 'align': 'right'})
        _in_fmt = {'money2': money2_in, 'money0': money0_in, 'num2': num2_in, 'pct': pct_in, 'int': int_in}

        # ── Header block (rows 1–6, 1-indexed) ───────────────────────────────────
        prof = client_profile
        client_name = (getattr(prof, 'client_name', '') or '').strip() or "Fleet Electrification Analysis"
        meeting_date = (getattr(prof, 'meeting_date', '') or '').strip()
        presenter = (getattr(prof, 'presenter_name', '') or '').strip()
        presenter_co = (getattr(prof, 'presenter_company', '') or '').strip()
        report_date = datetime.datetime.now().strftime("%B %d, %Y")

        ws.set_row(0, 30)
        ws.merge_range('A1:F1', "Fleet Electrification Analysis", big_title_fmt)
        ws.set_row(1, 24)
        ws.merge_range('A2:F2', client_name, client_fmt)
        ws.write(2, 0, f"Report generated: {report_date}", meta_fmt)
        prepared = "Prepared by: " + (f"{presenter}" + (f", {presenter_co}" if presenter_co else "") if presenter else (presenter_co or "—"))
        ws.write(3, 0, prepared, meta_fmt)
        if meeting_date:
            ws.write(4, 0, f"Meeting date: {meeting_date}", meta_fmt)
        ws.write(5, 0, f"Fleet size: {len(vehicles)} vehicles   |   Incentive region: {self._state_code}", meta_fmt)

        # ── Global assumptions (header row 8; values rows 9..) ────────────────────
        ws.merge_range('A8:C8', "Global Assumptions  (edit yellow cells — drives the TCO & Financials model)", section_fmt)

        # Seed values: prefer the analysis object, else settings defaults.
        def _a(attr, default):
            return getattr(analysis, attr, default) if analysis is not None else default
        seeds = {
            "gas_price":           _a('gas_price', DEFAULT_GAS_PRICE),
            "electricity_price":   _a('electricity_price', DEFAULT_ELECTRICITY_PRICE),
            "ev_efficiency":       _a('ev_efficiency', DEFAULT_EV_EFFICIENCY),
            "ice_maintenance":     DEFAULT_ICE_MAINTENANCE,
            "ev_maintenance":      DEFAULT_EV_MAINTENANCE,
            "fuel_escalation":     DEFAULT_FUEL_ESCALATION_RATE,
            "discount_rate":       _a('discount_rate', 5.0),
            "battery_degradation": DEFAULT_BATTERY_DEGRADATION,
            "infra_cost_per_veh":  DEFAULT_INFRASTRUCTURE_COST_PER_VEHICLE,
            "analysis_period":     int(_a('analysis_period', 12) or 12),
        }
        for key, label, kind in _ASSUMPTION_DEFS:
            r0 = _ASSUMPTION_ROW[key] - 1  # 0-indexed
            ws.write(r0, 0, label, label_fmt)
            ws.write_number(r0, 1, float(seeds[key]), _in_fmt[kind])
        note_row = _ASSUMPTION_FIRST_ROW + len(_ASSUMPTION_DEFS)  # 1-indexed row after block
        ws.write(note_row, 0,
                 "Yellow cells are the single source of truth; the TCO & Financials sheet mirrors them.",
                 note_fmt)

        # ── Data quality summary KPIs ────────────────────────────────────────────
        row = note_row + 2  # 0-indexed cursor (note_row is 1-indexed → already one below)
        ws.merge_range(row, 0, row, 5, "Data Quality Summary", section_fmt)
        row += 1

        q_scores = [getattr(v, 'data_quality_score', 0.0) for v in vehicles]
        avg_quality = sum(q_scores) / len(q_scores) if q_scores else 0.0
        with_mpg = [v for v in vehicles if getattr(v.fuel_economy, 'combined_mpg', 0)]
        est_count = sum(1 for v in with_mpg if getattr(v.fuel_economy, 'mpg_is_estimate', False))
        pct_est = (est_count / len(with_mpg) * 100) if with_mpg else 0.0
        unresolved = sum(1 for v in vehicles
                         if not getattr(v, 'processing_success', True) or not v.vehicle_id.make)
        conf_vals = [getattr(v, 'match_confidence', 0.0) for v in vehicles
                     if getattr(v, 'match_confidence', 0.0) > 0]
        avg_conf = sum(conf_vals) / len(conf_vals) if conf_vals else 0.0

        kpis = [
            ("Avg Data Quality", f"{avg_quality:.0f}%"),
            ("MPG Estimated", f"{pct_est:.0f}%"),
            ("Unresolved VINs", f"{unresolved}"),
            ("Avg Match Conf.", f"{avg_conf:.0f}%"),
        ]
        for col, (label, value) in enumerate(kpis):
            ws.write(row, col, label, kpi_label_fmt)
            ws.write(row + 1, col, value, kpi_value_fmt)
        ws.write(row + 2, 0, "See the 'Infrastructure & Action Items' sheet for the vehicle-level data-gap punch list.", note_fmt)
        row += 4

        # ── Methodology & data sources ───────────────────────────────────────────
        ws.merge_range(row, 0, row, 5, "Methodology & Data Sources", section_fmt)
        row += 1
        methodology = (
            "VINs are decoded via the NHTSA vPIC API (make, model, year, GVWR, body class, fuel type). "
            "MPG is resolved through a tiered cascade: FuelEconomy.gov (light-duty) → commercial sources "
            "(Fuelly / EPA SmartWay / OEM) → an analyst-maintained SQLite reference DB → an EPA class-average "
            "estimate by GVWR bucket (flagged as an estimate). Each vehicle is classified for CARB Advanced "
            "Clean Fleets (ACF) and assigned a proposed electrification year via a score-based even-spread "
            "model. Financials (TCO, savings, payback) come from a year-by-year cash-flow model driven by the "
            "assumptions above. Incentives reflect the selected region and are applied to net cost."
        )
        ws.merge_range(row, 0, row + 4, 5, methodology, body_fmt)
        row += 6

        # ── ACF glossary ─────────────────────────────────────────────────────────
        ws.merge_range(row, 0, row, 5, "ACF Classification Glossary", section_fmt)
        row += 1
        glossary = [
            ("ZEV", "Already zero-emission — no replacement required."),
            ("A",   "Exempt, light-duty (GVWR ≤ 8,500 lbs)."),
            ("B",   "Mandate-subject, medium/heavy-duty — drives the compliance timeline."),
            ("C",   "Exempt by body type (dump truck, crane, concrete mixer, etc.)."),
            ("D",   "Emergency vehicle (PPV/SSV, ambulance, fire apparatus)."),
        ]
        for code, desc in glossary:
            ws.write(row, 0, code, workbook.add_format({'border': 1, 'bold': True, 'align': 'center'}))
            ws.merge_range(row, 1, row, 5, desc, label_fmt)
            row += 1
        row += 1

        # ── Disclaimer ───────────────────────────────────────────────────────────
        ws.merge_range(row, 0, row + 1, 5,
                       "Estimates are for planning purposes only and depend on input data quality and "
                       "assumptions. Figures flagged as estimates (e.g. EPA class-average MPG, synthetic "
                       "emissions projections) should be validated before procurement decisions.",
                       note_fmt)

    def _create_vehicle_data_sheet(self, workbook, vehicles, fields, title_format, header_format, cell_format, number_format):
        """Create the Vehicle Data sheet with all vehicle information.

        Data-quality columns (MPG Source, MPG Estimated, Match Confidence, Data
        Quality, Processing Status) come straight from to_row_dict(); this method
        highlights estimated/low-confidence/failed cells so analysts and clients
        can see at a glance which numbers are soft.
        """
        # Create worksheet
        worksheet = workbook.add_worksheet("Vehicle Data")

        # Quality-flag formats
        estimate_fmt = workbook.add_format({'border': 1, 'bg_color': '#FFF3E0'})   # amber — estimate
        warn_fmt     = workbook.add_format({'border': 1, 'bg_color': '#FDE0DC',
                                            'font_color': '#B71C1C'})              # red — low quality / failed

        def _pct(value):
            """Parse '85%' / '85' / 85 -> 85.0; None on failure."""
            try:
                return float(str(value).replace('%', '').strip())
            except (ValueError, TypeError):
                return None

        # Create user-friendly headers
        headers = [COLUMN_NAME_MAP.get(field, field) for field in fields]

        # Set column widths
        for i, field in enumerate(fields):
            worksheet.set_column(i, i, 15)  # Default width

            # Wider columns for specific fields
            if field in ["VIN", "Make", "Model", "BodyClass"]:
                worksheet.set_column(i, i, 20)
            elif field in ["Assumed Vehicle (Text)"]:
                worksheet.set_column(i, i, 30)

        # Add title
        worksheet.merge_range('A1:E1', "Fleet Vehicle Data", title_format)

        # Add headers
        for col, header in enumerate(headers):
            worksheet.write(2, col, header, header_format)

        # Add data
        for row, vehicle in enumerate(vehicles):
            data = vehicle.to_row_dict()

            for col, field in enumerate(fields):
                value = data.get(field, "")

                # Per-cell quality highlighting
                quality_fmt = None
                if field == "MPG Estimated" and str(value).strip().lower() == "yes":
                    quality_fmt = estimate_fmt
                elif field == "MPG Source" and getattr(vehicle.fuel_economy, "mpg_is_estimate", False):
                    quality_fmt = estimate_fmt
                elif field == "Processing Status" and str(value).strip().lower() == "failed":
                    quality_fmt = warn_fmt
                elif field == "Data Quality":
                    p = _pct(value)
                    if p is not None and p < 60:
                        quality_fmt = warn_fmt
                elif field == "Match Confidence":
                    p = _pct(value)
                    if p is not None and p < 60:
                        quality_fmt = estimate_fmt

                # Format numbers appropriately
                if field in ["MPG City", "MPG Highway", "MPG Combined", "CO2 emissions", "co2A"]:
                    try:
                        value = float(value)
                        worksheet.write(row + 3, col, value, quality_fmt or number_format)
                    except (ValueError, TypeError):
                        worksheet.write(row + 3, col, value, quality_fmt or cell_format)
                else:
                    worksheet.write(row + 3, col, value, quality_fmt or cell_format)

        # Freeze header row
        worksheet.freeze_panes(3, 0)

        # Auto-filter
        worksheet.autofilter(2, 0, 2 + len(vehicles), len(headers) - 1)
    
    def _create_fleet_overview_sheet(self, workbook, vehicles, analysis, charging, emissions,
                                     title_format, header_format, cell_format, number_format):
        """Fleet Overview — merges the old Summary + Summary Dashboard into one sheet.

        KPI band on top, then composition tables (ACF / fuel / make) each paired
        with a native Excel chart. Department-emissions and EV-year distributions
        are intentionally NOT duplicated here — they live on their canonical
        Emissions & Scenarios and Replacement & Capital Plan sheets.
        """
        SHEET = "Fleet Overview"
        ws = workbook.add_worksheet(SHEET)
        ws.set_column('A:A', 26)
        ws.set_column('B:C', 14)
        ws.hide_gridlines(2)

        kpi_label_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#EAF2FB', 'border': 1, 'align': 'center', 'font_size': 9})
        kpi_value_fmt = workbook.add_format({
            'bold': True, 'font_size': 16, 'align': 'center', 'border': 1})
        kpi_money_fmt = workbook.add_format({
            'bold': True, 'font_size': 16, 'align': 'center', 'border': 1, 'num_format': '$#,##0'})
        count_fmt = workbook.add_format({'border': 1, 'align': 'right'})

        ws.merge_range('A1:H1', "Fleet Electrification — Overview", title_format)

        # ── KPI band ─────────────────────────────────────────────────────────────
        mpg_vals = [v.fuel_economy.combined_mpg for v in vehicles
                    if getattr(v.fuel_economy, 'combined_mpg', 0)]
        avg_mpg = sum(mpg_vals) / len(mpg_vals) if mpg_vals else 0.0
        coverage = (len(mpg_vals) / len(vehicles) * 100) if vehicles else 0.0

        def _acf(v):
            return (v.custom_fields.get('_acf_code')
                    or v.custom_fields.get('ACF Category', '') or '').strip()
        acf_b = sum(1 for v in vehicles if _acf(v) == 'B')

        kpis = [
            ("Fleet Size", len(vehicles), kpi_value_fmt),
            ("Avg MPG", f"{avg_mpg:.1f}", kpi_value_fmt),
            ("MPG Coverage", f"{coverage:.0f}%", kpi_value_fmt),
            ("ACF-B Count", acf_b, kpi_value_fmt),
        ]
        if analysis is not None:
            kpis.append(("Total Savings", getattr(analysis, 'total_savings', 0) or 0, kpi_money_fmt))
            kpis.append(("Payback (yr)", f"{getattr(analysis, 'payback_period', 0) or 0:.1f}", kpi_value_fmt))
        for col, (label, value, vfmt) in enumerate(kpis):
            ws.write(2, col, label, kpi_label_fmt)
            ws.write(3, col, value, vfmt)

        # ── Composition tables + native charts ───────────────────────────────────
        def _dist(getter):
            d = {}
            for v in vehicles:
                k = getter(v)
                if k:
                    d[k] = d.get(k, 0) + 1
            return sorted(d.items(), key=lambda x: x[1], reverse=True)

        acf_dist  = _dist(_acf)
        fuel_dist = _dist(lambda v: v.vehicle_id.fuel_type)
        make_dist = _dist(lambda v: v.vehicle_id.make)[:10]

        def _write_table(start_row, title, rows, chart_type):
            """Write a 2-col table (category, count) and return chart anchored to the right."""
            ws.write(start_row, 0, title, header_format)
            ws.write(start_row, 1, "Count", header_format)
            r = start_row + 1
            for name, count in rows:
                ws.write(r, 0, name, cell_format)
                ws.write(r, 1, count, count_fmt)
                r += 1
            last = r - 1
            if rows:
                chart = workbook.add_chart({'type': chart_type})
                chart.add_series({
                    'name': title,
                    'categories': [SHEET, start_row + 1, 0, last, 0],
                    'values':     [SHEET, start_row + 1, 1, last, 1],
                    'data_labels': {'value': True} if chart_type == 'pie' else {},
                })
                chart.set_title({'name': title})
                chart.set_size({'width': 360, 'height': 220})
                if chart_type != 'pie':
                    chart.set_legend({'none': True})
                ws.insert_chart(start_row, 3, chart)
            return r + 1  # next free row (one blank separator)

        row = 6
        row = _write_table(row, "ACF Classification", acf_dist, 'pie')
        row = _write_table(row + 11, "Fuel Type", fuel_dist, 'pie')
        row = _write_table(row + 11, "Top Makes", make_dist, 'column')

    def _create_summary_sheet(self, workbook, vehicles, title_format, header_format):
        """Create a summary sheet with fleet statistics and charts."""
        # Create worksheet
        worksheet = workbook.add_worksheet("Summary")
        
        # Add title
        worksheet.merge_range('A1:D1', "Fleet Summary", title_format)
        
        # Set column widths
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 15)
        worksheet.set_column('C:D', 10)
        
        # Basic statistics
        row = 3
        worksheet.write(row, 0, "Total Vehicles:", header_format)
        worksheet.write(row, 1, len(vehicles))
        
        # Count makes and models
        makes = {}
        models = {}
        fuel_types = {}
        body_classes = {}
        
        # Calculate stats
        mpg_values = []
        co2_values = []
        
        for vehicle in vehicles:
            # Count makes
            make = vehicle.vehicle_id.make
            if make:
                makes[make] = makes.get(make, 0) + 1
            
            # Count models
            model = vehicle.vehicle_id.model
            if model:
                models[model] = models.get(model, 0) + 1
            
            # Count fuel types
            fuel_type = vehicle.vehicle_id.fuel_type
            if fuel_type:
                fuel_types[fuel_type] = fuel_types.get(fuel_type, 0) + 1
            
            # Count body classes
            body_class = vehicle.vehicle_id.body_class
            if body_class:
                body_classes[body_class] = body_classes.get(body_class, 0) + 1
            
            # MPG stats
            mpg = vehicle.fuel_economy.combined_mpg
            if mpg and mpg > 0:
                mpg_values.append(mpg)
            
            # CO2 stats
            co2 = vehicle.fuel_economy.co2_primary
            if co2 and co2 > 0:
                co2_values.append(co2)
        
        # Add make distribution
        row += 2
        worksheet.write(row, 0, "Make Distribution:", header_format)
        row += 1
        
        # Sort makes by count
        sorted_makes = sorted(makes.items(), key=lambda x: x[1], reverse=True)
        for make, count in sorted_makes[:10]:  # Top 10
            worksheet.write(row, 0, make)
            worksheet.write(row, 1, count)
            # Add simple bar using cell background
            for i in range(min(count, 10)):
                worksheet.write(row, 2 + i, "", workbook.add_format({'bg_color': '#5B7553'}))  # PRIMARY_HEX_3
            row += 1
        
        # Add fuel type distribution
        row += 2
        worksheet.write(row, 0, "Fuel Type Distribution:", header_format)
        row += 1
        
        sorted_fuel_types = sorted(fuel_types.items(), key=lambda x: x[1], reverse=True)
        for fuel_type, count in sorted_fuel_types:
            worksheet.write(row, 0, fuel_type)
            worksheet.write(row, 1, count)
            # Add simple bar using cell background
            for i in range(min(count, 10)):
                worksheet.write(row, 2 + i, "", workbook.add_format({'bg_color': '#C45911'}))  # SECONDARY_HEX_1
            row += 1
        
        # Add MPG stats
        row += 2
        worksheet.write(row, 0, "MPG Statistics:", header_format)
        row += 1
        
        if mpg_values:
            avg_mpg = sum(mpg_values) / len(mpg_values)
            min_mpg = min(mpg_values)
            max_mpg = max(mpg_values)
            
            worksheet.write(row, 0, "Average MPG:")
            worksheet.write(row, 1, avg_mpg)
            row += 1
            
            worksheet.write(row, 0, "Minimum MPG:")
            worksheet.write(row, 1, min_mpg)
            row += 1
            
            worksheet.write(row, 0, "Maximum MPG:")
            worksheet.write(row, 1, max_mpg)
            row += 1
        else:
            worksheet.write(row, 0, "No MPG data available")
            row += 1
        
        # Add CO2 stats
        row += 2
        worksheet.write(row, 0, "CO2 Emissions Statistics:", header_format)
        row += 1
        
        if co2_values:
            avg_co2 = sum(co2_values) / len(co2_values)
            min_co2 = min(co2_values)
            max_co2 = max(co2_values)
            
            worksheet.write(row, 0, "Average CO2 (g/mile):")
            worksheet.write(row, 1, avg_co2)
            row += 1
            
            worksheet.write(row, 0, "Minimum CO2 (g/mile):")
            worksheet.write(row, 1, min_co2)
            row += 1
            
            worksheet.write(row, 0, "Maximum CO2 (g/mile):")
            worksheet.write(row, 1, max_co2)
            row += 1
        else:
            worksheet.write(row, 0, "No CO2 data available")
            row += 1
    
    def _create_electrification_sheet(self, workbook, analysis, title_format, header_format, cell_format, number_format):
        """Create Electrification Analysis sheet."""
        # Create worksheet
        worksheet = workbook.add_worksheet("Electrification Analysis")
        
        # Add title
        worksheet.merge_range('A1:F1', "Fleet Electrification Analysis", title_format)
        
        # Set column widths
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:F', 15)
        
        # Add summary section
        row = 3
        worksheet.write(row, 0, "Analysis Parameters:", header_format)
        row += 1
        
        worksheet.write(row, 0, "Gas Price ($/gal):")
        worksheet.write(row, 1, analysis.gas_price)
        row += 1
        
        worksheet.write(row, 0, "Electricity Price ($/kWh):")
        worksheet.write(row, 1, analysis.electricity_price)
        row += 1
        
        worksheet.write(row, 0, "EV Efficiency (kWh/mile):")
        worksheet.write(row, 1, analysis.ev_efficiency)
        row += 1
        
        worksheet.write(row, 0, "Analysis Period (years):")
        worksheet.write(row, 1, analysis.analysis_period)
        row += 1
        
        worksheet.write(row, 0, "Discount Rate (%):")
        worksheet.write(row, 1, analysis.discount_rate)
        row += 1
        
        # Add results section
        row += 2
        worksheet.write(row, 0, "Analysis Results:", header_format)
        row += 1
        
        worksheet.write(row, 0, "Total CO2 Savings (tons):")
        worksheet.write(row, 1, analysis.co2_savings)
        row += 1
        
        worksheet.write(row, 0, "Fuel Cost Savings ($):")
        worksheet.write(row, 1, analysis.fuel_cost_savings)
        row += 1
        
        worksheet.write(row, 0, "Maintenance Savings ($):")
        worksheet.write(row, 1, analysis.maintenance_savings)
        row += 1
        
        worksheet.write(row, 0, "Total Savings ($):")
        worksheet.write(row, 1, analysis.total_savings)
        row += 1
        
        worksheet.write(row, 0, "Payback Period (years):")
        worksheet.write(row, 1, analysis.payback_period)
        row += 1
        
        # Add vehicle results table
        row += 2
        worksheet.write(row, 0, "Vehicle-Level Results:", header_format)
        row += 1
        
        # Create headers
        headers = [
            "Vehicle", "Annual Mileage", "MPG", "Annual Fuel Savings ($)", 
            "Lifetime Fuel Savings ($)", "CO2 Reduction (tons)", "Payback Period (years)"
        ]
        
        for col, header in enumerate(headers):
            worksheet.write(row, col, header, header_format)
        
        row += 1
        
        # Sort vehicles by savings (if we have the prioritized list)
        vehicles_to_show = []
        for vin in analysis.prioritized_vehicles:
            if vin in analysis.vehicle_results:
                vehicles_to_show.append((vin, analysis.vehicle_results[vin]))
        
        # If no prioritized list, just sort by savings
        if not vehicles_to_show:
            vehicles_to_show = sorted(
                analysis.vehicle_results.items(),
                key=lambda x: x[1].get("total_npv_savings", 0),
                reverse=True
            )
        
        # Add data for top vehicles
        for vin, data in vehicles_to_show:
            display_name = data.get("display_name", vin)
            annual_mileage = data.get("annual_mileage", 0)
            mpg = data.get("mpg", 0)
            annual_savings = data.get("annual_fuel_savings", 0)
            total_savings = data.get("total_fuel_savings", 0)
            co2_reduction = data.get("total_co2_reduction", 0)
            
            # Calculate payback (assuming $15k premium)
            ev_premium = 15000
            annual_total_savings = annual_savings + data.get("annual_maintenance_savings", 0)
            payback = ev_premium / annual_total_savings if annual_total_savings > 0 else float('inf')
            
            worksheet.write(row, 0, display_name, cell_format)
            worksheet.write(row, 1, annual_mileage, number_format)
            worksheet.write(row, 2, mpg, number_format)
            worksheet.write(row, 3, annual_savings, number_format)
            worksheet.write(row, 4, total_savings, number_format)
            worksheet.write(row, 5, co2_reduction, number_format)
            
            if payback == float('inf'):
                worksheet.write(row, 6, "N/A", cell_format)
            else:
                worksheet.write(row, 6, payback, number_format)
            
            row += 1
    
    def _create_charging_sheet(self, workbook, charging, title_format, header_format, cell_format, number_format):
        """Create Charging Infrastructure sheet."""
        # Create worksheet
        worksheet = workbook.add_worksheet("Charging Infrastructure")
        
        # Add title
        worksheet.merge_range('A1:D1', "Charging Infrastructure Analysis", title_format)
        
        # Set column widths
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:D', 15)
        
        # Add parameters section
        row = 3
        worksheet.write(row, 0, "Analysis Parameters:", header_format)
        row += 1
        
        worksheet.write(row, 0, "Usage Pattern:")
        worksheet.write(row, 1, charging.daily_usage_pattern.capitalize())
        row += 1
        
        worksheet.write(row, 0, "Charging Window:")
        window_text = f"{charging.charging_window[0]}:00 to {charging.charging_window[1]}:00"
        worksheet.write(row, 1, window_text)
        row += 1
        
        # Add results section
        row += 2
        worksheet.write(row, 0, "Infrastructure Requirements:", header_format)
        row += 1
        
        worksheet.write(row, 0, "Level 2 Chargers:")
        worksheet.write(row, 1, charging.level2_chargers_needed)
        row += 1
        
        worksheet.write(row, 0, "DC Fast Chargers:")
        worksheet.write(row, 1, charging.dcfc_chargers_needed)
        row += 1
        
        worksheet.write(row, 0, "Maximum Power Required (kW):")
        worksheet.write(row, 1, charging.max_power_required)
        row += 1
        
        worksheet.write(row, 0, "Estimated Cost ($):")
        worksheet.write(row, 1, charging.estimated_installation_cost)
        row += 1
        
        # Add recommended layout
        if charging.recommended_layout and "zones" in charging.recommended_layout:
            row += 2
            worksheet.write(row, 0, "Recommended Layout:", header_format)
            row += 1
            
            for i, zone in enumerate(charging.recommended_layout["zones"]):
                worksheet.write(row, 0, f"Zone {i+1}: {zone.get('name', '')}")
                row += 1
                
                worksheet.write(row, 0, "Level 2 Chargers:")
                worksheet.write(row, 1, zone.get("level2_chargers", 0))
                row += 1
                
                worksheet.write(row, 0, "DC Fast Chargers:")
                worksheet.write(row, 1, zone.get("dcfc_chargers", 0))
                row += 1
                
                worksheet.write(row, 0, "Power Required (kW):")
                worksheet.write(row, 1, zone.get("power_required", 0))
                row += 1
        
        # Add phasing plan
        if charging.recommended_layout and "phasing" in charging.recommended_layout:
            row += 2
            worksheet.write(row, 0, "Implementation Phasing:", header_format)
            row += 1
            
            headers = ["Phase", "Level 2 Chargers", "DC Fast Chargers", "Estimated Cost ($)"]
            for col, header in enumerate(headers):
                worksheet.write(row, col, header, header_format)
            row += 1
            
            for phase in charging.recommended_layout["phasing"]:
                worksheet.write(row, 0, f"Phase {phase.get('phase', '')}")
                worksheet.write(row, 1, phase.get("level2_chargers", 0))
                worksheet.write(row, 2, phase.get("dcfc_chargers", 0))
                worksheet.write(row, 3, phase.get("estimated_cost", 0))
                row += 1
    
    def _create_emissions_sheet(self, workbook, emissions, title_format, header_format, cell_format, number_format):
        """Create Emissions Inventory sheet."""
        # Create worksheet
        worksheet = workbook.add_worksheet("Emissions Inventory")
        
        # Add title
        worksheet.merge_range('A1:D1', "Fleet Emissions Inventory", title_format)
        
        # Set column widths
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:D', 15)
        
        # Add summary section
        row = 3
        worksheet.write(row, 0, "Inventory Summary:", header_format)
        row += 1
        
        worksheet.write(row, 0, "Inventory Year:")
        worksheet.write(row, 1, emissions.inventory_year)
        row += 1
        
        worksheet.write(row, 0, "Total Emissions (tons CO2e):")
        worksheet.write(row, 1, emissions.total_emissions)
        row += 1
        
        worksheet.write(row, 0, "Baseline Year:")
        worksheet.write(row, 1, emissions.baseline_year)
        row += 1
        
        worksheet.write(row, 0, "Target Year:")
        worksheet.write(row, 1, emissions.target_year)
        row += 1
        
        worksheet.write(row, 0, "Reduction Target (%):")
        worksheet.write(row, 1, emissions.reduction_target)
        row += 1
        
        # Add historical data
        if emissions.historical_data:
            row += 2
            worksheet.write(row, 0, "Historical Emissions:", header_format)
            row += 1
            
            headers = ["Year", "Emissions (tons CO2e)"]
            for col, header in enumerate(headers):
                worksheet.write(row, col, header, header_format)
            row += 1
            
            for year, value in sorted(emissions.historical_data.items()):
                worksheet.write(row, 0, year)
                worksheet.write(row, 1, value, number_format)
                row += 1
        
        # Add projected data
        if emissions.projected_emissions:
            row += 2
            worksheet.write(row, 0, "Projected Emissions:", header_format)
            row += 1
            
            headers = ["Year", "Emissions (tons CO2e)"]
            for col, header in enumerate(headers):
                worksheet.write(row, col, header, header_format)
            row += 1
            
            for year, value in sorted(emissions.projected_emissions.items()):
                worksheet.write(row, 0, year)
                worksheet.write(row, 1, value, number_format)
                row += 1
        
        # Add emissions by department
        if emissions.by_department:
            row += 2
            worksheet.write(row, 0, "Emissions by Department:", header_format)
            row += 1
            
            headers = ["Department", "Emissions (tons CO2e)", "Percentage"]
            for col, header in enumerate(headers):
                worksheet.write(row, col, header, header_format)
            row += 1
            
            for dept, value in sorted(emissions.by_department.items(), key=lambda x: x[1], reverse=True):
                percentage = (value / emissions.total_emissions * 100) if emissions.total_emissions > 0 else 0
                
                worksheet.write(row, 0, dept)
                worksheet.write(row, 1, value, number_format)
                worksheet.write(row, 2, f"{percentage:.1f}%")
                row += 1
        
        # Add emissions by vehicle type
        if emissions.by_vehicle_type:
            row += 2
            worksheet.write(row, 0, "Emissions by Vehicle Type:", header_format)
            row += 1
            
            headers = ["Vehicle Type", "Emissions (tons CO2e)", "Percentage"]
            for col, header in enumerate(headers):
                worksheet.write(row, col, header, header_format)
            row += 1
            
            for vtype, value in sorted(emissions.by_vehicle_type.items(), key=lambda x: x[1], reverse=True):
                percentage = (value / emissions.total_emissions * 100) if emissions.total_emissions > 0 else 0
                
                worksheet.write(row, 0, vtype)
                worksheet.write(row, 1, value, number_format)
                worksheet.write(row, 2, f"{percentage:.1f}%")
                row += 1
        
        # Add emissions by fuel type
        if emissions.by_fuel_type:
            row += 2
            worksheet.write(row, 0, "Emissions by Fuel Type:", header_format)
            row += 1
            
            headers = ["Fuel Type", "Emissions (tons CO2e)", "Percentage"]
            for col, header in enumerate(headers):
                worksheet.write(row, col, header, header_format)
            row += 1
            
            for ftype, value in sorted(emissions.by_fuel_type.items(), key=lambda x: x[1], reverse=True):
                percentage = (value / emissions.total_emissions * 100) if emissions.total_emissions > 0 else 0
                
                worksheet.write(row, 0, ftype)
                worksheet.write(row, 1, value, number_format)
                worksheet.write(row, 2, f"{percentage:.1f}%")
                row += 1


    def _create_tco_model_sheet(self, workbook, analysis, vehicles, title_format, header_format, cell_format, number_format):
        """Create TCO Model sheet with a live-formula assumptions block and year-by-year cash flows.

        Layout (1-indexed rows as Excel sees them):
          Row 1       : Title (merged A1:J1)
          Row 2       : blank
          Row 3       : "Assumptions" header
          Rows 4-14   : Editable assumption cells in column B (yellow), labels in column A
          Rows 15-16  : Read-only helper cells (avg fleet MPG, avg annual mileage)
          Row 17      : blank
          Row 18      : Column headers for the cash-flow table
          Rows 19+    : One row per year; all cost/savings cells use =formulas referencing B4:B16
          Last rows   : Summary KPIs (also formula-based)

        Assumption cells (all in column B):
          B4  = Gas price ($/gal)               ← editable (yellow)
          B5  = Electricity price ($/kWh)        ← editable
          B6  = EV efficiency (kWh/mile)         ← editable
          B7  = ICE maintenance ($/mile)         ← editable
          B8  = EV maintenance ($/mile)          ← editable
          B9  = Fuel escalation rate (% / yr)    ← editable
          B10 = Discount rate (%)                ← editable
          B11 = Battery degradation (% / yr)     ← editable
          B12 = Infrastructure cost/vehicle ($)  ← editable
          B13 = Fleet size (vehicles)            ← editable
          B14 = Analysis period (years)          ← editable
          B15 = Avg fleet MPG                    ← read-only (fleet aggregate)
          B16 = Avg annual mileage (miles)       ← read-only (fleet aggregate)

        Anchor cells K–O hold Year-1 base cost formulas that reference B4:B16.
        Changing any editable assumption recalculates the entire model.
        """
        from settings import (
            DEFAULT_GAS_PRICE, DEFAULT_ELECTRICITY_PRICE, DEFAULT_EV_EFFICIENCY,
            DEFAULT_ICE_MAINTENANCE, DEFAULT_EV_MAINTENANCE,
            DEFAULT_FUEL_ESCALATION_RATE, DEFAULT_BATTERY_DEGRADATION,
            DEFAULT_INFRASTRUCTURE_COST_PER_VEHICLE
        )

        ws = workbook.add_worksheet("TCO Model")

        # ── Formats ─────────────────────────────────────────────────────────────
        currency_fmt = workbook.add_format({'border': 1, 'num_format': '$#,##0', 'align': 'right'})
        pct_fmt      = workbook.add_format({'border': 1, 'num_format': '0.00%', 'align': 'right'})
        num_fmt      = workbook.add_format({'border': 1, 'num_format': '#,##0.00', 'align': 'right'})
        input_fmt    = workbook.add_format({
            'border': 2, 'num_format': '#,##0.00', 'align': 'right',
            'bg_color': '#EAF2FB', 'bold': False
        })
        input_pct_fmt = workbook.add_format({
            'border': 2, 'num_format': '0.00', 'align': 'right',
            'bg_color': '#EAF2FB'
        })
        input_int_fmt = workbook.add_format({
            'border': 2, 'num_format': '#,##0', 'align': 'right',
            'bg_color': '#EAF2FB'
        })
        label_fmt    = workbook.add_format({'border': 1, 'bold': False, 'align': 'left'})
        note_fmt     = workbook.add_format({'border': 1, 'italic': True, 'font_color': '#2471A3'})
        payback_fmt  = workbook.add_format({
            'border': 1, 'bold': True,
            'bg_color': '#D5F5E3', 'font_color': '#1E8449'
        })
        hint_fmt     = workbook.add_format({
            'italic': True, 'font_size': 9, 'font_color': '#7F8C8D'
        })
        kpi_title_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#3C465A',
            'font_color': 'white', 'border': 1, 'align': 'center'
        })

        # ── Column widths ────────────────────────────────────────────────────────
        ws.set_column('A:A', 30)   # Label
        ws.set_column('B:B', 18)   # Input / Year
        ws.set_column('C:C', 18)   # ICE Annual
        ws.set_column('D:D', 18)   # EV Annual
        ws.set_column('E:E', 18)   # Annual Savings
        ws.set_column('F:F', 18)   # ICE Cumulative
        ws.set_column('G:G', 18)   # EV Cumulative
        ws.set_column('H:H', 18)   # Cumulative Savings
        ws.set_column('I:I', 20)   # NPV Savings
        ws.set_column('J:J', 22)   # Notes

        # ── Title ────────────────────────────────────────────────────────────────
        ws.merge_range('A1:J1', "Fleet TCO Model — Year-by-Year Cash Flows (Live Formula Model)", title_format)

        # ── Assumptions block (rows 3–15, 1-indexed) ─────────────────────────────
        # In xlsxwriter, row/col are 0-indexed: row 3 (1-idx) = row 2 (0-idx)
        ASSUMP_HEADER_ROW = 2   # 0-idx → Excel row 3
        ASSUMP_START_ROW  = 3   # 0-idx → Excel row 4  (first input row)

        ws.merge_range(ASSUMP_HEADER_ROW, 0, ASSUMP_HEADER_ROW, 9,
                       "⚙ Assumptions — Edit yellow cells to recalculate the model", header_format)
        ws.write(ASSUMP_HEADER_ROW + 1, 0,
                 "Changes to yellow cells instantly update all formula cells below.",
                 hint_fmt)

        # Derive baseline values from the first vehicle's cash flow data if possible
        cash_flows = analysis.fleet_cash_flows or []
        # Read actual params from first year-1 flow if present; fall back to defaults
        y1 = next((cf for cf in cash_flows if cf.get('year') == 1), {})
        gas_price_val   = y1.get('gas_price', DEFAULT_GAS_PRICE)
        elec_price_val  = y1.get('electricity_price', DEFAULT_ELECTRICITY_PRICE)
        n_years         = max((cf.get('year', 0) for cf in cash_flows), default=0)

        # Compute fleet-level aggregates needed for anchor cell formulas
        vehicles_with_mpg = [
            v for v in vehicles
            if v.fuel_economy and v.fuel_economy.combined_mpg and v.fuel_economy.combined_mpg > 0
        ]
        avg_mpg = (
            sum(v.fuel_economy.combined_mpg for v in vehicles_with_mpg) / len(vehicles_with_mpg)
            if vehicles_with_mpg else 20.0
        )
        vehicles_with_mileage = [
            v for v in vehicles
            if getattr(v, 'annual_mileage_miles', None) and v.annual_mileage_miles > 0
        ]
        avg_mileage = (
            sum(v.annual_mileage_miles for v in vehicles_with_mileage) / len(vehicles_with_mileage)
            if vehicles_with_mileage else 15000.0
        )

        # Assumption rows: (label, value, format, cell_hint)
        # First 11 rows are editable (yellow); last 2 are read-only fleet aggregates.
        assumptions_def = [
            ("Gas Price ($/gal)",              gas_price_val,                              input_fmt,     "B4"),
            ("Electricity Price ($/kWh)",       elec_price_val,                             input_fmt,     "B5"),
            ("EV Efficiency (kWh/mile)",        DEFAULT_EV_EFFICIENCY,                      input_fmt,     "B6"),
            ("ICE Maintenance ($/mile)",        DEFAULT_ICE_MAINTENANCE,                    input_fmt,     "B7"),
            ("EV Maintenance ($/mile)",         DEFAULT_EV_MAINTENANCE,                     input_fmt,     "B8"),
            ("Fuel Escalation Rate (%/yr)",     DEFAULT_FUEL_ESCALATION_RATE,               input_pct_fmt, "B9"),
            ("Discount Rate (%/yr)",            analysis.discount_rate if hasattr(analysis, 'discount_rate') else 5.0, input_pct_fmt, "B10"),
            ("Battery Degradation (%/yr)",      DEFAULT_BATTERY_DEGRADATION,                input_pct_fmt, "B11"),
            ("Infrastructure Cost/Vehicle ($)", DEFAULT_INFRASTRUCTURE_COST_PER_VEHICLE,    input_fmt,     "B12"),
            ("Fleet Size (vehicles)",           len(vehicles),                              input_int_fmt, "B13"),
            ("Analysis Period (years)",         n_years or 12,                              input_int_fmt, "B14"),
            # Read-only fleet aggregates — used by anchor formulas in columns K–O
            ("Avg Fleet MPG (read-only)",       round(avg_mpg, 2),                          input_fmt,     "B15"),
            ("Avg Annual Mileage (read-only)",  round(avg_mileage, 0),                      input_int_fmt, "B16"),
        ]

        # Named cells for formula references (0-indexed row, col 1 = column B)
        ASSUMP_ROWS = {}  # key → 0-indexed row number
        for i, (label, value, fmt, cell_hint) in enumerate(assumptions_def):
            r = ASSUMP_START_ROW + i
            ws.write(r, 0, label, label_fmt)
            ws.write(r, 1, value, fmt)
            ASSUMP_ROWS[cell_hint] = r  # store for formula building

        blank_row = ASSUMP_START_ROW + len(assumptions_def)  # one blank separator row

        # ── Cash-flow table ──────────────────────────────────────────────────────
        TABLE_HEADER_ROW = blank_row + 1
        TABLE_DATA_START = TABLE_HEADER_ROW + 1

        col_headers = [
            "Year", "ICE Annual Cost", "EV Annual Cost", "Annual Savings",
            "ICE Cumulative", "EV Cumulative", "Cumulative Savings",
            "NPV Savings (yr)", "Notes"
        ]
        for col, h in enumerate(col_headers):
            ws.write(TABLE_HEADER_ROW, col + 1, h, header_format)
        # col A (0) left blank on header row — used for row labels in assumption block
        ws.write(TABLE_HEADER_ROW, 0, "", header_format)

        # Helper: Excel column letter from 0-indexed col number
        def col_letter(c):
            letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            return letters[c] if c < 26 else letters[c // 26 - 1] + letters[c % 26]

        # Build assumption cell references (1-indexed Excel addresses)
        def aref(key):
            """Return absolute $B$N reference for an assumption cell."""
            r0 = ASSUMP_ROWS[key]
            return f"$B${r0 + 1}"   # +1 because Excel rows are 1-indexed

        gas_ref   = aref("B4")
        elec_ref  = aref("B5")
        eff_ref   = aref("B6")
        ice_m_ref = aref("B7")
        ev_m_ref  = aref("B8")
        esc_ref   = aref("B9")
        disc_ref  = aref("B10")
        batt_ref  = aref("B11")
        infra_ref = aref("B12")
        size_ref  = aref("B13")
        yrs_ref   = aref("B14")
        mpg_ref   = aref("B15")   # avg fleet MPG (read-only helper)
        mi_ref    = aref("B16")   # avg annual mileage (read-only helper)

        # We also need per-vehicle MPG average to compute fuel costs.
        # Since we aggregate across the fleet, use the actual fleet_cash_flows values
        # for Year 0 purchase and Years 1-N operating formulas.
        #
        # Strategy: write the year-0 row as static purchase amounts (no per-unit MPG
        # formula possible at fleet level), then for years 1-N write formulas that
        # scale the Year-1 actuals by the escalation / degradation factors.
        # This gives consultants live escalation / discount sensitivity.
        #
        # For ICE and EV base costs we use the Year-1 actuals as anchors; the
        # formula for year Y scales them by (1 + esc/100)^(Y-1) for fuel and
        # keeps maintenance flat (no formula for mileage change — KISS).

        # Find year-0 and year-1 cash flows
        cf0 = next((cf for cf in cash_flows if cf.get('year') == 0), {})
        cf1 = next((cf for cf in cash_flows if cf.get('year') == 1), {})

        ice_y0_total = cf0.get('ice_total', 0)
        ev_y0_total  = cf0.get('ev_total', 0)

        # Year-1 component breakdown (fleet aggregate)
        ice_y1_fuel   = cf1.get('ice_fuel', 0)
        ice_y1_maint  = cf1.get('ice_maintenance', 0)
        ev_y1_fuel    = cf1.get('ev_fuel', 0)
        ev_y1_maint   = cf1.get('ev_maintenance', 0)
        ev_y1_infra   = cf1.get('ev_infrastructure', 0)

        # Anchor cells: Year-1 base costs as live formulas referencing the assumptions block.
        # Stored off-screen in columns K-O so they update when consultants change B4:B16.
        #
        # Formula logic (fleet-level totals for year 1):
        #   ICE fuel     = (avg_mileage / avg_mpg) × gas_price × fleet_size
        #   ICE maint    = avg_mileage × ICE_maintenance_rate × fleet_size
        #   EV fuel      = avg_mileage × ev_efficiency × electricity_price × fleet_size
        #   EV maint     = avg_mileage × EV_maintenance_rate × fleet_size
        #   EV infra     = infra_cost_per_vehicle × fleet_size
        #
        # B15 = avg fleet MPG, B16 = avg annual mileage (both stored as read-only helpers above)
        ANCHOR_COL = 10  # column K (0-indexed)
        ws.write(TABLE_HEADER_ROW, ANCHOR_COL,     "⟵ Anchor: ICE fuel yr1 (live formula)",  hint_fmt)
        ws.write(TABLE_HEADER_ROW, ANCHOR_COL + 1, "⟵ Anchor: ICE maint yr1 (live formula)", hint_fmt)
        ws.write(TABLE_HEADER_ROW, ANCHOR_COL + 2, "⟵ Anchor: EV fuel yr1 (live formula)",   hint_fmt)
        ws.write(TABLE_HEADER_ROW, ANCHOR_COL + 3, "⟵ Anchor: EV maint yr1 (live formula)",  hint_fmt)
        ws.write(TABLE_HEADER_ROW, ANCHOR_COL + 4, "⟵ Anchor: EV infra yr1 (live formula)",  hint_fmt)

        ANCHOR_ROW = TABLE_DATA_START  # Year-0 row; anchor formulas go here
        ws.write_formula(ANCHOR_ROW, ANCHOR_COL,
                         f"={mi_ref}/{mpg_ref}*{gas_ref}*{size_ref}",
                         input_fmt, ice_y1_fuel)
        ws.write_formula(ANCHOR_ROW, ANCHOR_COL + 1,
                         f"={mi_ref}*{ice_m_ref}*{size_ref}",
                         input_fmt, ice_y1_maint)
        ws.write_formula(ANCHOR_ROW, ANCHOR_COL + 2,
                         f"={mi_ref}*{eff_ref}*{elec_ref}*{size_ref}",
                         input_fmt, ev_y1_fuel)
        ws.write_formula(ANCHOR_ROW, ANCHOR_COL + 3,
                         f"={mi_ref}*{ev_m_ref}*{size_ref}",
                         input_fmt, ev_y1_maint)
        ws.write_formula(ANCHOR_ROW, ANCHOR_COL + 4,
                         f"={infra_ref}*{size_ref}",
                         input_fmt, ev_y1_infra)

        # Helper cell references
        def acref(offset):
            """Absolute reference to anchor cell at ANCHOR_ROW, ANCHOR_COL+offset."""
            r_excel = ANCHOR_ROW + 1   # 1-indexed
            c_letter = col_letter(ANCHOR_COL + offset)
            return f"${c_letter}${r_excel}"

        ice_fuel_anchor   = acref(0)
        ice_maint_anchor  = acref(1)
        ev_fuel_anchor    = acref(2)
        ev_maint_anchor   = acref(3)
        ev_infra_anchor   = acref(4)

        # ── Write year-by-year rows ──────────────────────────────────────────────
        # Year 0 row: static purchase costs (no operating formulas)
        r0_excel = TABLE_DATA_START
        ws.write(r0_excel, 0, "Year 0 — Purchase", label_fmt)
        ws.write(r0_excel, 1, 0,           cell_format)          # Year number
        ws.write(r0_excel, 2, ice_y0_total, currency_fmt)         # ICE Annual (purchase)
        ws.write(r0_excel, 3, ev_y0_total,  currency_fmt)         # EV Annual (purchase)
        ws.write_formula(r0_excel, 4, f"={col_letter(2)}{r0_excel+1}-{col_letter(3)}{r0_excel+1}", currency_fmt)  # Savings
        ws.write(r0_excel, 5, ice_y0_total, currency_fmt)         # ICE Cumulative
        ws.write(r0_excel, 6, ev_y0_total,  currency_fmt)         # EV Cumulative
        ws.write_formula(r0_excel, 7, f"={col_letter(5)}{r0_excel+1}-{col_letter(6)}{r0_excel+1}", currency_fmt)
        ws.write(r0_excel, 8, 0,            currency_fmt)         # NPV
        ws.write(r0_excel, 9, "Year 0: Purchase costs", note_fmt)

        # Column letter shortcuts for the cash-flow columns (B=1, C=2, ..., I=8)
        # col index in worksheet: B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9
        C_YEAR    = 1
        C_ICE_ANN = 2
        C_EV_ANN  = 3
        C_ANN_SAV = 4
        C_ICE_CUM = 5
        C_EV_CUM  = 6
        C_CUM_SAV = 7
        C_NPV     = 8
        C_NOTE    = 9

        prev_ice_cum_ref = f"{col_letter(C_ICE_CUM)}{r0_excel + 1}"
        prev_ev_cum_ref  = f"{col_letter(C_EV_CUM)}{r0_excel + 1}"
        prev_cum_sav_ref = f"{col_letter(C_CUM_SAV)}{r0_excel + 1}"

        n_data_rows = len([cf for cf in cash_flows if cf.get('year', -1) >= 1]) or (n_years or 12)

        for i in range(1, n_data_rows + 1):
            row = TABLE_DATA_START + i
            row_excel = row + 1  # 1-indexed

            prev_row_excel = row_excel - 1

            # Escalation factor: (1 + esc%/100)^(year-1)
            esc_factor = f"(1+{esc_ref}/100)^({i}-1)"

            # Degradation factor for EV efficiency: (1 + batt%/100)*(year-1)
            # EV fuel = anchor_ev_fuel * (1 + batt%/100)*(year-1) * esc_factor / esc_factor_yr1
            # Simpler: EV fuel yr Y = (ev_fuel_anchor * (1 + batt_ref/100*(Y-1))) * (1+esc_ref/100)^(Y-1)
            # anchor is already at yr1 gas/elec prices; we scale fuel by esc and efficiency by batt deg
            batt_factor = f"(1+{batt_ref}/100*({i}-1))"

            # ICE annual = (ice_fuel_anchor * esc_factor) + ice_maint_anchor
            ice_ann_formula = f"={ice_fuel_anchor}*{esc_factor}+{ice_maint_anchor}"

            # EV annual = (ev_fuel_anchor * batt_factor * esc_factor) + ev_maint_anchor + ev_infra_anchor
            ev_ann_formula  = f"={ev_fuel_anchor}*{batt_factor}*{esc_factor}+{ev_maint_anchor}+{ev_infra_anchor}"

            # Annual savings
            ice_ann_ref = f"{col_letter(C_ICE_ANN)}{row_excel}"
            ev_ann_ref  = f"{col_letter(C_EV_ANN)}{row_excel}"
            ann_sav_formula = f"={ice_ann_ref}-{ev_ann_ref}"

            # Cumulative
            ice_cum_formula = f"={col_letter(C_ICE_CUM)}{prev_row_excel}+{ice_ann_ref}"
            ev_cum_formula  = f"={col_letter(C_EV_CUM)}{prev_row_excel}+{ev_ann_ref}"

            ice_cum_ref = f"{col_letter(C_ICE_CUM)}{row_excel}"
            ev_cum_ref  = f"{col_letter(C_EV_CUM)}{row_excel}"
            cum_sav_formula = f"={ice_cum_ref}-{ev_cum_ref}"

            # NPV of annual savings: ann_sav / (1 + disc%)^year
            npv_formula = f"={col_letter(C_ANN_SAV)}{row_excel}/(1+{disc_ref}/100)^{i}"

            # Payback note: flag when cumulative savings crosses from negative to positive
            prev_cum_sav_cell = f"{col_letter(C_CUM_SAV)}{prev_row_excel}"
            cum_sav_cell      = f"{col_letter(C_CUM_SAV)}{row_excel}"
            note_formula = (
                f'=IF(AND({prev_cum_sav_cell}<0,{cum_sav_cell}>=0),"✓ PAYBACK YEAR","")'
            )

            ws.write(row, 0, f"Year {i}", label_fmt)
            ws.write(row, C_YEAR,    i,                cell_format)
            ws.write_formula(row, C_ICE_ANN, ice_ann_formula, currency_fmt)
            ws.write_formula(row, C_EV_ANN,  ev_ann_formula,  currency_fmt)
            ws.write_formula(row, C_ANN_SAV, ann_sav_formula, currency_fmt)
            ws.write_formula(row, C_ICE_CUM, ice_cum_formula, currency_fmt)
            ws.write_formula(row, C_EV_CUM,  ev_cum_formula,  currency_fmt)
            ws.write_formula(row, C_CUM_SAV, cum_sav_formula, currency_fmt)
            ws.write_formula(row, C_NPV,     npv_formula,     currency_fmt)
            ws.write_formula(row, C_NOTE,    note_formula,    payback_fmt)

        # ── Summary KPIs (formula-based) ──────────────────────────────────────────
        last_data_row = TABLE_DATA_START + n_data_rows + 1  # 1-indexed Excel row of last data row
        last_data_row_excel = TABLE_DATA_START + n_data_rows  # 0-indexed

        SUMMARY_START = last_data_row_excel + 2
        ws.merge_range(SUMMARY_START, 0, SUMMARY_START, 9,
                       "Summary — Formula-Linked KPIs", header_format)

        ice_cum_last = f"{col_letter(C_ICE_CUM)}{TABLE_DATA_START + n_data_rows + 1}"
        ev_cum_last  = f"{col_letter(C_EV_CUM)}{TABLE_DATA_START + n_data_rows + 1}"
        npv_col_range = f"{col_letter(C_NPV)}{TABLE_DATA_START + 2}:{col_letter(C_NPV)}{TABLE_DATA_START + n_data_rows + 1}"

        kpis = [
            ("Total ICE TCO (fleet)",     f"={ice_cum_last}",                 currency_fmt),
            ("Total EV TCO (fleet)",      f"={ev_cum_last}",                  currency_fmt),
            ("Total Savings (fleet)",     f"={ice_cum_last}-{ev_cum_last}",   currency_fmt),
            ("Total NPV Savings",         f"=SUM({npv_col_range})",           currency_fmt),
        ]
        for j, (lbl, formula, fmt) in enumerate(kpis):
            r = SUMMARY_START + 1 + j
            ws.write(r, 0, lbl, label_fmt)
            ws.write_formula(r, 1, formula, fmt)

    def _create_replacement_schedule_sheet(
        self, workbook, vehicles, title_format, header_format, cell_format, number_format,
        timeline_options: Optional[dict] = None,
    ):
        """Create Replacement Schedule sheet — Gantt-style table.

        timeline_options: dict of {scenario_key: bool} indicating which scenario
        year columns to append (from the Timelines to Include dialog).
        """
        from analysis.scenarios import get_scenario_year_assignments, PRESET_SCENARIOS

        ws = workbook.add_worksheet("Replacement Schedule")
        currency_fmt = workbook.add_format({'border': 1, 'num_format': '$#,##0'})
        year_fmt     = workbook.add_format({'border': 1, 'align': 'center',
                                            'bg_color': '#D6EAF8'})
        override_fmt = workbook.add_format({'border': 1, 'align': 'center',
                                            'bg_color': '#FFF3E0'})
        yes_fmt      = workbook.add_format({'border': 1, 'align': 'center',
                                            'bg_color': '#FFE0B2', 'bold': True})

        ws.merge_range('A1:F1', "Vehicle Replacement Schedule", title_format)

        # Compute scenario year assignments for selected timelines
        active_scenarios: list = []
        if timeline_options:
            for key, selected in timeline_options.items():
                if selected and key in PRESET_SCENARIOS:
                    assignments = get_scenario_year_assignments(vehicles, key)
                    active_scenarios.append(
                        (PRESET_SCENARIOS[key].name, assignments)
                    )

        # Gather schedulable vehicles (keyed by vin for scenario lookup)
        scheduled = []
        for v in vehicles:
            ev_year = v.custom_fields.get('Proposed EV Year', '')
            if ev_year and ev_year not in ('N/A', 'Exempt', ''):
                try:
                    year_int = int(ev_year)
                except (ValueError, TypeError):
                    continue
                make = v.vehicle_id.make or ''
                model = v.vehicle_id.model or ''
                dept = v.custom_fields.get('department', '')
                ev_equiv = v.custom_fields.get('EV Equivalent', '')
                ev_cost = float(v.custom_fields.get('_ev_purchase_price', 0) or 0)
                is_override = v.custom_fields.get('EV Year Overridden', '') == 'Yes'
                scheduled.append({
                    'vehicle': f"{v.vehicle_id.year or ''} {make} {model}".strip(),
                    'dept': dept,
                    'ev_year': year_int,
                    'ev_equiv': ev_equiv,
                    'ev_cost': ev_cost,
                    'vin': v.vin[:8] + '...' if v.vin else '',
                    'vin_full': v.vin or '',
                    'is_override': is_override,
                })

        if not scheduled:
            ws.write(2, 0, "No vehicles scheduled for replacement")
            return

        scheduled.sort(key=lambda x: (x['ev_year'], x['vehicle']))

        # Get year range
        all_years = sorted(set(s['ev_year'] for s in scheduled))

        # Base headers + override flag column
        base_headers = ["Vehicle", "VIN", "Department", "EV Equivalent",
                        "Est. Cost", "Overridden?"]
        for col, h in enumerate(base_headers):
            ws.write(2, col, h, header_format)
        ws.set_column(0, 0, 25)
        ws.set_column(1, 1, 12)
        ws.set_column(2, 2, 18)
        ws.set_column(3, 3, 22)
        ws.set_column(4, 4, 14)
        ws.set_column(5, 5, 12)

        base_col_count = len(base_headers)

        # Year columns (current/override timeline)
        for i, yr in enumerate(all_years):
            col = base_col_count + i
            ws.write(2, col, str(yr), header_format)
            ws.set_column(col, col, 8)

        # Scenario year columns — one per selected scenario
        scen_col_start = base_col_count + len(all_years)
        for scen_idx, (scen_name, _assignments) in enumerate(active_scenarios):
            col = scen_col_start + scen_idx
            ws.write(2, col, f"{scen_name} Year", header_format)
            ws.set_column(col, col, 14)

        # Data rows
        for row_idx, s in enumerate(scheduled):
            row = 3 + row_idx
            cell_fmt = override_fmt if s['is_override'] else cell_format

            ws.write(row, 0, s['vehicle'], cell_fmt)
            ws.write(row, 1, s['vin'], cell_fmt)
            ws.write(row, 2, s['dept'], cell_fmt)
            ws.write(row, 3, s['ev_equiv'], cell_fmt)
            ws.write(row, 4, s['ev_cost'], currency_fmt)
            ws.write(row, 5, "Yes" if s['is_override'] else "",
                     yes_fmt if s['is_override'] else cell_format)

            # Mark the replacement year (highlight override rows)
            for i, yr in enumerate(all_years):
                col = base_col_count + i
                if yr == s['ev_year']:
                    fmt = override_fmt if s['is_override'] else year_fmt
                    ws.write(row, col, "X", fmt)
                else:
                    ws.write(row, col, "", cell_format)

            # Scenario year values (one column per selected scenario)
            for scen_idx, (_scen_name, assignments) in enumerate(active_scenarios):
                col = scen_col_start + scen_idx
                scen_year = assignments.get(s['vin_full'], "—")
                ws.write(row, col, scen_year, cell_format)

        # Summary row
        summary_row = 3 + len(scheduled) + 1
        ws.write(summary_row, 0, "Vehicles per year:", header_format)
        for i, yr in enumerate(all_years):
            col = base_col_count + i
            count = sum(1 for s in scheduled if s['ev_year'] == yr)
            ws.write(summary_row, col, count, header_format)

        total_cost = sum(s['ev_cost'] for s in scheduled)
        ws.write(summary_row + 1, 0, "Total estimated cost:", header_format)
        ws.write(summary_row + 1, 4, total_cost, currency_fmt)

    def _create_summary_dashboard_sheet(self, workbook, vehicles, analysis, charging, emissions, title_format, header_format, cell_format, number_format):
        """Create Summary Dashboard sheet with KPI reference cells."""
        ws = workbook.add_worksheet("Summary Dashboard")
        kpi_title_fmt = workbook.add_format({
            'bold': True, 'font_size': 11, 'bg_color': '#2C3E50',
            'font_color': 'white', 'border': 1, 'align': 'center'
        })
        kpi_value_fmt = workbook.add_format({
            'bold': True, 'font_size': 16, 'align': 'center',
            'border': 1, 'num_format': '$#,##0'
        })
        kpi_text_fmt = workbook.add_format({
            'bold': True, 'font_size': 16, 'align': 'center', 'border': 1
        })

        ws.merge_range('A1:F1', "Fleet Electrification — Executive Summary Dashboard", title_format)
        ws.set_column(0, 5, 22)

        # Row 3-4: Fleet KPIs
        fleet_kpis = [
            ("Fleet Size", str(len(vehicles)), False),
            ("Avg MPG", f"{sum(v.fuel_economy.combined_mpg or 0 for v in vehicles) / max(1, sum(1 for v in vehicles if v.fuel_economy.combined_mpg)):,.1f}", False),
        ]

        if analysis:
            fleet_kpis.extend([
                ("Annual Savings", analysis.total_savings, True),
                ("CO₂ Reduction", f"{analysis.co2_savings:,.1f} MT", False),
                ("Payback Period", f"{analysis.payback_period:.1f} yr", False),
            ])

        if charging:
            fleet_kpis.append(
                ("Infrastructure Cost", charging.estimated_installation_cost, True)
            )

        for col, (title, value, is_currency) in enumerate(fleet_kpis):
            ws.write(2, col, title, kpi_title_fmt)
            if is_currency:
                ws.write(3, col, value, kpi_value_fmt)
            else:
                ws.write(3, col, str(value), kpi_text_fmt)

        # Row 6+: Breakdown tables
        row = 6
        if emissions and emissions.by_department:
            ws.write(row, 0, "Emissions by Department", header_format)
            ws.write(row, 1, "MT CO2e", header_format)
            row += 1
            for dept, em in sorted(emissions.by_department.items(), key=lambda x: x[1], reverse=True):
                ws.write(row, 0, dept, cell_format)
                ws.write(row, 1, em, number_format)
                row += 1
            row += 1

        # ACF breakdown
        acf_counts = {}
        for v in vehicles:
            code = v.custom_fields.get('ACF Category', '')
            if code:
                acf_counts[code] = acf_counts.get(code, 0) + 1

        if acf_counts:
            ws.write(row, 0, "ACF Classification", header_format)
            ws.write(row, 1, "Count", header_format)
            row += 1
            for cat, count in sorted(acf_counts.items()):
                ws.write(row, 0, cat, cell_format)
                ws.write(row, 1, count, cell_format)
                row += 1
            row += 1

        # EV year distribution
        year_counts = {}
        for v in vehicles:
            yr = v.custom_fields.get('Proposed EV Year', '')
            if yr and yr not in ('N/A', 'Exempt'):
                year_counts[yr] = year_counts.get(yr, 0) + 1

        if year_counts:
            ws.write(row, 0, "Replacement Year", header_format)
            ws.write(row, 1, "Vehicles", header_format)
            row += 1
            for yr, count in sorted(year_counts.items()):
                ws.write(row, 0, yr, cell_format)
                ws.write(row, 1, count, cell_format)
                row += 1



###############################################################################
# PDF Export — REMOVED (Phase 5 Fix 40 / Phase 13 Fix G)
###############################################################################

class PdfReportGenerator(ReportGenerator):
    """PDF export removed in Phase 5 (Fix 40). Stub so factory does not raise
    NameError if called with a .pdf path; always returns False."""

    def generate(self, fleet, **kwargs) -> bool:  # type: ignore[override]
        logger.error("PDF export is not supported. Only CSV and Excel exports are available.")
        return False


###############################################################################
# JSON Export
###############################################################################

class JsonReportGenerator(ReportGenerator):
    """Generate JSON export of fleet data and analysis."""
    
    def generate(self, fleet: Union[Fleet, List[FleetVehicle]], 
               analysis: Optional[ElectrificationAnalysis] = None,
               charging: Optional[ChargingAnalysis] = None,
               emissions: Optional[EmissionsInventory] = None) -> bool:
        """
        Generate a JSON export with all data and analysis.
        
        Args:
            fleet: Fleet object or list of vehicles
            analysis: Optional electrification analysis
            charging: Optional charging analysis
            emissions: Optional emissions inventory
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Extract vehicles from fleet if needed
            vehicles = fleet.vehicles if isinstance(fleet, Fleet) else fleet
            
            if not vehicles:
                logger.warning("No vehicles to export")
                return False
            
            # Build JSON structure
            export_data = {
                "metadata": {
                    "generated_at": datetime.datetime.now().isoformat(),
                    "fleet_name": fleet.name if isinstance(fleet, Fleet) else "Fleet Export",
                    "vehicle_count": len(vehicles)
                },
                "vehicles": []
            }
            
            # Add vehicles
            for vehicle in vehicles:
                export_data["vehicles"].append(vehicle.to_dict())
            
            # Add analysis if available
            if analysis:
                export_data["electrification_analysis"] = analysis.to_dict()
            
            # Add charging if available
            if charging:
                export_data["charging_analysis"] = charging.to_dict()
            
            # Add emissions if available
            if emissions:
                export_data["emissions_inventory"] = emissions.to_dict()
            
            # Write to file
            with open(self.output_path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, indent=2)
            
            logger.info(f"JSON export generated successfully: {self.output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error generating JSON export: {e}")
            return False


###############################################################################
# Report Generator Factory
###############################################################################

class ReportGeneratorFactory:
    """Factory for creating report generators based on file extension."""
    
    @staticmethod
    def create_generator(output_path: str) -> Optional[ReportGenerator]:
        """
        Create appropriate report generator based on file extension.
        
        Args:
            output_path: Path to output file
            
        Returns:
            ReportGenerator instance or None if format not supported
        """
        _, ext = os.path.splitext(output_path.lower())
        
        if ext == '.csv':
            return CsvReportGenerator(output_path)
        
        elif ext == '.xlsx':
            if not EXCEL_AVAILABLE:
                logger.error("Excel export is not available (xlsxwriter not installed)")
                return None
            return ExcelReportGenerator(output_path)
        
        
        elif ext == '.json':
            return JsonReportGenerator(output_path)
        
        else:
            logger.error(f"Unsupported export format: {ext}")
            return None


###############################################################################
# Export Coordinator
###############################################################################

class ExportCoordinator:
    """
    Coordinates the export of fleet data and analysis to various formats.
    Handles file management and provides a unified interface for exports.
    """
    
    def __init__(self, export_dir: str = "exports"):
        """
        Initialize the export coordinator.
        
        Args:
            export_dir: Directory for saving exports
        """
        self.export_dir = export_dir
        
        # Ensure directory exists
        os.makedirs(self.export_dir, exist_ok=True)
    
    def export_to_format(self, format_ext: str, fleet: Union[Fleet, List[FleetVehicle]],
                       analysis: Optional[ElectrificationAnalysis] = None,
                       charging: Optional[ChargingAnalysis] = None,
                       emissions: Optional[EmissionsInventory] = None,
                       custom_filename: Optional[str] = None) -> Optional[str]:
        """
        Export fleet data to the specified format.
        
        Args:
            format_ext: File extension without dot (e.g., 'csv', 'xlsx')
            fleet: Fleet object or list of vehicles
            analysis: Optional electrification analysis
            charging: Optional charging analysis
            emissions: Optional emissions inventory
            custom_filename: Optional custom filename
            
        Returns:
            Path to the generated file or None if export failed
        """
        # Ensure format has leading dot
        if not format_ext.startswith('.'):
            format_ext = f".{format_ext}"
        
        # Generate filename
        if custom_filename:
            # If user provided a filename, use it (ensuring it has the right extension)
            if not custom_filename.lower().endswith(format_ext.lower()):
                filename = f"{custom_filename}{format_ext}"
            else:
                filename = custom_filename
        else:
            # Generate a default filename with timestamp
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            fleet_name = fleet.name if isinstance(fleet, Fleet) else "fleet"
            # Replace any invalid characters in fleet_name
            fleet_name = ''.join(c if c.isalnum() or c in ' _-' else '_' for c in fleet_name)
            filename = f"{fleet_name}_{timestamp}{format_ext}"
        
        # Create complete path
        output_path = os.path.join(self.export_dir, filename)
        
        # Create appropriate generator
        generator = ReportGeneratorFactory.create_generator(output_path)
        
        if generator is None:
            return None
        
        # Generate the report
        success = generator.generate(
            fleet=fleet,
            analysis=analysis,
            charging=charging,
            emissions=emissions
        )
        
        if success:
            return output_path
        else:
            return None
    
    def export_to_all_formats(self, fleet: Union[Fleet, List[FleetVehicle]],
                            analysis: Optional[ElectrificationAnalysis] = None,
                            charging: Optional[ChargingAnalysis] = None,
                            emissions: Optional[EmissionsInventory] = None,
                            base_filename: Optional[str] = None) -> Dict[str, Optional[str]]:
        """
        Export fleet data to all available formats.
        
        Args:
            fleet: Fleet object or list of vehicles
            analysis: Optional electrification analysis
            charging: Optional charging analysis
            emissions: Optional emissions inventory
            base_filename: Optional base filename (without extension)
            
        Returns:
            Dictionary mapping formats to output paths (or None if export failed)
        """
        results = {}
        
        for format_name, format_ext in EXPORT_FORMATS.items():
            # Skip formats that require unavailable libraries
            if format_ext == '.xlsx' and not EXCEL_AVAILABLE:
                results[format_name] = None
                continue
                
            # Export to the format
            if base_filename:
                custom_name = f"{base_filename}{format_ext}"
            else:
                custom_name = None
                
            result = self.export_to_format(
                format_ext=format_ext,
                fleet=fleet,
                analysis=analysis,
                charging=charging,
                emissions=emissions,
                custom_filename=custom_name
            )
            
            results[format_name] = result
        
        return results
