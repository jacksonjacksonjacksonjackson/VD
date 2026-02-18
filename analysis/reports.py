"""
reports.py

Report generation and exports for the Fleet Electrification Analyzer.
Provides functions to create various report formats (PDF, Excel, etc.).
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

try:
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False
    logger.warning("ReportLab not available. PDF export will be disabled.")

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
               fields: Optional[List[str]] = None) -> bool:
        """
        Generate an Excel report with data, charts, and analysis.
        
        Args:
            fleet: Fleet object or list of vehicles
            analysis: Optional electrification analysis
            charging: Optional charging analysis
            emissions: Optional emissions inventory
            fields: List of fields to include (None for default fields)
            
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
            
            # Create Vehicle Data sheet
            self._create_vehicle_data_sheet(workbook, vehicles, fields, title_format, header_format, cell_format, number_format)
            
            # Create Summary sheet
            self._create_summary_sheet(workbook, vehicles, title_format, header_format)
            
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

            self._create_replacement_schedule_sheet(workbook, vehicles, title_format, header_format, cell_format, number_format)

            self._create_summary_dashboard_sheet(workbook, vehicles, analysis, charging, emissions, title_format, header_format, cell_format, number_format)

            # Close workbook to save changes
            workbook.close()
            
            logger.info(f"Excel report generated successfully: {self.output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error generating Excel report: {e}")
            return False
    
    def _create_vehicle_data_sheet(self, workbook, vehicles, fields, title_format, header_format, cell_format, number_format):
        """Create the Vehicle Data sheet with all vehicle information."""
        # Create worksheet
        worksheet = workbook.add_worksheet("Vehicle Data")
        
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
                
                # Format numbers appropriately
                if field in ["MPG City", "MPG Highway", "MPG Combined", "CO2 emissions", "co2A"]:
                    try:
                        value = float(value)
                        worksheet.write(row + 3, col, value, number_format)
                    except (ValueError, TypeError):
                        worksheet.write(row + 3, col, value, cell_format)
                else:
                    worksheet.write(row + 3, col, value, cell_format)
        
        # Freeze header row
        worksheet.freeze_panes(3, 0)
        
        # Auto-filter
        worksheet.autofilter(2, 0, 2 + len(vehicles), len(headers) - 1)
    
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
          Rows 4-14   : Assumption input cells in column B, labels in column A
          Row 15      : blank
          Row 16      : Column headers for the cash-flow table
          Rows 17+    : One row per year; all cost/savings cells use =formulas referencing B4:B14
          Last rows   : Summary KPIs (also formula-based)

        Assumption cells (all in column B, named via xlsxwriter cell reference):
          B4  = Gas price ($/gal)
          B5  = Electricity price ($/kWh)
          B6  = EV efficiency (kWh/mile)
          B7  = ICE maintenance ($/mile)
          B8  = EV maintenance ($/mile)
          B9  = Fuel escalation rate (% / yr)
          B10 = Discount rate (%)
          B11 = Battery degradation (% / yr)
          B12 = Infrastructure cost per vehicle ($)
          B13 = Fleet size (vehicles)
          B14 = Analysis period (years)

        Consultants can change B4:B14 and every formula cell recalculates automatically.
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

        # Assumption rows: (label, value, format, cell_hint)
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

        # Anchor cells: write Year-1 base values as hidden helpers in columns K-O
        # (off-screen, labelled so consultants can find them)
        ANCHOR_COL = 10  # column K (0-indexed)
        ws.write(TABLE_HEADER_ROW, ANCHOR_COL,     "⟵ Anchor: ICE fuel yr1",  hint_fmt)
        ws.write(TABLE_HEADER_ROW, ANCHOR_COL + 1, "⟵ Anchor: ICE maint yr1", hint_fmt)
        ws.write(TABLE_HEADER_ROW, ANCHOR_COL + 2, "⟵ Anchor: EV fuel yr1",   hint_fmt)
        ws.write(TABLE_HEADER_ROW, ANCHOR_COL + 3, "⟵ Anchor: EV maint yr1",  hint_fmt)
        ws.write(TABLE_HEADER_ROW, ANCHOR_COL + 4, "⟵ Anchor: EV infra yr1",  hint_fmt)

        ANCHOR_ROW = TABLE_DATA_START  # Year-0 row; anchor values go here
        ws.write(ANCHOR_ROW, ANCHOR_COL,     ice_y1_fuel,  input_fmt)
        ws.write(ANCHOR_ROW, ANCHOR_COL + 1, ice_y1_maint, input_fmt)
        ws.write(ANCHOR_ROW, ANCHOR_COL + 2, ev_y1_fuel,   input_fmt)
        ws.write(ANCHOR_ROW, ANCHOR_COL + 3, ev_y1_maint,  input_fmt)
        ws.write(ANCHOR_ROW, ANCHOR_COL + 4, ev_y1_infra,  input_fmt)

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

    def _create_replacement_schedule_sheet(self, workbook, vehicles, title_format, header_format, cell_format, number_format):
        """Create Replacement Schedule sheet — Gantt-style table."""
        ws = workbook.add_worksheet("Replacement Schedule")
        currency_fmt = workbook.add_format({'border': 1, 'num_format': '$#,##0'})
        year_fmt = workbook.add_format({'border': 1, 'align': 'center', 'bg_color': '#D6EAF8'})

        ws.merge_range('A1:F1', "Vehicle Replacement Schedule", title_format)

        # Gather schedulable vehicles
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
                scheduled.append({
                    'vehicle': f"{v.vehicle_id.year or ''} {make} {model}".strip(),
                    'dept': dept,
                    'ev_year': year_int,
                    'ev_equiv': ev_equiv,
                    'ev_cost': ev_cost,
                    'vin': v.vin[:8] + '...' if v.vin else '',
                })

        if not scheduled:
            ws.write(2, 0, "No vehicles scheduled for replacement")
            return

        scheduled.sort(key=lambda x: (x['ev_year'], x['vehicle']))

        # Get year range
        all_years = sorted(set(s['ev_year'] for s in scheduled))

        # Headers
        base_headers = ["Vehicle", "VIN", "Department", "EV Equivalent", "Est. Cost"]
        for col, h in enumerate(base_headers):
            ws.write(2, col, h, header_format)
        ws.set_column(0, 0, 25)
        ws.set_column(1, 1, 12)
        ws.set_column(2, 2, 18)
        ws.set_column(3, 3, 22)
        ws.set_column(4, 4, 14)

        # Year columns
        for i, yr in enumerate(all_years):
            col = len(base_headers) + i
            ws.write(2, col, str(yr), header_format)
            ws.set_column(col, col, 8)

        # Data rows
        for row_idx, s in enumerate(scheduled):
            row = 3 + row_idx
            ws.write(row, 0, s['vehicle'], cell_format)
            ws.write(row, 1, s['vin'], cell_format)
            ws.write(row, 2, s['dept'], cell_format)
            ws.write(row, 3, s['ev_equiv'], cell_format)
            ws.write(row, 4, s['ev_cost'], currency_fmt)

            # Mark the replacement year
            for i, yr in enumerate(all_years):
                col = len(base_headers) + i
                if yr == s['ev_year']:
                    ws.write(row, col, "X", year_fmt)
                else:
                    ws.write(row, col, "", cell_format)

        # Summary row
        summary_row = 3 + len(scheduled) + 1
        ws.write(summary_row, 0, "Vehicles per year:", header_format)
        for i, yr in enumerate(all_years):
            col = len(base_headers) + i
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
# PDF Export
###############################################################################

class PdfReportGenerator(ReportGenerator):
    """Generate PDF reports from fleet data with charts and analysis."""
    
    def generate(self, fleet: Union[Fleet, List[FleetVehicle]], 
               analysis: Optional[ElectrificationAnalysis] = None,
               charging: Optional[ChargingAnalysis] = None,
               emissions: Optional[EmissionsInventory] = None,
               include_charts: bool = True,
               include_vehicle_table: bool = True) -> bool:
        """
        Generate a PDF report with data, charts, and analysis.
        
        Args:
            fleet: Fleet object or list of vehicles
            analysis: Optional electrification analysis
            charging: Optional charging analysis
            emissions: Optional emissions inventory
            include_charts: Whether to include charts
            include_vehicle_table: Whether to include the vehicle data table
            
        Returns:
            True if successful, False otherwise
        """
        if not PDF_AVAILABLE:
            logger.error("PDF export is not available (reportlab not installed)")
            return False
        
        try:
            # Extract vehicles from fleet if needed
            vehicles = fleet.vehicles if isinstance(fleet, Fleet) else fleet
            
            if not vehicles:
                logger.warning("No vehicles to export")
                return False
            
            # Create PDF document
            doc = SimpleDocTemplate(
                self.output_path,
                pagesize=letter,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=72
            )
            
            # Get styles
            styles = getSampleStyleSheet()
            title_style = styles['Title']
            heading1_style = styles['Heading1']
            heading2_style = styles['Heading2']
            normal_style = styles['Normal']
            
            # Create custom paragraph styles
            subtitle_style = ParagraphStyle(
                'Subtitle',
                parent=styles['Heading2'],
                fontSize=12,
                spaceAfter=12
            )
            
            table_header_style = ParagraphStyle(
                'TableHeader',
                parent=styles['Normal'],
                fontSize=9,
                textColor=colors.white,
                alignment=1  # Center
            )
            
            table_cell_style = ParagraphStyle(
                'TableCell',
                parent=styles['Normal'],
                fontSize=8
            )
            
            # Create story (list of elements to add to the PDF)
            story = []
            
            # Add title
            fleet_name = fleet.name if isinstance(fleet, Fleet) else "Fleet Analysis"
            report_date = datetime.datetime.now().strftime("%Y-%m-%d")
            
            story.append(Paragraph(f"{fleet_name}", title_style))
            story.append(Paragraph(f"Fleet Electrification Analysis Report", subtitle_style))
            story.append(Paragraph(f"Generated on {report_date}", normal_style))
            story.append(Spacer(1, 12))
            
            # Add fleet summary
            story.append(Paragraph("Fleet Summary", heading1_style))
            story.append(Spacer(1, 6))
            
            # Basic statistics
            summary_data = [
                ["Total Vehicles:", f"{len(vehicles)}"],
                ["Average MPG:", f"{self._calculate_avg_mpg(vehicles):.1f}"],
                ["Average CO2 Emissions:", f"{self._calculate_avg_co2(vehicles):.1f} g/mile"],
                ["Average Annual Mileage:", f"{self._calculate_avg_mileage(vehicles):.0f} miles"]
            ]
            
            # Create table for summary stats
            summary_table = Table(summary_data, colWidths=[2*inch, 1.5*inch])
            summary_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
                ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ]))
            
            story.append(summary_table)
            story.append(Spacer(1, 12))
            
            # Add charts if requested
            if include_charts and len(vehicles) > 0:
                story.append(Paragraph("Fleet Composition", heading2_style))
                story.append(Spacer(1, 6))
                
                # Create and add make frequency chart
                if self._add_chart_to_story(story, "Make Frequency", vehicles):
                    story.append(Spacer(1, 12))
                
                # Create and add body class distribution chart
                if self._add_chart_to_story(story, "Body Class Distribution", vehicles):
                    story.append(Spacer(1, 12))
                
                # Create and add fuel type distribution chart
                if self._add_chart_to_story(story, "Fuel Type Distribution", vehicles):
                    story.append(Spacer(1, 12))
                
                story.append(Paragraph("Fleet Performance", heading2_style))
                story.append(Spacer(1, 6))
                
                # Create and add MPG distribution chart
                if self._add_chart_to_story(story, "MPG Distribution", vehicles):
                    story.append(Spacer(1, 12))
                
                # Create and add CO2 distribution chart
                if self._add_chart_to_story(story, "CO2 Emissions Distribution", vehicles):
                    story.append(Spacer(1, 12))
                
                # Create and add CO2 vs MPG chart
                if self._add_chart_to_story(story, "CO2 vs MPG Correlation", vehicles):
                    story.append(Spacer(1, 12))
            
            # Add electrification analysis if available
            if analysis:
                story.append(Paragraph("Electrification Analysis", heading1_style))
                story.append(Spacer(1, 6))
                
                # Add analysis parameters
                story.append(Paragraph("Analysis Parameters:", heading2_style))
                
                params_data = [
                    ["Gas Price:", f"${analysis.gas_price:.2f}/gal"],
                    ["Electricity Price:", f"${analysis.electricity_price:.2f}/kWh"],
                    ["EV Efficiency:", f"{analysis.ev_efficiency:.2f} kWh/mile"],
                    ["Analysis Period:", f"{analysis.analysis_period} years"],
                    ["Discount Rate:", f"{analysis.discount_rate:.1f}%"]
                ]
                
                params_table = Table(params_data, colWidths=[2*inch, 1.5*inch])
                params_table.setStyle(TableStyle([
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
                    ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ]))
                
                story.append(params_table)
                story.append(Spacer(1, 12))
                
                # Add analysis results
                story.append(Paragraph("Analysis Results:", heading2_style))
                
                results_data = [
                    ["Total CO2 Savings:", f"{analysis.co2_savings:.1f} tons"],
                    ["Fuel Cost Savings:", f"${analysis.fuel_cost_savings:,.2f}"],
                    ["Maintenance Savings:", f"${analysis.maintenance_savings:,.2f}"],
                    ["Total Savings:", f"${analysis.total_savings:,.2f}"],
                    ["Payback Period:", f"{analysis.payback_period:.1f} years"]
                ]
                
                results_table = Table(results_data, colWidths=[2*inch, 1.5*inch])
                results_table.setStyle(TableStyle([
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
                    ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ]))
                
                story.append(results_table)
                story.append(Spacer(1, 12))
                
                # Add electrification potential chart
                if include_charts:
                    if self._add_chart_to_story(story, "Electrification Potential", analysis):
                        story.append(Spacer(1, 12))
            
            # Add charging analysis if available
            if charging:
                story.append(Paragraph("Charging Infrastructure", heading1_style))
                story.append(Spacer(1, 6))
                
                # Add infrastructure requirements
                infra_data = [
                    ["Level 2 Chargers:", f"{charging.level2_chargers_needed}"],
                    ["DC Fast Chargers:", f"{charging.dcfc_chargers_needed}"],
                    ["Maximum Power Required:", f"{charging.max_power_required:.1f} kW"],
                    ["Estimated Cost:", f"${charging.estimated_installation_cost:,.2f}"]
                ]
                
                infra_table = Table(infra_data, colWidths=[2*inch, 1.5*inch])
                infra_table.setStyle(TableStyle([
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
                    ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ]))
                
                story.append(infra_table)
                story.append(Spacer(1, 12))
                
                # Add charging infrastructure chart
                if include_charts:
                    if self._add_chart_to_story(story, "Charging Infrastructure", charging):
                        story.append(Spacer(1, 12))
            
            # Add vehicle data table if requested
            if include_vehicle_table:
                story.append(Paragraph("Vehicle Data", heading1_style))
                story.append(Spacer(1, 6))
                
                # Get selected fields
                fields = ["VIN", "Year", "Make", "Model", "MPG Combined", "CO2 emissions", "Annual Mileage"]
                headers = [COLUMN_NAME_MAP.get(field, field) for field in fields]
                
                # Create table data
                table_data = [headers]  # First row is headers
                
                for vehicle in vehicles:
                    row_dict = vehicle.to_row_dict()
                    table_data.append([row_dict.get(field, "") for field in fields])
                
                # Create table
                vehicle_table = Table(table_data, repeatRows=1)
                
                # Style the table
                style = TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
                    ('FONTSIZE', (0, 1), (-1, -1), 7),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
                ])
                
                vehicle_table.setStyle(style)
                
                # Add table to story
                story.append(vehicle_table)
            
            # Build the PDF
            doc.build(story)
            
            logger.info(f"PDF report generated successfully: {self.output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error generating PDF report: {e}")
            return False
    
    def _calculate_avg_mpg(self, vehicles):
        """Calculate average MPG for fleet."""
        mpg_values = [v.fuel_economy.combined_mpg for v in vehicles 
                     if v.fuel_economy.combined_mpg and v.fuel_economy.combined_mpg > 0]
        
        if not mpg_values:
            return 0.0
        
        return sum(mpg_values) / len(mpg_values)
    
    def _calculate_avg_co2(self, vehicles):
        """Calculate average CO2 emissions for fleet."""
        co2_values = [v.fuel_economy.co2_primary for v in vehicles 
                     if v.fuel_economy.co2_primary and v.fuel_economy.co2_primary > 0]
        
        if not co2_values:
            return 0.0
        
        return sum(co2_values) / len(co2_values)
    
    def _calculate_avg_mileage(self, vehicles):
        """Calculate average annual mileage for fleet."""
        mileage_values = [v.annual_mileage for v in vehicles 
                         if v.annual_mileage and v.annual_mileage > 0]
        
        if not mileage_values:
            return 0.0
        
        return sum(mileage_values) / len(mileage_values)
    
    def _add_chart_to_story(self, story, chart_type, data):
        """Add a chart to the PDF story."""
        try:
            # Create the chart using matplotlib
            figure = ChartFactory.create_chart(chart_type, data)
            
            # Save chart to a BytesIO object
            buf = BytesIO()
            figure.savefig(buf, format='png', dpi=150, bbox_inches='tight')
            buf.seek(0)
            
            # Add chart to story
            img = Image(buf, width=6*inch, height=4*inch)
            story.append(img)
            
            plt.close(figure)
            return True
            
        except Exception as e:
            logger.error(f"Error adding chart to PDF: {e}")
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
        
        elif ext == '.pdf':
            if not PDF_AVAILABLE:
                logger.error("PDF export is not available (reportlab not installed)")
                return None
            return PdfReportGenerator(output_path)
        
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
                
            if format_ext == '.pdf' and not PDF_AVAILABLE:
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
