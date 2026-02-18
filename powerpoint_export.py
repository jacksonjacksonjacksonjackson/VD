"""
powerpoint_export_enhanced.py

Enhanced PowerPoint export functionality for the Fleet Electrification Analyzer.
Implements complete slide builders with real data, charts, and professional formatting.
"""

import os
import io
import logging
import tempfile
import datetime
from typing import Dict, List, Any, Optional, Union
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.dml import MSO_THEME_COLOR
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

# Import matplotlib for chart generation
try:
    import matplotlib
    matplotlib.use('Agg')  # Use non-interactive backend for server/background use
    import matplotlib.pyplot as plt
    from matplotlib.figure import Figure
    import numpy as np
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

from data.models import Fleet, FleetVehicle
from settings import (
    APP_NAME, APP_VERSION, 
    PRIMARY_HEX_1, PRIMARY_HEX_2, PRIMARY_HEX_3, 
    SECONDARY_HEX_1, SECONDARY_HEX_2
)
from analysis.calculations import (
    analyze_fleet_electrification, create_emissions_inventory,
    analyze_charging_needs, calculate_electrification_savings
)
from powerpoint_charts import (
    SlideConfiguration, add_fleet_composition_chart, add_emissions_timeline_chart,
    add_emissions_by_weight_class_chart, add_electrification_timeline_by_weight_chart,
    add_electrification_timeline_by_body_type_chart, add_age_distribution_chart,
    add_tco_comparison_chart, add_payback_timeline_chart,
    add_scenario_comparison_chart
)

# Set up module logger
logger = logging.getLogger(__name__)

###############################################################################
# Main Export Function
###############################################################################

def export_prelim_deck(data: dict, template_path: Optional[str] = None, out_path: Optional[str] = None,
                      slide_config: Optional[SlideConfiguration] = None) -> str:
    """
    Builds an enhanced preliminary .pptx from `data`. Returns absolute path.
    
    Args:
        data: Dictionary containing fleet data (see data contract)
        template_path: Path to .potx template file (optional)
        out_path: Output path for .pptx file (optional)
        slide_config: SlideConfiguration object for customizing slides/charts (optional)
        
    Returns:
        Absolute path to the generated .pptx file
        
    Raises:
        RuntimeError: If python-pptx is not available or required shapes are missing
        ValueError: If data is invalid or insufficient
    """
    if not PPTX_AVAILABLE:
        raise RuntimeError("python-pptx is required but not installed. Run: pip install python-pptx")
    
    logger.info("[prelim-deck] Starting enhanced PowerPoint export")
    
    # Validate data
    if not data or not isinstance(data, dict):
        raise ValueError("Data must be a non-empty dictionary")
    
    # Get template presentation
    prs = _get_template_presentation(template_path)
    
    # Extract fleet data
    fleet_data = _extract_fleet_data(data)
    
    # Use default configuration if none provided
    if slide_config is None:
        slide_config = SlideConfiguration()
    
    # Build slides based on configuration
    slides_created = []
    
    # Build selected slides
    slide_builders = {
        'cover': _add_cover_slide,
        'fleet_snapshot': _add_fleet_snapshot_slide,
        'fleet_composition': _add_fleet_composition_slide,
        'financial_summary': _add_financial_summary_slide,
        'emissions_timeline': _add_emissions_timeline_slide,
        'emissions_by_weight': _add_emissions_by_weight_slide,
        'electrification_timeline_weight': _add_electrification_timeline_weight_slide,
        'electrification_timeline_body': _add_electrification_timeline_body_slide,
        'replacement_schedule': _add_replacement_schedule_slide,
        'executive_recommendations': _add_executive_recommendations_slide,
        'scenario_comparison': _add_scenario_comparison_slide,
        'age_analysis': _add_age_analysis_slide,
        'data_quality': _add_data_completeness_slide,
        'next_steps': _add_next_steps_slide
    }
    
    for slide_id in slide_config.get_selected_slides():
        if slide_id in slide_builders:
            slide_func = slide_builders[slide_id]
            if slide_func(prs, fleet_data):
                slide_info = slide_config.available_slides.get(slide_id, {})
                slides_created.append(slide_info.get('name', slide_id))
    
    # Determine output path
    if not out_path:
        out_path = _generate_output_path(fleet_data)
    
    # Ensure output directory exists
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    
    # Save presentation
    prs.save(out_path)
    
    abs_path = os.path.abspath(out_path)
    logger.info(f"[prelim-deck] Enhanced PowerPoint exported successfully: {abs_path}")
    logger.info(f"[prelim-deck] Slides created: {', '.join(slides_created)}")
    
    return abs_path

###############################################################################
# New Slide Functions with Native Charts
###############################################################################

def _get_brand_colors() -> Dict[str, str]:
    """
    Get brand color palette from settings.
    
    Returns:
        Dictionary mapping color names to hex values
    """
    return {
        'primary_dark': PRIMARY_HEX_1,      # Charcoal
        'primary_light': PRIMARY_HEX_2,     # White
        'primary_green': PRIMARY_HEX_3,     # Reseda green
        'secondary_orange': SECONDARY_HEX_1, # Deep orange
        'secondary_grey': SECONDARY_HEX_2    # Light grey
    }

def _hex_to_rgb(hex_color: str) -> tuple:
    """
    Convert hex color to RGB tuple.
    
    Args:
        hex_color: Hex color string (e.g., '#3C465A')
        
    Returns:
        RGB tuple (r, g, b)
    """
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def _apply_brand_formatting(text_frame, font_size: int = 12, 
                          color: str = PRIMARY_HEX_1, bold: bool = False):
    """
    Apply consistent brand formatting to text.
    
    Args:
        text_frame: PowerPoint text frame object
        font_size: Font size in points
        color: Hex color string
        bold: Whether to make text bold
    """
    try:
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                font = run.font
                font.name = 'Calibri'
                font.size = Pt(font_size)
                font.bold = bold
                
                # Convert hex to RGB and apply color
                r, g, b = _hex_to_rgb(color)
                font.color.rgb = RGBColor(r, g, b)
                
    except Exception as e:
        logger.warning(f"[prelim-deck] Failed to apply formatting: {e}")

###############################################################################
# Template Handling (Simplified)
###############################################################################

def _get_template_presentation(template_path: Optional[str] = None) -> Presentation:
    """Get template presentation, loading from template_path if provided and valid."""
    if template_path:
        path = Path(template_path)
        if path.exists() and path.suffix.lower() in ('.pptx', '.potx'):
            try:
                return Presentation(str(path))
            except Exception as e:
                logger.warning(f"[prelim-deck] Failed to load template '{template_path}': {e}. Using blank presentation.")
        else:
            logger.warning(f"[prelim-deck] Template not found or unsupported: '{template_path}'. Using blank presentation.")
    return Presentation()

###############################################################################
# Data Extraction
###############################################################################

def _extract_fleet_data(data: dict) -> dict:
    """Extract and normalize fleet data from various input formats."""
    
    # Handle Fleet object
    if hasattr(data, 'vehicles') and hasattr(data, 'name'):
        fleet = data
        vehicles = fleet.vehicles
        fleet_name = fleet.name
    # Handle dictionary with vehicles
    elif 'vehicles' in data:
        vehicles = data['vehicles']
        fleet_name = data.get('name', 'Fleet Analysis')
    # Handle list of vehicles
    elif isinstance(data, list):
        vehicles = data
        fleet_name = 'Fleet Analysis'
    # Handle dictionary format from analysis
    elif 'fleet' in data:
        fleet_data = data['fleet']
        if hasattr(fleet_data, 'vehicles'):
            vehicles = fleet_data.vehicles
            fleet_name = fleet_data.name
        else:
            vehicles = fleet_data.get('vehicles', [])
            fleet_name = fleet_data.get('name', 'Fleet Analysis')
    else:
        # Try to extract vehicles from data
        vehicles = []
        fleet_name = 'Fleet Analysis'
        
        # Look for vehicle-like objects in the data
        for key, value in data.items():
            if isinstance(value, list) and value:
                # Check if this looks like a list of vehicles
                first_item = value[0]
                if hasattr(first_item, 'vin') or (isinstance(first_item, dict) and 'vin' in first_item):
                    vehicles = value
                    break
    
    # Calculate summary statistics
    summary = _calculate_fleet_summary(vehicles)
    
    # Prepare normalized data structure
    normalized_data = {
        'fleet_name': fleet_name,
        'vehicles': vehicles,
        'summary': summary,
        'generation_date': datetime.datetime.now().strftime("%Y-%m-%d"),
        'client_name': data.get('client_name', 'Client'),
        'stage': data.get('stage', 'Preliminary')
    }
    
    return normalized_data

def _calculate_fleet_summary(vehicles: List) -> dict:
    """Calculate summary statistics for the fleet."""
    if not vehicles:
        return {
            'total_units': 0,
            'ld_count': 0,
            'md_count': 0,
            'hd_count': 0,
            'median_age': 0,
            'annual_miles': 0,
            'baseline_fuel_cost': 0,
            'baseline_co2e': 0
        }
    
    total_units = len(vehicles)
    ld_count = md_count = hd_count = 0
    ages = []
    annual_miles_list = []
    mpg_values = []
    co2_values = []
    
    current_year = datetime.datetime.now().year
    
    for vehicle in vehicles:
        # Handle different vehicle formats
        if hasattr(vehicle, 'vehicle_id'):
            # FleetVehicle object
            year = vehicle.vehicle_id.year
            gvwr_pounds = vehicle.vehicle_id.gvwr_pounds
            annual_mileage = vehicle.annual_mileage
            mpg = vehicle.fuel_economy.combined_mpg
            co2 = vehicle.fuel_economy.co2_primary
        elif isinstance(vehicle, dict):
            # Dictionary format
            year = vehicle.get('year', '')
            gvwr_pounds = vehicle.get('gvwr_pounds', 0)
            annual_mileage = vehicle.get('annual_mileage', 0)
            mpg = vehicle.get('combined_mpg', 0)
            co2 = vehicle.get('co2_primary', 0)
        else:
            continue
        
        # Calculate age
        try:
            if year:
                age = current_year - int(year)
                if age >= 0:
                    ages.append(age)
        except (ValueError, TypeError):
            pass
        
        # Classify by GVWR
        if gvwr_pounds:
            if gvwr_pounds <= 8500:
                ld_count += 1
            elif gvwr_pounds <= 19500:
                md_count += 1
            else:
                hd_count += 1
        
        # Collect mileage data
        if annual_mileage and annual_mileage > 0:
            annual_miles_list.append(annual_mileage)
        
        # Collect fuel economy data
        if mpg and mpg > 0:
            mpg_values.append(mpg)
        
        # Collect CO2 data
        if co2 and co2 > 0:
            co2_values.append(co2)
    
    # Calculate statistics
    median_age = sorted(ages)[len(ages)//2] if ages else 0
    avg_annual_miles = sum(annual_miles_list) / len(annual_miles_list) if annual_miles_list else 12000
    avg_mpg = sum(mpg_values) / len(mpg_values) if mpg_values else 20
    avg_co2 = sum(co2_values) / len(co2_values) if co2_values else 400
    
    # Estimate baseline fuel cost (simplified)
    gas_price = 3.50  # $/gallon
    baseline_fuel_cost = (avg_annual_miles / avg_mpg * gas_price * total_units) if avg_mpg > 0 else 0
    
    # Estimate baseline CO2e (simplified - convert g/mile to tons/year)
    baseline_co2e = (avg_co2 * avg_annual_miles * total_units / 1000000 * 1.1023) if avg_co2 > 0 else 0  # Convert to metric tons
    
    return {
        'total_units': total_units,
        'ld_count': ld_count,
        'md_count': md_count,
        'hd_count': hd_count,
        'median_age': median_age,
        'annual_miles': int(avg_annual_miles),
        'baseline_fuel_cost': int(baseline_fuel_cost),
        'baseline_co2e': int(baseline_co2e)
    }

###############################################################################
# Basic Slide Builder Functions
###############################################################################

def _add_cover_slide(prs: Presentation, data: dict) -> bool:
    """Add cover slide with client, stage, and date."""
    try:
        blank_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
        title_frame = title_box.text_frame
        title_frame.text = f"{data['fleet_name']}"
        title_frame.paragraphs[0].font.size = Pt(32)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Subtitle
        subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(3.5), Inches(8), Inches(1))
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.text = data.get('subtitle', f"{data['stage']} Fleet Electrification Analysis")
        subtitle_frame.paragraphs[0].font.size = Pt(20)
        subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Footer date
        date_box = slide.shapes.add_textbox(Inches(1), Inches(7), Inches(3), Inches(0.5))
        date_frame = date_box.text_frame
        date_frame.text = data['generation_date']
        date_frame.paragraphs[0].font.size = Pt(12)
        
        # Footer client
        client_box = slide.shapes.add_textbox(Inches(6), Inches(7), Inches(3), Inches(0.5))
        client_frame = client_box.text_frame
        client_frame.text = f"Prepared for {data['client_name']}"
        client_frame.paragraphs[0].font.size = Pt(12)
        client_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
        
        # Optony branding (bottom of slide — 7.5" tall, place at 6.9")
        brand_box = slide.shapes.add_textbox(Inches(4), Inches(6.9), Inches(2), Inches(0.5))
        brand_frame = brand_box.text_frame
        brand_frame.text = "Powered by Optony"
        brand_frame.paragraphs[0].font.size = Pt(10)
        brand_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        brand_frame.paragraphs[0].font.color.rgb = RGBColor(128, 128, 128)
        
        logger.debug("[prelim-deck] Added cover slide")
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create cover slide: {e}")
        return False

def _add_fleet_snapshot_slide(prs: Presentation, data: dict) -> bool:
    """Add Fleet Snapshot KPIs slide."""
    if "summary" not in data:
        logger.debug("[prelim-deck] Skipping Fleet Snapshot - no summary data")
        return False
    
    try:
        blank_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "Fleet Snapshot"
        title_frame.paragraphs[0].font.size = Pt(24)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        summary = data['summary']
        
        # Create 6 KPI boxes in 2 rows of 3
        kpis = [
            ("Total Units", f"{summary['total_units']:,}"),
            ("Light Duty", f"{summary['ld_count']:,}"),
            ("Medium Duty", f"{summary['md_count']:,}"),
            ("Heavy Duty", f"{summary['hd_count']:,}"),
            ("Median Age", f"{summary['median_age']} years"),
            ("Annual Miles", f"{summary['annual_miles']:,}")
        ]
        
        # Position KPIs in a 2x3 grid
        for i, (label, value) in enumerate(kpis):
            row = i // 3
            col = i % 3
            
            x = Inches(0.5 + col * 3)
            y = Inches(2 + row * 2)
            
            # KPI box
            kpi_box = slide.shapes.add_textbox(x, y, Inches(2.5), Inches(1.5))
            kpi_frame = kpi_box.text_frame
            
            # Value (large)
            kpi_frame.text = value
            kpi_frame.paragraphs[0].font.size = Pt(28)
            kpi_frame.paragraphs[0].font.bold = True
            kpi_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            
            # Label (small, below)
            p = kpi_frame.add_paragraph()
            p.text = label
            p.font.size = Pt(14)
            p.alignment = PP_ALIGN.CENTER
        
        # Add baseline metrics at the bottom
        baseline_box = slide.shapes.add_textbox(Inches(1), Inches(6), Inches(8), Inches(1.5))
        baseline_frame = baseline_box.text_frame
        baseline_frame.text = f"Baseline Annual Fuel Cost: ${summary['baseline_fuel_cost']:,}"
        baseline_frame.paragraphs[0].font.size = Pt(16)
        baseline_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        p = baseline_frame.add_paragraph()
        p.text = f"Baseline CO₂e Emissions: {summary['baseline_co2e']:,} metric tons"
        p.font.size = Pt(16)
        p.alignment = PP_ALIGN.CENTER
        
        logger.debug("[prelim-deck] Added Fleet Snapshot slide")
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Fleet Snapshot slide: {e}")
        return False

def _add_fleet_composition_slide(prs: Presentation, data: dict) -> bool:
    """Add Fleet Composition slide with body type distribution."""
    if not data.get('vehicles'):
        logger.debug("[prelim-deck] Skipping Fleet Composition - no vehicle data")
        return False
    
    try:
        blank_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "Fleet Composition"
        title_frame.paragraphs[0].font.size = Pt(24)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Count body types
        body_types = {}
        for vehicle in data['vehicles']:
            if hasattr(vehicle, 'vehicle_id'):
                body_class = vehicle.vehicle_id.body_class or 'Unknown'
            elif isinstance(vehicle, dict):
                body_class = vehicle.get('body_class', 'Unknown')
            else:
                body_class = 'Unknown'
            
            body_types[body_class] = body_types.get(body_class, 0) + 1
        
        # Add native PowerPoint pie chart
        if not add_fleet_composition_chart(slide, data['vehicles'], 1, 1.8, 8, 4.5):
            # Fallback to placeholder text
            placeholder_box = slide.shapes.add_textbox(Inches(2), Inches(3), Inches(6), Inches(2))
            placeholder_frame = placeholder_box.text_frame
            placeholder_frame.text = "Fleet composition data will be displayed here once vehicle body class information is available."
            placeholder_frame.paragraphs[0].font.size = Pt(16)
            placeholder_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        logger.debug("[prelim-deck] Added Fleet Composition slide")
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Fleet Composition slide: {e}")
        return False

def _add_emissions_timeline_slide(prs: Presentation, data: dict) -> bool:
    """Add CO2 Emissions Reduction Timeline slide."""
    try:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "CO₂ Emissions Reduction Timeline"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        vehicles = data.get('vehicles', [])
        
        # Add native PowerPoint line chart
        if not add_emissions_timeline_chart(slide, vehicles, 1, 1.8, 8, 4.5):
            # Fallback to placeholder text
            placeholder_box = slide.shapes.add_textbox(Inches(2), Inches(3), Inches(6), Inches(2))
            placeholder_frame = placeholder_box.text_frame
            placeholder_frame.text = "CO₂ emissions reduction timeline will be displayed here once sufficient vehicle data is available."
            placeholder_frame.paragraphs[0].font.size = Pt(16)
            placeholder_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        logger.debug("[prelim-deck] Added Emissions Timeline slide")
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Emissions Timeline slide: {e}")
        return False

def _add_emissions_by_weight_slide(prs: Presentation, data: dict) -> bool:
    """Add CO2 Emissions by Weight Class slide."""
    try:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "CO₂ Emissions by Vehicle Weight Class"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        vehicles = data.get('vehicles', [])
        
        # Add native PowerPoint pie chart
        if not add_emissions_by_weight_class_chart(slide, vehicles, 1, 1.8, 8, 4.5):
            # Fallback to placeholder text
            placeholder_box = slide.shapes.add_textbox(Inches(2), Inches(3), Inches(6), Inches(2))
            placeholder_frame = placeholder_box.text_frame
            placeholder_frame.text = "CO₂ emissions by weight class will be displayed here once GVWR data is available."
            placeholder_frame.paragraphs[0].font.size = Pt(16)
            placeholder_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        logger.debug("[prelim-deck] Added Emissions by Weight slide")
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Emissions by Weight slide: {e}")
        return False

def _add_electrification_timeline_weight_slide(prs: Presentation, data: dict) -> bool:
    """Add Electrification Timeline by Weight Class slide."""
    try:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "Electrification Timeline by Weight Class"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        vehicles = data.get('vehicles', [])
        
        # Add native PowerPoint stacked bar chart
        if not add_electrification_timeline_by_weight_chart(slide, vehicles, 1, 1.8, 8, 4.5):
            # Fallback to placeholder text
            placeholder_box = slide.shapes.add_textbox(Inches(2), Inches(3), Inches(6), Inches(2))
            placeholder_frame = placeholder_box.text_frame
            placeholder_frame.text = "Electrification timeline by weight class will be displayed here once GVWR data is available."
            placeholder_frame.paragraphs[0].font.size = Pt(16)
            placeholder_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        logger.debug("[prelim-deck] Added Electrification Timeline by Weight slide")
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Electrification Timeline by Weight slide: {e}")
        return False

def _add_electrification_timeline_body_slide(prs: Presentation, data: dict) -> bool:
    """Add Electrification Timeline by Body Type slide."""
    try:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "Electrification Timeline by Body Type"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        vehicles = data.get('vehicles', [])
        
        # Add native PowerPoint stacked bar chart
        if not add_electrification_timeline_by_body_type_chart(slide, vehicles, 1, 1.8, 8, 4.5):
            # Fallback to placeholder text
            placeholder_box = slide.shapes.add_textbox(Inches(2), Inches(3), Inches(6), Inches(2))
            placeholder_frame = placeholder_box.text_frame
            placeholder_frame.text = "Electrification timeline by body type will be displayed here once body class data is available."
            placeholder_frame.paragraphs[0].font.size = Pt(16)
            placeholder_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        logger.debug("[prelim-deck] Added Electrification Timeline by Body Type slide")
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Electrification Timeline by Body Type slide: {e}")
        return False

def _add_age_analysis_slide(prs: Presentation, data: dict) -> bool:
    """Add Fleet Age Analysis slide with native chart."""
    try:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "Fleet Age Analysis"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        vehicles = data.get('vehicles', [])
        
        # Add native PowerPoint column chart
        if not add_age_distribution_chart(slide, vehicles, 1, 1.8, 6, 4):
            # Fallback to placeholder text
            placeholder_box = slide.shapes.add_textbox(Inches(2), Inches(3), Inches(6), Inches(2))
            placeholder_frame = placeholder_box.text_frame
            placeholder_frame.text = "Fleet age analysis will be displayed here once vehicle year data is available."
            placeholder_frame.paragraphs[0].font.size = Pt(16)
            placeholder_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        else:
            # Add statistics alongside chart
            stats_box = slide.shapes.add_textbox(Inches(7.5), Inches(1.8), Inches(2), Inches(4))
            stats_frame = stats_box.text_frame
            
            # Calculate age statistics
            current_year = datetime.datetime.now().year
            ages = []
            for vehicle in vehicles:
                try:
                    year = int(getattr(vehicle.vehicle_id, 'year', 0))
                    if year > 0:
                        ages.append(current_year - year)
                except (ValueError, TypeError):
                    continue
            
            if ages:
                avg_age = sum(ages) / len(ages)
                median_age = sorted(ages)[len(ages)//2]
                stats_text = f"Age Statistics:\n\n"
                stats_text += f"• Average: {avg_age:.1f} years\n"
                stats_text += f"• Median: {median_age} years\n"
                stats_text += f"• Range: {min(ages)}-{max(ages)} years\n"
                stats_text += f"• Vehicles: {len(ages)}"
            else:
                stats_text = "Age Statistics:\n\nNo age data available"
            
            stats_frame.text = stats_text
            _apply_brand_formatting(stats_frame, 11, PRIMARY_HEX_1)
        
        logger.debug("[prelim-deck] Added Age Analysis slide")
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Age Analysis slide: {e}")
        return False

###############################################################################
# Enhanced Slide Functions (Simplified Implementation)
###############################################################################

def _add_data_completeness_slide(prs: Presentation, data: dict) -> bool:
    """Add Data Sources & Completeness slide."""
    try:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "Data Sources & Completeness"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Simplified completeness analysis
        vehicles = data.get('vehicles', [])
        total_vehicles = len(vehicles)
        
        content_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(5))
        content_frame = content_box.text_frame
        
        if total_vehicles > 0:
            # Calculate basic completeness
            vin_complete = sum(1 for v in vehicles if hasattr(v, 'vin') and v.vin)
            basic_complete = sum(1 for v in vehicles if hasattr(v, 'vehicle_id') and 
                               v.vehicle_id.year and v.vehicle_id.make and v.vehicle_id.model)
            mpg_complete = sum(1 for v in vehicles if hasattr(v, 'fuel_economy') and 
                             v.fuel_economy.combined_mpg > 0)
            
            content_text = f"Data Quality Summary:\n\n"
            content_text += f"• Total Vehicles: {total_vehicles:,}\n"
            content_text += f"• VIN Data: {vin_complete}/{total_vehicles} ({vin_complete/total_vehicles*100:.1f}%)\n"
            content_text += f"• Basic Info (Year/Make/Model): {basic_complete}/{total_vehicles} ({basic_complete/total_vehicles*100:.1f}%)\n"
            content_text += f"• Fuel Economy Data: {mpg_complete}/{total_vehicles} ({mpg_complete/total_vehicles*100:.1f}%)\n\n"
            content_text += f"• Data Sources: NHTSA VIN Database, FuelEconomy.gov API\n"
            content_text += f"• Processing Status: {sum(1 for v in vehicles if getattr(v, 'processing_success', True))} successful"
        else:
            content_text = "No vehicle data available for completeness analysis."
        
        content_frame.text = content_text
        _apply_brand_formatting(content_frame, 14, PRIMARY_HEX_1)
        
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Data Completeness slide: {e}")
        return False


# Orphaned slide builders removed in Phase 5 (Fix 40):
# _add_duty_age_slide, _add_screening_rules_slide, _add_candidates_slide,
# _add_example_replacements_slide, _add_energy_charging_slide, _add_cost_emissions_slide
# These were never registered in slide_builders and unreachable from the UI.


###############################################################################
# Phase 9D: Financial, Executive, and Replacement Schedule Slides
###############################################################################

def _add_financial_summary_slide(prs: Presentation, data: dict) -> bool:
    """Add Financial Summary slide with TCO KPI boxes and comparison chart.

    Shows key financial metrics (Total Investment, Annual Savings, Simple Payback,
    Lifetime CO2 Reduction) and a TCO comparison chart.
    """
    try:
        vehicles = data.get('vehicles', [])
        if not vehicles:
            return False

        from analysis.calculations import (
            calculate_annual_fuel_cost, calculate_annual_ev_cost,
            calculate_electrification_savings, calculate_emissions_reduction,
            DEFAULT_ICE_MAINTENANCE, DEFAULT_EV_MAINTENANCE,
            DEFAULT_ANNUAL_MILEAGE, DEFAULT_VEHICLE_LIFESPAN
        )

        # Compute fleet-wide financial metrics
        total_ev_investment = 0.0
        total_ice_replacement = 0.0
        total_annual_fuel_savings = 0.0
        total_annual_maint_savings = 0.0
        total_annual_co2_reduction = 0.0
        total_npv_savings = 0.0
        counted = 0

        for v in vehicles:
            ev_price = v.custom_fields.get("_ev_purchase_price")
            ice_price = v.custom_fields.get("_ice_purchase_price")
            if not ev_price or not ice_price:
                continue
            if not v.fuel_economy.combined_mpg:
                continue

            counted += 1
            mileage = v.annual_mileage or DEFAULT_ANNUAL_MILEAGE

            total_ev_investment += float(ev_price)
            total_ice_replacement += float(ice_price)

            savings = calculate_electrification_savings(v)
            total_annual_fuel_savings += savings.get("annual_fuel_savings", 0)
            total_annual_maint_savings += savings.get("annual_maintenance_savings", 0)
            total_npv_savings += savings.get("total_npv_savings", 0)
            total_annual_co2_reduction += savings.get("annual_co2_reduction", 0)

        if counted == 0:
            return False

        total_annual_savings = total_annual_fuel_savings + total_annual_maint_savings
        investment_premium = total_ev_investment - total_ice_replacement
        payback_years = investment_premium / total_annual_savings if total_annual_savings > 0 else float('inf')
        lifetime_co2 = total_annual_co2_reduction * DEFAULT_VEHICLE_LIFESPAN

        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.7))
        title_frame = title_box.text_frame
        title_frame.text = "Financial Summary"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        # Subtitle with fleet context
        sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.9), Inches(9), Inches(0.4))
        sub_frame = sub_box.text_frame
        sub_frame.text = f"Based on {counted} vehicles with EV equivalents identified | {DEFAULT_VEHICLE_LIFESPAN}-year analysis period"
        _apply_brand_formatting(sub_frame, 11, '#666666')
        sub_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        # KPI boxes row — 5 boxes across the top
        kpis = [
            ("EV Investment", f"${total_ev_investment:,.0f}"),
            ("Investment Premium", f"${investment_premium:,.0f}"),
            ("Annual Savings", f"${total_annual_savings:,.0f}"),
            ("Simple Payback", f"{payback_years:.1f} years" if payback_years < 100 else "N/A"),
            ("Lifetime CO₂ Saved", f"{lifetime_co2:,.0f} MT"),
        ]

        box_width = 1.7
        box_gap = 0.1
        start_x = 0.35

        for i, (label, value) in enumerate(kpis):
            x = start_x + i * (box_width + box_gap)

            # KPI value
            val_box = slide.shapes.add_textbox(Inches(x), Inches(1.5), Inches(box_width), Inches(0.5))
            val_frame = val_box.text_frame
            val_frame.text = value
            val_frame.paragraphs[0].font.size = Pt(16)
            val_frame.paragraphs[0].font.bold = True
            val_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            r, g, b = _hex_to_rgb(PRIMARY_HEX_3)
            val_frame.paragraphs[0].font.color.rgb = RGBColor(r, g, b)

            # KPI label
            lbl_box = slide.shapes.add_textbox(Inches(x), Inches(1.95), Inches(box_width), Inches(0.35))
            lbl_frame = lbl_box.text_frame
            lbl_frame.text = label
            lbl_frame.paragraphs[0].font.size = Pt(9)
            lbl_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            r, g, b = _hex_to_rgb('#666666')
            lbl_frame.paragraphs[0].font.color.rgb = RGBColor(r, g, b)

        # TCO Comparison chart (left half)
        add_tco_comparison_chart(slide, vehicles, 0.3, 2.5, 4.5, 3.5)

        # Payback timeline chart (right half)
        add_payback_timeline_chart(slide, vehicles, 5.0, 2.5, 4.5, 3.5)

        # Bottom footnote
        foot_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.3), Inches(9), Inches(0.5))
        foot_frame = foot_box.text_frame
        foot_frame.text = (
            f"Assumptions: $3.50/gal gas, $0.13/kWh electricity, "
            f"3% annual fuel price escalation, {DEFAULT_VEHICLE_LIFESPAN}-year ownership period. "
            f"NPV at 5% discount rate: ${total_npv_savings:,.0f}."
        )
        _apply_brand_formatting(foot_frame, 8, '#999999')

        logger.debug("[prelim-deck] Added Financial Summary slide")
        return True

    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Financial Summary slide: {e}")
        return False


def _add_executive_recommendations_slide(prs: Presentation, data: dict) -> bool:
    """Add Executive Recommendations slide with data-driven insights.

    Shows a narrative recommendation paragraph based on actual fleet data,
    followed by a summary table of top 5 priority replacements and
    risk/opportunity callouts.
    """
    try:
        vehicles = data.get('vehicles', [])
        if not vehicles:
            return False

        # Gather fleet metrics for narrative
        total = len(vehicles)
        ev_matched = sum(1 for v in vehicles if v.custom_fields.get("EV Equivalent"))
        acf_subject = sum(1 for v in vehicles if v.custom_fields.get("_acf_code") == "B")

        # Count vehicles by proposed year phase
        current_year = datetime.datetime.now().year
        phase1 = []  # Next 2 years
        phase2 = []  # Years 3-5
        phase3 = []  # Years 6+
        no_ev_year = 0

        for v in vehicles:
            ev_year_str = v.custom_fields.get("Proposed EV Year", "")
            try:
                ev_year = int(ev_year_str)
            except (ValueError, TypeError):
                no_ev_year += 1
                continue
            if ev_year <= current_year + 2:
                phase1.append(v)
            elif ev_year <= current_year + 5:
                phase2.append(v)
            else:
                phase3.append(v)

        # Top vehicle types in Phase 1
        phase1_types = {}
        for v in phase1:
            vtype = f"{v.vehicle_id.make or ''} {v.vehicle_id.model or ''}".strip()
            phase1_types[vtype] = phase1_types.get(vtype, 0) + 1
        top_phase1_types = sorted(phase1_types.items(), key=lambda x: x[1], reverse=True)[:3]

        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.7))
        title_frame = title_box.text_frame
        title_frame.text = "Executive Recommendations"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        # Build narrative recommendation
        narrative_parts = [
            f"Based on analysis of {total} fleet vehicles, "
            f"{ev_matched} have been matched to available EV replacements."
        ]

        if phase1:
            type_list = ", ".join(f"{name} ({count})" for name, count in top_phase1_types) if top_phase1_types else "various types"
            narrative_parts.append(
                f"We recommend replacing {len(phase1)} vehicles in Phase 1 (FY{current_year+1}-{current_year+2}), "
                f"prioritizing {type_list}."
            )

        if acf_subject > 0:
            narrative_parts.append(
                f"{acf_subject} vehicles are subject to CARB ACF compliance requirements "
                f"and should be prioritized for electrification."
            )

        if phase2:
            narrative_parts.append(
                f"Phase 2 ({current_year+3}-{current_year+5}) targets {len(phase2)} additional vehicles."
            )

        narrative = " ".join(narrative_parts)

        narr_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(1.2))
        narr_frame = narr_box.text_frame
        narr_frame.word_wrap = True
        narr_frame.text = narrative
        _apply_brand_formatting(narr_frame, 12, PRIMARY_HEX_1)

        # Top 5 priority replacements table
        table_label = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(0.4))
        tl_frame = table_label.text_frame
        tl_frame.text = "Top Priority Replacements"
        _apply_brand_formatting(tl_frame, 14, PRIMARY_HEX_1, True)

        # Collect priority vehicles (with EV match and financial data)
        priority_vehicles = []
        for v in vehicles:
            ev_equiv = v.custom_fields.get("EV Equivalent", "")
            ev_year = v.custom_fields.get("Proposed EV Year", "")
            if ev_equiv and ev_year and ev_year not in ("N/A", "Exempt", ""):
                priority_vehicles.append(v)

        # Sort by proposed year then by annual mileage (higher first)
        priority_vehicles.sort(key=lambda v: (
            int(v.custom_fields.get("Proposed EV Year", "9999")),
            -(v.annual_mileage or 0)
        ))
        top5 = priority_vehicles[:5]

        if top5:
            # Create table: 5 rows + header, 5 columns
            cols = 5
            rows = len(top5) + 1
            table_shape = slide.shapes.add_table(
                rows, cols,
                Inches(0.3), Inches(3.0),
                Inches(9.4), Inches(0.3 + rows * 0.35)
            )
            table = table_shape.table

            # Set column widths
            table.columns[0].width = Inches(2.5)  # Current Vehicle
            table.columns[1].width = Inches(2.5)  # Proposed EV
            table.columns[2].width = Inches(1.3)  # Proposed Year
            table.columns[3].width = Inches(1.5)  # EV MSRP Range
            table.columns[4].width = Inches(1.6)  # ACF Status

            # Header row
            headers = ["Current Vehicle", "Proposed EV", "Target Year", "EV MSRP Range", "ACF Status"]
            for j, header in enumerate(headers):
                cell = table.cell(0, j)
                cell.text = header
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.size = Pt(9)
                    paragraph.font.bold = True
                    r, g, b = _hex_to_rgb(PRIMARY_HEX_2)
                    paragraph.font.color.rgb = RGBColor(r, g, b)
                # Header background
                r, g, b = _hex_to_rgb(PRIMARY_HEX_1)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(r, g, b)

            # Data rows
            for i, v in enumerate(top5):
                row_idx = i + 1
                current = f"{v.vehicle_id.year or ''} {v.vehicle_id.make or ''} {v.vehicle_id.model or ''}".strip()
                ev_equiv = v.custom_fields.get("EV Equivalent", "")
                ev_year = v.custom_fields.get("Proposed EV Year", "")
                ev_msrp = v.custom_fields.get("EV MSRP Range", "")
                acf_detail = v.custom_fields.get("ACF Detail", "")

                row_data = [current, ev_equiv, ev_year, ev_msrp, acf_detail]
                for j, val in enumerate(row_data):
                    cell = table.cell(row_idx, j)
                    cell.text = str(val)
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.size = Pt(9)

                # Alternate row shading
                if i % 2 == 1:
                    for j in range(cols):
                        cell = table.cell(row_idx, j)
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(245, 245, 245)

        # Risk/Opportunity callouts at bottom
        callout_y = 3.0 + (len(top5) + 1) * 0.35 + 0.3 if top5 else 3.5
        if callout_y > 6.5:
            callout_y = 6.2

        callouts = []
        vehicles_no_ev = sum(1 for v in vehicles if not v.custom_fields.get("EV Equivalent"))
        if vehicles_no_ev > 0:
            callouts.append(f"⚠ {vehicles_no_ev} vehicles have no available EV equivalent — consider operational changes or future models")
        if acf_subject > 0:
            callouts.append(f"📋 {acf_subject} vehicles require CARB ACF compliance — these must be prioritized")
        if phase1:
            callouts.append(f"✅ {len(phase1)} vehicles recommended for immediate replacement (highest ROI and/or oldest)")

        if callouts:
            callout_box = slide.shapes.add_textbox(Inches(0.5), Inches(callout_y), Inches(9), Inches(1.0))
            cf = callout_box.text_frame
            cf.word_wrap = True
            cf.text = callouts[0]
            _apply_brand_formatting(cf, 10, '#555555')
            for callout in callouts[1:]:
                p = cf.add_paragraph()
                p.text = callout
                p.font.size = Pt(10)
                r, g, b = _hex_to_rgb('#555555')
                p.font.color.rgb = RGBColor(r, g, b)

        logger.debug("[prelim-deck] Added Executive Recommendations slide")
        return True

    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Executive Recommendations slide: {e}")
        return False


def _add_replacement_schedule_slide(prs: Presentation, data: dict) -> bool:
    """Add Replacement Schedule slide with a table of top 12-15 priority vehicles.

    Shows: Current Vehicle, Department, Proposed EV, Annual Savings, Payback, Target Year.
    Vehicles sorted by proposed year then NPV savings.
    """
    try:
        vehicles = data.get('vehicles', [])
        if not vehicles:
            return False

        from analysis.calculations import (
            calculate_electrification_savings, DEFAULT_VEHICLE_LIFESPAN
        )

        # Collect vehicles with EV matches and proposed years
        candidates = []
        for v in vehicles:
            ev_equiv = v.custom_fields.get("EV Equivalent", "")
            ev_year = v.custom_fields.get("Proposed EV Year", "")
            if not ev_equiv or ev_year in ("N/A", "Exempt", ""):
                continue

            # Calculate savings for each vehicle
            savings = calculate_electrification_savings(v)
            annual_savings = savings.get("annual_fuel_savings", 0) + savings.get("annual_maintenance_savings", 0)
            npv_savings = savings.get("total_npv_savings", 0)

            ev_price = v.custom_fields.get("_ev_purchase_price", 0)
            ice_price = v.custom_fields.get("_ice_purchase_price", 0)
            premium = float(ev_price or 0) - float(ice_price or 0)
            payback = premium / annual_savings if annual_savings > 0 else float('inf')

            candidates.append({
                "vehicle": v,
                "current": f"{v.vehicle_id.year or ''} {v.vehicle_id.make or ''} {v.vehicle_id.model or ''}".strip(),
                "department": v.department or "—",
                "ev_equiv": ev_equiv,
                "annual_savings": annual_savings,
                "npv_savings": npv_savings,
                "payback": payback,
                "ev_year": ev_year,
            })

        if not candidates:
            return False

        # Sort: by proposed year (ascending), then by NPV savings (descending)
        candidates.sort(key=lambda c: (
            int(c["ev_year"]) if c["ev_year"].isdigit() else 9999,
            -c["npv_savings"]
        ))

        # Limit to 12 rows to fit on one slide
        top_vehicles = candidates[:12]

        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
        title_frame = title_box.text_frame
        title_frame.text = "Replacement Schedule — Priority Vehicles"
        _apply_brand_formatting(title_frame, 22, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        # Subtitle
        sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.85), Inches(9), Inches(0.3))
        sub_frame = sub_box.text_frame
        sub_frame.text = f"Top {len(top_vehicles)} vehicles ranked by replacement priority | {len(candidates)} total candidates identified"
        _apply_brand_formatting(sub_frame, 10, '#666666')
        sub_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        # Table
        cols = 6
        rows = len(top_vehicles) + 1  # +1 for header

        # Calculate available height
        table_top = 1.3
        table_height = min(0.3 + rows * 0.35, 5.8)

        table_shape = slide.shapes.add_table(
            rows, cols,
            Inches(0.2), Inches(table_top),
            Inches(9.6), Inches(table_height)
        )
        table = table_shape.table

        # Set column widths
        table.columns[0].width = Inches(2.2)  # Current Vehicle
        table.columns[1].width = Inches(1.2)  # Department
        table.columns[2].width = Inches(2.2)  # Proposed EV
        table.columns[3].width = Inches(1.3)  # Annual Savings
        table.columns[4].width = Inches(1.3)  # Payback
        table.columns[5].width = Inches(1.0)  # Target Year

        # Header row
        headers = ["Current Vehicle", "Department", "Proposed EV", "Annual Savings", "Payback (yrs)", "Target Year"]
        for j, header in enumerate(headers):
            cell = table.cell(0, j)
            cell.text = header
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(9)
                paragraph.font.bold = True
                r, g, b = _hex_to_rgb(PRIMARY_HEX_2)
                paragraph.font.color.rgb = RGBColor(r, g, b)
            r, g, b = _hex_to_rgb(PRIMARY_HEX_1)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(r, g, b)

        # Data rows
        total_annual_savings = 0
        for i, c in enumerate(top_vehicles):
            row_idx = i + 1
            payback_str = f"{c['payback']:.1f}" if c['payback'] < 100 else "—"
            savings_str = f"${c['annual_savings']:,.0f}"
            total_annual_savings += c['annual_savings']

            row_data = [
                c["current"], c["department"], c["ev_equiv"],
                savings_str, payback_str, c["ev_year"]
            ]

            for j, val in enumerate(row_data):
                cell = table.cell(row_idx, j)
                cell.text = str(val)
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.size = Pt(8)
                    # Right-align financial columns
                    if j in (3, 4, 5):
                        paragraph.alignment = PP_ALIGN.RIGHT

            # Alternate row shading
            if i % 2 == 1:
                for j in range(cols):
                    cell = table.cell(row_idx, j)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(245, 245, 245)

        # Summary footer below table
        footer_y = table_top + table_height + 0.15
        if footer_y < 7.0:
            foot_box = slide.shapes.add_textbox(Inches(0.5), Inches(footer_y), Inches(9), Inches(0.4))
            foot_frame = foot_box.text_frame
            foot_frame.text = (
                f"Combined annual savings for these {len(top_vehicles)} vehicles: "
                f"${total_annual_savings:,.0f}/year"
            )
            _apply_brand_formatting(foot_frame, 11, PRIMARY_HEX_3, True)
            foot_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        logger.debug("[prelim-deck] Added Replacement Schedule slide")
        return True

    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Replacement Schedule slide: {e}")
        return False


def _add_scenario_comparison_slide(prs: Presentation, data: dict) -> bool:
    """Add Scenario Comparison slide showing multiple electrification timelines.

    Runs 3-4 preset scenarios and shows:
    - Vehicles electrified over time (multi-line chart)
    - Side-by-side metrics table (investment, savings, CO₂, payback)
    """
    try:
        vehicles = data.get('vehicles', [])
        if not vehicles:
            return False

        from analysis.scenarios import compare_scenarios

        # Use consultant-selected scenarios from the Present panel, or all 4 presets
        scenario_names = data.get(
            'selected_scenarios',
            ["aggressive", "moderate", "conservative", "acf_compliance"]
        )
        comparison = compare_scenarios(
            vehicles,
            scenario_names=scenario_names
        )

        scenario_results = comparison.get("scenarios", [])
        if not scenario_results or all(r["total_vehicles"] == 0 for r in scenario_results):
            return False

        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
        title_frame = title_box.text_frame
        title_frame.text = "Scenario Comparison — Electrification Timelines"
        _apply_brand_formatting(title_frame, 22, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        # Vehicles electrified chart (left)
        add_scenario_comparison_chart(
            slide, scenario_results,
            left=0.3, top=1.1, width=4.6, height=3.0,
            metric="vehicles"
        )

        # Investment chart (right)
        add_scenario_comparison_chart(
            slide, scenario_results,
            left=5.1, top=1.1, width=4.6, height=3.0,
            metric="cost"
        )

        # Summary comparison table below charts
        comp_table = comparison.get("comparison_table", [])
        if comp_table:
            valid_rows = [c for c in comp_table if c["vehicles"] > 0]
            if valid_rows:
                cols = 6
                rows = len(valid_rows) + 1

                table_shape = slide.shapes.add_table(
                    rows, cols,
                    Inches(0.3), Inches(4.3),
                    Inches(9.4), Inches(0.3 + rows * 0.35)
                )
                table = table_shape.table

                table.columns[0].width = Inches(1.8)
                table.columns[1].width = Inches(1.0)
                table.columns[2].width = Inches(1.5)
                table.columns[3].width = Inches(1.8)
                table.columns[4].width = Inches(1.8)
                table.columns[5].width = Inches(1.5)

                headers = ["Scenario", "Vehicles", "End Year", "Total Investment", "Annual Savings", "Payback (yrs)"]
                for j, header in enumerate(headers):
                    cell = table.cell(0, j)
                    cell.text = header
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.size = Pt(9)
                        paragraph.font.bold = True
                        r, g, b = _hex_to_rgb(PRIMARY_HEX_2)
                        paragraph.font.color.rgb = RGBColor(r, g, b)
                    r, g, b = _hex_to_rgb(PRIMARY_HEX_1)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(r, g, b)

                for i, c in enumerate(valid_rows):
                    row_idx = i + 1
                    payback = c.get("payback_years", float('inf'))
                    payback_str = f"{payback:.1f}" if payback < 100 else "—"
                    row_data = [
                        c["name"],
                        str(c["vehicles"]),
                        str(c["end_year"]),
                        f"${c['total_investment']:,.0f}",
                        f"${c['total_annual_savings']:,.0f}/yr",
                        payback_str,
                    ]
                    for j, val in enumerate(row_data):
                        cell = table.cell(row_idx, j)
                        cell.text = val
                        for paragraph in cell.text_frame.paragraphs:
                            paragraph.font.size = Pt(9)
                            if j >= 2:
                                paragraph.alignment = PP_ALIGN.RIGHT

                    if i % 2 == 1:
                        for j in range(cols):
                            cell = table.cell(row_idx, j)
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = RGBColor(245, 245, 245)

        # Best-in-class callouts
        best_roi = comparison.get("best_roi", "")
        fastest = comparison.get("fastest", "")
        lowest_cost = comparison.get("lowest_cost", "")

        callouts = []
        if best_roi:
            callouts.append(f"Best ROI: {best_roi}")
        if fastest:
            callouts.append(f"Fastest completion: {fastest}")
        if lowest_cost:
            callouts.append(f"Lowest total cost: {lowest_cost}")

        if callouts:
            footer_y = 4.3 + (len(valid_rows) + 1) * 0.35 + 0.2 if valid_rows else 6.0
            if footer_y > 6.8:
                footer_y = 6.8
            foot_box = slide.shapes.add_textbox(Inches(0.5), Inches(footer_y), Inches(9), Inches(0.4))
            foot_frame = foot_box.text_frame
            foot_frame.text = " | ".join(callouts)
            _apply_brand_formatting(foot_frame, 10, PRIMARY_HEX_3, True)
            foot_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        logger.debug("[prelim-deck] Added Scenario Comparison slide")
        return True

    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Scenario Comparison slide: {e}")
        return False


def _add_next_steps_slide(prs: Presentation, data: dict) -> bool:
    """Add data-driven Next Steps roadmap slide.

    Uses actual fleet analysis results to generate specific, actionable
    next steps instead of generic boilerplate.
    """
    try:
        vehicles = data.get('vehicles', [])
        summary = data.get('summary', {})
        current_year = datetime.datetime.now().year

        # Count vehicles by phase
        phase1_count = 0
        phase2_count = 0
        phase3_count = 0
        acf_count = 0
        ev_matched = 0

        for v in vehicles:
            ev_year_str = v.custom_fields.get("Proposed EV Year", "") if hasattr(v, 'custom_fields') else ""
            acf_code = v.custom_fields.get("_acf_code", "") if hasattr(v, 'custom_fields') else ""
            ev_equiv = v.custom_fields.get("EV Equivalent", "") if hasattr(v, 'custom_fields') else ""

            if ev_equiv:
                ev_matched += 1
            if acf_code == "B":
                acf_count += 1

            try:
                ev_year = int(ev_year_str)
                if ev_year <= current_year + 2:
                    phase1_count += 1
                elif ev_year <= current_year + 5:
                    phase2_count += 1
                else:
                    phase3_count += 1
            except (ValueError, TypeError):
                pass

        total = len(vehicles)

        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.7))
        title_frame = title_box.text_frame
        title_frame.text = "Next Steps & Implementation Roadmap"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        # Build data-driven roadmap
        roadmap_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(5.5))
        roadmap_frame = roadmap_box.text_frame
        roadmap_frame.word_wrap = True

        # Phase 1: Immediate (data-driven)
        lines = []
        lines.append(f"Phase 1: Immediate Actions (FY{current_year + 1})")
        if phase1_count > 0:
            lines.append(f"  • Begin procurement for {phase1_count} priority vehicle replacements")
        lines.append(f"  • Validate operational data and confirm fleet utilization patterns")
        if acf_count > 0:
            lines.append(f"  • Develop CARB ACF compliance strategy for {acf_count} regulated vehicles")
        lines.append(f"  • Conduct site assessments for charging infrastructure")
        lines.append(f"  • Engage utility provider for electrical capacity planning")
        lines.append("")

        # Phase 2: Near-term
        lines.append(f"Phase 2: Near-Term Deployment (FY{current_year + 2}-{current_year + 3})")
        if phase1_count > 0:
            lines.append(f"  • Deploy first {min(phase1_count, 20)} EVs with charging infrastructure")
        else:
            lines.append(f"  • Deploy pilot EV vehicles with charging infrastructure")
        lines.append(f"  • Monitor performance, energy costs, and driver feedback")
        if phase2_count > 0:
            lines.append(f"  • Prepare procurement for {phase2_count} Phase 2 replacements")
        lines.append(f"  • Evaluate and apply for available federal/state incentives")
        lines.append("")

        # Phase 3: Expansion
        lines.append(f"Phase 3: Fleet-Wide Expansion (FY{current_year + 4}-{current_year + 7})")
        lines.append(f"  • Scale EV deployment based on pilot learnings")
        if phase3_count > 0:
            lines.append(f"  • Complete remaining {phase3_count} vehicle transitions")
        lines.append(f"  • Optimize charging schedules and energy management")
        lines.append(f"  • Report on emissions reductions and cost savings")
        lines.append("")

        # Immediate 30-day actions
        lines.append("Immediate Next Steps (30 days):")
        lines.append(f"  • Stakeholder alignment meeting to review {total}-vehicle analysis")
        if phase1_count > 0:
            lines.append(f"  • Budget development for Phase 1 ({phase1_count} vehicles)")
        else:
            lines.append(f"  • Budget development for pilot program")
        lines.append(f"  • Utility preliminary discussions for site capacity")
        if acf_count > 0:
            lines.append(f"  • ACF compliance timeline review with regulatory team")

        roadmap_frame.text = "\n".join(lines)
        _apply_brand_formatting(roadmap_frame, 11, PRIMARY_HEX_1)

        # Bold the phase headers
        for paragraph in roadmap_frame.paragraphs:
            text = paragraph.text
            if text.startswith("Phase ") or text.startswith("Immediate Next"):
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.size = Pt(13)
                    r, g, b = _hex_to_rgb(PRIMARY_HEX_3)
                    run.font.color.rgb = RGBColor(r, g, b)

        return True

    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Next Steps slide: {e}")
        return False

###############################################################################
# Utility Functions
###############################################################################

def _generate_output_path(data: dict) -> str:
    """Generate a default output path for the presentation."""
    # Clean the fleet name for use in filename
    fleet_name = data.get('fleet_name', 'Fleet')
    clean_name = ''.join(c if c.isalnum() or c in ' _-' else '_' for c in fleet_name)
    clean_name = clean_name.replace(' ', '_')
    
    # Add timestamp
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    
    # Create filename
    filename = f"{clean_name}_{data.get('stage', 'Prelim')}_{timestamp}.pptx"
    
    # Try to use exports directory
    exports_dir = Path(__file__).parent / "data" / "exports"
    if exports_dir.exists():
        return str(exports_dir / filename)
    else:
        # Try to create exports directory in current directory
        local_exports = Path("exports")
        local_exports.mkdir(exist_ok=True)
        return str(local_exports / filename)
