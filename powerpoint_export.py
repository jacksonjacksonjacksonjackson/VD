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
    add_electrification_timeline_by_body_type_chart, add_age_distribution_chart
)

# Set up module logger
logger = logging.getLogger(__name__)

###############################################################################
# Main Export Function
###############################################################################

def export_prelim_deck(data: dict, template_path: str | None = None, out_path: str | None = None, 
                      slide_config: SlideConfiguration | None = None) -> str:
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
        'emissions_timeline': _add_emissions_timeline_slide,
        'emissions_by_weight': _add_emissions_by_weight_slide,
        'electrification_timeline_weight': _add_electrification_timeline_weight_slide,
        'electrification_timeline_body': _add_electrification_timeline_body_slide,
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

def _get_template_presentation(template_path: str | None = None) -> Presentation:
    """Get template presentation - simplified version."""
    # For now, just create a blank presentation
    # In production, this would implement the full template discovery logic
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
        subtitle_frame.text = f"{data['stage']} Fleet Electrification Analysis"
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
        
        # Optony branding
        brand_box = slide.shapes.add_textbox(Inches(4), Inches(8), Inches(2), Inches(0.5))
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

def _add_duty_age_slide(prs: Presentation, data: dict) -> bool:
    """Add Duty & Age Analysis slide."""
    try:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "Fleet Duty & Age Analysis"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        vehicles = data.get('vehicles', [])
        
        # Try to export age distribution chart
        chart_path = _export_chart_to_image("Fleet Age Distribution", vehicles)
        if chart_path:
            _add_chart_to_slide(slide, chart_path, 1, 1.8, 5, 3)
            try:
                os.unlink(chart_path)
            except:
                pass
        
        # Add statistics
        stats_box = slide.shapes.add_textbox(Inches(6.5), Inches(1.8), Inches(3), Inches(3))
        stats_frame = stats_box.text_frame
        
        if vehicles:
            ages = [getattr(v, 'age', 0) for v in vehicles if hasattr(v, 'age') and v.age > 0]
            mileages = [getattr(v, 'annual_mileage', 0) for v in vehicles 
                       if hasattr(v, 'annual_mileage') and v.annual_mileage > 0]
            
            if ages:
                avg_age = sum(ages) / len(ages)
                stats_text = f"Fleet Statistics:\n\n"
                stats_text += f"• Average Age: {avg_age:.1f} years\n"
                stats_text += f"• Age Range: {min(ages)}-{max(ages)} years\n\n"
            else:
                stats_text = "Fleet Statistics:\n\n• No age data available\n\n"
            
            if mileages:
                avg_mileage = sum(mileages) / len(mileages)
                stats_text += f"• Avg Annual Mileage: {avg_mileage:,.0f}\n"
                stats_text += f"• Mileage Range: {min(mileages):,.0f}-{max(mileages):,.0f}"
            else:
                stats_text += "• No mileage data available"
        else:
            stats_text = "No vehicle data available for analysis."
        
        stats_frame.text = stats_text
        _apply_brand_formatting(stats_frame, 12, PRIMARY_HEX_1)
        
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Duty & Age slide: {e}")
        return False

def _add_screening_rules_slide(prs: Presentation, data: dict) -> bool:
    """Add Screening Rules slide."""
    try:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "Electrification Screening Rules"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Screening criteria
        criteria_box = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(8), Inches(5))
        criteria_frame = criteria_box.text_frame
        
        criteria_text = "Phase-Based Screening Approach:\n\n"
        criteria_text += "Phase 1 - Priority Candidates:\n"
        criteria_text += "• Light duty vehicles (GVWR ≤ 8,500 lbs)\n"
        criteria_text += "• High annual mileage (≥ 12,000 miles/year)\n"
        criteria_text += "• Poor fuel economy (< 25 MPG)\n"
        criteria_text += "• Older vehicles (≥ 8 years)\n\n"
        
        criteria_text += "Phase 2 - Secondary Candidates:\n"
        criteria_text += "• Medium duty vehicles (8,501-19,500 lbs GVWR)\n"
        criteria_text += "• Predictable daily routes\n"
        criteria_text += "• Depot-based operations\n\n"
        
        criteria_text += "Phase 3 - Future Consideration:\n"
        criteria_text += "• Heavy duty vehicles (> 19,500 lbs GVWR)\n"
        criteria_text += "• Specialized applications\n"
        criteria_text += "• Awaiting technology maturity"
        
        criteria_frame.text = criteria_text
        _apply_brand_formatting(criteria_frame, 12, PRIMARY_HEX_1)
        
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Screening Rules slide: {e}")
        return False

def _add_candidates_slide(prs: Presentation, data: dict) -> bool:
    """Add Electrification Candidates slide."""
    try:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "Electrification Candidates & Phasing"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        vehicles = data.get('vehicles', [])
        
        # Simplified candidate analysis
        total_vehicles = len(vehicles)
        light_duty = sum(1 for v in vehicles if hasattr(v, 'vehicle_id') and 
                        v.vehicle_id.gvwr_pounds <= 8500)
        medium_duty = sum(1 for v in vehicles if hasattr(v, 'vehicle_id') and 
                         8500 < v.vehicle_id.gvwr_pounds <= 19500)
        heavy_duty = sum(1 for v in vehicles if hasattr(v, 'vehicle_id') and 
                        v.vehicle_id.gvwr_pounds > 19500)
        
        content_box = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(8), Inches(5))
        content_frame = content_box.text_frame
        
        content_text = f"Candidate Analysis Summary:\n\n"
        content_text += f"Total Fleet Size: {total_vehicles:,} vehicles\n\n"
        
        if total_vehicles > 0:
            content_text += f"Phase 1 Candidates (Light Duty): {light_duty} vehicles ({light_duty/total_vehicles*100:.1f}%)\n"
            content_text += f"Phase 2 Candidates (Medium Duty): {medium_duty} vehicles ({medium_duty/total_vehicles*100:.1f}%)\n"
            content_text += f"Phase 3 Candidates (Heavy Duty): {heavy_duty} vehicles ({heavy_duty/total_vehicles*100:.1f}%)\n\n"
            
            # Estimate savings potential
            est_annual_savings = light_duty * 2000 + medium_duty * 1500 + heavy_duty * 1000  # Simplified
            content_text += f"Estimated Annual Savings Potential: ${est_annual_savings:,}\n\n"
            
            content_text += "Recommendations:\n"
            content_text += f"• Start with Phase 1: Focus on {light_duty} light-duty vehicles\n"
            content_text += "• Consider pilot program with 10-20 highest-priority vehicles\n"
            content_text += "• Develop infrastructure plan for depot charging"
        else:
            content_text += "No vehicle data available for candidate analysis."
        
        content_frame.text = content_text
        _apply_brand_formatting(content_frame, 12, PRIMARY_HEX_1)
        
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Candidates slide: {e}")
        return False

def _add_example_replacements_slide(prs: Presentation, data: dict) -> bool:
    """Add Example EV Replacements slide."""
    try:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "Example EV Replacements"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Example replacements content
        content_box = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(8), Inches(5))
        content_frame = content_box.text_frame
        
        content_text = "Common Fleet Vehicle → EV Replacements:\n\n"
        content_text += "Light-Duty Examples:\n"
        content_text += "• Ford F-150 → Ford F-150 Lightning\n"
        content_text += "• Chevrolet Silverado → Chevrolet Silverado EV\n"
        content_text += "• Ford Transit Van → Ford E-Transit\n\n"
        
        content_text += "Medium-Duty Examples:\n"
        content_text += "• Isuzu NPR → Isuzu NPR-EV\n"
        content_text += "• Freightliner Sprinter → eSprinter\n"
        content_text += "• Ford Transit 350 → Ford E-Transit HD\n\n"
        
        content_text += "Analysis Considerations:\n"
        content_text += "• Range requirements vs. daily usage patterns\n"
        content_text += "• Payload capacity impact\n"
        content_text += "• Total cost of ownership over vehicle lifecycle\n"
        content_text += "• Charging infrastructure compatibility\n\n"
        
        content_text += "Next Steps: Vehicle-specific analysis with operational requirements"
        
        content_frame.text = content_text
        _apply_brand_formatting(content_frame, 12, PRIMARY_HEX_1)
        
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Example Replacements slide: {e}")
        return False

def _add_energy_charging_slide(prs: Presentation, data: dict) -> bool:
    """Add Energy & Charging Infrastructure slide."""
    try:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "Energy & Charging Infrastructure"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        vehicles = data.get('vehicles', [])
        
        # Simplified charging analysis
        content_box = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(8), Inches(5))
        content_frame = content_box.text_frame
        
        if vehicles:
            total_vehicles = len(vehicles)
            # Simplified calculations
            est_daily_miles = sum(getattr(v, 'annual_mileage', 12000) for v in vehicles) / 365
            est_daily_energy = est_daily_miles * 0.30  # kWh/mile
            est_l2_chargers = max(1, int(total_vehicles * 0.8))  # 80% of fleet
            est_dcfc_chargers = max(1, int(total_vehicles * 0.1))  # 10% of fleet
            est_cost = est_l2_chargers * 4000 + est_dcfc_chargers * 50000
            
            content_text = f"Infrastructure Requirements Estimate:\n\n"
            content_text += f"Fleet Size: {total_vehicles} vehicles\n"
            content_text += f"Estimated Daily Energy Need: {est_daily_energy:,.0f} kWh\n\n"
            
            content_text += f"Charging Infrastructure:\n"
            content_text += f"• Level 2 Chargers (7.2 kW): {est_l2_chargers} units\n"
            content_text += f"• DC Fast Chargers (50 kW): {est_dcfc_chargers} units\n"
            content_text += f"• Estimated Installation Cost: ${est_cost:,}\n\n"
            
            content_text += "Key Assumptions:\n"
            content_text += "• Average EV efficiency: 0.30 kWh/mile\n"
            content_text += "• Overnight charging for 80% of fleet\n"
            content_text += "• DC fast charging for operational flexibility\n"
            content_text += "• Installation costs: $4K per L2, $50K per DCFC\n\n"
            
            content_text += "Recommendation: Detailed site assessment required"
        else:
            content_text = "No vehicle data available for charging infrastructure analysis."
        
        content_frame.text = content_text
        _apply_brand_formatting(content_frame, 12, PRIMARY_HEX_1)
        
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Energy & Charging slide: {e}")
        return False

def _add_cost_emissions_slide(prs: Presentation, data: dict) -> bool:
    """Add Cost & Emissions Analysis slide."""
    try:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "Cost & Emissions Analysis"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        vehicles = data.get('vehicles', [])
        
        # Try to add cost comparison chart
        chart_path = _export_chart_to_image("Annual Cost Comparison", vehicles)
        if chart_path:
            _add_chart_to_slide(slide, chart_path, 1, 1.8, 5, 3)
            try:
                os.unlink(chart_path)
            except:
                pass
        
        # Financial analysis
        analysis_box = slide.shapes.add_textbox(Inches(6.5), Inches(1.8), Inches(3), Inches(5))
        analysis_frame = analysis_box.text_frame
        
        if vehicles:
            # Simplified financial analysis
            total_vehicles = len(vehicles)
            est_annual_fuel_savings = total_vehicles * 1500  # Simplified estimate
            est_co2_reduction = total_vehicles * 4.6  # metric tons per vehicle
            
            analysis_text = f"Financial Impact:\n\n"
            analysis_text += f"• Fleet Size: {total_vehicles} vehicles\n"
            analysis_text += f"• Est. Annual Fuel Savings: ${est_annual_fuel_savings:,}\n"
            analysis_text += f"• 10-Year NPV Savings: ${est_annual_fuel_savings * 8:,}\n\n"
            
            analysis_text += f"Environmental Impact:\n\n"
            analysis_text += f"• Est. CO₂ Reduction: {est_co2_reduction:,.1f} tons/year\n"
            analysis_text += f"• Equivalent to removing {est_co2_reduction/4.6:.0f} cars\n\n"
            
            analysis_text += "Key Assumptions:\n"
            analysis_text += "• Gas: $3.50/gal\n"
            analysis_text += "• Electricity: $0.13/kWh\n"
            analysis_text += "• 10-year analysis period"
        else:
            analysis_text = "No vehicle data available for cost/emissions analysis."
        
        analysis_frame.text = analysis_text
        _apply_brand_formatting(analysis_frame, 11, PRIMARY_HEX_1)
        
        return True
        
    except Exception as e:
        logger.error(f"[prelim-deck] Failed to create Cost & Emissions slide: {e}")
        return False

def _add_next_steps_slide(prs: Presentation, data: dict) -> bool:
    """Add Next Steps roadmap slide."""
    try:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = "Next Steps & Implementation Roadmap"
        _apply_brand_formatting(title_frame, 24, PRIMARY_HEX_1, True)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Implementation roadmap
        roadmap_box = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(8), Inches(5))
        roadmap_frame = roadmap_box.text_frame
        
        roadmap_text = "Implementation Roadmap:\n\n"
        roadmap_text += "Phase 1: Data Validation (Months 1-2)\n"
        roadmap_text += "• Confirm fleet data accuracy and operational patterns\n"
        roadmap_text += "• Refine vehicle selection criteria\n"
        roadmap_text += "• Validate electrification candidates\n\n"
        
        roadmap_text += "Phase 2: Pilot Planning (Months 2-3)\n"
        roadmap_text += "• Select 10-20 pilot vehicles\n"
        roadmap_text += "• Conduct detailed site assessments\n"
        roadmap_text += "• Coordinate with utility providers\n\n"
        
        roadmap_text += "Phase 3: Pilot Implementation (Months 4-9)\n"
        roadmap_text += "• Install charging infrastructure\n"
        roadmap_text += "• Deploy pilot electric vehicles\n"
        roadmap_text += "• Monitor performance and collect data\n\n"
        
        roadmap_text += "Phase 4: Fleet Rollout (Months 10+)\n"
        roadmap_text += "• Scale successful pilot models\n"
        roadmap_text += "• Implement full fleet conversion plan\n"
        roadmap_text += "• Ongoing optimization and reporting\n\n"
        
        roadmap_text += "Immediate Next Steps (30 days):\n"
        roadmap_text += "• Stakeholder alignment meeting\n"
        roadmap_text += "• Pilot program budget development\n"
        roadmap_text += "• Utility preliminary discussions"
        
        roadmap_frame.text = roadmap_text
        _apply_brand_formatting(roadmap_frame, 11, PRIMARY_HEX_1)
        
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
