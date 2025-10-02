"""
powerpoint_charts.py

Native PowerPoint chart generation for the Fleet Electrification Analyzer.
Creates editable charts directly in PowerPoint instead of static images.
"""

import logging
from typing import Dict, List, Any, Optional, Union, Tuple
from datetime import datetime, timedelta

try:
    from pptx.chart.data import CategoryChartData, ChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

from data.models import FleetVehicle, Fleet
from settings import PRIMARY_HEX_1, PRIMARY_HEX_3, SECONDARY_HEX_1

logger = logging.getLogger(__name__)

###############################################################################
# Native PowerPoint Chart Functions
###############################################################################

def add_fleet_composition_chart(slide, vehicles: List[FleetVehicle], 
                               left: float = 1, top: float = 2, 
                               width: float = 8, height: float = 4) -> bool:
    """Add native PowerPoint pie chart for fleet composition by body class."""
    try:
        if not vehicles:
            return False
        
        # Count body classes
        body_counts = {}
        for vehicle in vehicles:
            body_class = getattr(vehicle.vehicle_id, 'body_class', 'Unknown') or 'Unknown'
            body_counts[body_class] = body_counts.get(body_class, 0) + 1
        
        # Sort and limit to top categories
        sorted_counts = sorted(body_counts.items(), key=lambda x: x[1], reverse=True)
        if len(sorted_counts) > 8:
            top_counts = sorted_counts[:7]
            other_count = sum(count for _, count in sorted_counts[7:])
            sorted_counts = top_counts + [("Other", other_count)]
        
        # Create chart data
        chart_data = CategoryChartData()
        chart_data.categories = [item[0] for item in sorted_counts]
        chart_data.add_series('Vehicle Count', [item[1] for item in sorted_counts])
        
        # Add chart to slide
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.PIE, 
            Inches(left), Inches(top), 
            Inches(width), Inches(height), 
            chart_data
        ).chart
        
        # Format chart with improved styling
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.font.size = Pt(10)
        
        # Format data labels
        chart.plots[0].has_data_labels = True
        data_labels = chart.plots[0].data_labels
        data_labels.font.size = Pt(9)
        data_labels.font.bold = True
        
        # Apply brand colors
        series = chart.series[0]
        for i, point in enumerate(series.points):
            colors = [PRIMARY_HEX_1, PRIMARY_HEX_3, SECONDARY_HEX_1, '#4ECDC4', '#95E1D3', '#F38BA8']
            color_hex = colors[i % len(colors)]
            r, g, b = tuple(int(color_hex.lstrip('#')[j:j+2], 16) for j in (0, 2, 4))
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = RGBColor(r, g, b)
        
        return True
        
    except Exception as e:
        logger.error(f"Failed to create fleet composition chart: {e}")
        return False

def add_emissions_timeline_chart(slide, vehicles: List[FleetVehicle],
                               left: float = 1, top: float = 2,
                               width: float = 8, height: float = 4) -> bool:
    """Add CO2 emissions reduction timeline chart showing reduction over years."""
    try:
        if not vehicles:
            return False
        
        # Calculate baseline emissions and projected reductions
        current_year = datetime.now().year
        years = list(range(current_year, current_year + 11))  # 10-year projection
        
        # Calculate current total emissions
        baseline_emissions = 0
        for vehicle in vehicles:
            annual_mileage = getattr(vehicle, 'annual_mileage', 12000)
            co2_per_mile = getattr(vehicle.fuel_economy, 'co2_primary', 400)  # g/mile
            if co2_per_mile > 0:
                # Convert to metric tons per year
                vehicle_emissions = (co2_per_mile * annual_mileage) / 1000000 * 1.1023
                baseline_emissions += vehicle_emissions
        
        # Project emissions reductions (assuming gradual electrification)
        emissions_data = []
        reduction_rates = [0, 0.05, 0.10, 0.18, 0.28, 0.40, 0.55, 0.68, 0.78, 0.85, 0.90]
        
        for i, year in enumerate(years):
            reduction_rate = reduction_rates[i] if i < len(reduction_rates) else 0.90
            remaining_emissions = baseline_emissions * (1 - reduction_rate)
            emissions_data.append(round(remaining_emissions, 1))
        
        # Create chart data
        chart_data = CategoryChartData()
        chart_data.categories = [str(year) for year in years]
        chart_data.add_series('Fleet CO₂ Emissions (MT)', emissions_data)
        
        # Add chart to slide
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.LINE,
            Inches(left), Inches(top),
            Inches(width), Inches(height),
            chart_data
        ).chart
        
        # Format chart with improved styling
        chart.has_legend = False
        chart.chart_title.text_frame.text = "Fleet CO₂ Emissions Reduction Timeline"
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(14)
        chart.chart_title.text_frame.paragraphs[0].font.bold = True
        
        # Format axes
        chart.value_axis.axis_title.text_frame.text = "CO₂ Emissions (Metric Tons)"
        chart.value_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(11)
        chart.category_axis.axis_title.text_frame.text = "Year"
        chart.category_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(11)
        
        # Format series
        series = chart.series[0]
        series.format.line.color.rgb = RGBColor(*tuple(int(PRIMARY_HEX_1.lstrip('#')[j:j+2], 16) for j in (0, 2, 4)))
        series.format.line.width = Pt(3)
        
        return True
        
    except Exception as e:
        logger.error(f"Failed to create emissions timeline chart: {e}")
        return False

def add_emissions_by_weight_class_chart(slide, vehicles: List[FleetVehicle],
                                       left: float = 1, top: float = 2,
                                       width: float = 8, height: float = 4) -> bool:
    """Add pie chart showing CO2 emissions by vehicle weight class."""
    try:
        if not vehicles:
            return False
        
        # Calculate emissions by weight class
        weight_class_emissions = {
            'Light Duty (≤8,500 lbs)': 0,
            'Medium Duty (8,501-19,500 lbs)': 0,
            'Heavy Duty (19,501-33,000 lbs)': 0,
            'Extra Heavy Duty (>33,000 lbs)': 0
        }
        
        for vehicle in vehicles:
            gvwr = getattr(vehicle.vehicle_id, 'gvwr_pounds', 0)
            annual_mileage = getattr(vehicle, 'annual_mileage', 12000)
            co2_per_mile = getattr(vehicle.fuel_economy, 'co2_primary', 400)
            
            if co2_per_mile > 0:
                # Convert to metric tons per year
                vehicle_emissions = (co2_per_mile * annual_mileage) / 1000000 * 1.1023
                
                # Classify by weight
                if gvwr <= 8500:
                    weight_class_emissions['Light Duty (≤8,500 lbs)'] += vehicle_emissions
                elif gvwr <= 19500:
                    weight_class_emissions['Medium Duty (8,501-19,500 lbs)'] += vehicle_emissions
                elif gvwr <= 33000:
                    weight_class_emissions['Heavy Duty (19,501-33,000 lbs)'] += vehicle_emissions
                else:
                    weight_class_emissions['Extra Heavy Duty (>33,000 lbs)'] += vehicle_emissions
        
        # Filter out zero emissions classes
        filtered_emissions = {k: round(v, 1) for k, v in weight_class_emissions.items() if v > 0}
        
        if not filtered_emissions:
            return False
        
        # Create chart data
        chart_data = CategoryChartData()
        chart_data.categories = list(filtered_emissions.keys())
        chart_data.add_series('CO₂ Emissions (MT/year)', list(filtered_emissions.values()))
        
        # Add chart to slide
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.PIE,
            Inches(left), Inches(top),
            Inches(width), Inches(height),
            chart_data
        ).chart
        
        # Format chart with improved styling
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.font.size = Pt(10)
        chart.chart_title.text_frame.text = "CO₂ Emissions by Vehicle Weight Class"
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(14)
        chart.chart_title.text_frame.paragraphs[0].font.bold = True
        
        # Format data labels
        chart.plots[0].has_data_labels = True
        data_labels = chart.plots[0].data_labels
        data_labels.font.size = Pt(9)
        data_labels.font.bold = True
        
        # Apply brand colors to pie slices
        series = chart.series[0]
        colors = [PRIMARY_HEX_1, PRIMARY_HEX_3, SECONDARY_HEX_1, '#4ECDC4']
        for i, point in enumerate(series.points):
            color_hex = colors[i % len(colors)]
            r, g, b = tuple(int(color_hex.lstrip('#')[j:j+2], 16) for j in (0, 2, 4))
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = RGBColor(r, g, b)
        
        return True
        
    except Exception as e:
        logger.error(f"Failed to create emissions by weight class chart: {e}")
        return False

def add_electrification_timeline_by_weight_chart(slide, vehicles: List[FleetVehicle],
                                               left: float = 1, top: float = 2,
                                               width: float = 8, height: float = 4) -> bool:
    """Add stacked bar chart showing electrification timeline by weight class."""
    try:
        if not vehicles:
            return False
        
        # Count vehicles by weight class
        weight_classes = {
            'Light Duty': sum(1 for v in vehicles if getattr(v.vehicle_id, 'gvwr_pounds', 0) <= 8500),
            'Medium Duty': sum(1 for v in vehicles if 8500 < getattr(v.vehicle_id, 'gvwr_pounds', 0) <= 19500),
            'Heavy Duty': sum(1 for v in vehicles if getattr(v.vehicle_id, 'gvwr_pounds', 0) > 19500)
        }
        
        # Create electrification timeline (10 years)
        current_year = datetime.now().year
        years = [str(year) for year in range(current_year + 1, current_year + 11)]
        
        # Electrification schedule by weight class (cumulative percentages)
        schedules = {
            'Light Duty': [0.1, 0.25, 0.45, 0.65, 0.80, 0.90, 0.95, 0.98, 0.99, 1.0],
            'Medium Duty': [0.05, 0.15, 0.30, 0.50, 0.70, 0.85, 0.95, 0.98, 0.99, 1.0],
            'Heavy Duty': [0.02, 0.08, 0.18, 0.32, 0.50, 0.68, 0.82, 0.92, 0.97, 1.0]
        }
        
        # Create chart data
        chart_data = CategoryChartData()
        chart_data.categories = years
        
        for weight_class, total_vehicles in weight_classes.items():
            if total_vehicles > 0:
                schedule = schedules[weight_class]
                yearly_counts = []
                prev_cumulative = 0
                
                for pct in schedule:
                    cumulative_vehicles = int(total_vehicles * pct)
                    yearly_new = cumulative_vehicles - prev_cumulative
                    yearly_counts.append(yearly_new)
                    prev_cumulative = cumulative_vehicles
                
                chart_data.add_series(weight_class, yearly_counts)
        
        # Add chart to slide
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_STACKED,
            Inches(left), Inches(top),
            Inches(width), Inches(height),
            chart_data
        ).chart
        
        # Format chart with improved styling
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(10)
        chart.chart_title.text_frame.text = "Electrification Timeline by Weight Class"
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(14)
        chart.chart_title.text_frame.paragraphs[0].font.bold = True
        
        # Format axes
        chart.value_axis.axis_title.text_frame.text = "Vehicles Electrified"
        chart.value_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(11)
        chart.category_axis.axis_title.text_frame.text = "Year"
        chart.category_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(11)
        
        # Apply brand colors to series
        colors = [PRIMARY_HEX_1, PRIMARY_HEX_3, SECONDARY_HEX_1]
        for i, series in enumerate(chart.series):
            color_hex = colors[i % len(colors)]
            r, g, b = tuple(int(color_hex.lstrip('#')[j:j+2], 16) for j in (0, 2, 4))
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = RGBColor(r, g, b)
        
        return True
        
    except Exception as e:
        logger.error(f"Failed to create electrification timeline by weight chart: {e}")
        return False

def add_electrification_timeline_by_body_type_chart(slide, vehicles: List[FleetVehicle],
                                                  left: float = 1, top: float = 2,
                                                  width: float = 8, height: float = 4) -> bool:
    """Add stacked bar chart showing electrification timeline by body type."""
    try:
        if not vehicles:
            return False
        
        # Count vehicles by body type (top 5 most common)
        body_counts = {}
        for vehicle in vehicles:
            body_type = getattr(vehicle.vehicle_id, 'body_class', 'Unknown') or 'Unknown'
            body_counts[body_type] = body_counts.get(body_type, 0) + 1
        
        # Get top 5 body types
        top_body_types = dict(sorted(body_counts.items(), key=lambda x: x[1], reverse=True)[:5])
        
        if not top_body_types:
            return False
        
        # Create electrification timeline (10 years)
        current_year = datetime.now().year
        years = [str(year) for year in range(current_year + 1, current_year + 11)]
        
        # Electrification schedule by body type (based on typical fleet priorities)
        body_type_schedules = {
            'Sedan': [0.15, 0.35, 0.55, 0.75, 0.90, 0.98, 1.0, 1.0, 1.0, 1.0],
            'SUV': [0.12, 0.30, 0.50, 0.70, 0.85, 0.95, 1.0, 1.0, 1.0, 1.0],
            'Pickup': [0.08, 0.20, 0.35, 0.55, 0.75, 0.90, 0.98, 1.0, 1.0, 1.0],
            'Van': [0.05, 0.15, 0.30, 0.50, 0.70, 0.85, 0.95, 1.0, 1.0, 1.0],
            'Truck': [0.03, 0.10, 0.22, 0.38, 0.58, 0.75, 0.88, 0.96, 1.0, 1.0]
        }
        
        # Default schedule for unknown body types
        default_schedule = [0.06, 0.18, 0.32, 0.50, 0.68, 0.82, 0.92, 0.98, 1.0, 1.0]
        
        # Create chart data
        chart_data = CategoryChartData()
        chart_data.categories = years
        
        for body_type, total_vehicles in top_body_types.items():
            if total_vehicles > 0:
                # Find appropriate schedule
                schedule = None
                for key, sched in body_type_schedules.items():
                    if key.lower() in body_type.lower():
                        schedule = sched
                        break
                
                if schedule is None:
                    schedule = default_schedule
                
                yearly_counts = []
                prev_cumulative = 0
                
                for pct in schedule:
                    cumulative_vehicles = int(total_vehicles * pct)
                    yearly_new = cumulative_vehicles - prev_cumulative
                    yearly_counts.append(yearly_new)
                    prev_cumulative = cumulative_vehicles
                
                chart_data.add_series(body_type, yearly_counts)
        
        # Add chart to slide
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_STACKED,
            Inches(left), Inches(top),
            Inches(width), Inches(height),
            chart_data
        ).chart
        
        # Format chart with improved styling
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(9)  # Smaller for more body types
        chart.chart_title.text_frame.text = "Electrification Timeline by Body Type"
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(14)
        chart.chart_title.text_frame.paragraphs[0].font.bold = True
        
        # Format axes
        chart.value_axis.axis_title.text_frame.text = "Vehicles Electrified"
        chart.value_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(11)
        chart.category_axis.axis_title.text_frame.text = "Year"
        chart.category_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(11)
        
        # Apply diverse colors for body types
        colors = [PRIMARY_HEX_1, PRIMARY_HEX_3, SECONDARY_HEX_1, '#4ECDC4', '#95E1D3', '#F38BA8', '#A8E6CF', '#FFD93D']
        for i, series in enumerate(chart.series):
            color_hex = colors[i % len(colors)]
            r, g, b = tuple(int(color_hex.lstrip('#')[j:j+2], 16) for j in (0, 2, 4))
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = RGBColor(r, g, b)
        
        return True
        
    except Exception as e:
        logger.error(f"Failed to create electrification timeline by body type chart: {e}")
        return False

def add_age_distribution_chart(slide, vehicles: List[FleetVehicle],
                              left: float = 1, top: float = 2,
                              width: float = 8, height: float = 4) -> bool:
    """Add column chart showing fleet age distribution."""
    try:
        if not vehicles:
            return False
        
        # Calculate ages and group into bins
        current_year = datetime.now().year
        age_bins = {}
        
        for vehicle in vehicles:
            try:
                year = int(getattr(vehicle.vehicle_id, 'year', 0))
                if year > 0:
                    age = current_year - year
                    age_group = f"{age} years"
                    age_bins[age_group] = age_bins.get(age_group, 0) + 1
            except (ValueError, TypeError):
                continue
        
        if not age_bins:
            return False
        
        # Sort by age
        sorted_ages = sorted(age_bins.items(), key=lambda x: int(x[0].split()[0]))
        
        # Create chart data
        chart_data = CategoryChartData()
        chart_data.categories = [item[0] for item in sorted_ages]
        chart_data.add_series('Vehicle Count', [item[1] for item in sorted_ages])
        
        # Add chart to slide
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(left), Inches(top),
            Inches(width), Inches(height),
            chart_data
        ).chart
        
        # Format chart with improved styling
        chart.has_legend = False
        chart.chart_title.text_frame.text = "Fleet Age Distribution"
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(14)
        chart.chart_title.text_frame.paragraphs[0].font.bold = True
        
        # Format axes
        chart.value_axis.axis_title.text_frame.text = "Number of Vehicles"
        chart.value_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(11)
        chart.category_axis.axis_title.text_frame.text = "Vehicle Age"
        chart.category_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(11)
        
        # Format series with brand color
        series = chart.series[0]
        r, g, b = tuple(int(PRIMARY_HEX_3.lstrip('#')[j:j+2], 16) for j in (0, 2, 4))
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = RGBColor(r, g, b)
        
        return True
        
    except Exception as e:
        logger.error(f"Failed to create age distribution chart: {e}")
        return False

###############################################################################
# Chart Configuration and Customization
###############################################################################

class SlideConfiguration:
    """Configuration class for customizable slide generation."""
    
    def __init__(self):
        self.available_slides = {
            'cover': {
                'name': 'Cover Slide',
                'description': 'Fleet name, client info, and date',
                'required': True,
                'charts': []
            },
            'fleet_snapshot': {
                'name': 'Fleet Snapshot KPIs',
                'description': '6 key metrics with baseline costs and emissions',
                'required': True,
                'charts': []
            },
            'fleet_composition': {
                'name': 'Fleet Composition',
                'description': 'Vehicle composition by body class',
                'required': False,
                'charts': ['pie_body_class']
            },
            'emissions_timeline': {
                'name': 'Emissions Reduction Timeline',
                'description': 'CO₂ emissions reduction over time',
                'required': False,
                'charts': ['line_emissions_timeline']
            },
            'emissions_by_weight': {
                'name': 'Emissions by Weight Class',
                'description': 'CO₂ emissions breakdown by vehicle weight',
                'required': False,
                'charts': ['pie_emissions_weight']
            },
            'electrification_timeline_weight': {
                'name': 'Electrification Timeline by Weight',
                'description': 'Electrification schedule by weight class',
                'required': False,
                'charts': ['stacked_bar_weight_timeline']
            },
            'electrification_timeline_body': {
                'name': 'Electrification Timeline by Body Type',
                'description': 'Electrification schedule by body type',
                'required': False,
                'charts': ['stacked_bar_body_timeline']
            },
            'age_analysis': {
                'name': 'Fleet Age Analysis',
                'description': 'Age distribution and statistics',
                'required': False,
                'charts': ['column_age_distribution']
            },
            'data_quality': {
                'name': 'Data Quality & Completeness',
                'description': 'Data completeness analysis',
                'required': False,
                'charts': []
            },
            'next_steps': {
                'name': 'Next Steps & Roadmap',
                'description': 'Implementation timeline and recommendations',
                'required': False,
                'charts': []
            }
        }
        
        self.available_charts = {
            'pie_body_class': {
                'name': 'Fleet Composition (Pie)',
                'description': 'Pie chart of vehicles by body class',
                'function': add_fleet_composition_chart
            },
            'line_emissions_timeline': {
                'name': 'Emissions Timeline (Line)',
                'description': 'Line chart showing CO₂ reduction over time',
                'function': add_emissions_timeline_chart
            },
            'pie_emissions_weight': {
                'name': 'Emissions by Weight (Pie)',
                'description': 'Pie chart of emissions by weight class',
                'function': add_emissions_by_weight_class_chart
            },
            'stacked_bar_weight_timeline': {
                'name': 'Electrification by Weight (Stacked Bar)',
                'description': 'Stacked bar chart of electrification timeline by weight',
                'function': add_electrification_timeline_by_weight_chart
            },
            'stacked_bar_body_timeline': {
                'name': 'Electrification by Body Type (Stacked Bar)',
                'description': 'Stacked bar chart of electrification timeline by body type',
                'function': add_electrification_timeline_by_body_type_chart
            },
            'column_age_distribution': {
                'name': 'Age Distribution (Column)',
                'description': 'Column chart showing fleet age distribution',
                'function': add_age_distribution_chart
            }
        }
        
        # Default configuration
        self.selected_slides = [
            'cover',
            'fleet_snapshot', 
            'fleet_composition',
            'emissions_timeline',
            'emissions_by_weight',
            'electrification_timeline_weight',
            'age_analysis',
            'next_steps'
        ]
    
    def get_slide_options(self) -> Dict[str, Dict[str, Any]]:
        """Get available slide options for user selection."""
        return self.available_slides
    
    def get_chart_options(self) -> Dict[str, Dict[str, Any]]:
        """Get available chart options for user selection."""
        return self.available_charts
    
    def set_selected_slides(self, slide_ids: List[str]) -> bool:
        """Set which slides to include in the presentation."""
        # Validate slide IDs
        invalid_slides = [sid for sid in slide_ids if sid not in self.available_slides]
        if invalid_slides:
            logger.error(f"Invalid slide IDs: {invalid_slides}")
            return False
        
        # Ensure required slides are included
        required_slides = [sid for sid, info in self.available_slides.items() if info['required']]
        final_slides = list(set(slide_ids + required_slides))
        
        self.selected_slides = final_slides
        return True
    
    def get_selected_slides(self) -> List[str]:
        """Get the currently selected slides."""
        return self.selected_slides
