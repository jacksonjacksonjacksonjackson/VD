"""
charts.py

Chart generation and visualization for the Fleet Electrification Analyzer.
Provides functions to create various charts for fleet analysis and reports.
"""

import logging
import numpy as np
from typing import Dict, List, Any, Optional, Union, Tuple, Set

import matplotlib
matplotlib.use("TkAgg")  # Use TkAgg backend for tkinter compatibility
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
import matplotlib.patches as mpatches
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.colors as mcolors

from settings import (
    PRIMARY_HEX_1, 
    PRIMARY_HEX_2, 
    PRIMARY_HEX_3, 
    SECONDARY_HEX_1, 
    SECONDARY_HEX_2
)
from data.models import FleetVehicle, Fleet, ElectrificationAnalysis, ChargingAnalysis, EmissionsInventory

# Set up module logger
logger = logging.getLogger(__name__)

# Configure matplotlib style
plt.style.use('seaborn-v0_8-whitegrid')
matplotlib.rcParams.update({
    # Font settings
    'font.family': 'sans-serif',
    'font.sans-serif': ['Arial', 'Helvetica', 'DejaVu Sans'],
    'font.size': 10,
    
    # Figure settings
    'figure.figsize': (10, 6),
    'figure.dpi': 100,
    'figure.autolayout': True,
    'figure.facecolor': 'white',
    
    # Axes settings
    'axes.titlesize': 14,
    'axes.labelsize': 12,
    'axes.labelweight': 'bold',
    'axes.spines.top': False,
    'axes.spines.right': False,
    'axes.grid': True,
    'grid.alpha': 0.3,
    'axes.facecolor': 'white',
    
    # Tick settings
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
    'xtick.major.size': 4,
    'ytick.major.size': 4,
    'xtick.minor.size': 2,
    'ytick.minor.size': 2,
    
    # Legend settings
    'legend.fontsize': 10,
    'legend.frameon': False,
    'legend.loc': 'best',
    
    # Line settings
    'lines.linewidth': 2,
    'lines.markersize': 8,
    
    # Save settings
    'savefig.dpi': 300,
    'savefig.bbox': 'tight',
    'savefig.pad_inches': 0.1
})

# Professional color palette inspired by McKinsey
COLORS = [
    '#006BA2',  # McKinsey blue
    '#FF8F1C',  # Orange
    '#598C7E',  # Teal
    '#B55A30',  # Rust
    '#4B4B4B',  # Dark gray
    '#8DC63F',  # Green
    '#2C5985',  # Navy
    '#F2BE1A',  # Yellow
    '#AA1F73',  # Purple
    '#009A9B',  # Turquoise
]

# Create professional color maps
SEQUENTIAL_COLORS = ['#DEEBF7', '#6BAED6', '#2171B5']  # Blue sequential
DIVERGING_COLORS = ['#D73027', '#FFFFBF', '#1A9850']   # Red to Green

###############################################################################
# Chart Factory
###############################################################################

class ChartFactory:
    """Factory class for creating different types of charts."""
    
    @staticmethod
    def create_chart(chart_type: str, data: Any, figure: Optional[Figure] = None, 
                   **kwargs) -> Figure:
        """
        Create a chart based on chart type and data.
        
        Args:
            chart_type: Type of chart to create
            data: Data for the chart (fleet, vehicles list, or analysis object)
            figure: Existing figure to use (or None to create new)
            **kwargs: Additional arguments for specific chart types
            
        Returns:
            Matplotlib Figure object
        """
        # Create figure if not provided
        if figure is None:
            figure = plt.figure(figsize=(8, 6), dpi=100)
            figure.tight_layout(pad=3.0)
        else:
            figure.clear()
        
        # Create the requested chart
        if chart_type == "Body Class Distribution":
            FleetCharts.body_class_distribution(data, figure)
        
        elif chart_type == "MPG Distribution":
            FleetCharts.mpg_distribution(data, figure)
        
        elif chart_type == "CO2 Emissions Distribution":
            FleetCharts.co2_distribution(data, figure)
        
        elif chart_type == "CO2 vs MPG Correlation":
            FleetCharts.co2_vs_mpg(data, figure)
        
        elif chart_type == "CO2 Comparison (Primary vs Alt)":
            FleetCharts.co2_primary_vs_alt(data, figure)
        
        elif chart_type == "EV Range Distribution":
            FleetCharts.ev_range_distribution(data, figure)
        
        elif chart_type == "Make Frequency":
            FleetCharts.make_frequency(data, figure)
        
        elif chart_type == "Model Distribution":
            FleetCharts.model_distribution(data, figure)
        
        elif chart_type == "Fuel Type Distribution":
            FleetCharts.fuel_type_distribution(data, figure)
        
        elif chart_type == "Annual Cost Comparison":
            ElectrificationCharts.annual_cost_comparison(data, figure, **kwargs)
        
        elif chart_type == "Fleet Age Distribution":
            FleetCharts.age_distribution(data, figure)
        
        elif chart_type == "Electrification Potential":
            ElectrificationCharts.electrification_potential(data, figure, **kwargs)
        
        elif chart_type == "Emissions Reduction":
            ElectrificationCharts.emissions_reduction(data, figure, **kwargs)
        
        elif chart_type == "ROI Analysis":
            ElectrificationCharts.roi_analysis(data, figure, **kwargs)
        
        elif chart_type == "Charging Infrastructure":
            ElectrificationCharts.charging_infrastructure(data, figure, **kwargs)
        
        elif chart_type == "Emissions Inventory":
            EmissionsCharts.emissions_inventory(data, figure)
        
        elif chart_type == "Emissions Trends":
            EmissionsCharts.emissions_trends(data, figure)
        
        elif chart_type == "Emissions by Department":
            EmissionsCharts.emissions_by_department(data, figure)
        
        elif chart_type == "Emissions by Vehicle Type":
            EmissionsCharts.emissions_by_vehicle_type(data, figure)
        
        else:
            # Create empty chart with error message
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, f"Unknown chart type: {chart_type}", 
                   ha='center', va='center', fontsize=12)
            ax.axis('off')
        
        return figure


###############################################################################
# Fleet Charts
###############################################################################

class FleetCharts:
    """Charts for visualizing fleet composition and metrics."""
    
    @staticmethod
    def body_class_distribution(fleet_data: Union[Fleet, List[FleetVehicle]], figure: Figure) -> None:
        """
        Create a horizontal bar chart of vehicle body classes.
        
        Args:
            fleet_data: Fleet object or list of vehicles
            figure: Figure to draw on
        """
        # Extract vehicles from fleet if needed
        vehicles = fleet_data.vehicles if isinstance(fleet_data, Fleet) else fleet_data
        
        # Count body classes
        body_classes = {}
        for vehicle in vehicles:
            body_class = vehicle.vehicle_id.body_class
            if body_class:
                body_classes[body_class] = body_classes.get(body_class, 0) + 1
        
        # Sort by count and get top 10
        sorted_classes = sorted(body_classes.items(), key=lambda x: x[1], reverse=True)
        if len(sorted_classes) > 10:
            sorted_classes = sorted_classes[:10]
            other_count = sum(count for _, count in sorted_classes[10:])
            sorted_classes.append(("Other", other_count))
        
        # Extract data for chart
        class_names, counts = zip(*sorted_classes)
        
        # Create chart
        ax = figure.add_subplot(111)
        bars = ax.barh(class_names, counts, color=PRIMARY_HEX_3)
        
        # Add count labels to bars
        for bar in bars:
            width = bar.get_width()
            ax.text(width + 0.5, bar.get_y() + bar.get_height()/2, 
                   f"{int(width)}", ha='left', va='center')
        
        # Configure chart
        ax.set_title("Vehicle Body Class Distribution")
        ax.set_xlabel("Count")
        ax.set_ylabel("Body Class")
        
        # Tight layout
        figure.tight_layout()
    
    @staticmethod
    def mpg_distribution(fleet_data: Union[Fleet, List[FleetVehicle]], figure: Figure) -> None:
        """
        Create a histogram of MPG distribution.
        
        Args:
            fleet_data: Fleet object or list of vehicles
            figure: Figure to draw on
        """
        # Extract vehicles from fleet if needed
        vehicles = fleet_data.vehicles if isinstance(fleet_data, Fleet) else fleet_data
        
        # Extract MPG values
        mpg_values = []
        for vehicle in vehicles:
            mpg = vehicle.fuel_economy.combined_mpg
            if mpg and mpg > 0:
                mpg_values.append(mpg)
        
        # Validate data
        if not mpg_values:
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, "No MPG data available", ha='center', va='center')
            ax.axis('off')
            return
        
        # Create chart
        ax = figure.add_subplot(111)
        
        # Determine bins based on data range
        max_mpg = max(mpg_values)
        if max_mpg <= 30:
            bins = np.arange(0, max_mpg + 5, 5)
        elif max_mpg <= 60:
            bins = np.arange(0, max_mpg + 10, 10)
        else:
            bins = 10
        
        # Create histogram
        n, bins, patches = ax.hist(mpg_values, bins=bins, edgecolor='white', alpha=0.7, color=PRIMARY_HEX_3)
        
        # Configure chart
        ax.set_title("Fuel Economy Distribution")
        ax.set_xlabel("MPG (Combined)")
        ax.set_ylabel("Number of Vehicles")
        
        # Add mean line
        mean_mpg = np.mean(mpg_values)
        ax.axvline(mean_mpg, color=SECONDARY_HEX_1, linestyle='dashed', linewidth=2, 
                  label=f'Mean: {mean_mpg:.1f} MPG')
        
        # Add legend
        ax.legend()
        
        # Tight layout
        figure.tight_layout()
    
    @staticmethod
    def co2_distribution(fleet_data: Union[Fleet, List[FleetVehicle]], figure: Figure) -> None:
        """
        Create a histogram of CO2 emissions distribution.
        
        Args:
            fleet_data: Fleet object or list of vehicles
            figure: Figure to draw on
        """
        # Extract vehicles from fleet if needed
        vehicles = fleet_data.vehicles if isinstance(fleet_data, Fleet) else fleet_data
        
        # Extract CO2 values
        co2_values = []
        for vehicle in vehicles:
            co2 = vehicle.fuel_economy.co2_primary
            if co2 and co2 > 0:
                co2_values.append(co2)
        
        # Validate data
        if not co2_values:
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, "No CO2 data available", ha='center', va='center')
            ax.axis('off')
            return
        
        # Create chart
        ax = figure.add_subplot(111)
        
        # Create histogram
        n, bins, patches = ax.hist(co2_values, bins=10, edgecolor='white', alpha=0.7, color=SECONDARY_HEX_1)
        
        # Configure chart
        ax.set_title("CO2 Emissions Distribution")
        ax.set_xlabel("CO2 Emissions (g/mile)")
        ax.set_ylabel("Number of Vehicles")
        
        # Add mean line
        mean_co2 = np.mean(co2_values)
        ax.axvline(mean_co2, color='red', linestyle='dashed', linewidth=2, 
                  label=f'Mean: {mean_co2:.1f} g/mile')
        
        # Add legend
        ax.legend()
        
        # Tight layout
        figure.tight_layout()
    
    @staticmethod
    def co2_vs_mpg(fleet_data: Union[Fleet, List[FleetVehicle]], figure: Figure) -> None:
        """
        Create a scatter plot of CO2 vs MPG.
        
        Args:
            fleet_data: Fleet object or list of vehicles
            figure: Figure to draw on
        """
        # Extract vehicles from fleet if needed
        vehicles = fleet_data.vehicles if isinstance(fleet_data, Fleet) else fleet_data
        
        # Extract data points
        mpg_values = []
        co2_values = []
        for vehicle in vehicles:
            mpg = vehicle.fuel_economy.combined_mpg
            co2 = vehicle.fuel_economy.co2_primary
            if mpg and mpg > 0 and co2 and co2 > 0:
                mpg_values.append(mpg)
                co2_values.append(co2)
        
        # Validate data
        if not mpg_values or not co2_values:
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, "Insufficient data for CO2 vs MPG comparison", 
                   ha='center', va='center')
            ax.axis('off')
            return
        
        # Create chart
        ax = figure.add_subplot(111)
        
        # Create scatter plot
        scatter = ax.scatter(mpg_values, co2_values, c='red', alpha=0.6, edgecolors='none')
        
        # Add best fit line
        if len(mpg_values) > 1:
            try:
                z = np.polyfit(mpg_values, co2_values, 1)
                p = np.poly1d(z)
                x_range = np.linspace(min(mpg_values), max(mpg_values), 100)
                ax.plot(x_range, p(x_range), linestyle='--', color='blue')
            except Exception as e:
                logger.warning(f"Error calculating trend line: {e}")
        
        # Configure chart
        ax.set_title("CO2 Emissions vs. Fuel Economy")
        ax.set_xlabel("MPG (Combined)")
        ax.set_ylabel("CO2 Emissions (g/mile)")
        
        # Add grid
        ax.grid(True, linestyle='--', alpha=0.7)
        
        # Tight layout
        figure.tight_layout()
    
    @staticmethod
    def co2_primary_vs_alt(fleet_data: Union[Fleet, List[FleetVehicle]], figure: Figure) -> None:
        """
        Create a scatter plot comparing primary and alternative fuel CO2 emissions.
        
        Args:
            fleet_data: Fleet object or list of vehicles
            figure: Figure to draw on
        """
        # Extract vehicles from fleet if needed
        vehicles = fleet_data.vehicles if isinstance(fleet_data, Fleet) else fleet_data
        
        # Extract data points (only for vehicles with both values)
        co2_primary = []
        co2_alt = []
        makes = []
        
        for vehicle in vehicles:
            primary = vehicle.fuel_economy.co2_primary
            alt = vehicle.fuel_economy.co2_alt
            if primary and primary > 0 and alt and alt > 0:
                co2_primary.append(primary)
                co2_alt.append(alt)
                makes.append(vehicle.vehicle_id.make)
        
        # Validate data
        if not co2_primary or not co2_alt:
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, "No alternative fuel CO2 data available", 
                   ha='center', va='center')
            ax.axis('off')
            return
        
        # Create chart
        ax = figure.add_subplot(111)
        
        # Create scatter plot with colors by make
        unique_makes = list(set(makes))
        colors = plt.cm.tab10(np.linspace(0, 1, len(unique_makes)))
        
        for i, make in enumerate(unique_makes):
            indices = [j for j, m in enumerate(makes) if m == make]
            ax.scatter(
                [co2_primary[j] for j in indices],
                [co2_alt[j] for j in indices],
                alpha=0.7,
                c=[colors[i]],
                label=make,
                edgecolors='none'
            )
        
        # Add equity line
        max_val = max(max(co2_primary), max(co2_alt))
        ax.plot([0, max_val], [0, max_val], linestyle='--', color='gray', alpha=0.7)
        
        # Configure chart
        ax.set_title("Primary vs. Alternative Fuel CO2 Emissions")
        ax.set_xlabel("Primary Fuel CO2 (g/mile)")
        ax.set_ylabel("Alternative Fuel CO2 (g/mile)")
        
        # Add legend if not too many makes
        if len(unique_makes) <= 10:
            ax.legend(fontsize='small')
        
        # Tight layout
        figure.tight_layout()
    
    @staticmethod
    def ev_range_distribution(fleet_data: Union[Fleet, List[FleetVehicle]], figure: Figure) -> None:
        """
        Create a histogram of EV/Alternative fuel range distribution.
        
        Args:
            fleet_data: Fleet object or list of vehicles
            figure: Figure to draw on
        """
        # Extract vehicles from fleet if needed
        vehicles = fleet_data.vehicles if isinstance(fleet_data, Fleet) else fleet_data
        
        # Extract range values
        range_values = []
        
        for vehicle in vehicles:
            range_val = vehicle.fuel_economy.alt_range
            if range_val and range_val > 0:
                range_values.append(range_val)
        
        # Validate data
        if not range_values:
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, "No alternative fuel range data available", 
                   ha='center', va='center')
            ax.axis('off')
            return
        
        # Create chart
        ax = figure.add_subplot(111)
        
        # Create histogram
        n, bins, patches = ax.hist(range_values, bins=10, edgecolor='white', alpha=0.7, color=PRIMARY_HEX_3)
        
        # Configure chart
        ax.set_title("Alternative Fuel Range Distribution")
        ax.set_xlabel("Range (miles)")
        ax.set_ylabel("Number of Vehicles")
        
        # Add mean line
        mean_range = np.mean(range_values)
        ax.axvline(mean_range, color=SECONDARY_HEX_1, linestyle='dashed', linewidth=2, 
                  label=f'Mean: {mean_range:.1f} miles')
        
        # Add legend
        ax.legend()
        
        # Tight layout
        figure.tight_layout()
    
    @staticmethod
    def make_frequency(fleet_data: Union[Fleet, List[FleetVehicle]], figure: Figure) -> None:
        """
        Create a bar chart of vehicle makes.
        
        Args:
            fleet_data: Fleet object or list of vehicles
            figure: Figure to draw on
        """
        # Extract vehicles from fleet if needed
        vehicles = fleet_data.vehicles if isinstance(fleet_data, Fleet) else fleet_data
        
        # Count makes
        makes = {}
        for vehicle in vehicles:
            make = vehicle.vehicle_id.make
            if make:
                makes[make] = makes.get(make, 0) + 1
        
        # Sort by count and get top 10
        sorted_makes = sorted(makes.items(), key=lambda x: x[1], reverse=True)
        if len(sorted_makes) > 10:
            top_makes = sorted_makes[:10]
            other_count = sum(count for _, count in sorted_makes[10:])
            sorted_makes = top_makes + [("Other", other_count)]
        
        # Extract data for chart
        make_names, counts = zip(*sorted_makes)
        
        # Create chart
        ax = figure.add_subplot(111)
        bars = ax.bar(make_names, counts, color=COLORS[:len(make_names)])
        
        # Add count labels above bars
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, height + 0.1, 
                   f"{int(height)}", ha='center', va='bottom')
        
        # Configure chart
        ax.set_title("Vehicle Make Distribution")
        ax.set_xlabel("Make")
        ax.set_ylabel("Count")
        
        # Rotate x-axis labels for readability
        plt.xticks(rotation=45, ha='right')
        
        # Tight layout
        figure.tight_layout()
    
    @staticmethod
    def model_distribution(fleet_data: Union[Fleet, List[FleetVehicle]], figure: Figure) -> None:
        """
        Create a bar chart of top vehicle models.
        
        Args:
            fleet_data: Fleet object or list of vehicles
            figure: Figure to draw on
        """
        # Extract vehicles from fleet if needed
        vehicles = fleet_data.vehicles if isinstance(fleet_data, Fleet) else fleet_data
        
        # Count models
        models = {}
        for vehicle in vehicles:
            model = vehicle.vehicle_id.model
            make = vehicle.vehicle_id.make
            if model and make:
                # Combine make and model for better identification
                key = f"{make} {model}"
                models[key] = models.get(key, 0) + 1
        
        # Sort by count and get top 15
        sorted_models = sorted(models.items(), key=lambda x: x[1], reverse=True)
        if len(sorted_models) > 15:
            top_models = sorted_models[:15]
            other_count = sum(count for _, count in sorted_models[15:])
            sorted_models = top_models + [("Other", other_count)]
        
        # Extract data for chart
        model_names, counts = zip(*sorted_models)
        
        # Create chart
        ax = figure.add_subplot(111)
        bars = ax.bar(model_names, counts, color=COLORS[:len(model_names)])
        
        # Configure chart
        ax.set_title("Top Vehicle Models")
        ax.set_xlabel("Model")
        ax.set_ylabel("Count")
        
        # Rotate x-axis labels for readability
        plt.xticks(rotation=45, ha='right')
        
        # Tight layout
        figure.tight_layout()
    
    @staticmethod
    def fuel_type_distribution(fleet_data: Union[Fleet, List[FleetVehicle]], figure: Figure) -> None:
        """
        Create a pie chart of fuel types.
        
        Args:
            fleet_data: Fleet object or list of vehicles
            figure: Figure to draw on
        """
        # Extract vehicles from fleet if needed
        vehicles = fleet_data.vehicles if isinstance(fleet_data, Fleet) else fleet_data
        
        # Count fuel types
        fuel_types = {}
        for vehicle in vehicles:
            fuel_type = vehicle.vehicle_id.fuel_type
            if fuel_type:
                fuel_types[fuel_type] = fuel_types.get(fuel_type, 0) + 1
        
        # Sort by count
        sorted_types = sorted(fuel_types.items(), key=lambda x: x[1], reverse=True)
        
        # Extract data for chart
        labels, sizes = zip(*sorted_types)
        
        # Create chart
        ax = figure.add_subplot(111)
        
        # Create pie chart
        wedges, texts, autotexts = ax.pie(
            sizes, 
            labels=None, 
            autopct='%1.1f%%',
            startangle=90,
            colors=COLORS[:len(labels)]
        )
        
        # Configure chart
        ax.set_title("Fuel Type Distribution")
        ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle
        
        # Add legend
        ax.legend(wedges, labels, title="Fuel Types", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
        
        # Set text properties for better readability
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontsize(9)
        
        # Tight layout
        figure.tight_layout()
    
    @staticmethod
    def age_distribution(fleet_data: Union[Fleet, List[FleetVehicle]], figure: Figure) -> None:
        """
        Create a histogram of vehicle ages.
        
        Args:
            fleet_data: Fleet object or list of vehicles
            figure: Figure to draw on
        """
        # Extract vehicles from fleet if needed
        vehicles = fleet_data.vehicles if isinstance(fleet_data, Fleet) else fleet_data
        
        # Extract ages
        ages = []
        for vehicle in vehicles:
            age = vehicle.age
            if age > 0:
                ages.append(age)
        
        # Validate data
        if not ages:
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, "No age data available", ha='center', va='center')
            ax.axis('off')
            return
        
        # Create chart
        ax = figure.add_subplot(111)
        
        # Create histogram
        n, bins, patches = ax.hist(ages, bins=range(0, int(max(ages)) + 2), 
                                  edgecolor='white', alpha=0.7, color=PRIMARY_HEX_3)
        
        # Configure chart
        ax.set_title("Fleet Age Distribution")
        ax.set_xlabel("Age (years)")
        ax.set_ylabel("Number of Vehicles")
        
        # Add mean line
        mean_age = np.mean(ages)
        ax.axvline(mean_age, color=SECONDARY_HEX_1, linestyle='dashed', linewidth=2, 
                  label=f'Mean: {mean_age:.1f} years')
        
        # Add legend
        ax.legend()
        
        # Tight layout
        figure.tight_layout()


###############################################################################
# Electrification Charts
###############################################################################

class ElectrificationCharts:
    """Charts for visualizing fleet electrification analysis."""
    
    @staticmethod
    def annual_cost_comparison(fleet_data: Union[Fleet, List[FleetVehicle]], figure: Figure,
                             gas_price: float = 3.50, electricity_price: float = 0.13,
                             ev_efficiency: float = 0.30) -> None:
        """
        Create a bar chart comparing annual fuel costs for ICE vs EV.
        
        Args:
            fleet_data: Fleet object or list of vehicles
            figure: Figure to draw on
            gas_price: Price of gasoline in $/gallon
            electricity_price: Price of electricity in $/kWh
            ev_efficiency: EV energy usage in kWh/mile
        """
        # Extract vehicles from fleet if needed
        vehicles = fleet_data.vehicles if isinstance(fleet_data, Fleet) else fleet_data
        
        # Calculate costs for each vehicle
        ice_costs = []
        ev_costs = []
        labels = []
        
        for i, vehicle in enumerate(vehicles):
            # Skip vehicles without MPG data
            if not vehicle.fuel_economy.combined_mpg or vehicle.fuel_economy.combined_mpg <= 0:
                continue
            
            # Get annual mileage (use default if not available)
            annual_mileage = vehicle.annual_mileage or 12000
            
            # Calculate ICE cost
            mpg = vehicle.fuel_economy.combined_mpg
            ice_cost = (annual_mileage / mpg) * gas_price
            
            # Calculate EV cost
            ev_cost = annual_mileage * ev_efficiency * electricity_price
            
            # Add to data
            ice_costs.append(ice_cost)
            ev_costs.append(ev_cost)
            
            # Create label
            make = vehicle.vehicle_id.make
            model = vehicle.vehicle_id.model
            labels.append(f"{make} {model}")
            
            # Limit to top 10 vehicles
            if len(ice_costs) >= 10:
                break
        
        # Validate data
        if not ice_costs or not ev_costs:
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, "Insufficient data for cost comparison", 
                   ha='center', va='center')
            ax.axis('off')
            return
        
        # Create chart
        ax = figure.add_subplot(111)
        
        # Set up x-axis
        x = np.arange(len(labels))
        width = 0.35
        
        # Create grouped bars
        ax.bar(x - width/2, ice_costs, width, label='ICE', color=SECONDARY_HEX_1)
        ax.bar(x + width/2, ev_costs, width, label='EV', color=PRIMARY_HEX_3)
        
        # Configure chart
        ax.set_title("Annual Fuel Cost Comparison: ICE vs. Electric")
        ax.set_ylabel("Annual Cost ($)")
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45, ha='right')
        
        # Add legend
        ax.legend()
        
        # Add cost labels above bars
        for i, cost in enumerate(ice_costs):
            ax.text(i - width/2, cost + 20, f"${cost:.0f}", ha='center', va='bottom')
        
        for i, cost in enumerate(ev_costs):
            ax.text(i + width/2, cost + 20, f"${cost:.0f}", ha='center', va='bottom')
        
        # Add savings percentages
        for i in range(len(ice_costs)):
            savings_pct = 100 * (ice_costs[i] - ev_costs[i]) / ice_costs[i]
            ax.text(i, max(ice_costs[i], ev_costs[i]) + 100, 
                   f"{savings_pct:.0f}% savings", ha='center', va='bottom')
        
        # Tight layout
        figure.tight_layout()
    
    @staticmethod
    def electrification_potential(analysis: ElectrificationAnalysis, figure: Figure,
                                top_n: int = 10) -> None:
        """
        Create a professional chart showing electrification potential by vehicle type.
        
        Args:
            analysis: Electrification analysis results
            figure: Figure to draw on
            top_n: Number of top vehicle types to show
        """
        # Get data
        potential_by_type = analysis.potential_by_vehicle_type
        sorted_types = sorted(potential_by_type.items(), key=lambda x: x[1], reverse=True)
        
        if len(sorted_types) > top_n:
            other_sum = sum(score for _, score in sorted_types[top_n:])
            sorted_types = sorted_types[:top_n] + [("Other", other_sum)]
        
        # Extract data for chart
        types, scores = zip(*sorted_types)
        
        # Create chart
        gs = figure.add_gridspec(1, 2, width_ratios=[3, 1])
        
        # Main bar chart
        ax1 = figure.add_subplot(gs[0])
        bars = ax1.barh(types, scores, color=COLORS[0], alpha=0.8)
        
        # Add value labels on bars
        for bar in bars:
            width = bar.get_width()
            ax1.text(width, bar.get_y() + bar.get_height()/2,
                    f'{width:.1f}%', 
                    ha='left', va='center',
                    fontsize=9, fontweight='bold',
                    bbox=dict(facecolor='white', edgecolor='none', alpha=0.7, pad=2))
        
        # Configure main chart
        ax1.set_title('Electrification Potential by Vehicle Type', 
                     pad=20, fontsize=14, fontweight='bold')
        ax1.set_xlabel('Potential Score (%)', fontsize=12)
        ax1.grid(True, axis='x', alpha=0.3)
        ax1.set_axisbelow(True)
        
        # Summary statistics on the right
        ax2 = figure.add_subplot(gs[1])
        ax2.axis('off')
        
        # Add summary box
        summary_text = [
            'Summary Statistics',
            '─' * 20,
            f'Fleet Average: {analysis.fleet_average_potential:.1f}%',
            f'Highest: {max(scores):.1f}%',
            f'Lowest: {min(scores):.1f}%',
            f'Vehicles Analyzed: {len(analysis.vehicles)}',
            '\nRecommendations:',
            '• Focus on top 3 types',
            '• Consider pilot program',
            '• Review infrastructure'
        ]
        
        ax2.text(0.1, 0.95, '\n'.join(summary_text),
                 transform=ax2.transAxes,
                 bbox=dict(facecolor='#F5F5F5', edgecolor='#CCCCCC', 
                          alpha=0.5, pad=10, boxstyle='round'),
                 verticalalignment='top',
                 fontsize=10)
        
        # Adjust layout
        figure.set_tight_layout(True)
    
    @staticmethod
    def emissions_reduction(analysis: ElectrificationAnalysis, figure: Figure,
                         top_n: int = 10) -> None:
        """
        Create a bar chart showing CO2 emissions reduction potential.
        
        Args:
            analysis: ElectrificationAnalysis object
            figure: Figure to draw on
            top_n: Number of top vehicles to display
        """
        # Validate data
        if not analysis.vehicle_results:
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, "No emissions reduction data available", 
                   ha='center', va='center')
            ax.axis('off')
            return
        
        # Get top vehicles by CO2 reduction
        top_vehicles = sorted(
            analysis.vehicle_results.items(), 
            key=lambda x: x[1]["total_co2_reduction"], 
            reverse=True
        )[:top_n]
        
        # Extract data for chart
        vehicle_names = [f"{v['year']} {v['make']} {v['model']}" for _, v in top_vehicles]
        reductions = [v["total_co2_reduction"] for _, v in top_vehicles]
        
        # Create chart
        ax = figure.add_subplot(111)
        
        # Create horizontal bar chart
        bars = ax.barh(vehicle_names, reductions, color='green')
        
        # Add reduction labels
        for bar in bars:
            width = bar.get_width()
            ax.text(width + 0.5, bar.get_y() + bar.get_height()/2, 
                   f"{width:.1f} tons", ha='left', va='center')
        
        # Configure chart
        ax.set_title("CO2 Emissions Reduction Potential")
        ax.set_xlabel("Lifetime CO2 Reduction (metric tons)")
        
        # Tight layout
        figure.tight_layout()
    
    @staticmethod
    def roi_analysis(analysis: ElectrificationAnalysis, figure: Figure,
                   ev_premium: float = 15000) -> None:
        """
        Create an insightful ROI analysis visualization with quadrant analysis.
        
        Args:
            analysis: ElectrificationAnalysis object
            figure: Figure to draw on
            ev_premium: EV price premium over comparable ICE vehicle
        """
        # Create subplots with specific layout
        gs = figure.add_gridspec(2, 2, width_ratios=[3, 1], height_ratios=[3, 1])
        
        # Main scatter plot
        ax_main = figure.add_subplot(gs[0, 0])
        ax_summary = figure.add_subplot(gs[0, 1])
        ax_hist_x = figure.add_subplot(gs[1, 0])
        
        # Extract and validate data
        data = []
        for vin, vehicle in analysis.vehicle_results.items():
            miles = vehicle.get("annual_mileage", 0)
            fuel_savings = vehicle.get("annual_fuel_savings", 0)
            maint_savings = vehicle.get("annual_maintenance_savings", 0)
            total_savings = fuel_savings + maint_savings
            
            if miles > 0 and total_savings > 0:
                payback = ev_premium / total_savings
                data.append({
                    'mileage': miles,
                    'savings': total_savings,
                    'payback': payback,
                    'name': f"{vehicle.get('year', '')} {vehicle.get('make', '')} {vehicle.get('model', '')}"
                })
        
        if not data:
            ax_main.text(0.5, 0.5, "Insufficient data for ROI analysis", 
                        ha='center', va='center')
            ax_main.axis('off')
            return
        
        # Convert to arrays for plotting
        mileages = np.array([d['mileage'] for d in data])
        savings = np.array([d['savings'] for d in data])
        paybacks = np.array([d['payback'] for d in data])
        
        # Calculate median values for quadrant analysis
        median_mileage = np.median(mileages)
        median_savings = np.median(savings)
        
        # Create scatter plot with quadrant analysis
        scatter = ax_main.scatter(
            mileages, savings, c=paybacks,
            cmap='RdYlGn_r',  # Red to Yellow to Green (reversed)
            s=100, alpha=0.7,
            norm=plt.Normalize(vmin=0, vmax=10)  # 10-year max payback
        )
        
        # Add quadrant lines
        ax_main.axvline(median_mileage, color='gray', linestyle='--', alpha=0.5)
        ax_main.axhline(median_savings, color='gray', linestyle='--', alpha=0.5)
        
        # Label quadrants
        padding = 0.02
        bbox_props = dict(boxstyle="round,pad=0.3", fc="white", ec="gray", alpha=0.8)
        
        ax_main.text(max(mileages), max(savings), "Priority\nConversion",
                    ha='right', va='top', bbox=bbox_props)
        ax_main.text(min(mileages), max(savings), "High\nEfficiency",
                    ha='left', va='top', bbox=bbox_props)
        ax_main.text(max(mileages), min(savings), "High\nUtilization",
                    ha='right', va='bottom', bbox=bbox_props)
        ax_main.text(min(mileages), min(savings), "Low\nPriority",
                    ha='left', va='bottom', bbox=bbox_props)
        
        # Configure main plot
        ax_main.set_title('Fleet Electrification ROI Analysis', pad=20, fontsize=14, fontweight='bold')
        ax_main.set_xlabel('Annual Mileage', fontsize=12)
        ax_main.set_ylabel('Annual Savings ($)', fontsize=12)
        
        # Add colorbar
        cbar = figure.colorbar(scatter, ax=ax_main)
        cbar.set_label('Payback Period (years)', fontsize=10)
        
        # Add target lines
        target_payback_years = [3, 5, 7]
        for years in target_payback_years:
            x = np.linspace(min(mileages), max(mileages), 100)
            y = ev_premium / years
            ax_main.plot(x, [y] * 100, '--', color='gray', alpha=0.5,
                        label=f'{years}-year payback')
        
        ax_main.legend(loc='upper left')
        
        # Summary statistics
        summary_text = [
            'ROI Summary',
            '─' * 15,
            f'Median Payback: {np.median(paybacks):.1f} years',
            f'Best Payback: {np.min(paybacks):.1f} years',
            f'Vehicles < 5yr: {np.sum(paybacks < 5)}',
            '\nRecommendations:',
            '• Focus on upper-right quadrant',
            '• Target < 5 year payback',
            '• Consider pilot program'
        ]
        
        ax_summary.text(0.1, 0.95, '\n'.join(summary_text),
                       transform=ax_summary.transAxes,
                       bbox=dict(facecolor='#F5F5F5', edgecolor='#CCCCCC',
                               alpha=0.5, pad=10, boxstyle='round'),
                       verticalalignment='top',
                       fontsize=10)
        ax_summary.axis('off')
        
        # Distribution of mileages
        ax_hist_x.hist(mileages, bins=20, color=COLORS[0], alpha=0.5)
        ax_hist_x.set_xlabel('Annual Mileage Distribution')
        ax_hist_x.set_ylabel('Count')
        
        # Adjust layout
        figure.tight_layout()
    
    @staticmethod
    def charging_infrastructure(analysis: ChargingAnalysis, figure: Figure) -> None:
        """
        Create a chart showing charging infrastructure requirements.
        
        Args:
            analysis: ChargingAnalysis object
            figure: Figure to draw on
        """
        # Create chart
        ax = figure.add_subplot(111)
        
        # If no data, show message
        if not analysis.level2_chargers_needed and not analysis.dcfc_chargers_needed:
            ax.text(0.5, 0.5, "No charging infrastructure data available", 
                   ha='center', va='center')
            ax.axis('off')
            return
        
        # Create data for chart
        categories = ['Level 2', 'DC Fast Chargers']
        values = [analysis.level2_chargers_needed, analysis.dcfc_chargers_needed]
        
        # Create bar chart
        bars = ax.bar(categories, values, color=[PRIMARY_HEX_3, SECONDARY_HEX_1])
        
        # Add value labels
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, height + 0.1, 
                   f"{int(height)}", ha='center', va='bottom')
        
        # Configure chart
        ax.set_title("Charging Infrastructure Requirements")
        ax.set_ylabel("Number of Chargers")
        
        # Add power requirement text
        power_text = f"Maximum Power Required: {analysis.max_power_required:.1f} kW"
        ax.text(0.5, 0.9, power_text, transform=ax.transAxes, ha='center')
        
        # Add cost text
        cost_text = f"Estimated Installation Cost: ${analysis.estimated_installation_cost:,.0f}"
        ax.text(0.5, 0.85, cost_text, transform=ax.transAxes, ha='center')
        
        # Add usage pattern info
        pattern_text = f"Usage Pattern: {analysis.daily_usage_pattern.capitalize()}"
        window_text = f"Charging Window: {analysis.charging_window[0]}:00 to {analysis.charging_window[1]}:00"
        
        ax.text(0.5, 0.8, pattern_text, transform=ax.transAxes, ha='center')
        ax.text(0.5, 0.75, window_text, transform=ax.transAxes, ha='center')
        
        # Tight layout
        figure.tight_layout()


###############################################################################
# Emissions Charts
###############################################################################

class EmissionsCharts:
    """Charts for visualizing emissions inventory and analysis."""
    
    @staticmethod
    def emissions_inventory(inventory: EmissionsInventory, figure: Figure) -> None:
        """
        Create a comprehensive emissions analysis dashboard.
        
        Args:
            inventory: EmissionsInventory object
            figure: Figure to draw on
        """
        # Create grid layout
        gs = figure.add_gridspec(2, 2, height_ratios=[2, 1.2])
        
        # Main pie chart
        ax_pie = figure.add_subplot(gs[0, 0])
        ax_bars = figure.add_subplot(gs[0, 1])
        ax_trends = figure.add_subplot(gs[1, :])
        
        # 1. Pie Chart - Emissions by Type
        if inventory.by_vehicle_type:
            # Sort and prepare data
            sorted_data = sorted(inventory.by_vehicle_type.items(), 
                               key=lambda x: x[1], reverse=True)
            labels, sizes = zip(*sorted_data)
            
            # Create pie chart
            wedges, texts, autotexts = ax_pie.pie(
                sizes,
                labels=labels,
                autopct='%1.1f%%',
                startangle=90,
                colors=[COLORS[i % len(COLORS)] for i in range(len(sizes))],
                wedgeprops=dict(width=0.5)  # Create a donut chart
            )
            
            # Enhance text readability
            plt.setp(autotexts, size=8, weight="bold")
            plt.setp(texts, size=9)
            
            # Add center text with total
            total_emissions = sum(sizes)
            center_text = f'Total\n{total_emissions:.1f}\ntons CO₂e'
            ax_pie.text(0, 0, center_text, ha='center', va='center',
                       fontsize=10, fontweight='bold')
            
            ax_pie.set_title('Emissions by Vehicle Type', pad=20)
        
        # 2. Bar Chart - Top Contributors
        if inventory.by_department:
            # Sort departments by emissions
            sorted_depts = sorted(inventory.by_department.items(),
                                key=lambda x: x[1], reverse=True)[:8]
            dept_names, emissions = zip(*sorted_depts)
            
            # Create horizontal bars
            bars = ax_bars.barh(dept_names, emissions,
                              color=COLORS[1], alpha=0.7)
            
            # Add value labels
            for bar in bars:
                width = bar.get_width()
                ax_bars.text(width, bar.get_y() + bar.get_height()/2,
                           f'{width:.1f}',
                           ha='left', va='center',
                           fontsize=8, fontweight='bold',
                           bbox=dict(facecolor='white', edgecolor='none',
                                   alpha=0.7, pad=2))
            
            ax_bars.set_title('Top Emitting Departments', pad=20)
            ax_bars.set_xlabel('Emissions (tons CO₂e)')
        
        # 3. Trend Analysis
        if inventory.historical_data:
            years = sorted(inventory.historical_data.keys())
            emissions = [inventory.historical_data[year] for year in years]
            
            # Plot historical trend
            ax_trends.plot(years, emissions, 'o-', color=COLORS[0],
                         label='Historical', linewidth=2)
            
            # Add trend line
            z = np.polyfit(range(len(years)), emissions, 1)
            p = np.poly1d(z)
            ax_trends.plot(years, p(range(len(years))), '--',
                         color='gray', alpha=0.5, label='Trend')
            
            # Add projections if available
            if inventory.projected_emissions:
                proj_years = sorted(inventory.projected_emissions.keys())
                proj_emissions = [inventory.projected_emissions[year] 
                                for year in proj_years]
                ax_trends.plot(proj_years, proj_emissions, 'o--',
                             color=COLORS[2], label='Projected')
            
            # Add reduction target if available
            if inventory.reduction_target:
                baseline = emissions[0]
                target = baseline * (1 - inventory.reduction_target/100)
                target_year = inventory.target_year or max(years) + 5
                ax_trends.plot([min(years), target_year],
                             [baseline, target],
                             'r--', label=f'{inventory.reduction_target}% Target')
            
            ax_trends.set_title('Emissions Trend Analysis', pad=20)
            ax_trends.set_xlabel('Year')
            ax_trends.set_ylabel('Emissions (tons CO₂e)')
            ax_trends.legend()
            
            # Add annotations for key insights
            current = emissions[-1]
            change = ((current - emissions[0]) / emissions[0]) * 100
            trend_text = (f"YoY Change: {change:.1f}%\n"
                         f"Current: {current:.1f} tons")
            ax_trends.text(0.02, 0.95, trend_text,
                         transform=ax_trends.transAxes,
                         bbox=dict(facecolor='white', alpha=0.8),
                         verticalalignment='top')
        
        # Adjust layout
        figure.suptitle('Fleet Emissions Analysis Dashboard',
                       fontsize=16, fontweight='bold', y=0.95)
        figure.tight_layout()
    
    @staticmethod
    def emissions_trends(inventory: EmissionsInventory, figure: Figure) -> None:
        """
        Create a line chart showing emissions trends and targets.
        
        Args:
            inventory: EmissionsInventory object
            figure: Figure to draw on
        """
        # Create chart
        ax = figure.add_subplot(111)
        
        # If no data, show message
        if not inventory.historical_data and not inventory.projected_emissions:
            ax.text(0.5, 0.5, "No emissions trend data available", 
                   ha='center', va='center')
            ax.axis('off')
            return
        
        # Combine historical and projected data
        all_data = {**inventory.historical_data, **inventory.projected_emissions}
        
        # Sort by year
        years = sorted(all_data.keys())
        emissions = [all_data[year] for year in years]
        
        # Split into historical (solid line) and projected (dashed line)
        current_year = inventory.inventory_year
        
        historical_years = [y for y in years if y <= current_year]
        historical_emissions = [all_data[y] for y in historical_years]
        
        projected_years = [y for y in years if y > current_year]
        projected_emissions = [all_data[y] for y in projected_years]
        
        # Create line chart
        if historical_years:
            ax.plot(historical_years, historical_emissions, 'o-', 
                   color=PRIMARY_HEX_1, linewidth=2, label="Historical")
        
        if projected_years:
            ax.plot(projected_years, projected_emissions, 'o--', 
                   color=PRIMARY_HEX_3, linewidth=2, label="Projected")
        
        # Add target
        if inventory.target_year and inventory.baseline_year:
            # Draw a straight line from baseline to target
            baseline_emissions = all_data.get(inventory.baseline_year)
            if baseline_emissions:
                target_emissions = baseline_emissions * (1 - inventory.reduction_target / 100)
                
                ax.plot(
                    [inventory.baseline_year, inventory.target_year],
                    [baseline_emissions, target_emissions],
                    'r--', linewidth=1.5, label=f"{inventory.reduction_target}% Reduction Target"
                )
        
        # Configure chart
        ax.set_title("Fleet Emissions Trends and Targets")
        ax.set_xlabel("Year")
        ax.set_ylabel("Emissions (metric tons CO2e)")
        
        # Add legend
        ax.legend()
        
        # Add grid
        ax.grid(True, linestyle='--', alpha=0.7)
        
        # Tight layout
        figure.tight_layout()
    
    @staticmethod
    def emissions_by_department(inventory: EmissionsInventory, figure: Figure) -> None:
        """
        Create a horizontal bar chart showing emissions by department.
        
        Args:
            inventory: EmissionsInventory object
            figure: Figure to draw on
        """
        # Create chart
        ax = figure.add_subplot(111)
        
        # If no data, show message
        if not inventory.by_department:
            ax.text(0.5, 0.5, "No department emissions data available", 
                   ha='center', va='center')
            ax.axis('off')
            return
        
        # Sort departments by emissions
        sorted_depts = sorted(
            inventory.by_department.items(), 
            key=lambda x: x[1], 
            reverse=True
        )
        
        # Extract data for chart
        dept_names = [d[0] for d in sorted_depts]
        emissions = [d[1] for d in sorted_depts]
        
        # Create horizontal bar chart
        bars = ax.barh(dept_names, emissions, color=PRIMARY_HEX_3)
        
        # Add emissions labels
        for bar in bars:
            width = bar.get_width()
            ax.text(width + 0.5, bar.get_y() + bar.get_height()/2, 
                   f"{width:.1f}", ha='left', va='center')
        
        # Configure chart
        ax.set_title(f"Emissions by Department ({inventory.inventory_year})")
        ax.set_xlabel("Emissions (metric tons CO2e)")
        
        # Tight layout
        figure.tight_layout()
    
    @staticmethod
    def emissions_by_vehicle_type(inventory: EmissionsInventory, figure: Figure) -> None:
        """
        Create a horizontal bar chart showing emissions by vehicle type.
        
        Args:
            inventory: EmissionsInventory object
            figure: Figure to draw on
        """
        # Create chart
        ax = figure.add_subplot(111)
        
        # If no data, show message
        if not inventory.by_vehicle_type:
            ax.text(0.5, 0.5, "No vehicle type emissions data available", 
                   ha='center', va='center')
            ax.axis('off')
            return
        
        # Sort vehicle types by emissions
        sorted_types = sorted(
            inventory.by_vehicle_type.items(), 
            key=lambda x: x[1], 
            reverse=True
        )
        
        # Extract data for chart
        type_names = [t[0] for t in sorted_types]
        emissions = [t[1] for t in sorted_types]
        
        # Create horizontal bar chart
        bars = ax.barh(type_names, emissions, color=SECONDARY_HEX_1)
        
        # Add emissions labels
        for bar in bars:
            width = bar.get_width()
            ax.text(width + 0.5, bar.get_y() + bar.get_height()/2, 
                   f"{width:.1f}", ha='left', va='center')
        
        # Configure chart
        ax.set_title(f"Emissions by Vehicle Type ({inventory.inventory_year})")
        ax.set_xlabel("Emissions (metric tons CO2e)")
        
        # Tight layout
        figure.tight_layout()

