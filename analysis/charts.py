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
import matplotlib.rcsetup

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

# Standard figure sizes: 16:9 for slide-native rendering
FIG_SIZE_SLIDE = (13.33, 7.5)    # Full slide (16:9)
FIG_SIZE_HALF = (10, 5.6)        # Half-slide or default


###############################################################################
# Presentation-Quality Helper (Phase 9F)
###############################################################################

def apply_presentation_style(fig, ax, title: str = "", subtitle: str = "",
                             footnote: str = ""):
    """Apply consistent presentation-quality formatting to a matplotlib chart.

    Standardizes visual output across all chart types so they look
    professional when pasted into slides or exported to PDF/PNG.

    Args:
        fig: matplotlib Figure
        ax: matplotlib Axes (or primary axes if multiple)
        title: Main chart title
        subtitle: Insight subtitle computed from data (shown in smaller gray text)
        footnote: Footnote/disclaimer text (bottom of chart)
    """
    # White background
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')

    # Remove top/right spines, keep left/bottom light
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#CCCCCC')
    ax.spines['bottom'].set_color('#CCCCCC')

    # Subtle gridlines
    ax.grid(True, axis='y', alpha=0.2, linestyle='-', color='#CCCCCC')
    ax.set_axisbelow(True)

    # Title and subtitle
    if title:
        ax.set_title(title, fontsize=15, fontweight='bold', color='#333333',
                     pad=20 if subtitle else 12, loc='left')
    if subtitle:
        ax.text(0.0, 1.03, subtitle, transform=ax.transAxes,
                fontsize=10, color='#777777', style='italic',
                ha='left', va='bottom')

    # Footnote
    if footnote:
        fig.text(0.05, 0.01, footnote, fontsize=7, color='#AAAAAA',
                 style='italic', ha='left', va='bottom')

    # Tight layout with room for footnote
    fig.tight_layout(rect=[0, 0.03 if footnote else 0, 1, 1])


def _compute_fleet_insight(vehicles):
    """Compute a one-line executive insight string from fleet data."""
    if not vehicles:
        return ""

    total = len(vehicles)
    mpg_values = [v.fuel_economy.combined_mpg for v in vehicles
                  if v.fuel_economy.combined_mpg and v.fuel_economy.combined_mpg > 0]
    avg_mpg = sum(mpg_values) / len(mpg_values) if mpg_values else 0

    # Top emitter
    emissions = []
    for v in vehicles:
        co2 = v.fuel_economy.co2_primary or 0
        mileage = v.annual_mileage or 12000
        annual = (co2 * mileage) / 1_000_000 if co2 > 0 else 0
        emissions.append((annual, v))

    emissions.sort(reverse=True)
    top5_pct = (sum(e[0] for e in emissions[:5]) /
                sum(e[0] for e in emissions) * 100) if sum(e[0] for e in emissions) > 0 else 0

    parts = []
    if avg_mpg > 0:
        parts.append(f"Fleet average: {avg_mpg:.1f} MPG")
    if top5_pct > 0:
        parts.append(f"Top 5 vehicles account for {top5_pct:.0f}% of emissions")

    return " — ".join(parts)

###############################################################################
# Chart Factory
###############################################################################

    # Map of style names to matplotlib style sheets
STYLE_MAP = {
    'default': 'seaborn-v0_8-whitegrid',
    'minimal': 'seaborn-v0_8-white',
    'dark': 'dark_background',
    'colorful': 'seaborn-v0_8-colorblind',
}

# Map of color scheme names to matplotlib colormaps / color cycles
COLOR_SCHEME_MAP = {
    'default': None,       # Use the module-level COLORS palette
    'viridis': 'viridis',
    'magma': 'magma',
    'plasma': 'plasma',
    'inferno': 'inferno',
}


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
                chart_style: Style name (default, minimal, dark, colorful)
                color_scheme: Color scheme name (default, viridis, magma, plasma, inferno)
                fig_size: (width, height) tuple; defaults to FIG_SIZE_HALF (10×5.6).
                          Pass FIG_SIZE_SLIDE (13.33×7.5) for full-bleed slide export.

        Returns:
            Matplotlib Figure object
        """
        # Extract style/color/size kwargs (not forwarded to individual chart methods)
        chart_style = kwargs.pop('chart_style', 'default')
        color_scheme = kwargs.pop('color_scheme', 'default')
        fig_size = kwargs.pop('fig_size', FIG_SIZE_HALF)

        # Resolve matplotlib style sheet
        mpl_style = STYLE_MAP.get(chart_style, STYLE_MAP['default'])

        # Apply color scheme — swap prop_cycle when a colormap is requested
        cmap_name = COLOR_SCHEME_MAP.get(color_scheme)
        style_overrides = {}
        if cmap_name:
            cmap = matplotlib.colormaps[cmap_name]
            cycle_colors = [mcolors.to_hex(cmap(i / 9)) for i in range(10)]
            style_overrides['axes.prop_cycle'] = matplotlib.rcsetup.cycler('color', cycle_colors)

        # Create figure if not provided
        if figure is None:
            figure = plt.figure(figsize=fig_size, dpi=100)
            figure.tight_layout(pad=3.0)
        else:
            figure.clear()
        
        # Apply style context and render the requested chart
        with plt.style.context([mpl_style, style_overrides] if style_overrides else mpl_style):
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

            elif chart_type == "Fleet Cash Flow":
                DecisionCharts.fleet_cashflow_chart(data, figure)

            elif chart_type == "Replacement Priority":
                DecisionCharts.replacement_priority_chart(data, figure)

            elif chart_type == "Scenario Comparison":
                DecisionCharts.scenario_comparison_chart(data, figure)

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
            other_count = sum(count for _, count in sorted_classes[10:])
            sorted_classes = sorted_classes[:10]
            sorted_classes.append(("Other", other_count))
        
        # Extract data for chart
        class_names, counts = zip(*sorted_classes)
        
        # Create chart
        ax = figure.add_subplot(111)
        bars = ax.barh(class_names, counts, color=PRIMARY_HEX_3, edgecolor='none',
                       height=0.6)

        # Add count labels to bars
        total = sum(counts)
        for bar in bars:
            width = bar.get_width()
            pct = width / total * 100 if total > 0 else 0
            ax.text(width + max(total * 0.01, 0.3), bar.get_y() + bar.get_height()/2,
                    f"{int(width)}  ({pct:.0f}%)", ha='left', va='center',
                    fontsize=9, color='#555555')

        # Insight subtitle
        top_class = class_names[0] if class_names else "N/A"
        top_pct = counts[0] / total * 100 if total > 0 else 0
        subtitle = f"{top_class} represents {top_pct:.0f}% of the fleet ({total} vehicles)"

        ax.set_xlabel("Count")
        ax.set_ylabel("")
        apply_presentation_style(figure, ax,
                                 title="Vehicle Body Class Distribution",
                                 subtitle=subtitle)
    
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
        n, bins_arr, patches = ax.hist(mpg_values, bins=bins, edgecolor='white',
                                       alpha=0.85, color=PRIMARY_HEX_3)

        ax.set_xlabel("MPG (Combined)")
        ax.set_ylabel("Number of Vehicles")

        # Add mean line
        mean_mpg = np.mean(mpg_values)
        median_mpg = np.median(mpg_values)
        ax.axvline(mean_mpg, color=SECONDARY_HEX_1, linestyle='dashed', linewidth=2,
                  label=f'Mean: {mean_mpg:.1f} MPG')
        ax.axvline(median_mpg, color=COLORS[2], linestyle='dotted', linewidth=2,
                  label=f'Median: {median_mpg:.1f} MPG')
        ax.legend(frameon=True, facecolor='white', edgecolor='#CCCCCC')

        # Insight subtitle
        below_15 = sum(1 for v in mpg_values if v < 15)
        below_pct = below_15 / len(mpg_values) * 100 if mpg_values else 0
        subtitle = f"Fleet average: {mean_mpg:.1f} MPG"
        if below_pct > 0:
            subtitle += f" — {below_pct:.0f}% of vehicles below 15 MPG"

        apply_presentation_style(figure, ax,
                                 title="Fuel Economy Distribution",
                                 subtitle=subtitle)
    
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
        n, bins_arr, patches = ax.hist(co2_values, bins=10, edgecolor='white',
                                       alpha=0.85, color=SECONDARY_HEX_1)

        ax.set_xlabel("CO2 Emissions (g/mile)")
        ax.set_ylabel("Number of Vehicles")

        # Add mean line
        mean_co2 = np.mean(co2_values)
        ax.axvline(mean_co2, color=COLORS[3], linestyle='dashed', linewidth=2,
                  label=f'Mean: {mean_co2:.1f} g/mile')
        ax.legend(frameon=True, facecolor='white', edgecolor='#CCCCCC')

        # Insight
        high_emitters = sum(1 for v in co2_values if v > 500)
        subtitle = f"Fleet average: {mean_co2:.0f} g/mile"
        if high_emitters > 0:
            subtitle += f" — {high_emitters} vehicle{'s' if high_emitters != 1 else ''} above 500 g/mile"

        apply_presentation_style(figure, ax,
                                 title="CO2 Emissions Distribution",
                                 subtitle=subtitle)
    
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
        scatter = ax.scatter(mpg_values, co2_values, c=COLORS[0], alpha=0.6,
                            edgecolors='white', linewidth=0.5, s=60)

        # Add best fit line
        if len(mpg_values) > 1:
            try:
                z = np.polyfit(mpg_values, co2_values, 1)
                p = np.poly1d(z)
                x_range = np.linspace(min(mpg_values), max(mpg_values), 100)
                ax.plot(x_range, p(x_range), linestyle='--', color=COLORS[4],
                       alpha=0.7, label='Trend')
            except Exception as e:
                logger.warning(f"Error calculating trend line: {e}")

        ax.set_xlabel("MPG (Combined)")
        ax.set_ylabel("CO2 Emissions (g/mile)")
        if len(mpg_values) > 1:
            ax.legend(frameon=True, facecolor='white', edgecolor='#CCCCCC')

        # Annotate worst emitter
        if co2_values:
            max_idx = co2_values.index(max(co2_values))
            ax.annotate(f"{max(co2_values):.0f} g/mi",
                       xy=(mpg_values[max_idx], co2_values[max_idx]),
                       xytext=(10, 10), textcoords='offset points',
                       fontsize=8, color=COLORS[3],
                       arrowprops=dict(arrowstyle='->', color=COLORS[3], lw=1))

        # Correlation insight
        if len(mpg_values) > 2:
            corr = np.corrcoef(mpg_values, co2_values)[0, 1]
            subtitle = f"R = {corr:.2f} correlation — {len(mpg_values)} vehicles with both MPG and CO2 data"
        else:
            subtitle = f"{len(mpg_values)} vehicles with both MPG and CO2 data"

        apply_presentation_style(figure, ax,
                                 title="CO2 Emissions vs. Fuel Economy",
                                 subtitle=subtitle)
    
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
        
        ax.set_xlabel("Primary Fuel CO2 (g/mile)")
        ax.set_ylabel("Alternative Fuel CO2 (g/mile)")

        # Add legend if not too many makes
        if len(unique_makes) <= 10:
            ax.legend(fontsize='small', frameon=True, facecolor='white',
                     edgecolor='#CCCCCC')

        subtitle = f"{len(co2_primary)} dual-fuel vehicles across {len(unique_makes)} makes"
        apply_presentation_style(figure, ax,
                                 title="Primary vs. Alternative Fuel CO2 Emissions",
                                 subtitle=subtitle)
    
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
        n, bins_arr, patches = ax.hist(range_values, bins=10, edgecolor='white',
                                       alpha=0.85, color=PRIMARY_HEX_3)

        ax.set_xlabel("Range (miles)")
        ax.set_ylabel("Number of Vehicles")

        # Add mean line
        mean_range = np.mean(range_values)
        ax.axvline(mean_range, color=SECONDARY_HEX_1, linestyle='dashed', linewidth=2,
                  label=f'Mean: {mean_range:.1f} miles')
        ax.legend(frameon=True, facecolor='white', edgecolor='#CCCCCC')

        subtitle = f"Average range: {mean_range:.0f} miles across {len(range_values)} vehicles"
        apply_presentation_style(figure, ax,
                                 title="Alternative Fuel Range Distribution",
                                 subtitle=subtitle)
    
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
        bars = ax.bar(make_names, counts, color=COLORS[:len(make_names)],
                     edgecolor='none')

        # Add count + percentage labels above bars
        total = sum(counts)
        for bar in bars:
            height = bar.get_height()
            pct = height / total * 100 if total > 0 else 0
            ax.text(bar.get_x() + bar.get_width()/2, height + max(total * 0.005, 0.1),
                   f"{int(height)} ({pct:.0f}%)", ha='center', va='bottom',
                   fontsize=8, color='#555555')

        ax.set_xlabel("")
        ax.set_ylabel("Count")

        # Rotate x-axis labels for readability
        plt.xticks(rotation=45, ha='right')

        # Insight
        top_make = make_names[0] if make_names else "N/A"
        top_pct = counts[0] / total * 100 if total > 0 else 0
        subtitle = f"{top_make} leads with {top_pct:.0f}% of fleet — {len(make_names)} makes total"

        apply_presentation_style(figure, ax,
                                 title="Vehicle Make Distribution",
                                 subtitle=subtitle)
    
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
        bars = ax.bar(model_names, counts, color=COLORS[:len(model_names)],
                     edgecolor='none')

        # Add count labels
        total = sum(counts)
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, height + max(total * 0.005, 0.1),
                   f"{int(height)}", ha='center', va='bottom',
                   fontsize=8, color='#555555')

        ax.set_xlabel("")
        ax.set_ylabel("Count")

        # Rotate x-axis labels for readability
        plt.xticks(rotation=45, ha='right')

        top_model = model_names[0] if model_names else "N/A"
        subtitle = f"Top {min(len(model_names), 15)} models — {top_model} is most common ({counts[0]})"

        apply_presentation_style(figure, ax,
                                 title="Top Vehicle Models",
                                 subtitle=subtitle)
    
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

        # Create donut chart for a more modern look
        wedges, texts, autotexts = ax.pie(
            sizes,
            labels=None,
            autopct='%1.1f%%',
            startangle=90,
            colors=COLORS[:len(labels)],
            wedgeprops=dict(width=0.55, edgecolor='white', linewidth=2),
            pctdistance=0.75,
        )
        ax.axis('equal')

        # Add legend
        ax.legend(wedges, labels, title="Fuel Types", loc="center left",
                 bbox_to_anchor=(1, 0, 0.5, 1), frameon=False)

        # Set text properties
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontsize(9)
            autotext.set_fontweight('bold')

        # Center text with total
        total = sum(sizes)
        ax.text(0, 0, f"{total}\nvehicles", ha='center', va='center',
                fontsize=12, fontweight='bold', color='#333333')

        # Title
        dominant = labels[0] if labels else "N/A"
        dom_pct = sizes[0] / total * 100 if total > 0 else 0
        ax.set_title("Fuel Type Distribution", fontsize=15, fontweight='bold',
                     color='#333333', loc='left', pad=20)
        ax.text(0.0, 1.03, f"{dominant} dominates at {dom_pct:.0f}% of fleet",
                transform=ax.transAxes, fontsize=10, color='#777777',
                style='italic', ha='left', va='bottom')

        figure.patch.set_facecolor('white')
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
        n, bins_arr, patches = ax.hist(ages, bins=range(0, int(max(ages)) + 2),
                                       edgecolor='white', alpha=0.85, color=PRIMARY_HEX_3)

        ax.set_xlabel("Age (years)")
        ax.set_ylabel("Number of Vehicles")

        # Add mean line
        mean_age = np.mean(ages)
        ax.axvline(mean_age, color=SECONDARY_HEX_1, linestyle='dashed', linewidth=2,
                  label=f'Mean: {mean_age:.1f} years')
        ax.legend(frameon=True, facecolor='white', edgecolor='#CCCCCC')

        # Insight
        over_10 = sum(1 for a in ages if a > 10)
        over_pct = over_10 / len(ages) * 100 if ages else 0
        subtitle = f"Average age: {mean_age:.1f} years"
        if over_pct > 0:
            subtitle += f" — {over_pct:.0f}% of fleet is over 10 years old"

        apply_presentation_style(figure, ax,
                                 title="Fleet Age Distribution",
                                 subtitle=subtitle)


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
        bars_ice = ax.bar(x - width/2, ice_costs, width, label='ICE (Gasoline)',
                         color=SECONDARY_HEX_1, edgecolor='none')
        bars_ev = ax.bar(x + width/2, ev_costs, width, label='EV (Electric)',
                        color=PRIMARY_HEX_3, edgecolor='none')

        ax.set_ylabel("Annual Fuel Cost ($)")
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45, ha='right')
        ax.legend(frameon=True, facecolor='white', edgecolor='#CCCCCC')

        # Add savings percentage annotations
        for i in range(len(ice_costs)):
            savings_pct = 100 * (ice_costs[i] - ev_costs[i]) / ice_costs[i] if ice_costs[i] > 0 else 0
            ax.annotate(f"{savings_pct:.0f}%\nsaved",
                       xy=(i, max(ice_costs[i], ev_costs[i])),
                       xytext=(0, 12), textcoords='offset points',
                       ha='center', va='bottom', fontsize=7,
                       color=COLORS[2], fontweight='bold')

        # Insight
        total_ice = sum(ice_costs)
        total_ev = sum(ev_costs)
        total_savings = total_ice - total_ev
        avg_pct = total_savings / total_ice * 100 if total_ice > 0 else 0
        subtitle = f"Average {avg_pct:.0f}% fuel cost reduction — ${total_savings:,.0f}/yr potential savings across {len(labels)} vehicles"

        apply_presentation_style(figure, ax,
                                 title="Annual Fuel Cost: ICE vs. Electric",
                                 subtitle=subtitle)
    
    @staticmethod
    def electrification_potential(analysis: ElectrificationAnalysis, figure: Figure,
                                top_n: int = 10, **kwargs) -> None:
        """
        Create a chart showing electrification savings potential by vehicle type.

        Aggregates total NPV savings from vehicle_results by make/model and
        displays a horizontal bar chart of the top vehicle types.

        Args:
            analysis: Electrification analysis results
            figure: Figure to draw on
            top_n: Number of top vehicle types to show
        """
        if not analysis.vehicle_results:
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, "No electrification analysis data.\nRun analysis first.",
                    ha='center', va='center', fontsize=12)
            ax.axis('off')
            return

        # Aggregate savings by vehicle type (make + model)
        savings_by_type = {}   # type -> total NPV savings
        count_by_type = {}     # type -> vehicle count
        for _vin, result in analysis.vehicle_results.items():
            make = result.get("make", "Unknown")
            model = result.get("model", "Unknown")
            vtype = f"{make} {model}".strip() or "Unknown"
            npv = result.get("total_npv_savings", 0.0)
            savings_by_type[vtype] = savings_by_type.get(vtype, 0.0) + npv
            count_by_type[vtype] = count_by_type.get(vtype, 0) + 1

        sorted_types = sorted(savings_by_type.items(), key=lambda x: x[1], reverse=True)

        if len(sorted_types) > top_n:
            other_sum = sum(s for _, s in sorted_types[top_n:])
            sorted_types = sorted_types[:top_n] + [("Other", other_sum)]

        types, savings = zip(*sorted_types)

        # Layout: bar chart left, summary right
        gs = figure.add_gridspec(1, 2, width_ratios=[3, 1])

        # Main bar chart
        ax1 = figure.add_subplot(gs[0])
        bars = ax1.barh(types, savings, color=COLORS[0], alpha=0.8)

        # Value labels on bars
        for bar in bars:
            width = bar.get_width()
            label = f'${width:,.0f}'
            ax1.text(width, bar.get_y() + bar.get_height() / 2,
                     f' {label}',
                     ha='left', va='center',
                     fontsize=9, fontweight='bold',
                     bbox=dict(facecolor='white', edgecolor='none', alpha=0.7, pad=2))

        ax1.set_xlabel('Total NPV Savings ($)', fontsize=12)
        ax1.invert_yaxis()

        # Apply presentation style to the bar chart area
        subtitle = f"${sum(savings):,.0f} total savings across {len(analysis.vehicle_results)} vehicles"
        apply_presentation_style(figure, ax1,
                                 title='Electrification Savings by Vehicle Type',
                                 subtitle=subtitle)

        # Summary statistics on the right
        ax2 = figure.add_subplot(gs[1])
        ax2.axis('off')
        ax2.set_facecolor('white')

        total_savings = analysis.total_savings
        num_vehicles = len(analysis.vehicle_results)
        avg_savings = total_savings / num_vehicles if num_vehicles else 0

        summary_lines = [
            ('Vehicles', f'{num_vehicles}'),
            ('Total Savings', f'${total_savings:,.0f}'),
            ('Avg/Vehicle', f'${avg_savings:,.0f}'),
            ('CO₂ Reduction', f'{analysis.co2_savings:,.1f} tons'),
            ('Payback', f'{analysis.payback_period:.1f} yr'),
        ]

        y_pos = 0.90
        ax2.text(0.1, y_pos + 0.06, 'Summary',
                transform=ax2.transAxes, fontsize=12, fontweight='bold',
                color='#333333')
        for label, value in summary_lines:
            ax2.text(0.1, y_pos, label, transform=ax2.transAxes,
                    fontsize=9, color='#777777')
            ax2.text(0.95, y_pos, value, transform=ax2.transAxes,
                    fontsize=10, fontweight='bold', color='#333333', ha='right')
            y_pos -= 0.10
    
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
        bars = ax.barh(vehicle_names, reductions, color=COLORS[5], edgecolor='none',
                       height=0.6)

        # Add reduction labels
        for bar in bars:
            width = bar.get_width()
            ax.text(width + max(max(reductions) * 0.01, 0.3),
                   bar.get_y() + bar.get_height()/2,
                   f"{width:.1f} tons", ha='left', va='center',
                   fontsize=9, color='#555555')

        ax.set_xlabel("Lifetime CO2 Reduction (metric tons)")
        ax.set_ylabel("")

        total_reduction = sum(reductions)
        subtitle = f"Top {len(reductions)} vehicles — {total_reduction:,.1f} metric tons total lifetime reduction"

        apply_presentation_style(figure, ax,
                                 title="CO2 Emissions Reduction Potential",
                                 subtitle=subtitle)
    
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
        figure.patch.set_facecolor('white')
        ax_main.set_facecolor('white')
        ax_main.spines['top'].set_visible(False)
        ax_main.spines['right'].set_visible(False)
        ax_main.spines['left'].set_color('#CCCCCC')
        ax_main.spines['bottom'].set_color('#CCCCCC')
        ax_main.set_title('Fleet Electrification ROI Analysis', pad=20, fontsize=15,
                         fontweight='bold', color='#333333', loc='left')
        median_pb = np.median(paybacks)
        ax_main.text(0.0, 1.03,
                    f"Median payback: {median_pb:.1f} years — {np.sum(paybacks < 5)} vehicles under 5-year payback",
                    transform=ax_main.transAxes, fontsize=10, color='#777777',
                    style='italic', ha='left', va='bottom')
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
        ax_summary.axis('off')
        ax_summary.set_facecolor('white')

        summary_lines = [
            ('Median Payback', f'{np.median(paybacks):.1f} yr'),
            ('Best Payback', f'{np.min(paybacks):.1f} yr'),
            ('Vehicles < 5yr', f'{np.sum(paybacks < 5)}'),
            ('Avg Savings', f'${np.mean(savings):,.0f}/yr'),
        ]

        y_pos = 0.88
        ax_summary.text(0.1, y_pos + 0.08, 'ROI Summary',
                       transform=ax_summary.transAxes, fontsize=12,
                       fontweight='bold', color='#333333')
        for label, value in summary_lines:
            ax_summary.text(0.1, y_pos, label, transform=ax_summary.transAxes,
                          fontsize=9, color='#777777')
            ax_summary.text(0.95, y_pos, value, transform=ax_summary.transAxes,
                          fontsize=10, fontweight='bold', color='#333333', ha='right')
            y_pos -= 0.12

        # Distribution of mileages
        ax_hist_x.hist(mileages, bins=20, color=COLORS[0], alpha=0.6, edgecolor='white')
        ax_hist_x.set_xlabel('Annual Mileage Distribution')
        ax_hist_x.set_ylabel('Count')
        ax_hist_x.spines['top'].set_visible(False)
        ax_hist_x.spines['right'].set_visible(False)
        ax_hist_x.spines['left'].set_color('#CCCCCC')
        ax_hist_x.spines['bottom'].set_color('#CCCCCC')
        ax_hist_x.set_facecolor('white')

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
        categories = ['Level 2\nChargers', 'DC Fast\nChargers']
        values = [analysis.level2_chargers_needed, analysis.dcfc_chargers_needed]

        # Create bar chart
        bars = ax.bar(categories, values, color=[PRIMARY_HEX_3, SECONDARY_HEX_1],
                     edgecolor='none', width=0.5)

        # Add value labels
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, height + 0.2,
                   f"{int(height)}", ha='center', va='bottom',
                   fontsize=14, fontweight='bold', color='#333333')

        ax.set_ylabel("Number of Chargers")

        # Add infrastructure details as footnote
        footnote_parts = [
            f"Peak power: {analysis.max_power_required:.0f} kW",
            f"Est. cost: ${analysis.estimated_installation_cost:,.0f}",
            f"Pattern: {analysis.daily_usage_pattern}",
            f"Window: {analysis.charging_window[0]}:00-{analysis.charging_window[1]}:00",
        ]
        footnote = "  |  ".join(footnote_parts)

        total_chargers = sum(values)
        subtitle = f"{total_chargers} chargers required — ${analysis.estimated_installation_cost:,.0f} estimated installation"

        apply_presentation_style(figure, ax,
                                 title="Charging Infrastructure Requirements",
                                 subtitle=subtitle,
                                 footnote=footnote)


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
            
            ax_pie.set_title('Emissions by Vehicle Type', pad=15, fontsize=12,
                           fontweight='bold', color='#333333', loc='left')

        # 2. Bar Chart - Top Contributors
        if inventory.by_department:
            sorted_depts = sorted(inventory.by_department.items(),
                                key=lambda x: x[1], reverse=True)[:8]
            dept_names, emissions = zip(*sorted_depts)

            bars = ax_bars.barh(dept_names, emissions,
                              color=COLORS[1], alpha=0.85, edgecolor='none', height=0.6)

            for bar in bars:
                width = bar.get_width()
                ax_bars.text(width + max(max(emissions) * 0.01, 0.1),
                           bar.get_y() + bar.get_height()/2,
                           f'{width:.1f}',
                           ha='left', va='center',
                           fontsize=8, color='#555555')

            ax_bars.set_title('Top Emitting Departments', pad=15, fontsize=12,
                            fontweight='bold', color='#333333', loc='left')
            ax_bars.set_xlabel('Emissions (tons CO₂e)')
            ax_bars.spines['top'].set_visible(False)
            ax_bars.spines['right'].set_visible(False)
            ax_bars.spines['left'].set_color('#CCCCCC')
            ax_bars.spines['bottom'].set_color('#CCCCCC')
        
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
            
            ax_trends.set_title('Emissions Trend Analysis', pad=15, fontsize=12,
                              fontweight='bold', color='#333333', loc='left')
            ax_trends.set_xlabel('Year')
            ax_trends.set_ylabel('Emissions (tons CO₂e)')
            ax_trends.legend(frameon=True, facecolor='white', edgecolor='#CCCCCC')
            ax_trends.spines['top'].set_visible(False)
            ax_trends.spines['right'].set_visible(False)
            ax_trends.spines['left'].set_color('#CCCCCC')
            ax_trends.spines['bottom'].set_color('#CCCCCC')
            ax_trends.set_facecolor('white')

            # Add annotations for key insights
            current = emissions[-1]
            change = ((current - emissions[0]) / emissions[0]) * 100
            trend_text = (f"YoY Change: {change:.1f}%\n"
                         f"Current: {current:.1f} tons")
            ax_trends.text(0.02, 0.95, trend_text,
                         transform=ax_trends.transAxes,
                         bbox=dict(facecolor='white', alpha=0.9, edgecolor='#CCCCCC',
                                  boxstyle='round'),
                         verticalalignment='top', fontsize=9)

        # Overall figure styling
        figure.patch.set_facecolor('white')
        figure.suptitle('Fleet Emissions Analysis Dashboard',
                       fontsize=16, fontweight='bold', color='#333333', y=0.98)
        figure.tight_layout(rect=[0, 0, 1, 0.96])
    
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
        
        ax.set_xlabel("Year")
        ax.set_ylabel("Emissions (metric tons CO2e)")
        ax.legend(frameon=True, facecolor='white', edgecolor='#CCCCCC')

        # Subtitle and footnote
        subtitle = ""
        footnote = ""
        if historical_emissions:
            subtitle = f"Current: {historical_emissions[-1]:,.1f} MT CO2e ({inventory.inventory_year})"
        if getattr(inventory, 'is_synthetic', False):
            footnote = "Illustrative data — historical/projected values are estimates, not from actual records."

        apply_presentation_style(figure, ax,
                                 title="Fleet Emissions Trends and Targets",
                                 subtitle=subtitle,
                                 footnote=footnote)
    
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
        bars = ax.barh(dept_names, emissions, color=PRIMARY_HEX_3, edgecolor='none',
                       height=0.6)

        # Add emissions labels
        total = sum(emissions)
        for bar in bars:
            width = bar.get_width()
            pct = width / total * 100 if total > 0 else 0
            ax.text(width + max(total * 0.01, 0.3), bar.get_y() + bar.get_height()/2,
                   f"{width:.1f} ({pct:.0f}%)", ha='left', va='center',
                   fontsize=9, color='#555555')

        ax.set_xlabel("Emissions (metric tons CO2e)")
        ax.set_ylabel("")

        top_dept = dept_names[0] if dept_names else "N/A"
        top_pct = emissions[0] / total * 100 if total > 0 else 0
        subtitle = f"{top_dept} leads with {top_pct:.0f}% of total emissions ({total:,.1f} MT)"

        apply_presentation_style(figure, ax,
                                 title=f"Emissions by Department ({inventory.inventory_year})",
                                 subtitle=subtitle)

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
        bars = ax.barh(type_names, emissions, color=SECONDARY_HEX_1, edgecolor='none',
                       height=0.6)

        # Add emissions labels
        total = sum(emissions)
        for bar in bars:
            width = bar.get_width()
            pct = width / total * 100 if total > 0 else 0
            ax.text(width + max(total * 0.01, 0.3), bar.get_y() + bar.get_height()/2,
                   f"{width:.1f} ({pct:.0f}%)", ha='left', va='center',
                   fontsize=9, color='#555555')

        ax.set_xlabel("Emissions (metric tons CO2e)")
        ax.set_ylabel("")

        top_type = type_names[0] if type_names else "N/A"
        top_pct = emissions[0] / total * 100 if total > 0 else 0
        subtitle = f"{top_type} accounts for {top_pct:.0f}% — {total:,.1f} MT total"

        apply_presentation_style(figure, ax,
                                 title=f"Emissions by Vehicle Type ({inventory.inventory_year})",
                                 subtitle=subtitle)


###############################################################################
# Decision-Support Charts (Phase 9F)
###############################################################################

class DecisionCharts:
    """Charts for decision-support analysis — cash flows, priorities, scenarios."""

    @staticmethod
    def fleet_cashflow_chart(analysis: ElectrificationAnalysis, figure: Figure) -> None:
        """
        Side-by-side cumulative cost curves (ICE vs EV fleet) with crossover annotation.

        Uses fleet_cash_flows from the ElectrificationAnalysis to plot year-by-year
        cumulative costs for ICE and EV fleet-wide, highlighting the payback crossover.

        Args:
            analysis: ElectrificationAnalysis with fleet_cash_flows populated
            figure: Figure to draw on
        """
        cash_flows = getattr(analysis, 'fleet_cash_flows', [])

        if not cash_flows:
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, "No fleet cash flow data available.\nRun analysis with EV matching first.",
                    ha='center', va='center', fontsize=12)
            ax.axis('off')
            return

        ax = figure.add_subplot(111)

        years = [cf.get('year', i) for i, cf in enumerate(cash_flows)]
        ice_cumul = [cf.get('ice_cumulative', 0) for cf in cash_flows]
        ev_cumul = [cf.get('ev_cumulative', 0) for cf in cash_flows]

        ax.plot(years, [c / 1e6 for c in ice_cumul], '-', color=SECONDARY_HEX_1,
                linewidth=2.5, label='ICE Fleet (Cumulative)', marker='o', markersize=4)
        ax.plot(years, [c / 1e6 for c in ev_cumul], '-', color=PRIMARY_HEX_3,
                linewidth=2.5, label='EV Fleet (Cumulative)', marker='s', markersize=4)

        # Fill the savings area
        ax.fill_between(years,
                        [i / 1e6 for i in ice_cumul],
                        [e / 1e6 for e in ev_cumul],
                        alpha=0.1, color=COLORS[5],
                        where=[i > e for i, e in zip(ice_cumul, ev_cumul)])

        # Find crossover point
        crossover_year = None
        for i in range(1, len(cash_flows)):
            if ev_cumul[i - 1] > ice_cumul[i - 1] and ev_cumul[i] <= ice_cumul[i]:
                # Interpolate
                frac = (ev_cumul[i - 1] - ice_cumul[i - 1]) / (
                    (ev_cumul[i - 1] - ice_cumul[i - 1]) - (ev_cumul[i] - ice_cumul[i])
                )
                crossover_year = years[i - 1] + frac * (years[i] - years[i - 1])
                crossover_cost = (ice_cumul[i - 1] + frac * (ice_cumul[i] - ice_cumul[i - 1])) / 1e6
                ax.annotate(f"Payback: Year {crossover_year:.1f}",
                           xy=(crossover_year, crossover_cost),
                           xytext=(20, 20), textcoords='offset points',
                           fontsize=10, fontweight='bold', color=COLORS[5],
                           arrowprops=dict(arrowstyle='->', color=COLORS[5], lw=2),
                           bbox=dict(facecolor='white', edgecolor=COLORS[5],
                                    boxstyle='round,pad=0.3'))
                break

        ax.set_xlabel("Year")
        ax.set_ylabel("Cumulative Cost ($M)")
        ax.legend(frameon=True, facecolor='white', edgecolor='#CCCCCC', fontsize=10)

        # Format y-axis as millions
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'${x:.1f}M'))

        total_savings = (ice_cumul[-1] - ev_cumul[-1]) / 1e6 if cash_flows else 0
        subtitle = f"Fleet lifetime savings: ${total_savings:,.1f}M"
        if crossover_year:
            subtitle += f" — payback in Year {crossover_year:.1f}"

        apply_presentation_style(figure, ax,
                                 title="Fleet Total Cost of Ownership: ICE vs. Electric",
                                 subtitle=subtitle,
                                 footnote="Includes fuel escalation, maintenance, battery degradation, and residual values")

    @staticmethod
    def replacement_priority_chart(fleet_data: Union[Fleet, List[FleetVehicle]],
                                    figure: Figure, top_n: int = 12) -> None:
        """
        Horizontal bar chart of top vehicles by NPV savings, with EV labels and payback.

        Uses custom_fields populated by EV matching (Phase 9B) to show which
        vehicles should be replaced first based on financial benefit.

        Args:
            fleet_data: Fleet or list of vehicles (with EV matching data in custom_fields)
            figure: Figure to draw on
            top_n: Number of top vehicles to display
        """
        vehicles = fleet_data.vehicles if isinstance(fleet_data, Fleet) else fleet_data

        # Gather vehicles with EV match data
        candidates = []
        for v in vehicles:
            npv = v.custom_fields.get('_npv_savings', 0)
            ev_name = v.custom_fields.get('EV Equivalent', '')
            payback = v.custom_fields.get('_payback_years', 0)
            if npv and float(npv) > 0 and ev_name:
                make = v.vehicle_id.make or ''
                model = v.vehicle_id.model or ''
                year = v.vehicle_id.year or ''
                candidates.append({
                    'label': f"{year} {make} {model}".strip(),
                    'npv': float(npv),
                    'ev': ev_name,
                    'payback': float(payback) if payback else 0,
                })

        if not candidates:
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, "No EV replacement data available.\nRun analysis with EV matching first.",
                    ha='center', va='center', fontsize=12)
            ax.axis('off')
            return

        # Sort and take top N
        candidates.sort(key=lambda c: c['npv'], reverse=True)
        candidates = candidates[:top_n]
        candidates.reverse()  # For horizontal bars (top at top)

        ax = figure.add_subplot(111)

        labels = [c['label'] for c in candidates]
        npvs = [c['npv'] for c in candidates]

        bars = ax.barh(labels, npvs, color=COLORS[0], edgecolor='none', height=0.6)

        # Add NPV and EV labels
        for i, (bar, c) in enumerate(zip(bars, candidates)):
            width = bar.get_width()
            # NPV label
            ax.text(width + max(max(npvs) * 0.01, 50), bar.get_y() + bar.get_height() / 2,
                   f"${width:,.0f}  →  {c['ev']}  ({c['payback']:.1f}yr)",
                   ha='left', va='center', fontsize=8, color='#555555')

        ax.set_xlabel("Lifetime NPV Savings ($)")
        ax.set_ylabel("")

        total_npv = sum(c['npv'] for c in candidates)
        subtitle = f"Top {len(candidates)} vehicles — ${total_npv:,.0f} combined lifetime savings"

        apply_presentation_style(figure, ax,
                                 title="Priority Replacement Vehicles by NPV Savings",
                                 subtitle=subtitle,
                                 footnote="NPV includes fuel savings, maintenance savings, incentives, and residual values")

    @staticmethod
    def scenario_comparison_chart(fleet_data: Union[Fleet, List[FleetVehicle]],
                                   figure: Figure) -> None:
        """
        Overlaid area charts comparing 2-4 electrification scenarios.

        Runs preset scenarios and plots cumulative vehicles electrified per year
        as overlaid filled areas, with cumulative savings on secondary axis.

        Args:
            fleet_data: Fleet or list of vehicles
            figure: Figure to draw on
        """
        from analysis.scenarios import compare_scenarios

        vehicles = fleet_data.vehicles if isinstance(fleet_data, Fleet) else fleet_data

        comparison = compare_scenarios(vehicles)

        if not comparison.get('scenarios'):
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, "No scenario data available.\nRun analysis first.",
                    ha='center', va='center', fontsize=12)
            ax.axis('off')
            return

        ax = figure.add_subplot(111)

        all_years = comparison['all_years']
        scenario_colors = [COLORS[0], COLORS[1], COLORS[2], COLORS[3]]

        for i, scenario in enumerate(comparison['scenarios']):
            name = scenario['name']
            cumul = scenario['cumulative_vehicles']
            y_vals = [cumul.get(yr, 0) for yr in all_years]
            color = scenario_colors[i % len(scenario_colors)]

            ax.fill_between(all_years, y_vals, alpha=0.15, color=color)
            ax.plot(all_years, y_vals, '-o', color=color, linewidth=2,
                   markersize=4, label=name)

        ax.set_xlabel("Year")
        ax.set_ylabel("Cumulative Vehicles Electrified")
        ax.legend(frameon=True, facecolor='white', edgecolor='#CCCCCC', fontsize=10)

        # Subtitle with best scenario
        best = comparison.get('fastest', '')
        lowest = comparison.get('lowest_cost', '')
        parts = []
        if best:
            parts.append(f"Fastest: {best}")
        if lowest:
            parts.append(f"Lowest cost: {lowest}")
        subtitle = " — ".join(parts) if parts else ""

        apply_presentation_style(figure, ax,
                                 title="Electrification Scenario Comparison",
                                 subtitle=subtitle)

