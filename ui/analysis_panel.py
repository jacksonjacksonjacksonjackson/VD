"""
analysis_panel.py

Panel for analyzing fleet data and visualizing results in the
Fleet Electrification Analyzer.
"""

import os
import logging
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, List, Any, Optional, Callable, Tuple

import matplotlib
matplotlib.use("TkAgg")  # Use TkAgg backend for tkinter compatibility
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from settings import (
    PRIMARY_HEX_1,
    PRIMARY_HEX_2,
    PRIMARY_HEX_3,
    SECONDARY_HEX_1,
    DEFAULT_GAS_PRICE,
    DEFAULT_ELECTRICITY_PRICE,
    DEFAULT_EV_EFFICIENCY,
    CHART_TYPES
)
from utils import SimpleTooltip, ProgressDialog
from data.models import FleetVehicle, Fleet
from ui.theme import Colors, Fonts, Spacing
from analysis.calculations import (
    analyze_fleet_electrification,
    create_emissions_inventory,
    analyze_charging_needs
)
from analysis.charts import ChartFactory
from analysis.reports import ReportGeneratorFactory, ExportCoordinator
from analysis.rate_database import get_rates_for_state, get_all_incentives, get_available_states

# Set up module logger (must be before any try/except blocks that use it)
logger = logging.getLogger(__name__)

# Import PowerPoint export functionality
try:
    from powerpoint_export import export_prelim_deck
    PPTX_EXPORT_AVAILABLE = True
except ImportError:
    PPTX_EXPORT_AVAILABLE = False
    logger.warning("PowerPoint export not available - powerpoint_export module not found")

# Charging power level constants
CHARGING_POWER_LEVELS = {
    "LP": 7.2,    # Low Power (kW)
    "MP": 19.2,   # Medium Power (kW)
    "HP": 50.0,   # High Power (kW)
    "VHP": 150.0  # Very High Power (kW)
}

class AnalysisPanel(ttk.Frame):
    """
    Panel for analyzing fleet data and visualizing results.
    Features analysis options, parameter inputs, and chart display.
    """
    
    def __init__(self, parent, fleet=None, on_analysis_complete=None, on_report_generation=None):
        """
        Initialize the analysis panel.
        
        Args:
            parent: Parent widget
            fleet: Initial fleet data
            on_analysis_complete: Callback when analysis completes
            on_report_generation: Callback when report is generated
        """
        super().__init__(parent)
        
        # Store callbacks
        self.on_analysis_complete_callback = on_analysis_complete
        self.on_report_generation_callback = on_report_generation
        
        # Initialize variables
        self.fleet = fleet or Fleet(name="Empty Fleet")
        self.current_chart_type = tk.StringVar(value=CHART_TYPES[0] if CHART_TYPES else "")
        self.current_figure = None
        self.current_canvas = None
        
        # Analysis parameters
        self.gas_price_var = tk.DoubleVar(value=DEFAULT_GAS_PRICE)
        self.electricity_price_var = tk.DoubleVar(value=DEFAULT_ELECTRICITY_PRICE)
        self.ev_efficiency_var = tk.DoubleVar(value=DEFAULT_EV_EFFICIENCY)
        self.analysis_years_var = tk.IntVar(value=10)
        self.discount_rate_var = tk.DoubleVar(value=5.0)
        
        # Charging parameters
        self.charging_pattern_var = tk.StringVar(value="standard")
        self.charging_start_var = tk.IntVar(value=18)  # 6 PM
        self.charging_end_var = tk.IntVar(value=6)     # 6 AM
        self.power_level_var = tk.StringVar(value="LP")  # Default to Level 2
        
        # Analysis results
        self.electrification_analysis = None
        self.emissions_inventory = None
        self.charging_analysis = None
        
        # Export coordinator
        self.export_coordinator = ExportCoordinator()
        
        # Create UI components
        self._create_ui()
    
    def _create_ui(self):
        """Create the main UI components."""
        # Create paned window for resizable sections
        self.paned_window = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create left panel container (will hold scrollable canvas)
        left_panel_container = ttk.Frame(self.paned_window)
        self.paned_window.add(left_panel_container, weight=30)
        
        # Create scrollable left panel
        self._create_scrollable_left_panel(left_panel_container)
        
        # Create right panel (charts)
        self.right_panel = ttk.Frame(self.paned_window)
        self.paned_window.add(self.right_panel, weight=70)
        
        # Create controls in left panel
        self._create_left_panel()
        
        # Create chart area in right panel
        self._create_right_panel()
    
    def _create_scrollable_left_panel(self, container):
        """Create a scrollable container for the left panel controls."""
        # Create canvas and scrollbar
        self.left_canvas = tk.Canvas(container, bg="#f0f0f0", highlightthickness=0)
        left_scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.left_canvas.yview)
        
        # Create the actual left panel frame inside canvas
        self.left_panel = ttk.Frame(self.left_canvas)
        
        # Configure canvas scrolling
        self.left_panel.bind(
            "<Configure>",
            lambda e: self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))
        )
        
        # Create window in canvas
        self.left_canvas_window = self.left_canvas.create_window((0, 0), window=self.left_panel, anchor="nw")
        self.left_canvas.configure(yscrollcommand=left_scrollbar.set)
        
        # Pack canvas and scrollbar
        self.left_canvas.pack(side="left", fill="both", expand=True)
        left_scrollbar.pack(side="right", fill="y")
        
        # Enable mousewheel scrolling
        def on_mousewheel(event):
            self.left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        def on_enter(event):
            self.left_canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        def on_leave(event):
            self.left_canvas.unbind_all("<MouseWheel>")
        
        # Bind mousewheel only when mouse is over left panel
        self.left_canvas.bind("<Enter>", on_enter)
        self.left_canvas.bind("<Leave>", on_leave)
        
        # Bind canvas width to frame width for proper horizontal sizing
        def on_canvas_configure(event):
            self.left_canvas.itemconfig(self.left_canvas_window, width=event.width)
        
        self.left_canvas.bind("<Configure>", on_canvas_configure)
    
    def _create_left_panel(self):
        """Create analysis controls in the left panel."""
        # Create analysis parameters section
        self._create_parameters_section()
        
        # Create analysis buttons section
        self._create_analysis_buttons()
        
        # Create results summary section
        self._create_results_summary()
        
        # Create export section
        self._create_export_section()
    
    def _create_parameters_section(self):
        """Create the analysis parameters section."""
        # Parameters container using notebook for categorized parameters
        params_notebook = ttk.Notebook(self.left_panel)
        params_notebook.pack(fill=tk.X, pady=(0, 10))
        
        # Cost Parameters Tab
        cost_frame = ttk.LabelFrame(params_notebook, text="Cost Parameters")
        params_notebook.add(cost_frame, text="Costs")

        # State selector — auto-populates gas/electricity prices
        state_frame = ttk.Frame(cost_frame)
        state_frame.pack(fill=tk.X, padx=5, pady=(5, 2))
        state_label = ttk.Label(state_frame, text="State:")
        state_label.pack(side=tk.LEFT)
        self.state_var = tk.StringVar(value="")
        state_options = ["(National Avg)"] + get_available_states()
        state_combo = ttk.Combobox(
            state_frame, textvariable=self.state_var,
            values=state_options, state="readonly", width=12
        )
        state_combo.pack(side=tk.RIGHT)
        state_combo.bind("<<ComboboxSelected>>", self._on_state_selected)
        SimpleTooltip(state_label, "Select state to auto-populate gas and electricity prices")

        # Incentives display label (updated when state changes)
        self.incentive_label = ttk.Label(cost_frame, text="", font=("", 8),
                                         foreground="#555555", wraplength=220)
        self.incentive_label.pack(fill=tk.X, padx=5, pady=(0, 2))

        # Gas price
        gas_frame = ttk.Frame(cost_frame)
        gas_frame.pack(fill=tk.X, padx=5, pady=2)
        gas_label = ttk.Label(gas_frame, text="Gas Price ($/gal):")
        gas_label.pack(side=tk.LEFT)
        gas_entry = ttk.Entry(gas_frame, textvariable=self.gas_price_var, width=8)
        gas_entry.pack(side=tk.RIGHT)
        SimpleTooltip(gas_label, "Current price of gasoline per gallon")
        
        # Electricity price
        elec_frame = ttk.Frame(cost_frame)
        elec_frame.pack(fill=tk.X, padx=5, pady=2)
        elec_label = ttk.Label(elec_frame, text="Electricity ($/kWh):")
        elec_label.pack(side=tk.LEFT)
        elec_entry = ttk.Entry(elec_frame, textvariable=self.electricity_price_var, width=8)
        elec_entry.pack(side=tk.RIGHT)
        SimpleTooltip(elec_label, "Current price of electricity per kilowatt-hour")
        
        # Discount rate
        discount_frame = ttk.Frame(cost_frame)
        discount_frame.pack(fill=tk.X, padx=5, pady=2)
        discount_label = ttk.Label(discount_frame, text="Discount Rate (%):")
        discount_label.pack(side=tk.LEFT)
        discount_entry = ttk.Entry(discount_frame, textvariable=self.discount_rate_var, width=8)
        discount_entry.pack(side=tk.RIGHT)
        SimpleTooltip(discount_label, "Annual discount rate for future cost calculations")

        # Vehicle Parameters Tab
        vehicle_frame = ttk.LabelFrame(params_notebook, text="Vehicle Parameters")
        params_notebook.add(vehicle_frame, text="Vehicle")
        
        # EV efficiency
        eff_frame = ttk.Frame(vehicle_frame)
        eff_frame.pack(fill=tk.X, padx=5, pady=2)
        eff_label = ttk.Label(eff_frame, text="EV Efficiency (kWh/mi):")
        eff_label.pack(side=tk.LEFT)
        eff_entry = ttk.Entry(eff_frame, textvariable=self.ev_efficiency_var, width=8)
        eff_entry.pack(side=tk.RIGHT)
        SimpleTooltip(eff_label, "Average electricity consumption per mile for electric vehicles")
        
        # Analysis period
        years_frame = ttk.Frame(vehicle_frame)
        years_frame.pack(fill=tk.X, padx=5, pady=2)
        years_label = ttk.Label(years_frame, text="Analysis Period (yrs):")
        years_label.pack(side=tk.LEFT)
        years_spinbox = ttk.Spinbox(years_frame, from_=1, to=20, textvariable=self.analysis_years_var, width=5)
        years_spinbox.pack(side=tk.RIGHT)
        SimpleTooltip(years_label, "Number of years to consider in the analysis")

        # Charging Parameters Tab
        charging_frame = ttk.LabelFrame(params_notebook, text="Charging Parameters")
        params_notebook.add(charging_frame, text="Charging")
        
        # Power Level Selection
        power_frame = ttk.LabelFrame(charging_frame, text="Power Levels")
        power_frame.pack(fill=tk.X, padx=5, pady=(5, 2))
        
        # Power level variables and entries
        self.power_levels = {
            "LP": tk.DoubleVar(value=7.2),    # Low Power (kW)
            "MP": tk.DoubleVar(value=19.2),   # Medium Power (kW)
            "HP": tk.DoubleVar(value=50.0),   # High Power (kW)
            "VHP": tk.DoubleVar(value=150.0)  # Very High Power (kW)
        }
        
        # Create grid for power level entries
        power_labels = {
            "LP": "Low Power:",
            "MP": "Medium Power:",
            "HP": "High Power:",
            "VHP": "Very High Power:"
        }
        
        # Configure grid columns
        power_frame.grid_columnconfigure(1, weight=1, pad=10)
        
        # Create labeled entry fields for each power level
        for idx, (level, label) in enumerate(power_labels.items()):
            # Label
            ttk.Label(power_frame, text=label).grid(row=idx, column=0, sticky="e", padx=(10, 5), pady=2)
            
            # Entry field
            entry_frame = ttk.Frame(power_frame)
            entry_frame.grid(row=idx, column=1, sticky="ew", pady=2)
            entry_frame.grid_columnconfigure(0, weight=1)
            
            entry = ttk.Entry(entry_frame, textvariable=self.power_levels[level], width=10, justify="right")
            entry.grid(row=0, column=0, sticky="e", padx=(0, 5))
            
            # kW label
            ttk.Label(entry_frame, text="kW").grid(row=0, column=1, sticky="w")

            # Radio button for selecting active power level
            ttk.Radiobutton(
                entry_frame, text="", variable=self.power_level_var, value=level
            ).grid(row=0, column=2, padx=(5, 0))

        # Active level hint
        ttk.Label(power_frame, text="Select active charging level with radio buttons",
                  font=("", 8)).grid(row=len(power_labels), column=0, columnspan=2,
                                     sticky="w", padx=10, pady=(2, 5))

        # Charging Pattern
        pattern_frame = ttk.Frame(charging_frame)
        pattern_frame.pack(fill=tk.X, padx=5, pady=(10, 2))
        ttk.Label(pattern_frame, text="Charging Pattern:").pack(side=tk.LEFT, padx=(10, 0))
        pattern_combo = ttk.Combobox(
            pattern_frame,
            textvariable=self.charging_pattern_var,
            values=["standard", "extended", "24-hour"],
            state="readonly",
            width=15
        )
        pattern_combo.pack(side=tk.RIGHT, padx=5)
        
        # Charging Window
        window_frame = ttk.Frame(charging_frame)
        window_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(window_frame, text="Charging Window:").pack(side=tk.LEFT, padx=(10, 0))
        
        time_frame = ttk.Frame(window_frame)
        time_frame.pack(side=tk.RIGHT, padx=5)
        
        start_spinbox = ttk.Spinbox(
            time_frame,
            from_=0,
            to=23,
            textvariable=self.charging_start_var,
            width=5,
            justify="right"
        )
        start_spinbox.pack(side=tk.LEFT)
        
        ttk.Label(time_frame, text=" to ").pack(side=tk.LEFT, padx=5)
        
        end_spinbox = ttk.Spinbox(
            time_frame,
            from_=0,
            to=23,
            textvariable=self.charging_end_var,
            width=5,
            justify="right"
        )
        end_spinbox.pack(side=tk.LEFT)
        
        # Add Reset to Defaults button
        def reset_defaults():
            self.gas_price_var.set(DEFAULT_GAS_PRICE)
            self.electricity_price_var.set(DEFAULT_ELECTRICITY_PRICE)
            self.ev_efficiency_var.set(DEFAULT_EV_EFFICIENCY)
            self.analysis_years_var.set(10)
            self.discount_rate_var.set(5.0)
            self.charging_pattern_var.set("standard")
            self.charging_start_var.set(18)
            self.charging_end_var.set(6)

        reset_btn = ttk.Button(
            self.left_panel,
            text="↺ Reset to Defaults",
            command=reset_defaults,
            style="Secondary.TButton"
        )
        reset_btn.pack(fill=tk.X, padx=Spacing.SM, pady=(0, Spacing.MARGIN_ELEMENT))
        SimpleTooltip(reset_btn, "Reset all parameters to default values\nGas: $3.50/gal, Electricity: $0.13/kWh, EV Efficiency: 3.0 mi/kWh")
    
    def _create_analysis_buttons(self):
        """Create analysis action buttons."""
        # Analysis buttons container
        buttons_frame = ttk.LabelFrame(self.left_panel, text="Run Analysis")
        buttons_frame.pack(fill=tk.X, pady=(0, Spacing.MARGIN_ELEMENT))

        # Run Full Analysis button (ACCENT — primary action)
        full_btn = ttk.Button(
            buttons_frame,
            text="Run Full Analysis",
            command=self.run_full_analysis,
            style="Accent.TButton"
        )
        full_btn.pack(fill=tk.X, padx=Spacing.SM, pady=Spacing.SM)
        SimpleTooltip(full_btn, "Run all three analyses in sequence:\nElectrification → Emissions → Charging\nThen display KPI summary and charts")

        # Separator
        ttk.Separator(buttons_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=Spacing.SM, pady=2)

        # Individual analysis buttons (for running one at a time)
        electrify_btn = ttk.Button(
            buttons_frame,
            text="Electrification Analysis",
            command=self.run_electrification_analysis,
            style="Primary.TButton"
        )
        electrify_btn.pack(fill=tk.X, padx=Spacing.SM, pady=(Spacing.SM, 2))
        SimpleTooltip(electrify_btn, "TCO, ROI, and savings analysis")

        emissions_btn = ttk.Button(
            buttons_frame,
            text="Emissions Analysis",
            command=self.run_emissions_analysis,
            style="Secondary.TButton"
        )
        emissions_btn.pack(fill=tk.X, padx=Spacing.SM, pady=2)
        SimpleTooltip(emissions_btn, "CO₂e inventory and reduction potential")

        charging_btn = ttk.Button(
            buttons_frame,
            text="Charging Analysis",
            command=self.run_charging_analysis,
            style="Secondary.TButton"
        )
        charging_btn.pack(fill=tk.X, padx=Spacing.SM, pady=(2, Spacing.SM))
        SimpleTooltip(charging_btn, "Infrastructure requirements and costs")
    
    def _create_results_summary(self):
        """Create the KPI cards results summary section."""
        summary_frame = ttk.LabelFrame(self.left_panel, text="Key Metrics")
        summary_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # KPI cards container — 2 columns
        self.kpi_frame = ttk.Frame(summary_frame)
        self.kpi_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.kpi_frame.grid_columnconfigure(0, weight=1)
        self.kpi_frame.grid_columnconfigure(1, weight=1)

        # Define KPI slots (will be populated by _update_summary)
        self.kpi_labels = {}
        kpi_defs = [
            ("fleet_size", "Fleet Size", "—"),
            ("avg_mpg", "Avg MPG", "—"),
            ("annual_savings", "Annual Savings", "—"),
            ("co2_reduction", "CO₂ Reduction", "—"),
            ("payback", "Avg Payback", "—"),
            ("infra_cost", "Infrastructure", "—"),
        ]

        for idx, (key, title, default) in enumerate(kpi_defs):
            row, col = divmod(idx, 2)
            card = ttk.Frame(self.kpi_frame, relief="solid", borderwidth=1)
            card.grid(row=row, column=col, padx=3, pady=3, sticky="nsew")
            self.kpi_frame.grid_rowconfigure(row, weight=1)

            title_lbl = ttk.Label(card, text=title, font=("", 8),
                                 foreground="#777777")
            title_lbl.pack(anchor="w", padx=6, pady=(4, 0))

            value_lbl = ttk.Label(card, text=default, font=("", 13, "bold"),
                                 foreground=PRIMARY_HEX_1)
            value_lbl.pack(anchor="w", padx=6, pady=(0, 4))

            self.kpi_labels[key] = value_lbl

        # Detailed summary text (collapsed, smaller)
        self.summary_text = tk.Text(
            summary_frame,
            wrap=tk.WORD,
            width=30,
            height=6,
            bg=PRIMARY_HEX_2,
            state=tk.DISABLED,
            font=("", 9),
        )
        self.summary_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))

        summary_scrollbar = ttk.Scrollbar(summary_frame, command=self.summary_text.yview)
        summary_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.summary_text.config(yscrollcommand=summary_scrollbar.set)

        self.summary_text.tag_configure("heading", font=("", 9, "bold"), foreground=PRIMARY_HEX_1)
        self.summary_text.tag_configure("value", font=("", 8, ""), foreground=SECONDARY_HEX_1)
        self.summary_text.tag_configure("separator", foreground=PRIMARY_HEX_3)
    
    def _create_export_section(self):
        """Create the export options section."""
        export_frame = ttk.LabelFrame(self.left_panel, text="Export & Present")
        export_frame.pack(fill=tk.X, pady=(0, Spacing.MARGIN_ELEMENT))

        # Build Presentation button (navigates to Present tab)
        present_btn = ttk.Button(
            export_frame,
            text="Build Presentation",
            command=self._navigate_to_present,
            style="Accent.TButton"
        )
        present_btn.pack(fill=tk.X, padx=Spacing.SM, pady=Spacing.SM)
        SimpleTooltip(present_btn, "Switch to the Present tab to configure\nand generate a PowerPoint deck")

        ttk.Separator(export_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=Spacing.SM, pady=2)

        # Single Export button with format selection
        export_btn = ttk.Button(
            export_frame,
            text="Export",
            command=self._export_menu,
            style="Primary.TButton"
        )
        export_btn.pack(fill=tk.X, padx=Spacing.SM, pady=Spacing.SM)
        SimpleTooltip(export_btn, "Export analysis results:\n• PowerPoint, Excel, or chart image")
    
    def _create_right_panel(self):
        """Create the chart display area in the right panel."""
        # Chart controls container
        controls_frame = ttk.Frame(self.right_panel)
        controls_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Create top controls row
        top_controls = ttk.Frame(controls_frame)
        top_controls.pack(fill=tk.X, pady=(0, 5))
        
        # Chart type label and dropdown
        ttk.Label(top_controls, text="Chart Type:").pack(side=tk.LEFT, padx=(0, 5))
        chart_combo = ttk.Combobox(
            top_controls,
            textvariable=self.current_chart_type,
            values=CHART_TYPES,
            state="readonly",
            width=30
        )
        chart_combo.pack(side=tk.LEFT, padx=(0, 5))
        chart_combo.bind("<<ComboboxSelected>>", lambda e: self._update_chart())
        
        # Create visualization controls row
        viz_controls = ttk.Frame(controls_frame)
        viz_controls.pack(fill=tk.X)
        
        # Chart style
        style_frame = ttk.Frame(viz_controls)
        style_frame.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(style_frame, text="Style:").pack(side=tk.LEFT, padx=(0, 5))
        self.chart_style_var = tk.StringVar(value="default")
        style_combo = ttk.Combobox(
            style_frame,
            textvariable=self.chart_style_var,
            values=["default", "minimal", "dark", "colorful"],
            state="readonly",
            width=10
        )
        style_combo.pack(side=tk.LEFT)
        style_combo.bind("<<ComboboxSelected>>", lambda e: self._update_chart())
        
        # Color scheme
        color_frame = ttk.Frame(viz_controls)
        color_frame.pack(side=tk.LEFT, padx=10)
        
        ttk.Label(color_frame, text="Colors:").pack(side=tk.LEFT, padx=(0, 5))
        self.color_scheme_var = tk.StringVar(value="default")
        color_combo = ttk.Combobox(
            color_frame,
            textvariable=self.color_scheme_var,
            values=["default", "viridis", "magma", "plasma", "inferno"],
            state="readonly",
            width=10
        )
        color_combo.pack(side=tk.LEFT)
        color_combo.bind("<<ComboboxSelected>>", lambda e: self._update_chart())
        
        # Navigation buttons
        nav_frame = ttk.Frame(viz_controls)
        nav_frame.pack(side=tk.RIGHT)
        
        prev_btn = ttk.Button(
            nav_frame,
            text="Previous",
            command=self._previous_chart,
            width=10
        )
        prev_btn.pack(side=tk.LEFT, padx=2)
        
        next_btn = ttk.Button(
            nav_frame,
            text="Next",
            command=self._next_chart,
            width=10
        )
        next_btn.pack(side=tk.LEFT, padx=2)
        
        # Chart display container with zoom controls
        chart_container = ttk.LabelFrame(self.right_panel, text="Chart")
        chart_container.pack(fill=tk.BOTH, expand=True)
        
        # Add toolbar frame
        self.toolbar_frame = ttk.Frame(chart_container)
        self.toolbar_frame.pack(fill=tk.X)
        
        # Create chart frame
        self.chart_frame = ttk.Frame(chart_container)
        self.chart_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create initial empty chart
        self._create_initial_chart()
    
    def _create_initial_chart(self):
        """Create the initial empty chart."""
        # Create figure with tight layout
        self.current_figure = Figure(figsize=(10, 6), dpi=100)
        self.current_figure.set_tight_layout(True)
        
        # Add "No Data" message
        ax = self.current_figure.add_subplot(111)
        ax.text(0.5, 0.5, "No data available. Run an analysis to view charts.", 
                ha='center', va='center', fontsize=12)
        ax.axis('off')
        
        # Create canvas with support for pan/zoom
        self.current_canvas = FigureCanvasTkAgg(self.current_figure, master=self.chart_frame)
        self.current_canvas.draw()
        
        # Add matplotlib toolbar
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        self.toolbar = NavigationToolbar2Tk(self.current_canvas, self.toolbar_frame)
        self.toolbar.update()
        
        # Pack canvas
        canvas_widget = self.current_canvas.get_tk_widget()
        canvas_widget.pack(fill=tk.BOTH, expand=True)
        
        # Create right-click menu
        self.chart_menu = tk.Menu(self, tearoff=0)
        self.chart_menu.add_command(label="Copy Chart", command=self._copy_chart)
        self.chart_menu.add_command(label="Save Chart...", command=self._save_chart)
        
        # Bind right-click event
        canvas_widget.bind("<Button-3>", self._show_chart_menu)
    
    def _show_chart_menu(self, event):
        """Show the right-click menu for the chart."""
        try:
            self.chart_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.chart_menu.grab_release()
    
    def _copy_chart(self):
        """Copy the current chart to clipboard."""
        if not self.current_figure:
            return
            
        # Create a temporary file
        import tempfile
        import os
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
            # Save figure to temp file
            self.current_figure.savefig(
                tmp.name,
                dpi=300,
                bbox_inches='tight',
                facecolor='white',
                edgecolor='none'
            )
            
            # Copy to clipboard using platform-appropriate command
            import platform
            current_platform = platform.system()
            if current_platform == 'Darwin':  # macOS
                subprocess.run([
                    'osascript', '-e',
                    f'set the clipboard to (read (POSIX file "{tmp.name}") as «class PNGf»)'
                ], check=False)
            elif current_platform == 'Windows':
                try:
                    from PIL import Image
                    import io as _io
                    image = Image.open(tmp.name)
                    output = _io.BytesIO()
                    image.convert('RGB').save(output, 'BMP')
                    bmp_data = output.getvalue()[14:]
                    output.close()
                    import win32clipboard
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, bmp_data)
                    win32clipboard.CloseClipboard()
                except ImportError:
                    logger.warning("win32clipboard not available for clipboard copy")
            else:  # Linux
                subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'image/png', '-i', tmp.name], check=False)
        
        # Clean up temp file
        os.unlink(tmp.name)
    
    def _save_chart(self):
        """Save the current chart to a file."""
        if not self.current_figure:
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[
                ("PNG files", "*.png"),
                ("PDF files", "*.pdf"),
                ("SVG files", "*.svg"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.current_figure.savefig(
                file_path,
                dpi=300,
                bbox_inches='tight',
                facecolor='white',
                edgecolor='none'
            )
    
    def _update_chart(self):
        """Update the displayed chart based on the current chart type."""
        # Get current chart type
        chart_type = self.current_chart_type.get()
        
        if not chart_type:
            return
        
        # Check if we have data
        if not self.fleet or not self.fleet.vehicles:
            # Show no data message
            self.current_figure.clear()
            ax = self.current_figure.add_subplot(111)
            ax.text(0.5, 0.5, "No data available. Run an analysis to view charts.", 
                    ha='center', va='center', fontsize=12)
            ax.axis('off')
            self.current_canvas.draw()
            return
        
        # Determine what data to use based on chart type
        chart_data = self.fleet  # Default to fleet data
        extra_args = {}

        # Use specific analysis results for certain chart types
        if chart_type == "Fleet Cash Flow" and self.electrification_analysis:
            chart_data = self.electrification_analysis
        elif chart_type in ("Replacement Priority", "Scenario Comparison"):
            chart_data = self.fleet  # Decision charts use fleet vehicles
        elif chart_type in ("Annual Cost Comparison",):
            chart_data = self.fleet
            extra_args = {
                "gas_price": self.gas_price_var.get(),
                "electricity_price": self.electricity_price_var.get(),
                "ev_efficiency": self.ev_efficiency_var.get()
            }
        elif chart_type in ("Electrification Potential", "Emissions Reduction", "ROI Analysis"):
            if self.electrification_analysis:
                chart_data = self.electrification_analysis
                extra_args = {
                    "gas_price": self.gas_price_var.get(),
                    "electricity_price": self.electricity_price_var.get(),
                    "ev_efficiency": self.ev_efficiency_var.get()
                }
        elif "Emission" in chart_type and self.emissions_inventory:
            chart_data = self.emissions_inventory
        elif "Charging" in chart_type and self.charging_analysis:
            chart_data = self.charging_analysis
        
        # Create chart
        try:
            # Clear existing figure
            self.current_figure.clear()
            
            # Create new chart with selected style and color scheme
            ChartFactory.create_chart(
                chart_type=chart_type,
                data=chart_data,
                figure=self.current_figure,
                chart_style=self.chart_style_var.get(),
                color_scheme=self.color_scheme_var.get(),
                **extra_args
            )
            
            # Update canvas
            self.current_canvas.draw()
            
        except Exception as e:
            logger.error(f"Error creating chart: {e}")
            
            # Show error message
            self.current_figure.clear()
            ax = self.current_figure.add_subplot(111)
            ax.text(0.5, 0.5, f"Error creating chart:\n{str(e)}", 
                    ha='center', va='center', fontsize=10)
            ax.axis('off')
            self.current_canvas.draw()
    
    def _previous_chart(self):
        """Display the previous chart in the list."""
        if not CHART_TYPES:
            return
        
        # Get current index
        try:
            current_idx = CHART_TYPES.index(self.current_chart_type.get())
        except ValueError:
            current_idx = 0
        
        # Calculate previous index
        prev_idx = (current_idx - 1) % len(CHART_TYPES)
        
        # Update chart type
        self.current_chart_type.set(CHART_TYPES[prev_idx])
        
        # Update chart
        self._update_chart()
    
    def _next_chart(self):
        """Display the next chart in the list."""
        if not CHART_TYPES:
            return
        
        # Get current index
        try:
            current_idx = CHART_TYPES.index(self.current_chart_type.get())
        except ValueError:
            current_idx = 0
        
        # Calculate next index
        next_idx = (current_idx + 1) % len(CHART_TYPES)
        
        # Update chart type
        self.current_chart_type.set(CHART_TYPES[next_idx])
        
        # Update chart
        self._update_chart()
    
    def run_electrification_analysis(self):
        """Run electrification analysis on the current fleet."""
        # Check if we have fleet data
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Data", "No fleet data available for analysis.")
            return
        
        # Show progress dialog
        progress = ProgressDialog(
            self.master,
            "Running Electrification Analysis",
            "Analyzing electrification potential..."
        )
        
        # Run analysis in a background thread
        def analysis_task():
            try:
                # Get parameters
                gas_price = self.gas_price_var.get()
                electricity_price = self.electricity_price_var.get()
                ev_efficiency = self.ev_efficiency_var.get()
                analysis_years = self.analysis_years_var.get()
                discount_rate = self.discount_rate_var.get()
                
                # Update progress
                progress.update(20, "Running analysis...")
                
                # Run the analysis
                self.electrification_analysis = analyze_fleet_electrification(
                    fleet=self.fleet,
                    gas_price=gas_price,
                    electricity_price=electricity_price,
                    ev_efficiency=ev_efficiency,
                    analysis_years=analysis_years,
                    discount_rate=discount_rate
                )
                
                # Update progress
                progress.update(80, "Updating display...")
                
                # Update the UI
                self._update_summary()
                
                # Set chart to electrification-specific chart
                electrification_chart = next((c for c in CHART_TYPES if "Electrification" in c), CHART_TYPES[0])
                self.current_chart_type.set(electrification_chart)
                
                # Update chart
                self.master.after(100, self._update_chart)
                
                # Call completion callback
                if self.on_analysis_complete_callback:
                    self.on_analysis_complete_callback("Electrification", self.electrification_analysis)
                
            except Exception as e:
                logger.error(f"Error in electrification analysis: {e}")
                
                # Show error message
                self.master.after(
                    100,
                    lambda: messagebox.showerror(
                        "Analysis Error",
                        f"Error running electrification analysis:\n{str(e)}"
                    )
                )
                
            finally:
                # Close progress dialog
                self.master.after(100, progress.destroy)
        
        # Start analysis thread — UI updates happen via root.after() inside the thread
        threading.Thread(target=analysis_task, daemon=True).start()

    def run_emissions_analysis(self):
        """Run emissions analysis on the current fleet."""
        # Check if we have fleet data
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Data", "No fleet data available for analysis.")
            return
        
        # Show progress dialog
        progress = ProgressDialog(
            self.master,
            "Running Emissions Analysis",
            "Creating emissions inventory..."
        )
        
        # Run analysis in a background thread
        def analysis_task():
            try:
                # Update progress
                progress.update(20, "Calculating emissions...")
                
                # Run the analysis
                self.emissions_inventory = create_emissions_inventory(self.fleet)
                
                # Update progress
                progress.update(80, "Updating display...")
                
                # Update the UI
                self._update_summary()
                
                # Set chart to emissions-specific chart
                emissions_chart = next((c for c in CHART_TYPES if "Emission" in c), CHART_TYPES[0])
                self.current_chart_type.set(emissions_chart)
                
                # Update chart
                self.master.after(100, self._update_chart)
                
                # Call completion callback
                if self.on_analysis_complete_callback:
                    self.on_analysis_complete_callback("Emissions", self.emissions_inventory)
                
            except Exception as e:
                logger.error(f"Error in emissions analysis: {e}")
                
                # Show error message
                self.master.after(
                    100,
                    lambda: messagebox.showerror(
                        "Analysis Error",
                        f"Error running emissions analysis:\n{str(e)}"
                    )
                )
                
            finally:
                # Close progress dialog
                self.master.after(100, progress.destroy)
        
        # Start analysis thread — UI updates happen via root.after() inside the thread
        threading.Thread(target=analysis_task, daemon=True).start()

    def run_charging_analysis(self):
        """Run charging infrastructure analysis on the current fleet."""
        # Check if we have fleet data
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Data", "No fleet data available for analysis.")
            return
        
        # Show progress dialog
        progress = ProgressDialog(
            self.master,
            "Running Charging Analysis",
            "Analyzing charging infrastructure needs..."
        )
        
        # Run analysis in a background thread
        def analysis_task():
            try:
                # Get parameters
                charging_pattern = self.charging_pattern_var.get()
                charging_window = (self.charging_start_var.get(), self.charging_end_var.get())
                power_level = self.power_level_var.get()
                charging_power_kw = self.power_levels[power_level].get()

                # Update progress
                progress.update(20, "Calculating requirements...")

                # Run the analysis — map selected power level to the function's
                # level2/dcfc parameters based on whether it's L2 or DCFC class
                if power_level in ("LP", "MP"):
                    l2_rate = charging_power_kw
                    dcfc_rate = self.power_levels.get("HP", tk.DoubleVar(value=50.0)).get()
                else:
                    l2_rate = self.power_levels.get("MP", tk.DoubleVar(value=19.2)).get()
                    dcfc_rate = charging_power_kw

                self.charging_analysis = analyze_charging_needs(
                    fleet=self.fleet,
                    daily_usage_pattern=charging_pattern,
                    charging_window=charging_window,
                    level2_charging_rate=l2_rate,
                    dcfc_charging_rate=dcfc_rate
                )
                
                # Update progress
                progress.update(80, "Updating display...")
                
                # Update the UI
                self._update_summary()
                
                # Set chart to charging-specific chart
                charging_chart = next((c for c in CHART_TYPES if "Charging" in c), CHART_TYPES[0])
                self.current_chart_type.set(charging_chart)
                
                # Update chart
                self.master.after(100, self._update_chart)
                
                # Call completion callback
                if self.on_analysis_complete_callback:
                    self.on_analysis_complete_callback("Charging", self.charging_analysis)
                
            except Exception as e:
                logger.error(f"Error in charging analysis: {e}")
                
                # Show error message
                self.master.after(
                    100,
                    lambda: messagebox.showerror(
                        "Analysis Error",
                        f"Error running charging analysis:\n{str(e)}"
                    )
                )
                
            finally:
                # Close progress dialog
                self.master.after(100, progress.destroy)
        
        # Start analysis thread — UI updates happen via root.after() inside the thread
        threading.Thread(target=analysis_task, daemon=True).start()

    def run_full_analysis(self):
        """Run all three analyses sequentially: electrification → emissions → charging."""
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Data", "No fleet data available for analysis.")
            return

        progress = ProgressDialog(
            self.master,
            "Running Full Analysis",
            "Running electrification, emissions, and charging analyses..."
        )

        def full_task():
            try:
                # 1. Electrification
                progress.update(10, "Running electrification analysis...")
                self.electrification_analysis = analyze_fleet_electrification(
                    fleet=self.fleet,
                    gas_price=self.gas_price_var.get(),
                    electricity_price=self.electricity_price_var.get(),
                    ev_efficiency=self.ev_efficiency_var.get(),
                    analysis_years=self.analysis_years_var.get(),
                    discount_rate=self.discount_rate_var.get()
                )

                # 2. Emissions
                progress.update(45, "Creating emissions inventory...")
                self.emissions_inventory = create_emissions_inventory(self.fleet)

                # 3. Charging
                progress.update(70, "Analyzing charging needs...")
                power_level = self.power_level_var.get()
                charging_power_kw = self.power_levels[power_level].get()
                if power_level in ("LP", "MP"):
                    l2_rate = charging_power_kw
                    dcfc_rate = self.power_levels.get("HP", tk.DoubleVar(value=50.0)).get()
                else:
                    l2_rate = self.power_levels.get("MP", tk.DoubleVar(value=19.2)).get()
                    dcfc_rate = charging_power_kw

                self.charging_analysis = analyze_charging_needs(
                    fleet=self.fleet,
                    daily_usage_pattern=self.charging_pattern_var.get(),
                    charging_window=(self.charging_start_var.get(), self.charging_end_var.get()),
                    level2_charging_rate=l2_rate,
                    dcfc_charging_rate=dcfc_rate
                )

                progress.update(90, "Updating display...")

                # Update UI on main thread
                self.master.after(0, self._update_summary)
                self.master.after(100, lambda: self.current_chart_type.set("Fleet Cash Flow"))
                self.master.after(200, self._update_chart)

                if self.on_analysis_complete_callback:
                    self.on_analysis_complete_callback("Full", self.electrification_analysis)

            except Exception as e:
                logger.error(f"Error in full analysis: {e}")
                self.master.after(
                    100,
                    lambda: messagebox.showerror("Analysis Error",
                                                 f"Error running full analysis:\n{str(e)}")
                )
            finally:
                self.master.after(100, progress.destroy)

        threading.Thread(target=full_task, daemon=True).start()

    def _navigate_to_present(self):
        """Navigate to the Present tab in the main notebook."""
        try:
            main_window = self.winfo_toplevel()
            if hasattr(main_window, 'notebook'):
                # Present tab is index 3 (Process=0, Results=1, Analysis=2, Present=3)
                main_window.notebook.select(3)
        except Exception as e:
            logger.warning(f"Could not navigate to Present tab: {e}")

    def _on_state_selected(self, event=None):
        """Handle state selection — update gas/electricity prices and show incentives."""
        selected = self.state_var.get()
        if not selected or selected == "(National Avg)":
            self.gas_price_var.set(3.50)
            self.electricity_price_var.set(0.13)
            self.incentive_label.config(text="")
            return

        rates = get_rates_for_state(selected)
        self.gas_price_var.set(rates["gas_price"])
        self.electricity_price_var.set(rates["electricity_price"])

        # Show available incentives
        incentives = get_all_incentives(selected)
        parts = []
        if incentives["max_federal"] > 0:
            parts.append(f"Federal: up to ${incentives['max_federal']:,}")
        if incentives["max_state"] > 0:
            names = [i["name"] for i in incentives["state_incentives"]]
            parts.append(f"State: {', '.join(names)}")

        if parts:
            self.incentive_label.config(text="Incentives: " + " | ".join(parts))
        else:
            self.incentive_label.config(text=f"No state-specific incentives found for {selected}")

    def _export_menu(self):
        """Show export format selection menu."""
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="PowerPoint Presentation", command=self.export_preliminary_deck)
        menu.add_command(label="Excel / CSV Report", command=self.export_full_report)
        menu.add_separator()
        menu.add_command(label="Current Chart (PNG/PDF/SVG)", command=self.export_current_chart)

        # Position menu near the export button
        try:
            menu.tk_popup(self.winfo_pointerx(), self.winfo_pointery())
        finally:
            menu.grab_release()

    def _update_summary(self):
        """Update KPI cards and summary text with current analysis results."""
        vehicle_count = len(self.fleet.vehicles) if self.fleet and self.fleet.vehicles else 0

        # ── KPI Cards ──
        self.kpi_labels["fleet_size"].config(text=str(vehicle_count))

        if vehicle_count > 0:
            avg_mpg = self.fleet.avg_mpg
            self.kpi_labels["avg_mpg"].config(
                text=f"{avg_mpg:.1f}" if avg_mpg > 0 else "—")
        else:
            self.kpi_labels["avg_mpg"].config(text="—")

        if self.electrification_analysis:
            total_savings = self.electrification_analysis.total_savings
            self.kpi_labels["annual_savings"].config(
                text=f"${total_savings:,.0f}/yr" if total_savings > 0 else "—")
            co2 = self.electrification_analysis.co2_savings
            self.kpi_labels["co2_reduction"].config(
                text=f"{co2:,.1f} MT" if co2 > 0 else "—")
            payback = self.electrification_analysis.payback_period
            self.kpi_labels["payback"].config(
                text=f"{payback:.1f} yr" if payback > 0 else "—")
        else:
            self.kpi_labels["annual_savings"].config(text="—")
            self.kpi_labels["co2_reduction"].config(text="—")
            self.kpi_labels["payback"].config(text="—")

        if self.charging_analysis:
            cost = self.charging_analysis.estimated_installation_cost
            self.kpi_labels["infra_cost"].config(
                text=f"${cost:,.0f}" if cost > 0 else "—")
        else:
            self.kpi_labels["infra_cost"].config(text="—")

        # ── Detailed Summary Text ──
        self.summary_text.config(state=tk.NORMAL)
        self.summary_text.delete(1.0, tk.END)

        if self.electrification_analysis:
            ea = self.electrification_analysis
            self.summary_text.insert(tk.END, "Electrification\n", "heading")
            self.summary_text.insert(
                tk.END,
                f"Fuel savings: ${ea.fuel_cost_savings:,.0f}/yr\n"
                f"Total savings: ${ea.total_savings:,.0f}/yr\n"
                f"Payback: {ea.payback_period:.1f} yr\n\n",
                "value")

        if self.emissions_inventory:
            ei = self.emissions_inventory
            self.summary_text.insert(tk.END, "Emissions\n", "heading")
            self.summary_text.insert(
                tk.END,
                f"Total: {ei.total_emissions:.1f} MT CO2e\n",
                "value")
            if ei.by_department:
                top = sorted(ei.by_department.items(), key=lambda x: x[1], reverse=True)[:3]
                for dept, em in top:
                    self.summary_text.insert(tk.END, f"  {dept}: {em:.1f} MT\n", "value")
            self.summary_text.insert(tk.END, "\n")

        if self.charging_analysis:
            ca = self.charging_analysis
            self.summary_text.insert(tk.END, "Charging\n", "heading")
            self.summary_text.insert(
                tk.END,
                f"L2: {ca.level2_chargers_needed}  |  DCFC: {ca.dcfc_chargers_needed}\n"
                f"Peak power: {ca.max_power_required:.0f} kW\n"
                f"Install cost: ${ca.estimated_installation_cost:,.0f}\n",
                "value")

        self.summary_text.config(state=tk.DISABLED)
    
    def export_full_report(self):
        """Export a comprehensive report with all analyses."""
        # Check if we have fleet data
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Data", "No fleet data available for export.")
            return
        
        # Get file path
        filepath = filedialog.asksaveasfilename(
            title="Export Full Report",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Files", "*.xlsx"),
                ("PDF Files", "*.pdf"),
                ("All Files", "*.*")
            ]
        )
        
        if not filepath:
            return
        
        # Show progress dialog
        progress = ProgressDialog(
            self.master,
            "Generating Report",
            "Please wait while the report is being generated..."
        )
        
        # Run export in background thread
        def export_task():
            try:
                # Update progress
                progress.update(25, "Preparing data...")
                
                # Run any missing analyses
                if not self.electrification_analysis:
                    progress.update(40, "Running electrification analysis...")
                    self.electrification_analysis = analyze_fleet_electrification(
                        fleet=self.fleet,
                        gas_price=self.gas_price_var.get(),
                        electricity_price=self.electricity_price_var.get(),
                        ev_efficiency=self.ev_efficiency_var.get(),
                        analysis_years=self.analysis_years_var.get(),
                        discount_rate=self.discount_rate_var.get()
                    )
                
                if not self.emissions_inventory:
                    progress.update(55, "Creating emissions inventory...")
                    self.emissions_inventory = create_emissions_inventory(self.fleet)
                
                if not self.charging_analysis:
                    progress.update(70, "Analyzing charging needs...")
                    self.charging_analysis = analyze_charging_needs(
                        fleet=self.fleet,
                        daily_usage_pattern=self.charging_pattern_var.get(),
                        charging_window=(self.charging_start_var.get(), self.charging_end_var.get())
                    )
                
                # Update progress
                progress.update(85, "Generating report...")
                
                # Generate the report
                generator = ReportGeneratorFactory.create_generator(filepath)
                if generator:
                    success = generator.generate(
                        fleet=self.fleet,
                        analysis=self.electrification_analysis,
                        charging=self.charging_analysis,
                        emissions=self.emissions_inventory
                    )
                    
                    if success:
                        progress.update(100, "Report complete!")
                        
                        # Call callback
                        if self.on_report_generation_callback:
                            self.on_report_generation_callback(filepath)
                        else:
                            # Show success message
                            self.master.after(
                                500,
                                lambda: messagebox.showinfo(
                                    "Export Complete",
                                    f"Report has been exported to:\n{filepath}"
                                )
                            )
                    else:
                        # Show error message
                        self.master.after(
                            500,
                            lambda: messagebox.showerror(
                                "Export Failed",
                                "Failed to generate the report."
                            )
                        )
                else:
                    # Show error message
                    self.master.after(
                        500,
                        lambda: messagebox.showerror(
                            "Export Failed",
                            "Unsupported file format or missing dependencies."
                        )
                    )
                
            except Exception as e:
                logger.error(f"Error exporting report: {e}")
                
                # Show error message
                self.master.after(
                    500,
                    lambda: messagebox.showerror(
                        "Export Failed",
                        f"Error exporting report: {str(e)}"
                    )
                )
                
            finally:
                # Close progress dialog
                self.master.after(100, progress.destroy)
        
        # Start export thread
        threading.Thread(target=export_task, daemon=True).start()
    
    def export_current_chart(self):
        """Export the currently displayed chart."""
        # Check if we have a chart
        if not self.current_figure:
            messagebox.showinfo("No Chart", "No chart available for export.")
            return
        
        # Get file path
        filepath = filedialog.asksaveasfilename(
            title="Export Chart",
            defaultextension=".png",
            filetypes=[
                ("PNG Image", "*.png"),
                ("PDF File", "*.pdf"),
                ("SVG Image", "*.svg"),
                ("All Files", "*.*")
            ]
        )
        
        if not filepath:
            return
        
        try:
            # Save the figure
            self.current_figure.savefig(
                filepath,
                dpi=300,
                bbox_inches="tight"
            )
            
            # Show success message
            messagebox.showinfo(
                "Export Complete",
                f"Chart has been exported to:\n{filepath}"
            )
            
        except Exception as e:
            logger.error(f"Error exporting chart: {e}")
            
            # Show error message
            messagebox.showerror(
                "Export Failed",
                f"Error exporting chart: {str(e)}"
            )
    
    def export_preliminary_deck(self):
        """Export preliminary PowerPoint deck with fleet analysis."""
        # Check if PowerPoint export is available
        if not PPTX_EXPORT_AVAILABLE:
            messagebox.showerror(
                "PowerPoint Export Unavailable",
                "PowerPoint export is not available. Please ensure python-pptx is installed."
            )
            return
        
        # Check if we have fleet data
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Data", "No fleet data available for PowerPoint export.")
            return
        
        # Get file path
        filepath = filedialog.asksaveasfilename(
            title="Export Preliminary Deck",
            defaultextension=".pptx",
            filetypes=[
                ("PowerPoint Files", "*.pptx"),
                ("All Files", "*.*")
            ]
        )
        
        if not filepath:
            return
        
        # Show progress dialog
        progress = ProgressDialog(
            self.master,
            "Generating PowerPoint",
            "Please wait while the presentation is being generated..."
        )
        
        # Run export in background thread
        def export_task():
            try:
                # Update progress
                progress.update(25, "Preparing fleet data...")
                
                # Prepare data for PowerPoint export
                export_data = {
                    'fleet': self.fleet,
                    'client_name': getattr(self.fleet, 'client_name', 'Client'),
                    'stage': 'Preliminary'
                }
                
                # Update progress
                progress.update(50, "Creating presentation...")
                
                # Export the PowerPoint
                result_path = export_prelim_deck(
                    data=export_data,
                    template_path=None,  # Use auto-discovery
                    out_path=filepath
                )
                
                # Update progress
                progress.update(100, "Presentation complete!")
                
                # Show success message with copyable path
                def show_success():
                    success_dialog = tk.Toplevel(self.master)
                    success_dialog.title("Export Complete")
                    success_dialog.geometry("500x200")
                    success_dialog.transient(self.master)
                    success_dialog.grab_set()
                    
                    # Center the dialog
                    success_dialog.geometry("+{}+{}".format(
                        self.master.winfo_rootx() + 50,
                        self.master.winfo_rooty() + 50
                    ))
                    
                    # Success message
                    message_label = ttk.Label(
                        success_dialog,
                        text="PowerPoint presentation has been generated successfully!",
                        font=("", 12)
                    )
                    message_label.pack(pady=20)
                    
                    # Path display with copy functionality
                    path_frame = ttk.Frame(success_dialog)
                    path_frame.pack(fill=tk.X, padx=20, pady=10)
                    
                    path_label = ttk.Label(path_frame, text="Saved to:")
                    path_label.pack(anchor=tk.W)
                    
                    path_text = tk.Text(path_frame, height=2, wrap=tk.WORD)
                    path_text.insert(tk.END, result_path)
                    path_text.config(state=tk.DISABLED)
                    path_text.pack(fill=tk.X, pady=5)
                    
                    # Buttons
                    button_frame = ttk.Frame(success_dialog)
                    button_frame.pack(pady=20)
                    
                    def copy_path():
                        self.master.clipboard_clear()
                        self.master.clipboard_append(result_path)
                        copy_btn.config(text="Copied!")
                        self.master.after(1000, lambda: copy_btn.config(text="Copy Path"))
                    
                    copy_btn = ttk.Button(button_frame, text="Copy Path", command=copy_path)
                    copy_btn.pack(side=tk.LEFT, padx=5)
                    
                    close_btn = ttk.Button(button_frame, text="Close", command=success_dialog.destroy)
                    close_btn.pack(side=tk.LEFT, padx=5)
                
                # Schedule success dialog on main thread
                self.master.after(500, show_success)
                
            except Exception as e:
                logger.error(f"Error exporting PowerPoint: {e}")
                
                # Show error message
                self.master.after(
                    500,
                    lambda: messagebox.showerror(
                        "Export Failed",
                        f"Error exporting PowerPoint presentation:\n{str(e)}"
                    )
                )
                
            finally:
                # Close progress dialog
                self.master.after(100, progress.destroy)
        
        # Start export thread
        threading.Thread(target=export_task, daemon=True).start()
    
    def set_fleet(self, fleet):
        """
        Set the fleet data for analysis.
        
        Args:
            fleet: Fleet object
        """
        self.fleet = fleet
        
        # Reset analysis results
        self.electrification_analysis = None
        self.emissions_inventory = None
        self.charging_analysis = None
        
        # Update the UI
        self._update_summary()
        self._update_chart()
    
    def update_parameters(self, **kwargs):
        """
        Update analysis parameters.
        
        Args:
            **kwargs: Parameters to update
        """
        # Update gas price
        if "gas_price" in kwargs:
            self.gas_price_var.set(kwargs["gas_price"])
        
        # Update electricity price
        if "electricity_price" in kwargs:
            self.electricity_price_var.set(kwargs["electricity_price"])
        
        # Update EV efficiency
        if "ev_efficiency" in kwargs:
            self.ev_efficiency_var.set(kwargs["ev_efficiency"])
        
        # Update analysis years
        if "analysis_years" in kwargs:
            self.analysis_years_var.set(kwargs["analysis_years"])
        
        # Update discount rate
        if "discount_rate" in kwargs:
            self.discount_rate_var.set(kwargs["discount_rate"])
        
        # Update charging pattern
        if "charging_pattern" in kwargs:
            self.charging_pattern_var.set(kwargs["charging_pattern"])
        
        # Update charging window
        if "charging_window" in kwargs and len(kwargs["charging_window"]) == 2:
            self.charging_start_var.set(kwargs["charging_window"][0])
            self.charging_end_var.set(kwargs["charging_window"][1])
    
    def copy_selection(self):
        """Copy the summary text to clipboard."""
        try:
            # Enable text widget for selection
            self.summary_text.config(state=tk.NORMAL)
            
            # Select all text
            self.summary_text.tag_add(tk.SEL, "1.0", tk.END)
            
            # Copy to clipboard
            selected_text = self.summary_text.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.clipboard_clear()
            self.clipboard_append(selected_text)
            
            # Clear selection
            self.summary_text.tag_remove(tk.SEL, "1.0", tk.END)
            
            # Make text read-only again
            self.summary_text.config(state=tk.DISABLED)
            
        except tk.TclError:
            # No selection
            pass
    
    def refresh(self):
        """Refresh the display."""
        # Update the summary
        self._update_summary()
        
        # Update the chart
        self._update_chart()
    
    def on_resize(self):
        """Handle resize events for the panel."""
        # Update the chart if needed
        if self.current_canvas:
            self.current_canvas.draw()