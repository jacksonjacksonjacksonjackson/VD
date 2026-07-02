"""
main_window.py

Main window for the Fleet Electrification Analyzer application.
Implements the primary UI interface with panels for processing,
results viewing, and analysis.
"""

import datetime
import os
import time
import logging
import threading
from typing import Dict, List, Any, Optional, Callable, Tuple
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Menu

# Import drag-and-drop support
try:
    from tkinterdnd2 import TkinterDnD
    DRAG_DROP_AVAILABLE = True
except ImportError:
    DRAG_DROP_AVAILABLE = False

# Import settings and utilities
from settings import (
    APP_NAME,
    APP_VERSION,
    PRIMARY_HEX_1,
    PRIMARY_HEX_2,
    PRIMARY_HEX_3,
    SECONDARY_HEX_1,
    SECONDARY_HEX_2,
    DEFAULT_VISIBLE_COLUMNS,
    DEFAULT_DB_FILE,
    USER_SETTINGS,
    save_user_settings
)
from utils import StatusBar, SimpleTooltip, ProgressDialog, SafeDict, timestamp, ScrollableFrame

# Import UI panels
from ui.process_panel import ProcessPanel
from ui.results_panel import ResultsPanel
from ui.analysis_panel import AnalysisPanel
from ui.timeline_panel import TimelinePanel
from ui.present_panel import PresentPanel
from ui.charging_panel import ChargingPanel
from ui.database_panel import DatabasePanel

# Import data and analysis modules
from data.models import FleetVehicle, Fleet
from data.processor import BatchProcessor

# Set up module logger
logger = logging.getLogger(__name__)

###############################################################################
# Main Application Window
###############################################################################

class MainWindow:
    """
    Main application window for the Fleet Electrification Analyzer.
    Coordinates between different panels and manages the overall UI.
    """
    
    def __init__(self, root: tk.Tk):
        """
        Initialize the main window.
        
        Args:
            root: Root Tkinter window (may be TkinterDnD.Tk for drag-and-drop support)
        """
        self.root = root
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Set up window icon if available
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "../resources/icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            logger.warning(f"Could not load application icon: {e}")
        
        # Initialize shared data
        self.fleet = Fleet(name="New Fleet")
        self.results_data = []
        self.sharing_data = SafeDict()  # Thread-safe dictionary for sharing data between panels

        # Initialize vehicle reference database.
        # A single instance is shared across DatabasePanel and ResultsPanel so
        # analysts can manage entries and save from the results view.
        # The ProcessingPipeline opens its own separate connection (read-only
        # during runs) — two SQLite WAL-mode connections coexist safely.
        try:
            from data.vehicle_database import VehicleDatabaseManager
            self.db_manager = VehicleDatabaseManager(DEFAULT_DB_FILE)
            logger.info("Vehicle reference database initialized")
        except Exception as _db_err:
            logger.warning(f"Vehicle reference database unavailable: {_db_err}")
            self.db_manager = None

        # Create and configure styles
        self._create_styles()
        
        # Create main layout
        self._create_main_layout()
        
        # Create menu bar
        self._create_menu_bar()
        
        # Create status bar
        self.status_bar = StatusBar(self.root)
        self.status_bar.add_section("status")
        self.status_bar.add_section("vehicle_count", width=15)
        self.status_bar.add_section("version", width=15)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Set initial status
        self.status_bar.set("Ready - Drag & Drop CSV files or use Browse")
        self.status_bar.set(f"Version {APP_VERSION}", section="version")
        self.status_bar.set("No vehicles loaded", section="vehicle_count")
        
        # Create panels
        self._create_panels()
        
        # Initialize data processor
        self.processor = BatchProcessor()
        
        # Initialize processing state
        self.processing = False
        
        # Bind events
        self._bind_events()
        
        # Log startup
        drag_status = "with" if DRAG_DROP_AVAILABLE else "without"
        logger.info(f"{APP_NAME} v{APP_VERSION} started {drag_status} drag-and-drop support")
    
    def _create_styles(self):
        """Create custom styles for widgets."""
        style = ttk.Style()
        
        # Configure tab appearance
        style.configure(
            "TNotebook", 
            background=PRIMARY_HEX_2,
            tabmargins=[2, 5, 2, 0]
        )
        
        style.configure(
            "TNotebook.Tab",
            background="#E0E0E0",     # Light gray background
            foreground="#000000",     # Black text
            padding=[10, 5],
            font=("Segoe UI", 10, "bold")
        )
        
        style.map(
            "TNotebook.Tab",
            background=[
                ("selected", "#E0E0E0"),      # Keep light gray when selected
                ("!selected", "#CCCCCC")      # Slightly darker gray when not selected
            ],
            foreground=[
                ("selected", "#000000"),      # Black text when selected
                ("!selected", "#000000")      # Black text when not selected
            ],
            expand=[("selected", [1, 1, 1, 0])]
        )
        
        # Configure frame appearance
        style.configure(
            "Main.TFrame",
            background=PRIMARY_HEX_2
        )
    
    def _create_main_layout(self):
        """Create the main application layout."""
        # Main frame
        self.main_frame = ttk.Frame(self.root, style="Main.TFrame")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Main notebook (tabbed interface)
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    def _create_menu_bar(self):
        """Create the application menu bar."""
        menubar = Menu(self.root)
        
        # File menu
        file_menu = Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Fleet", command=self.new_fleet)
        file_menu.add_command(label="Open Input File...", command=self.open_input_file)
        file_menu.add_command(label="Open Project...", command=self.open_project)
        file_menu.add_separator()
        file_menu.add_command(label="Save Project...", command=self.save_project_file)
        file_menu.add_command(label="Save Results...", command=self.save_results)
        file_menu.add_command(label="Export Report...", command=self.export_report)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_closing)
        menubar.add_cascade(label="File", menu=file_menu)
        self._file_menu = file_menu
        
        # Edit menu
        edit_menu = Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Copy", command=self.copy_selection)
        edit_menu.add_separator()
        edit_menu.add_command(label="Preferences...", command=self.show_preferences)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        
        # View menu
        view_menu = Menu(menubar, tearoff=0)
        view_menu.add_command(label="Customize Columns...", command=self.customize_columns)
        view_menu.add_separator()
        view_menu.add_command(label="Refresh", command=self.refresh_view)
        menubar.add_cascade(label="View", menu=view_menu)
        
        # Tools menu
        tools_menu = Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Run Full Analysis",
                               command=self.run_full_analysis)
        tools_menu.add_separator()
        # Individual analysis sub-menu for power users
        individual_menu = Menu(tools_menu, tearoff=0)
        individual_menu.add_command(label="Emissions Only",
                                    command=self.analyze_emissions)
        individual_menu.add_command(label="Electrification Only",
                                    command=self.analyze_electrification)
        # Charging analysis hidden until engine is implemented
        # individual_menu.add_command(label="Charging Infrastructure Only",
        #                             command=self.analyze_charging)
        tools_menu.add_cascade(label="Individual Analyses ▸", menu=individual_menu)
        tools_menu.add_separator()
        tools_menu.add_command(label="Reset All EV Year Overrides",
                               command=self._reset_all_ev_overrides)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        
        # Help menu
        help_menu = Menu(menubar, tearoff=0)
        help_menu.add_command(label="Documentation", command=self.show_documentation)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        # Set the menu
        self.root.config(menu=menubar)
    
    def _create_panels(self):
        """Create the main application panels."""
        # Process panel
        self.process_frame = ttk.Frame(self.notebook)
        self.process_panel = ProcessPanel(
            self.process_frame,
            on_process=self.start_processing,
            on_stop=self.stop_processing,
            on_log=self.update_status
        )
        self.process_panel.pack(fill=tk.BOTH, expand=True)
        self.notebook.add(self.process_frame, text="Process")
        
        # Results panel
        self.results_frame = ttk.Frame(self.notebook)
        self.results_panel = ResultsPanel(
            self.results_frame,
            data=self.fleet.vehicles,
            visible_columns=DEFAULT_VISIBLE_COLUMNS,
            on_selection_change=self.on_result_selection_change,
            on_column_change=self.on_column_change,
            db_manager=self.db_manager,
            on_acf_override=self._on_acf_override,
        )
        self.results_panel.pack(fill=tk.BOTH, expand=True)
        self.notebook.add(self.results_frame, text="Results")
        
        # Analysis panel
        self.analysis_frame = ttk.Frame(self.notebook)
        self.analysis_panel = AnalysisPanel(
            self.analysis_frame,
            fleet=self.fleet,
            on_analysis_complete=self.on_analysis_complete,
            on_report_generation=self.on_report_generated,
            sharing_data=self.sharing_data,
        )
        self.analysis_panel.pack(fill=tk.BOTH, expand=True)
        self.notebook.add(self.analysis_frame, text="Analysis")

        # Timeline panel (tab 3 — inserted before Present)
        self.timeline_frame = ttk.Frame(self.notebook)
        self.timeline_panel = TimelinePanel(
            self.timeline_frame,
            fleet=self.fleet,
            on_year_changed=self._on_timeline_year_changed,
            on_acf_changed=self._on_acf_override,
        )
        self.timeline_panel.pack(fill=tk.BOTH, expand=True)
        self.notebook.add(self.timeline_frame, text="Timeline")

        # Present panel
        self.present_frame = ttk.Frame(self.notebook)
        self.present_panel = PresentPanel(
            self.present_frame,
            sharing_data=self.sharing_data
        )
        self.present_panel.get_panel_frame().pack(fill=tk.BOTH, expand=True)
        self.notebook.add(self.present_frame, text="Present")

        # Charging panel (tab 5)
        self.charging_frame = ttk.Frame(self.notebook)
        self.charging_panel = ChargingPanel(
            self.charging_frame,
            sharing_data=self.sharing_data,
            analysis_vars=self.analysis_panel.get_charging_vars(),
            on_run_analysis=self._run_charging_from_panel,
        )
        self.charging_panel.get_panel_frame().pack(fill=tk.BOTH, expand=True)
        self.notebook.add(self.charging_frame, text="Charging")
        self.notebook.hide(self.charging_frame)  # Hidden until analysis engine implemented

        # Database panel (tab 6)
        self.database_frame = ttk.Frame(self.notebook)
        self.database_panel = DatabasePanel(
            self.database_frame,
            db_manager=self.db_manager,
            root=self.root,
            status_bar=self.status_bar
        )
        self.database_panel.pack(fill=tk.BOTH, expand=True)
        self.notebook.add(self.database_frame, text="Database")

    def _bind_events(self):
        """Bind events for the main window."""
        # Notebook tab change
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        
        # Window resize
        self.root.bind("<Configure>", self.on_window_resize)
    
    def on_closing(self):
        """Handle application closing."""
        # Check if processing is in progress
        if self.processing:
            if not messagebox.askyesno("Confirm Exit", 
                                       "Processing is in progress. Are you sure you want to exit?"):
                return
            
            # Stop processing
            self.stop_processing()
        
        # Save user settings
        try:
            # Update settings from current state
            USER_SETTINGS["window_size"] = f"{self.root.winfo_width()}x{self.root.winfo_height()}"
            USER_SETTINGS["visible_columns"] = self.results_panel.get_visible_columns()
            
            # Save to file
            save_user_settings(USER_SETTINGS)
        except Exception as e:
            logger.error(f"Error saving user settings: {e}")
        
        # Close vehicle reference database connection
        if self.db_manager:
            try:
                self.db_manager.close()
            except Exception:
                pass

        # Close the application
        self.root.destroy()
    
    def on_tab_changed(self, event):
        """Handle notebook tab changed event."""
        selected = self.notebook.select()

        # Update status based on current tab (widget-based, resilient to hidden tabs)
        if selected == str(self.process_frame):
            self.status_bar.set("Process vehicles by VIN")
        elif selected == str(self.results_frame):
            self.status_bar.set("View and filter processed vehicles")
        elif selected == str(self.analysis_frame):
            self.status_bar.set("Analyze fleet electrification potential")
            # Keep the Gantt current if years were edited in the Timeline tab
            try:
                self.analysis_panel._update_gantt_section()
            except Exception:
                pass
        elif selected == str(self.timeline_frame):
            self.status_bar.set("Edit vehicle electrification years and view Gantt timeline")
        elif selected == str(self.present_frame):
            self.status_bar.set("Configure and export a PowerPoint presentation")
        elif selected == str(self.charging_frame):
            self.status_bar.set("Configure charging parameters and view infrastructure analysis")
            self.charging_panel.refresh_data()
        elif selected == str(self.database_frame):
            self.status_bar.set("Browse and manage the vehicle MPG reference database")
            self.database_panel.refresh()
    
    def on_window_resize(self, event):
        """Handle window resize event."""
        # Only respond to root window resizes, not child elements
        if event.widget == self.root:
            # Update panels that need resize handling
            self.results_panel.on_resize()
            self.analysis_panel.on_resize()
    
    def on_result_selection_change(self, selected_vehicles):
        """
        Handle selection change in results panel.
        
        Args:
            selected_vehicles: List of selected FleetVehicle objects
        """
        # Update shared data
        self.sharing_data.set("selected_vehicles", selected_vehicles)
        
        # Update status bar
        if selected_vehicles:
            self.status_bar.set(f"{len(selected_vehicles)} vehicles selected")
        else:
            self.status_bar.set("No vehicles selected")
    
    def on_column_change(self, visible_columns):
        """
        Handle column visibility change in results panel.
        
        Args:
            visible_columns: List of visible column IDs
        """
        # Update user settings
        USER_SETTINGS["visible_columns"] = visible_columns
        
        # No need to save here, will be saved on application close
    
    def _on_timeline_year_changed(self):
        """Called when the Timeline panel applies or resets a year override.

        Refreshes the Analysis tab's Gantt chart so both panels stay in sync.
        """
        try:
            self.analysis_panel._update_gantt_section()
        except Exception as exc:
            logger.warning(f"Could not refresh Analysis Gantt after year change: {exc}")

    def _on_acf_override(self, vehicle=None):
        """Called when an ACF category override is applied or reset (from either
        the Timeline tab or the Results tab right-click menu).

        Refreshes the Timeline panel and Analysis Gantt so all views reflect
        the new classification.  The ``vehicle`` argument is accepted for
        compatibility with the Results panel callback signature but is not used
        — the full fleet refresh is always performed.
        """
        try:
            self.timeline_panel.notify_year_changed()
        except Exception as exc:
            logger.warning(f"Could not refresh Timeline after ACF override: {exc}")
        try:
            self.analysis_panel._update_purchase_schedule_chart()
            self.analysis_panel._update_gantt_section()
            self.analysis_panel._update_acf_donut()
        except Exception as exc:
            logger.warning(f"Could not refresh Analysis charts after ACF override: {exc}")

    def _reset_all_ev_overrides(self):
        """Tools menu: reset all manual EV-year overrides in the current fleet."""
        if not self.fleet or not self.fleet.vehicles:
            messagebox.showinfo("No Fleet", "No fleet is currently loaded.")
            return
        overridden = sum(
            1 for v in self.fleet.vehicles
            if v.custom_fields.get("EV Year Overridden") == "Yes"
        )
        if overridden == 0:
            messagebox.showinfo("No Overrides",
                                "No manual EV year overrides are currently active.")
            return
        if messagebox.askyesno(
            "Reset All Overrides",
            f"Reset {overridden} manual EV year "
            f"override{'s' if overridden != 1 else ''}?\n\n"
            "All vehicles will revert to their system-recommended EV years.",
        ):
            self.analysis_panel._reset_all_overrides()
            self.timeline_panel.notify_year_changed()

    def on_analysis_complete(self, analysis_type, results):
        """
        Handle analysis completion.
        
        Args:
            analysis_type: Type of analysis completed
            results: Analysis results
        """
        # Update status
        self.status_bar.set(f"{analysis_type} analysis complete")

        # Store results in shared data
        self.sharing_data.set(f"{analysis_type.lower()}_results", results)

        # Sync fleet_type to the results panel so ACF override → recalculate
        # uses the same deadline table selected in the Analysis tab Fleet Settings.
        self.results_panel._fleet_type = self.fleet.fleet_type

        # Forward charging analysis results to the Charging tab
        if analysis_type == "Charging":
            try:
                self.charging_panel.show_results(results)
            except Exception:
                pass

    def _run_charging_from_panel(self):
        """Trigger charging analysis from the Charging tab's Run button."""
        try:
            self.analysis_panel.run_charging_analysis()
        except Exception as exc:
            logger.error(f"Error starting charging analysis: {exc}")

    def on_report_generated(self, report_path):
        """
        Handle report generation completion.

        Args:
            report_path: Path to the generated report
        """
        # Show success message
        messagebox.showinfo(
            "Report Generated",
            f"Report has been generated successfully.\n\nPath: {report_path}"
        )
    
    def update_status(self, message, section="status"):
        """
        Update status bar message.
        
        Args:
            message: Status message
            section: Status bar section to update
        """
        self.status_bar.set(message, section=section)
    
    def update_vehicle_count(self):
        """Update the vehicle count in the status bar."""
        count = len(self.fleet.vehicles)
        self.status_bar.set(f"{count} vehicles", section="vehicle_count")
    
    def new_fleet(self):
        """Create a new empty fleet."""
        # Confirm if there are existing vehicles
        if self.fleet.vehicles:
            if not messagebox.askyesno("New Fleet", 
                                       "This will clear all current vehicles. Continue?"):
                return
        
        # Reset fleet
        self.fleet = Fleet(name="New Fleet")
        self.results_data = []
        
        # Update UI
        self.results_panel.set_data(self.fleet.vehicles)
        self.analysis_panel.set_fleet(self.fleet)
        self.timeline_panel.set_fleet(
            self.fleet,
            on_year_changed=self._on_timeline_year_changed,
            on_acf_changed=self._on_acf_override,
        )
        self.sharing_data.set("fleet", self.fleet)  # B1: keep Present panel in sync
        self.present_panel.refresh_data()
        self.update_vehicle_count()

        # Update status
        self.status_bar.set("New fleet created")
    
    def open_input_file(self):
        """Open an input file dialog."""
        filepath = filedialog.askopenfilename(
            title="Open Input File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        
        if filepath:
            self.load_input_file(filepath)
    
    def load_input_file(self, filepath):
        """
        Load an input file into the process panel.
        
        Args:
            filepath: Path to the input file
        """
        # Set the input file in the process panel
        self.process_panel.set_input_file(filepath)
        
        # Switch to the process tab
        self.notebook.select(0)
    
    def save_results(self):
        """Save results to a file."""
        # Check if there are results to save
        if not self.fleet.vehicles:
            messagebox.showinfo("No Data", "There are no results to save.")
            return
        
        # Get file path
        filepath = filedialog.asksaveasfilename(
            title="Save Results",
            defaultextension=".csv",
            filetypes=[
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx"),
                ("All Files", "*.*")
            ]
        )
        
        if not filepath:
            return
        
        # Determine export format based on extension
        _, ext = os.path.splitext(filepath.lower())
        
        # Get visible columns from results panel
        visible_columns = self.results_panel.get_visible_columns()
        
        # Export the data
        from analysis.reports import ReportGeneratorFactory
        
        generator = ReportGeneratorFactory.create_generator(filepath)
        if generator:
            success = generator.generate(
                fleet=self.fleet,
                fields=visible_columns
            )
            
            if success:
                messagebox.showinfo(
                    "Export Complete",
                    f"Results have been saved to:\n{filepath}"
                )
            else:
                messagebox.showerror(
                    "Export Failed",
                    "An error occurred while saving the results."
                )
        else:
            messagebox.showerror(
                "Export Failed",
                f"Unsupported file format: {ext}"
            )
    
    def export_report(self):
        """Export comprehensive report — delegates to the Analysis panel.

        Routing all exports through analysis_panel.export_full_report() ensures
        the 'Timelines to Include' dialog, override flagging, and 8-tab structure
        are always applied consistently, regardless of which menu or button the
        user clicked.
        """
        self.analysis_panel.export_full_report()

    # ------------------------------------------------------------------
    # Project save / load (.fea)
    # ------------------------------------------------------------------

    def save_project_file(self):
        """Serialize the current Fleet + scenario results to a .fea file."""
        if not self.fleet.vehicles:
            messagebox.showinfo("No Data", "There are no vehicles to save.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Save Project",
            defaultextension=".fea",
            filetypes=[("Fleet Electrification Analyzer Project", "*.fea"),
                       ("All Files", "*.*")],
        )
        if not filepath:
            return

        try:
            from data.project_io import save_project
            scenario_results = None
            try:
                scenario_results = self.sharing_data.get("scenario_results")
            except Exception:
                pass
            save_project(filepath, self.fleet, scenario_results)
            messagebox.showinfo(
                "Project Saved",
                f"Project saved to:\n{filepath}",
            )
            self.status_bar.set(f"Project saved — {len(self.fleet.vehicles)} vehicles")
        except Exception as exc:
            logger.error("save_project_file error: %s", exc, exc_info=True)
            messagebox.showerror("Save Failed", f"Could not save project:\n{exc}")

    def open_project(self):
        """Load a .fea project file, bypassing the VIN-processing pipeline."""
        if self.fleet.vehicles:
            if not messagebox.askyesno(
                "Open Project",
                "This will replace the current fleet. Continue?",
            ):
                return

        filepath = filedialog.askopenfilename(
            title="Open Project",
            filetypes=[("Fleet Electrification Analyzer Project", "*.fea"),
                       ("All Files", "*.*")],
        )
        if not filepath:
            return

        try:
            from data.project_io import load_project
            fleet, scenario_results = load_project(filepath)
        except Exception as exc:
            logger.error("open_project error: %s", exc, exc_info=True)
            messagebox.showerror("Open Failed", f"Could not load project:\n{exc}")
            return

        # Install the fleet
        self.fleet = fleet
        self.results_data = list(fleet.vehicles)

        self.results_panel.set_data(fleet.vehicles)
        self.analysis_panel.set_fleet(fleet)
        self.timeline_panel.set_fleet(
            fleet,
            on_year_changed=self._on_timeline_year_changed,
            on_acf_changed=self._on_acf_override,
        )
        self.sharing_data.set("fleet", fleet)
        self.present_panel.refresh_data()

        # Restore scenario results if they were saved
        if scenario_results is not None:
            try:
                self.sharing_data.set("scenario_results", scenario_results)
                self.analysis_panel.scenario_results = scenario_results
            except Exception as exc:
                logger.warning("Could not restore scenario results: %s", exc)

        self.update_vehicle_count()
        self.notebook.select(1)  # jump to Results tab
        self.status_bar.set(
            f"Project loaded — {len(fleet.vehicles)} vehicles "
            f"(fleet type: {fleet.fleet_type})"
        )

    def copy_selection(self):
        """Copy the current selection to the clipboard."""
        current_tab = self.notebook.index(self.notebook.select())

        if current_tab == 0:    # Process tab
            self.process_panel.copy_selection()
        elif current_tab == 1:  # Results tab
            self.results_panel.copy_selection()
        elif current_tab == 2:  # Analysis tab
            self.analysis_panel.copy_selection()
        elif current_tab == 3:  # Timeline tab
            self.timeline_panel.copy_selection()
    
    def show_preferences(self):
        """Show the preferences dialog."""
        # Create a simple preferences dialog
        prefs_dialog = tk.Toplevel(self.root)
        prefs_dialog.title("Preferences")
        prefs_dialog.geometry("400x300")
        prefs_dialog.transient(self.root)  # Set as transient to main window
        prefs_dialog.grab_set()  # Make modal
        
        # Add a notebook for preference categories
        prefs_notebook = ttk.Notebook(prefs_dialog)
        prefs_notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # General preferences
        general_frame = ttk.Frame(prefs_notebook)
        prefs_notebook.add(general_frame, text="General")
        
        # Populate with some basic preferences
        ttk.Label(general_frame, text="Default Gas Price ($/gal):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        gas_price_var = tk.DoubleVar(value=USER_SETTINGS.get("gas_price", 3.50))
        ttk.Entry(general_frame, textvariable=gas_price_var, width=10).grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(general_frame, text="Default Electricity Price ($/kWh):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        elec_price_var = tk.DoubleVar(value=USER_SETTINGS.get("electricity_price", 0.13))
        ttk.Entry(general_frame, textvariable=elec_price_var, width=10).grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(general_frame, text="Maximum Worker Threads:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        threads_var = tk.IntVar(value=USER_SETTINGS.get("max_threads", 10))
        ttk.Spinbox(general_frame, from_=1, to=32, textvariable=threads_var, width=5).grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        # Appearance preferences
        appearance_frame = ttk.Frame(prefs_notebook)
        prefs_notebook.add(appearance_frame, text="Appearance")

        ttk.Label(
            appearance_frame,
            text="Appearance settings are not yet configurable.\n\n"
                 "The application uses the system default theme.\n"
                 "Future options may include font size and color themes.",
            justify=tk.LEFT,
            wraplength=350
        ).pack(padx=15, pady=15, anchor="nw")

        # Add buttons
        button_frame = ttk.Frame(prefs_dialog)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        
        ttk.Button(
            button_frame, 
            text="Cancel", 
            command=prefs_dialog.destroy
        ).pack(side=tk.RIGHT, padx=5)
        
        # Save preferences
        def save_preferences():
            # Update user settings
            USER_SETTINGS["gas_price"] = gas_price_var.get()
            USER_SETTINGS["electricity_price"] = elec_price_var.get()
            USER_SETTINGS["max_threads"] = threads_var.get()
            
            # Save settings
            save_user_settings(USER_SETTINGS)
            
            # Update UI components
            self.process_panel.set_max_threads(threads_var.get())
            self.analysis_panel.update_parameters(
                gas_price=gas_price_var.get(),
                electricity_price=elec_price_var.get()
            )
            
            # Close dialog
            prefs_dialog.destroy()
        
        ttk.Button(
            button_frame, 
            text="Save", 
            command=save_preferences
        ).pack(side=tk.RIGHT, padx=5)
    
    def customize_columns(self):
        """Show the column customization dialog."""
        from settings import DEFAULT_VISIBLE_COLUMNS

        # Create dialog
        columns_dialog = tk.Toplevel(self.root)
        columns_dialog.title("Customize Columns")
        columns_dialog.geometry("540x520")
        columns_dialog.minsize(420, 380)
        columns_dialog.resizable(True, True)
        columns_dialog.transient(self.root)
        columns_dialog.grab_set()

        # Get current column configuration
        all_columns = self.results_panel.get_all_columns()
        visible_columns = self.results_panel.get_visible_columns()

        # ── Header row: label + selected count ──────────────────────────────
        header_frame = ttk.Frame(columns_dialog)
        header_frame.pack(fill=tk.X, padx=12, pady=(10, 4))

        ttk.Label(
            header_frame,
            text="Select columns to display:",
            font=("", 10, "bold")
        ).pack(side=tk.LEFT)

        count_var = tk.StringVar()
        count_label = ttk.Label(header_frame, textvariable=count_var, foreground="gray")
        count_label.pack(side=tk.RIGHT)

        # ── Scrollable 2-column checkbox grid ───────────────────────────────
        sf_cols = ScrollableFrame(columns_dialog)
        sf_cols.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)
        inner_frame = sf_cols.scrollable_frame

        # Make both checkbox columns expand evenly
        inner_frame.columnconfigure(0, weight=1)
        inner_frame.columnconfigure(1, weight=1)

        checkbox_vars = {}

        def update_count():
            n = sum(1 for v in checkbox_vars.values() if v.get())
            count_var.set(f"{n} of {len(checkbox_vars)} selected")

        for i, (col_id, col_name) in enumerate(all_columns):
            var = tk.BooleanVar(value=col_id in visible_columns)
            var.trace_add("write", lambda *_: update_count())
            checkbox_vars[col_id] = var

            row, col = divmod(i, 2)
            ttk.Checkbutton(
                inner_frame,
                text=col_name,
                variable=var
            ).grid(row=row, column=col, sticky=tk.W, padx=8, pady=2)

        update_count()

        # ── Separator ────────────────────────────────────────────────────────
        ttk.Separator(columns_dialog, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=12, pady=(4, 0))

        # ── Button row ───────────────────────────────────────────────────────
        button_frame = ttk.Frame(columns_dialog)
        button_frame.pack(fill=tk.X, padx=12, pady=10)

        # Left side: bulk-selection helpers
        def select_all(select=True):
            for var in checkbox_vars.values():
                var.set(select)

        def reset_to_defaults():
            default_set = set(DEFAULT_VISIBLE_COLUMNS)
            for col_id, var in checkbox_vars.items():
                var.set(col_id in default_set)

        ttk.Button(button_frame, text="Select All",
                   command=lambda: select_all(True)).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(button_frame, text="Select None",
                   command=lambda: select_all(False)).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(button_frame, text="Reset to Defaults",
                   command=reset_to_defaults).pack(side=tk.LEFT)

        # Right side: Cancel / Apply
        def apply_columns():
            selected = [col_id for col_id, var in checkbox_vars.items() if var.get()]
            self.results_panel.set_visible_columns(selected)
            columns_dialog.destroy()

        ttk.Button(button_frame, text="Cancel",
                   command=columns_dialog.destroy).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(button_frame, text="Apply",
                   command=apply_columns).pack(side=tk.RIGHT, padx=(4, 0))
    
    def refresh_view(self):
        """Refresh the current view."""
        selected = self.notebook.select()

        if selected == str(self.process_frame):
            self.process_panel.refresh()
        elif selected == str(self.results_frame):
            self.results_panel.refresh()
        elif selected == str(self.analysis_frame):
            self.analysis_panel.refresh()
        elif selected == str(self.timeline_frame):
            self.timeline_panel.notify_year_changed()
        elif selected == str(self.present_frame):
            self.present_panel.refresh_data()
        elif selected == str(self.charging_frame):
            self.charging_panel.refresh_data()
        elif selected == str(self.database_frame):
            self.database_panel.refresh()
    
    def run_full_analysis(self):
        """Run all analyses — mirrors the Analysis tab's Run Full Analysis button."""
        self.notebook.select(2)
        self.analysis_panel.run_full_analysis()

    def analyze_emissions(self):
        """Run emissions analysis."""
        # Switch to analysis tab
        self.notebook.select(2)

        # Run emissions analysis
        self.analysis_panel.run_emissions_analysis()
    
    def analyze_electrification(self):
        """Run electrification analysis."""
        # Switch to analysis tab
        self.notebook.select(2)
        
        # Run electrification analysis
        self.analysis_panel.run_electrification_analysis()
    
    def analyze_charging(self):
        """Run charging infrastructure analysis."""
        # Switch to analysis tab
        self.notebook.select(2)
        
        # Run charging analysis
        self.analysis_panel.run_charging_analysis()
    
    def show_documentation(self):
        """Show documentation."""
        # Simple documentation dialog
        doc_dialog = tk.Toplevel(self.root)
        doc_dialog.title("Documentation")
        doc_dialog.geometry("600x400")
        
        # Create a notebook for documentation sections
        doc_notebook = ttk.Notebook(doc_dialog)
        doc_notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Overview section
        overview_frame = ttk.Frame(doc_notebook)
        doc_notebook.add(overview_frame, text="Overview")
        
        overview_text = """
Fleet Electrification Analyzer

This application helps fleet managers analyze their current fleet and plan
electrification strategies by:

1. Decoding VINs to identify vehicle details
2. Retrieving fuel economy and emissions data
3. Analyzing electrification potential and cost savings
4. Planning charging infrastructure requirements
5. Developing emissions reduction strategies

To get started, go to the Process tab and load a CSV file containing
vehicle VINs. The CSV should have a column named 'VINs' or 'VIN'.
        """
        
        overview_label = ttk.Label(
            overview_frame, 
            text=overview_text,
            wraplength=550,
            justify=tk.LEFT
        )
        overview_label.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Usage section
        usage_frame = ttk.Frame(doc_notebook)
        doc_notebook.add(usage_frame, text="Usage")
        
        usage_text = """
Usage Instructions (6-tab workflow):

1. Process Tab:
   - Load a CSV file with VINs
   - Set output path and processing options
   - Click "Start Processing" to begin

2. Results Tab:
   - View and search all processed vehicles
   - Customize visible columns, filter by quality/status
   - Right-click to save MPG to the reference database

3. Analysis Tab:
   - Click "Run Full Analysis" to compute TCO, emissions & charging
   - Review fleet KPIs, ACF donut, Top 5 priority vehicles
   - Compare electrification scenarios (auto-runs with analysis)
   - Export Excel report (8-tab) or navigate to Present

4. Timeline Tab:
   - View scheduled EV replacement years for every vehicle
   - Double-click "Proposed EV Year" to manually override
   - Filter by ACF category, EV year range, or free text
   - Live Gantt chart updates with each override

5. Present Tab:
   - Configure and export a PowerPoint presentation
   - Select slides, load a client template (.pptx/.potx)
   - Choose scenarios to include

6. Database Tab:
   - Browse and manage the analyst-maintained MPG reference DB
   - Add/edit/delete vehicle specs used as a fallback MPG source
        """
        
        usage_label = ttk.Label(
            usage_frame, 
            text=usage_text,
            wraplength=550,
            justify=tk.LEFT
        )
        usage_label.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Input Format section
        format_frame = ttk.Frame(doc_notebook)
        doc_notebook.add(format_frame, text="Input Format")
        
        format_text = """
Input CSV Format:

The input CSV should contain a column named 'VINs' or 'VIN' with valid
17-character Vehicle Identification Numbers.

Additional columns will be preserved and can include:
- Odometer: Current odometer readings
- Department: Vehicle department assignment
- Location: Vehicle location information
- Asset ID: Asset tracking number
- Annual Mileage: Expected annual mileage

Example:
VINs,Odometer,Department,Location,Annual_Mileage
1FTFW1ET5DFA92312,45000,Maintenance,North Depot,15000
2FMDK3JC4NBA12345,15200,Admin,Headquarters,8000

Commercial Vehicle Enhancement:
The system now automatically detects commercial vehicles (Class 3-8) and
enhances data through intelligent web scraping to provide:
- Payload and towing capacities
- Duty cycle classification
- Electrification suitability assessment
- Detailed specifications for fleet planning
        """
        
        format_label = ttk.Label(
            format_frame, 
            text=format_text,
            wraplength=550,
            justify=tk.LEFT
        )
        format_label.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Commercial Vehicle Features section
        commercial_frame = ttk.Frame(doc_notebook)
        doc_notebook.add(commercial_frame, text="Commercial Features")
        
        commercial_text = """
Commercial Vehicle Analysis:

Enhanced Data Collection:
The system automatically detects commercial vehicles (GVWR > 8,500 lbs) and 
uses intelligent web scraping to gather comprehensive specifications including:

• Payload Capacity: Maximum cargo weight capacity
• Towing Capacity: Maximum trailer weight capacity
• Duty Cycle: Operational classification (Urban, Regional, Long Haul, etc.)
• Electrification Suitability: Assessment for EV conversion potential

Key Commercial Classifications:
• Light Duty: ≤8,500 lbs GVWR (pickup trucks, small vans)
• Medium Duty: 8,501-19,500 lbs GVWR (box trucks, large vans)
• Heavy Duty: 19,501-33,000 lbs GVWR (delivery trucks, buses)
• Extra Heavy Duty: >33,000 lbs GVWR (semi-trucks, large buses)

Electrification Assessment:
The system evaluates each commercial vehicle's potential for electrification
based on duty cycle, range requirements, and operational patterns to help
prioritize fleet electrification planning.

Data Sources:
Information is gathered from manufacturer websites, government databases,
and industry sources to ensure comprehensive and accurate specifications.
        """
        
        commercial_label = ttk.Label(
            commercial_frame, 
            text=commercial_text,
            wraplength=550,
            justify=tk.LEFT
        )
        commercial_label.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Close button
        ttk.Button(
            doc_dialog, 
            text="Close", 
            command=doc_dialog.destroy
        ).pack(side=tk.BOTTOM, pady=10)
    
    def show_about(self):
        """Show about dialog."""
        messagebox.showinfo(
            "About Fleet Electrification Analyzer",
            f"{APP_NAME} v{APP_VERSION}\n\n"
            "A tool for analyzing fleet vehicles and planning electrification strategies.\n\n"
            "© 2026 Fleet Analytics"
        )
    
    def start_processing(self, input_path, output_path, options):
        """
        Start the processing pipeline.
        
        Args:
            input_path: Path to input CSV
            output_path: Path to output CSV
            options: Processing options
        """
        # Check if already processing
        if self.processing:
            messagebox.showinfo("Processing", "Processing is already in progress.")
            return
        
        # Validate paths
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "Input file does not exist.")
            return
        
        if not output_path:
            messagebox.showerror("Error", "Output path is required.")
            return
        
        # Store fleet path for profile sidecar
        self._fleet_input_path = input_path
        self.present_panel.set_fleet_path(input_path)

        # Update status
        self.status_bar.set("Processing started...")
        self.processing = True
        
        # Get processing options
        max_threads = options.get("max_threads", 10)
        cached_validation = options.get("cached_validation")
        
        # Define callbacks — marshal to main thread for Tkinter safety
        def log_callback(message):
            def _update_log(m=message):
                self.process_panel.add_log(m)
                self.status_bar.set(m)
            self.root.after(0, _update_log)

        def progress_callback(current, total):
            self.root.after(0, lambda c=current, t=total: self.process_panel.update_progress(c, t))
        
        def done_callback(vehicles):
            try:
                # Update processing state
                self.processing = False

                self.fleet = Fleet(
                    name=f"Fleet Analysis - {time.strftime('%Y-%m-%d')}",
                    vehicles=vehicles,
                    creation_date=datetime.datetime.now(),
                    last_modified=datetime.datetime.now()
                )

                # Update shared data for Present panel
                self.sharing_data.set("fleet", self.fleet)

                # UI updates - ensure they happen on main thread
                def update_ui():
                    try:
                        self.results_panel.set_data(vehicles)
                        self.analysis_panel.set_fleet(self.fleet)
                        self.timeline_panel.set_fleet(
                            self.fleet,
                            on_year_changed=self._on_timeline_year_changed,
                            on_acf_changed=self._on_acf_override,
                        )
                        self.present_panel.refresh_data()

                        # Load sidecar profile if one exists for this fleet
                        try:
                            fleet_path = getattr(self, "_fleet_input_path", None)
                            if fleet_path:
                                from data.processor import load_presentation_profile
                                profile = load_presentation_profile(fleet_path)
                                self.present_panel.load_profile(profile)
                                self.sharing_data.set("presentation_profile", profile)
                        except Exception as _pe:
                            logger.warning("Could not load presentation profile: %s", _pe)

                        self.update_vehicle_count()

                        success_count = sum(1 for v in vehicles if v.processing_success)
                        status_msg = f"Processing complete. {success_count}/{len(vehicles)} vehicles successful."
                        self.status_bar.set(status_msg)

                        # Reset process panel button states
                        self.process_panel.processing_complete()

                        # Switch to results tab if there are vehicles
                        if vehicles:
                            self.notebook.select(1)  # Results tab

                    except Exception as ui_e:
                        logger.error(f"Error updating UI after processing: {ui_e}")
                        import traceback
                        logger.error(traceback.format_exc())

                # Schedule UI update on main thread if we're not on it
                if threading.current_thread() == threading.main_thread():
                    update_ui()
                else:
                    self.root.after(0, update_ui)

            except Exception as e:
                logger.error(f"Error in processing completion: {e}")
                import traceback
                logger.error(traceback.format_exc())
        
        # Start processing
        self.processor.process_file(
            input_path=input_path,
            output_path=output_path,
            log_callback=log_callback,
            progress_callback=progress_callback,
            done_callback=done_callback,
            cached_validation=cached_validation,
        )
    
    def stop_processing(self):
        """Stop the processing pipeline."""
        if not self.processing:
            return
        
        # Confirm stop
        if messagebox.askyesno("Confirm", "Are you sure you want to stop processing?"):
            # Update status
            self.status_bar.set("Stopping processing...")
            
            # Stop processor
            self.processor.stop()
            
            # Update state
            self.processing = False

            # Reset process panel button states
            self.process_panel.processing_stopped()

            # Update status
            self.status_bar.set("Processing stopped.")