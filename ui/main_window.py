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
    USER_SETTINGS,
    save_user_settings
)
from utils import StatusBar, SimpleTooltip, ProgressDialog, SafeDict, timestamp

# Import UI panels
from ui.process_panel import ProcessPanel
from ui.results_panel import ResultsPanel
from ui.analysis_panel import AnalysisPanel

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
        file_menu.add_separator()
        file_menu.add_command(label="Save Results...", command=self.save_results)
        file_menu.add_command(label="Export Report...", command=self.export_report)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_closing)
        menubar.add_cascade(label="File", menu=file_menu)
        
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
        tools_menu.add_command(label="Analyze Fleet Emissions", command=self.analyze_emissions)
        tools_menu.add_command(label="Electrification Modeling", command=self.analyze_electrification)
        tools_menu.add_command(label="Charging Infrastructure", command=self.analyze_charging)
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
            on_column_change=self.on_column_change
        )
        self.results_panel.pack(fill=tk.BOTH, expand=True)
        self.notebook.add(self.results_frame, text="Results")
        
        # Analysis panel
        self.analysis_frame = ttk.Frame(self.notebook)
        self.analysis_panel = AnalysisPanel(
            self.analysis_frame,
            fleet=self.fleet,
            on_analysis_complete=self.on_analysis_complete,
            on_report_generation=self.on_report_generated
        )
        self.analysis_panel.pack(fill=tk.BOTH, expand=True)
        self.notebook.add(self.analysis_frame, text="Analysis")
    
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
        
        # Close the application
        self.root.destroy()
    
    def on_tab_changed(self, event):
        """Handle notebook tab changed event."""
        current_tab = self.notebook.index(self.notebook.select())
        
        # Update status based on current tab
        if current_tab == 0:  # Process tab
            self.status_bar.set("Process vehicles by VIN")
        elif current_tab == 1:  # Results tab
            self.status_bar.set("View and filter processed vehicles")
        elif current_tab == 2:  # Analysis tab
            self.status_bar.set("Analyze fleet electrification potential")
    
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
        """Export a comprehensive report with data and analysis."""
        # Check if there are results to export
        if not self.fleet.vehicles:
            messagebox.showinfo("No Data", "There are no results to export.")
            return
        
        # Get file path
        filepath = filedialog.asksaveasfilename(
            title="Export Report",
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
            self.root,
            "Generating Report",
            "Please wait while the report is being generated..."
        )
        
        # Run export in background thread
        def export_task():
            try:
                # Get analysis results from shared data
                electrification_analysis = self.sharing_data.get("electrification_results")
                charging_analysis = self.sharing_data.get("charging_results")
                emissions_analysis = self.sharing_data.get("emissions_results")
                
                # Run any missing analysis if needed
                if not electrification_analysis:
                    progress.update(25, message="Running electrification analysis...")
                    electrification_analysis = self.analysis_panel.run_electrification_analysis()
                
                progress.update(50, message="Generating charts...")
                
                # Export the report
                from analysis.reports import ReportGeneratorFactory
                
                generator = ReportGeneratorFactory.create_generator(filepath)
                if generator:
                    progress.update(75, message="Exporting report...")
                    
                    success = generator.generate(
                        fleet=self.fleet,
                        analysis=electrification_analysis,
                        charging=charging_analysis,
                        emissions=emissions_analysis
                    )
                    
                    if success:
                        progress.update(100, message="Report complete!")
                        
                        # Show success message after progress dialog is closed
                        self.root.after(
                            500,
                            lambda: messagebox.showinfo(
                                "Export Complete",
                                f"Report has been exported to:\n{filepath}"
                            )
                        )
                    else:
                        # Show error message after progress dialog is closed
                        self.root.after(
                            500,
                            lambda: messagebox.showerror(
                                "Export Failed",
                                "An error occurred while exporting the report."
                            )
                        )
                else:
                    # Show error message after progress dialog is closed
                    self.root.after(
                        500,
                        lambda: messagebox.showerror(
                            "Export Failed",
                            "Unsupported file format."
                        )
                    )
                
            except Exception as e:
                logger.error(f"Error exporting report: {e}")
                
                # Show error message after progress dialog is closed
                self.root.after(
                    500,
                    lambda: messagebox.showerror(
                        "Export Failed",
                        f"An error occurred: {str(e)}"
                    )
                )
            
            finally:
                # Close progress dialog
                self.root.after(100, progress.destroy)
        
        # Start export thread
        threading.Thread(target=export_task, daemon=True).start()
    
    def copy_selection(self):
        """Copy the current selection to the clipboard."""
        # Determine current tab and delegate to appropriate panel
        current_tab = self.notebook.index(self.notebook.select())
        
        if current_tab == 0:  # Process tab
            self.process_panel.copy_selection()
        elif current_tab == 1:  # Results tab
            self.results_panel.copy_selection()
        elif current_tab == 2:  # Analysis tab
            self.analysis_panel.copy_selection()
    
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
        # Create dialog
        columns_dialog = tk.Toplevel(self.root)
        columns_dialog.title("Customize Columns")
        columns_dialog.geometry("300x400")
        columns_dialog.transient(self.root)  # Set as transient to main window
        columns_dialog.grab_set()  # Make modal
        
        # Get current column configuration
        all_columns = self.results_panel.get_all_columns()
        visible_columns = self.results_panel.get_visible_columns()
        
        # Create column selection UI
        ttk.Label(
            columns_dialog, 
            text="Select columns to display:", 
            font=("", 10, "bold")
        ).pack(anchor=tk.W, padx=10, pady=5)
        
        # Create a frame with scrollbar for checkboxes
        checkbox_frame = ttk.Frame(columns_dialog)
        checkbox_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create canvas and scrollbar
        canvas = tk.Canvas(checkbox_frame)
        scrollbar = ttk.Scrollbar(checkbox_frame, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Create a frame inside the canvas for checkboxes
        inner_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=inner_frame, anchor=tk.NW)
        
        # Create checkbox variables
        checkbox_vars = {}
        
        # Create checkboxes for each column
        for i, (col_id, col_name) in enumerate(all_columns):
            var = tk.BooleanVar(value=col_id in visible_columns)
            checkbox_vars[col_id] = var
            
            ttk.Checkbutton(
                inner_frame,
                text=col_name,
                variable=var
            ).grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
        
        # Update canvas when frame size changes
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        inner_frame.bind("<Configure>", on_frame_configure)
        
        # Add buttons
        button_frame = ttk.Frame(columns_dialog)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        
        ttk.Button(
            button_frame, 
            text="Cancel", 
            command=columns_dialog.destroy
        ).pack(side=tk.RIGHT, padx=5)
        
        # Helper to select all/none
        def select_all(select=True):
            for var in checkbox_vars.values():
                var.set(select)
        
        ttk.Button(
            button_frame, 
            text="Select All", 
            command=lambda: select_all(True)
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="Select None", 
            command=lambda: select_all(False)
        ).pack(side=tk.LEFT, padx=5)
        
        # Apply changes
        def apply_columns():
            # Get selected columns
            selected = [col_id for col_id, var in checkbox_vars.items() if var.get()]
            
            # Update results panel
            self.results_panel.set_visible_columns(selected)
            
            # Close dialog
            columns_dialog.destroy()
        
        ttk.Button(
            button_frame, 
            text="Apply", 
            command=apply_columns
        ).pack(side=tk.RIGHT, padx=5)
    
    def refresh_view(self):
        """Refresh the current view."""
        # Determine current tab and refresh the appropriate panel
        current_tab = self.notebook.index(self.notebook.select())
        
        if current_tab == 0:  # Process tab
            self.process_panel.refresh()
        elif current_tab == 1:  # Results tab
            self.results_panel.refresh()
        elif current_tab == 2:  # Analysis tab
            self.analysis_panel.refresh()
    
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
Usage Instructions:

1. Process Tab:
   - Load a CSV file with VINs
   - Set output path and processing options
   - Click "Start Processing" to begin

2. Results Tab:
   - View processed vehicles in the table
   - Filter and sort data
   - Customize visible columns

3. Analysis Tab:
   - Run different analysis types
   - View charts and results
   - Export reports
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

Example:
VINs,Odometer,Department,Location
1FTFW1ET5DFA92312,45000,Maintenance,North Depot
2FMDK3JC4NBA12345,15200,Admin,Headquarters
        """
        
        format_label = ttk.Label(
            format_frame, 
            text=format_text,
            wraplength=550,
            justify=tk.LEFT
        )
        format_label.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
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
            "© 2023 Fleet Analytics"
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
        
        # Update status
        self.status_bar.set("Processing started...")
        self.processing = True
        
        # Get processing options
        max_threads = options.get("max_threads", 10)
        
        # Define callbacks
        def log_callback(message):
            # Update process panel log
            self.process_panel.add_log(message)
            
            # Update status bar
            self.status_bar.set(message)
        
        def progress_callback(current, total):
            # Update process panel progress
            self.process_panel.update_progress(current, total)
        
        def done_callback(vehicles):
            # Add detailed logging
            logger.info(f"🔧 DEBUG: done_callback received {len(vehicles)} vehicles")
            logger.info(f"🔧 DEBUG: done_callback called from thread: {threading.current_thread().name}")
            logger.info(f"🔧 DEBUG: Main thread check: {threading.current_thread() == threading.main_thread()}")
            
            try:
                # Update processing state
                self.processing = False
                
                # Log vehicle details for debugging
                success_count = sum(1 for v in vehicles if v.processing_success)
                failed_count = len(vehicles) - success_count
                logger.info(f"🔧 DEBUG: Vehicle breakdown - Success: {success_count}, Failed: {failed_count}")
                
                if failed_count > 0:
                    logger.info(f"🔧 DEBUG: First few failed vehicles:")
                    for i, v in enumerate([v for v in vehicles if not v.processing_success][:3]):
                        logger.info(f"🔧 DEBUG: Failed vehicle {i+1}: VIN={v.vin}, Error='{v.processing_error}'")
                
                # Log fleet creation attempt
                logger.info(f"🔧 DEBUG: Creating Fleet object...")
                self.fleet = Fleet(
                    name=f"Fleet Analysis - {time.strftime('%Y-%m-%d')}",
                    vehicles=vehicles,
                    creation_date=datetime.datetime.now(),
                    last_modified=datetime.datetime.now()
                )
                logger.info(f"🔧 DEBUG: Fleet created successfully with {len(vehicles)} vehicles")
                
                # UI updates - ensure they happen on main thread
                def update_ui():
                    try:
                        logger.info(f"🔧 DEBUG: Starting UI updates on thread: {threading.current_thread().name}")
                        
                        # Update results panel
                        logger.info(f"🔧 DEBUG: Calling results_panel.set_data() with {len(vehicles)} vehicles")
                        self.results_panel.set_data(vehicles)
                        logger.info(f"🔧 DEBUG: results_panel.set_data() completed")
                        
                        # Update analysis panel
                        logger.info(f"🔧 DEBUG: Calling analysis_panel.set_fleet()")
                        self.analysis_panel.set_fleet(self.fleet)
                        logger.info(f"🔧 DEBUG: analysis_panel.set_fleet() completed")
                        
                        # Update vehicle count
                        logger.info(f"🔧 DEBUG: Updating vehicle count")
                        self.update_vehicle_count()
                        
                        # Update status
                        status_msg = f"Processing complete. {len(vehicles)} vehicles processed."
                        logger.info(f"🔧 DEBUG: Setting status: {status_msg}")
                        self.status_bar.set(status_msg)
                        
                        # Switch to results tab if there are vehicles
                        if vehicles:
                            logger.info(f"🔧 DEBUG: Switching to results tab (notebook.select(1))")
                            self.notebook.select(1)  # Results tab
                            logger.info(f"🔧 DEBUG: Successfully switched to results tab")
                        else:
                            logger.warning(f"🔧 DEBUG: No vehicles to display, staying on process tab")
                        
                        logger.info(f"🔧 DEBUG: All UI updates completed successfully")
                        
                    except Exception as ui_e:
                        logger.error(f"🔧 DEBUG: Exception in UI update: {ui_e}")
                        logger.error(f"🔧 DEBUG: UI update exception type: {type(ui_e).__name__}")
                        import traceback
                        logger.error(f"🔧 DEBUG: UI update traceback: {traceback.format_exc()}")
                
                # Schedule UI update on main thread if we're not on it
                if threading.current_thread() == threading.main_thread():
                    logger.info(f"🔧 DEBUG: Already on main thread, updating UI directly")
                    update_ui()
                else:
                    logger.info(f"🔧 DEBUG: Not on main thread, scheduling UI update via root.after()")
                    self.root.after(0, update_ui)
                
            except Exception as e:
                logger.error(f"🔧 DEBUG: Exception in done_callback: {e}")
                logger.error(f"🔧 DEBUG: done_callback exception type: {type(e).__name__}")
                import traceback
                logger.error(f"🔧 DEBUG: done_callback traceback: {traceback.format_exc()}")
        
        # Start processing
        self.processor.process_file(
            input_path=input_path,
            output_path=output_path,
            log_callback=log_callback,
            progress_callback=progress_callback,
            done_callback=done_callback
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
            
            # Update status
            self.status_bar.set("Processing stopped.")