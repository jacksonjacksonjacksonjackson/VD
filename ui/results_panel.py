"""
results_panel.py

Panel for displaying and interacting with processed vehicle data in the
Fleet Electrification Analyzer.
"""

import logging
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List, Any, Optional, Callable, Tuple
import threading

from settings import (
    PRIMARY_HEX_1,
    PRIMARY_HEX_2, 
    PRIMARY_HEX_3,
    SECONDARY_HEX_1,
    COLUMN_NAME_MAP,
    DEFAULT_VISIBLE_COLUMNS
)
from utils import SimpleTooltip
from data.models import FleetVehicle

# Set up module logger
logger = logging.getLogger(__name__)

class ResultsPanel(ttk.Frame):
    """
    Panel for displaying and interacting with processed vehicle data.
    Features filterable and sortable treeview with customizable columns.
    """
    
    def __init__(self, parent, data=None, visible_columns=None,
               on_selection_change=None, on_column_change=None):
        """
        Initialize the results panel.
        
        Args:
            parent: Parent widget
            data: Initial data to display (list of FleetVehicle objects)
            visible_columns: Initially visible columns (or None for defaults)
            on_selection_change: Callback when selection changes
            on_column_change: Callback when column visibility changes
        """
        super().__init__(parent)
        
        # Store callbacks
        self.on_selection_change_callback = on_selection_change
        self.on_column_change_callback = on_column_change
        
        # Initialize variables
        self.data = data or []
        self.visible_columns = visible_columns or DEFAULT_VISIBLE_COLUMNS
        self.all_columns_map = {}  # Maps column IDs to display names
        self.search_var = tk.StringVar()  # For search/filter
        self.search_fields_var = tk.StringVar(value="all")  # Fields to search in
        self.status_filter_var = tk.StringVar(value="all")  # Status filter
        self.quality_filter_var = tk.StringVar(value="all")  # Quality filter
        self.data_map = {}  # Maps treeview IIDs to data objects
        
        # Create UI components
        self._create_toolbar()
        self._create_summary()
        self._create_treeview()
        
        # Populate data
        self.populate_data()
        
        # Create context menu
        self._create_context_menu()
    
    def _create_toolbar(self):
        """Create the toolbar with search, filters, and actions."""
        # Create toolbar frame
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        # Left side - search and filters
        search_frame = ttk.Frame(toolbar)
        search_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Search controls
        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT, padx=(0, 5))
        
        search_entry = ttk.Entry(
            search_frame, 
            textvariable=self.search_var,
            width=25
        )
        search_entry.pack(side=tk.LEFT, padx=(0, 5))
        SimpleTooltip(search_entry, "Enter text to filter results")
        
        # Bind search entry to filter data
        self.search_var.trace_add("write", lambda *args: self._apply_filter())
        
        # Search fields dropdown
        ttk.Label(search_frame, text="in:").pack(side=tk.LEFT, padx=(5, 2))
        
        search_fields = ttk.Combobox(
            search_frame,
            textvariable=self.search_fields_var,
            values=["all", "VIN", "Make", "Model", "Year", "Asset ID", "Department"],
            width=10,
            state="readonly"
        )
        search_fields.pack(side=tk.LEFT, padx=(0, 10))
        SimpleTooltip(search_fields, "Select fields to search in")
        
        # Bind search fields to apply filter
        search_fields.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())
        
        # Quick filters
        filters_frame = ttk.LabelFrame(search_frame, text="Quick Filters")
        filters_frame.pack(side=tk.LEFT, padx=(10, 0), pady=2)
        
        # Status filter
        status_filter = ttk.Combobox(
            filters_frame,
            textvariable=self.status_filter_var,
            values=["all", "successful", "failed"],
            width=10,
            state="readonly"
        )
        status_filter.pack(side=tk.LEFT, padx=5, pady=5)
        SimpleTooltip(status_filter, "Filter by processing status")
        status_filter.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())
        
        # Quality filter
        quality_filter = ttk.Combobox(
            filters_frame,
            textvariable=self.quality_filter_var,
            values=["all", "high quality (80%+)", "medium quality (50-80%)", "low quality (<50%)"],
            width=15,
            state="readonly"
        )
        quality_filter.pack(side=tk.LEFT, padx=5, pady=5)
        SimpleTooltip(quality_filter, "Filter by data quality score")
        quality_filter.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())
        
        # Right side - actions
        actions_frame = ttk.Frame(toolbar)
        actions_frame.pack(side=tk.RIGHT)
        
        # Export for Excel Analysis button (primary export)
        excel_export_btn = ttk.Button(
            actions_frame,
            text="Export for Excel Analysis",
            command=self._export_for_excel,
            style="Accent.TButton"
        )
        excel_export_btn.pack(side=tk.RIGHT, padx=5)
        SimpleTooltip(excel_export_btn, "One-click export optimized for Excel analysis with preserved order")
        
        # Standard export button
        export_btn = ttk.Button(
            actions_frame,
            text="Export...",
            command=self._export_data
        )
        export_btn.pack(side=tk.RIGHT, padx=5)
        SimpleTooltip(export_btn, "Export with custom format options")
        
        # Clear filter button
        clear_btn = ttk.Button(
            actions_frame,
            text="Clear Filters",
            command=self._clear_filter
        )
        clear_btn.pack(side=tk.RIGHT, padx=5)
        SimpleTooltip(clear_btn, "Clear all search filters")
        
        # Refresh button
        refresh_btn = ttk.Button(
            actions_frame,
            text="Refresh",
            command=self.refresh
        )
        refresh_btn.pack(side=tk.RIGHT, padx=5)
        SimpleTooltip(refresh_btn, "Refresh data display")
    
    def _create_summary(self):
        """Create the processing summary widget."""
        # Create summary frame
        self.summary_frame = ttk.Frame(self)
        self.summary_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
        
        # Configure summary frame style
        style = ttk.Style()
        style.configure("Summary.TFrame", relief="solid", borderwidth=1)
        self.summary_frame.configure(style="Summary.TFrame")
        
        # Create summary label
        self.summary_label = ttk.Label(
            self.summary_frame,
            text="No data loaded",
            font=("", 10, "bold"),
            foreground="#2E8B57"  # Sea green color
        )
        self.summary_label.pack(padx=10, pady=5)
    
    def _create_treeview(self):
        """Create the treeview for displaying vehicle data."""
        # Create frame for treeview and scrollbars
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Build all column mapping
        self._build_column_map()
        
        # Create scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Create treeview
        self.tree = ttk.Treeview(
            tree_frame,
            columns=self.visible_columns,
            show="headings",
            selectmode="extended",
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set
        )
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Configure scrollbars
        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)
        
        # Configure columns and headings
        for col in self.visible_columns:
            display_name = self.all_columns_map.get(col, col)
            self.tree.heading(
                col,
                text=display_name,
                command=lambda c=col: self._sort_by_column(c)
            )
            
            # Adjust column width based on content
            width = max(100, len(display_name) * 10)
            self.tree.column(col, width=width, minwidth=50)
        
        # Configure color tags for data quality indication
        self.tree.tag_configure("failed", background="#FFE4E1")  # Light red for failed
        self.tree.tag_configure("high_quality", background="#F0FFF0")  # Light green for high quality
        self.tree.tag_configure("medium_quality", background="#FFFACD")  # Light yellow for medium quality  
        self.tree.tag_configure("low_quality", background="#FFE4B5")  # Light orange for low quality
        
        # Bind events
        self.tree.bind("<<TreeviewSelect>>", self._on_selection_change)
        self.tree.bind("<Double-1>", self._show_details)
        self.tree.bind("<Button-3>", self._show_context_menu)
    
    def _create_context_menu(self):
        """Create context menu for the treeview."""
        self.context_menu = tk.Menu(self, tearoff=0)
        
        # Add menu items
        self.context_menu.add_command(
            label="View Details",
            command=self._show_details
        )
        
        self.context_menu.add_command(
            label="Copy Selected",
            command=self.copy_selection
        )
        
        self.context_menu.add_separator()
        
        self.context_menu.add_command(
            label="Analyze Selected",
            command=self._analyze_selected
        )
        
        self.context_menu.add_separator()
        
        self.context_menu.add_command(
            label="Select All",
            command=self._select_all
        )
        
        self.context_menu.add_command(
            label="Deselect All",
            command=self._deselect_all
        )
        
        # Bind right-click to show menu
        self.tree.bind("<Button-3>", self._show_context_menu)
    
    def _build_column_map(self):
        """Build the mapping of all possible columns to display names."""
        # Start with column map from settings
        self.all_columns_map = dict(COLUMN_NAME_MAP)
        
        # Add any missing columns based on data
        if self.data:
            sample = self.data[0].to_row_dict()
            for key in sample.keys():
                if key not in self.all_columns_map:
                    # Use the key itself as display name
                    self.all_columns_map[key] = key
    
    def _sort_by_column(self, column):
        """
        Sort treeview data by specified column.
        
        Args:
            column: Column identifier to sort by
        """
        # Get current sort direction
        if hasattr(self, '_sort_column') and self._sort_column == column:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_column = column
            self._sort_reverse = False
        
        # Update heading text to show sort direction
        heading_text = self.all_columns_map.get(column, column)
        if self._sort_reverse:
            heading_text += " ▼"
        else:
            heading_text += " ▲"
        
        # Update heading
        self.tree.heading(column, text=heading_text)
        
        # Get all items
        items = [(self.tree.set(iid, column), iid) for iid in self.tree.get_children("")]
        
        # Sort items
        items.sort(reverse=self._sort_reverse)
        
        # Try to convert to numeric for numeric sorting
        def try_numeric_sort():
            try:
                # Convert to numeric if possible
                numeric_items = [(float(value), iid) for value, iid in items]
                return sorted(numeric_items, reverse=self._sort_reverse)
            except ValueError:
                # Fall back to string sort
                return None
        
        # Try numeric sort first
        numeric_sort = try_numeric_sort()
        if numeric_sort is not None:
            items = numeric_sort
        
        # Rearrange items
        for idx, (_, iid) in enumerate(items):
            self.tree.move(iid, "", idx)
        
        # Restore default headings for other columns
        for col in self.visible_columns:
            if col != column:
                self.tree.heading(col, text=self.all_columns_map.get(col, col))
    
    def _apply_filter(self):
        """Apply search and filter criteria to the treeview."""
        # Get current filter criteria
        search_text = self.search_var.get().strip().lower()
        search_field = self.search_fields_var.get()
        status_filter = self.status_filter_var.get()
        quality_filter = self.quality_filter_var.get()
        
        # Clear the treeview
        self.tree.delete(*self.tree.get_children())
        self.data_map.clear()
        
        # If no data, return
        if not self.data:
            return
        
        # Filter and display matching vehicles
        filtered_count = 0
        for vehicle in self.data:
            # Apply status filter
            if status_filter != "all":
                vehicle_success = (vehicle.vehicle_id.year and 
                                 vehicle.vehicle_id.make and 
                                 vehicle.vehicle_id.model)
                if status_filter == "successful" and not vehicle_success:
                    continue
                if status_filter == "failed" and vehicle_success:
                    continue
            
            # Apply quality filter
            if quality_filter != "all":
                quality_score = getattr(vehicle, 'data_quality_score', 0)
                if quality_filter == "high quality (80%+)" and quality_score < 80:
                    continue
                if quality_filter == "medium quality (50-80%)" and (quality_score < 50 or quality_score >= 80):
                    continue
                if quality_filter == "low quality (<50%)" and quality_score >= 50:
                    continue
            
            # Apply search filter
            if search_text:
                match_found = False
                row_dict = vehicle.to_row_dict()
                
                if search_field == "all":
                    # Search all visible fields
                    for col in self.visible_columns:
                        value = str(row_dict.get(col, "")).lower()
                        if search_text in value:
                            match_found = True
                            break
                else:
                    # Search specific field
                    if search_field in row_dict:
                        value = str(row_dict[search_field]).lower()
                        match_found = search_text in value
                
                if not match_found:
                    continue
            
            # Vehicle passed all filters - add to display
            row_dict = vehicle.to_row_dict()
            values = [row_dict.get(col, "") for col in self.visible_columns]
            
            # Insert into treeview with color coding
            iid = self.tree.insert("", tk.END, values=values)
            self._apply_row_color_coding(iid, vehicle)
            
            # Map iid to vehicle object
            self.data_map[iid] = vehicle
            filtered_count += 1
        
        # Update summary to show filtered results
        self._update_summary_filtered(filtered_count)
    
    def _apply_row_color_coding(self, iid: str, vehicle: FleetVehicle) -> None:
        """
        Apply color coding to a row based on vehicle data quality.
        
        Args:
            iid: Treeview item ID
            vehicle: FleetVehicle object
        """
        # Determine colors based on processing success and data quality
        if not vehicle.processing_success:
            # Failed processing - red background
            self.tree.item(iid, tags=("failed",))
        elif vehicle.data_quality_score >= 80:
            # High quality - light green background
            self.tree.item(iid, tags=("high_quality",))
        elif vehicle.data_quality_score >= 50:
            # Medium quality - light yellow background
            self.tree.item(iid, tags=("medium_quality",))
        else:
            # Low quality - light orange background
            self.tree.item(iid, tags=("low_quality",))
    
    def _update_summary_filtered(self, filtered_count: int) -> None:
        """
        Update the summary display for filtered results.
        
        Args:
            filtered_count: Number of items after filtering
        """
        if not self.data:
            self.summary_label.config(text="No data loaded")
            return
        
        total_count = len(self.data)
        
        if filtered_count == total_count:
            # No filtering applied, use normal summary
            self._update_summary()
        else:
            # Show filtered results
            summary_text = f"📊 Showing {filtered_count} of {total_count} vehicles"
            text_color = "#4682B4"  # Steel blue for filtered results
            self.summary_label.config(text=summary_text, foreground=text_color)
    
    def _clear_filter(self):
        """Clear search filter and quick filters."""
        self.search_var.set("")
        self.status_filter_var.set("all")
        self.quality_filter_var.set("all")
        self.populate_data()
    
    def _export_data(self):
        """Export the currently displayed data."""
        # Get parent window to show dialog
        from tkinter import filedialog
        
        filepath = filedialog.asksaveasfilename(
            title="Export Data",
            defaultextension=".csv",
            filetypes=[
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx"),
                ("All Files", "*.*")
            ]
        )
        
        if not filepath:
            return
        
        # Get currently displayed data
        displayed_vehicles = []
        for iid in self.tree.get_children():
            if iid in self.data_map:
                displayed_vehicles.append(self.data_map[iid])
        
        # Get visible columns
        visible_columns = self.visible_columns
        
        # Export the data
        from analysis.reports import ReportGeneratorFactory
        
        generator = ReportGeneratorFactory.create_generator(filepath)
        if generator:
            success = generator.generate(
                fleet=displayed_vehicles,
                fields=visible_columns
            )
            
            if success:
                messagebox.showinfo(
                    "Export Complete",
                    f"Data has been exported to:\n{filepath}"
                )
            else:
                messagebox.showerror(
                    "Export Failed",
                    "An error occurred while exporting the data."
                )
        else:
            messagebox.showerror(
                "Export Failed",
                "Unsupported file format."
            )
    
    def _on_selection_change(self, event):
        """Handle selection change in treeview."""
        # Get selected items
        selected_iids = self.tree.selection()
        
        # Map to vehicle objects
        selected_vehicles = [self.data_map[iid] for iid in selected_iids if iid in self.data_map]
        
        # Call callback if provided
        if self.on_selection_change_callback:
            self.on_selection_change_callback(selected_vehicles)
    
    def _show_details(self, event=None):
        """Show details for the selected vehicle."""
        # Get selected item
        selected_iids = self.tree.selection()
        
        if not selected_iids:
            # If called from menu without selection, just return
            return
        
        # Get the first selected vehicle
        if selected_iids[0] in self.data_map:
            vehicle = self.data_map[selected_iids[0]]
            
            # Create details window
            self._create_details_window(vehicle)
    
    def _create_details_window(self, vehicle):
        """
        Create a details window for the specified vehicle.
        
        Args:
            vehicle: FleetVehicle object to display
        """
        # Create a new toplevel window
        details_window = tk.Toplevel(self)
        details_window.title(f"Vehicle Details: {vehicle.vin}")
        details_window.geometry("600x500")
        details_window.transient(self.master)  # Set as transient to main window
        
        # Make the window modal
        details_window.grab_set()
        
        # Create a notebook with tabs for different categories
        notebook = ttk.Notebook(details_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Overview tab
        overview_frame = ttk.Frame(notebook)
        notebook.add(overview_frame, text="Overview")
        
        # Create a grid of labels for overview data
        row = 0
        
        # Add VIN row
        ttk.Label(overview_frame, text="VIN:", font=("", 10, "bold")).grid(
            row=row, column=0, sticky=tk.W, padx=5, pady=2
        )
        ttk.Label(overview_frame, text=vehicle.vin).grid(
            row=row, column=1, sticky=tk.W, padx=5, pady=2
        )
        row += 1
        
        # Basic vehicle info section
        ttk.Separator(overview_frame, orient=tk.HORIZONTAL).grid(
            row=row, column=0, columnspan=2, sticky=tk.EW, pady=5
        )
        row += 1
        
        ttk.Label(overview_frame, text="Vehicle Information", font=("", 10, "bold", "underline")).grid(
            row=row, column=0, columnspan=2, sticky=tk.W, padx=5, pady=2
        )
        row += 1
        
        # Add basic vehicle info
        basic_info = [
            ("Year:", vehicle.vehicle_id.year),
            ("Make:", vehicle.vehicle_id.make),
            ("Model:", vehicle.vehicle_id.model),
            ("Fuel Type:", vehicle.vehicle_id.fuel_type),
            ("Body Class:", vehicle.vehicle_id.body_class),
            ("GVWR:", vehicle.vehicle_id.gvwr)
        ]
        
        for label, value in basic_info:
            ttk.Label(overview_frame, text=label).grid(
                row=row, column=0, sticky=tk.W, padx=5, pady=2
            )
            ttk.Label(overview_frame, text=value).grid(
                row=row, column=1, sticky=tk.W, padx=5, pady=2
            )
            row += 1
        
        # Fuel economy section
        ttk.Separator(overview_frame, orient=tk.HORIZONTAL).grid(
            row=row, column=0, columnspan=2, sticky=tk.EW, pady=5
        )
        row += 1
        
        ttk.Label(overview_frame, text="Fuel Economy", font=("", 10, "bold", "underline")).grid(
            row=row, column=0, columnspan=2, sticky=tk.W, padx=5, pady=2
        )
        row += 1
        
        # Add fuel economy info
        economy_info = [
            ("City MPG:", vehicle.fuel_economy.city_mpg),
            ("Highway MPG:", vehicle.fuel_economy.highway_mpg),
            ("Combined MPG:", vehicle.fuel_economy.combined_mpg),
            ("CO2 Emissions:", f"{vehicle.fuel_economy.co2_primary} g/mile"),
            ("Annual Fuel Cost:", f"${vehicle.fuel_economy.fuel_cost_primary}")
        ]
        
        for label, value in economy_info:
            ttk.Label(overview_frame, text=label).grid(
                row=row, column=0, sticky=tk.W, padx=5, pady=2
            )
            ttk.Label(overview_frame, text=value).grid(
                row=row, column=1, sticky=tk.W, padx=5, pady=2
            )
            row += 1
        
        # Fleet data section
        ttk.Separator(overview_frame, orient=tk.HORIZONTAL).grid(
            row=row, column=0, columnspan=2, sticky=tk.EW, pady=5
        )
        row += 1
        
        ttk.Label(overview_frame, text="Fleet Data", font=("", 10, "bold", "underline")).grid(
            row=row, column=0, columnspan=2, sticky=tk.W, padx=5, pady=2
        )
        row += 1
        
        # Add fleet data
        fleet_info = [
            ("Asset ID:", vehicle.asset_id or "N/A"),
            ("Department:", vehicle.department or "N/A"),
            ("Location:", vehicle.location or "N/A"),
            ("Odometer:", f"{vehicle.odometer:,.0f} miles" if vehicle.odometer else "N/A"),
            ("Annual Mileage:", f"{vehicle.annual_mileage:,.0f} miles" if vehicle.annual_mileage else "N/A")
        ]
        
        for label, value in fleet_info:
            ttk.Label(overview_frame, text=label).grid(
                row=row, column=0, sticky=tk.W, padx=5, pady=2
            )
            ttk.Label(overview_frame, text=value).grid(
                row=row, column=1, sticky=tk.W, padx=5, pady=2
            )
            row += 1
        
        # Technical details tab
        tech_frame = ttk.Frame(notebook)
        notebook.add(tech_frame, text="Technical Details")
        
        # Create a scrollable text widget for all technical details
        tech_text = tk.Text(
            tech_frame,
            wrap=tk.WORD,
            width=80,
            height=20,
            bg=PRIMARY_HEX_2,
            state=tk.DISABLED
        )
        tech_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add scrollbar
        tech_scrollbar = ttk.Scrollbar(tech_frame, command=tech_text.yview)
        tech_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tech_text.config(yscrollcommand=tech_scrollbar.set)
        
        # Populate technical details
        tech_text.config(state=tk.NORMAL)
        
        # Vehicle ID details
        tech_text.insert(tk.END, "Vehicle Identification\n", "heading")
        tech_text.insert(tk.END, "-" * 50 + "\n", "separator")
        
        vehicle_id_details = [
            ("Engine Displacement:", vehicle.vehicle_id.engine_displacement),
            ("Engine Cylinders:", vehicle.vehicle_id.engine_cylinders),
            ("Drive Type:", vehicle.vehicle_id.drive_type),
            ("Transmission:", vehicle.vehicle_id.transmission)
        ]
        
        for label, value in vehicle_id_details:
            tech_text.insert(tk.END, f"{label} {value}\n", "detail")
        
        tech_text.insert(tk.END, "\n")
        
        # Fuel Economy details
        tech_text.insert(tk.END, "Fuel Economy Details\n", "heading")
        tech_text.insert(tk.END, "-" * 50 + "\n", "separator")
        
        # Get all raw data
        for key, value in vehicle.fuel_economy.raw_data.items():
            if value:  # Only show non-empty values
                tech_text.insert(tk.END, f"{key}: {value}\n", "detail")
        
        tech_text.insert(tk.END, "\n")
        
        # Custom fields
        if vehicle.custom_fields:
            tech_text.insert(tk.END, "Custom Fields\n", "heading")
            tech_text.insert(tk.END, "-" * 50 + "\n", "separator")
            
            for key, value in vehicle.custom_fields.items():
                tech_text.insert(tk.END, f"{key}: {value}\n", "detail")
            
            tech_text.insert(tk.END, "\n")
        
        # Configure text styles
        tech_text.tag_configure("heading", font=("", 12, "bold"), foreground=PRIMARY_HEX_1)
        tech_text.tag_configure("separator", foreground=PRIMARY_HEX_3)
        tech_text.tag_configure("detail", font=("", 10, ""))
        
        tech_text.config(state=tk.DISABLED)
        
        # Raw Data tab
        raw_frame = ttk.Frame(notebook)
        notebook.add(raw_frame, text="Raw Data")
        
        # Create a scrollable text widget for all raw data
        raw_text = tk.Text(
            raw_frame,
            wrap=tk.WORD,
            width=80,
            height=20,
            bg=PRIMARY_HEX_2,
            state=tk.DISABLED
        )
        raw_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add scrollbar
        raw_scrollbar = ttk.Scrollbar(raw_frame, command=raw_text.yview)
        raw_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        raw_text.config(yscrollcommand=raw_scrollbar.set)
        
        # Populate raw data
        raw_text.config(state=tk.NORMAL)
        
        # Get the full data dictionary
        import json
        raw_data = json.dumps(vehicle.to_dict(), indent=2)
        raw_text.insert(tk.END, raw_data)
        
        raw_text.config(state=tk.DISABLED)
        
        # Close button
        close_button = ttk.Button(
            details_window,
            text="Close",
            command=details_window.destroy
        )
        close_button.pack(side=tk.BOTTOM, pady=10)
    
    def _show_context_menu(self, event):
        """Show the context menu for the treeview."""
        # Check if mouse was clicked on an item
        item = self.tree.identify_row(event.y)
        if item:
            # Select the item under the mouse if not already selected
            if item not in self.tree.selection():
                self.tree.selection_set(item)
            
            # Show the context menu
            self.context_menu.tk_popup(event.x_root, event.y_root)
    
    def _analyze_selected(self):
        """Analyze the selected vehicles."""
        # Get selected vehicles
        selected_iids = self.tree.selection()
        selected_vehicles = [self.data_map[iid] for iid in selected_iids if iid in self.data_map]
        
        if not selected_vehicles:
            messagebox.showinfo("No Selection", "Please select one or more vehicles to analyze.")
            return
        
        # Create an analysis popup
        analysis_window = tk.Toplevel(self)
        analysis_window.title(f"Analysis of {len(selected_vehicles)} Vehicles")
        analysis_window.geometry("600x500")
        analysis_window.transient(self.master)  # Set as transient to main window
        
        # Make the window modal
        analysis_window.grab_set()
        
        # Create a notebook with tabs for different analyses
        notebook = ttk.Notebook(analysis_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Summary tab
        summary_frame = ttk.Frame(notebook)
        notebook.add(summary_frame, text="Summary")
        
        # Calculate summary statistics
        mpg_values = [v.fuel_economy.combined_mpg for v in selected_vehicles if v.fuel_economy.combined_mpg > 0]
        co2_values = [v.fuel_economy.co2_primary for v in selected_vehicles if v.fuel_economy.co2_primary > 0]
        
        avg_mpg = sum(mpg_values) / len(mpg_values) if mpg_values else 0
        avg_co2 = sum(co2_values) / len(co2_values) if co2_values else 0
        
        # Create summary text
        summary_text = f"""
Summary of Selected Vehicles ({len(selected_vehicles)})

• Average Combined MPG: {avg_mpg:.1f}
• Average CO2 Emissions: {avg_co2:.1f} g/mile
        """
        
        # Add summary label
        ttk.Label(
            summary_frame,
            text=summary_text,
            font=("", 12, ""),
            wraplength=550,
            justify=tk.LEFT
        ).pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # TODO: Add more analysis tabs as needed
        
        # Close button
        close_button = ttk.Button(
            analysis_window,
            text="Close",
            command=analysis_window.destroy
        )
        close_button.pack(side=tk.BOTTOM, pady=10)
    
    def _select_all(self):
        """Select all items in the treeview."""
        self.tree.selection_set(self.tree.get_children())
    
    def _deselect_all(self):
        """Deselect all items in the treeview."""
        self.tree.selection_remove(self.tree.selection())
    
    def set_data(self, data):
        """
        Set the data for the results panel.
        
        Args:
            data: List of FleetVehicle objects or similar data
        """
        logger.info(f"🔧 DEBUG: ResultsPanel.set_data() called with {len(data) if data else 0} items")
        logger.info(f"🔧 DEBUG: set_data called from thread: {threading.current_thread().name}")
        
        try:
            self.data = data or []
            
            if self.data:
                # Log data breakdown for debugging
                success_count = sum(1 for v in self.data if hasattr(v, 'processing_success') and v.processing_success)
                failed_count = len(self.data) - success_count
                logger.info(f"🔧 DEBUG: ResultsPanel data breakdown - Success: {success_count}, Failed: {failed_count}")
                
                # Log first few items for debugging
                for i, item in enumerate(self.data[:3]):
                    if hasattr(item, 'vin') and hasattr(item, 'processing_success'):
                        logger.info(f"🔧 DEBUG: Item {i+1}: VIN={item.vin}, Success={item.processing_success}, Error='{getattr(item, 'processing_error', 'N/A')}'")
            
            # Update data display
            logger.info(f"🔧 DEBUG: Calling populate_data() to refresh display")
            self.populate_data()
            logger.info(f"🔧 DEBUG: populate_data() completed successfully")
            
            logger.info(f"🔧 DEBUG: Calling _update_summary()")
            self._update_summary()
            logger.info(f"🔧 DEBUG: _update_summary() completed successfully")
            
            # Trigger selection change callback
            if self.on_selection_change_callback:
                logger.info(f"🔧 DEBUG: Calling selection change callback")
                self.on_selection_change_callback([])
                logger.info(f"🔧 DEBUG: Selection change callback completed")
                
            logger.info(f"🔧 DEBUG: ResultsPanel.set_data() completed successfully")
            
        except Exception as e:
            logger.error(f"🔧 DEBUG: Exception in ResultsPanel.set_data(): {e}")
            logger.error(f"🔧 DEBUG: set_data exception type: {type(e).__name__}")
            import traceback
            logger.error(f"🔧 DEBUG: set_data traceback: {traceback.format_exc()}")
            raise  # Re-raise to see the full error chain
    
    def _update_summary(self):
        """Update the processing summary display."""
        if not self.data:
            self.summary_label.config(text="No data loaded")
            return
        
        total_count = len(self.data)
        successful_count = 0
        failed_count = 0
        
        # Analyze each vehicle to determine success/failure
        for vehicle in self.data:
            # Consider successful if basic vehicle info is present
            if (vehicle.vehicle_id.year and 
                vehicle.vehicle_id.make and 
                vehicle.vehicle_id.model):
                successful_count += 1
            else:
                failed_count += 1
        
        # Create summary text with color coding
        if failed_count == 0:
            summary_text = f"✓ {successful_count} successful, {total_count} total"
            text_color = "#2E8B57"  # Success green
        elif successful_count == 0:
            summary_text = f"✗ {failed_count} failed, {total_count} total"
            text_color = "#DC143C"  # Error red
        else:
            summary_text = f"⚠ {successful_count} successful, {failed_count} failed, {total_count} total"
            text_color = "#FF8C00"  # Warning orange
        
        self.summary_label.config(text=summary_text, foreground=text_color)
    
    def get_data(self):
        """
        Get the current data.
        
        Returns:
            List of FleetVehicle objects
        """
        return self.data
    
    def get_selected_data(self):
        """
        Get the currently selected data.
        
        Returns:
            List of selected FleetVehicle objects
        """
        selected_iids = self.tree.selection()
        return [self.data_map[iid] for iid in selected_iids if iid in self.data_map]
    
    def set_visible_columns(self, columns):
        """
        Set the visible columns.
        
        Args:
            columns: List of column IDs to display
        """
        # Validate columns
        valid_columns = [col for col in columns if col in self.all_columns_map]
        
        if not valid_columns:
            # Ensure at least some columns are visible
            valid_columns = DEFAULT_VISIBLE_COLUMNS
        
        self.visible_columns = valid_columns
        
        # Update the treeview
        self._update_columns()
        
        # Call callback
        if self.on_column_change_callback:
            self.on_column_change_callback(self.visible_columns)
    
    def get_visible_columns(self):
        """
        Get the currently visible columns.
        
        Returns:
            List of visible column IDs
        """
        return self.visible_columns
    
    def get_all_columns(self):
        """
        Get all available columns with their display names.
        
        Returns:
            List of (column_id, display_name) tuples
        """
        return [(col_id, display_name) for col_id, display_name in self.all_columns_map.items()]
    
    def _update_columns(self):
        """Update the treeview columns."""
        # Remember the current data
        current_data = self.data_map.values()
        
        # Clear the treeview
        self.tree.delete(*self.tree.get_children())
        
        # Configure new columns
        self.tree.config(columns=self.visible_columns)
        
        # Create new headers
        for col in self.visible_columns:
            display_name = self.all_columns_map.get(col, col)
            self.tree.heading(col, text=display_name, command=lambda c=col: self._sort_by_column(c))
            
            # Set a reasonable column width
            width = max(100, len(display_name) * 10)
            self.tree.column(col, width=width, minwidth=50)
        
        # Repopulate data
        self.populate_data()
    
    def populate_data(self):
        """
        Populate the treeview with data from self.data.
        Enhanced to handle both successful and failed vehicles properly.
        """
        logger.info(f"🔧 DEBUG: populate_data() called with {len(self.data)} items")
        logger.info(f"🔧 DEBUG: populate_data called from thread: {threading.current_thread().name}")
        
        try:
            # Clear existing data
            logger.info(f"🔧 DEBUG: Clearing existing treeview items")
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            self.data_map = {}
            logger.info(f"🔧 DEBUG: Treeview cleared, data_map reset")
            
            if not self.data:
                logger.info(f"🔧 DEBUG: No data to populate, returning early")
                return
            
            # Build column mapping (but don't update columns - that creates recursion)
            logger.info(f"🔧 DEBUG: Building column mapping")
            self._build_column_map()
            logger.info(f"🔧 DEBUG: Column mapping built")
            
            # Count different types of vehicles for logging
            successful_count = 0
            failed_count = 0
            missing_vin_count = 0
            empty_row_count = 0
            invalid_vin_count = 0
            
            logger.info(f"🔧 DEBUG: Starting to add vehicles to treeview")
            
            # Add data to treeview
            for i, vehicle in enumerate(self.data):
                try:
                    logger.info(f"🔧 DEBUG: Processing vehicle {i+1}/{len(self.data)}: VIN={getattr(vehicle, 'vin', 'Unknown')}")
                    
                    # Create row data
                    row_data = []
                    
                    for col_id in self.visible_columns:
                        if col_id in self.all_columns_map:
                            # Get value from vehicle using the proper method
                            value = self._get_vehicle_value(vehicle, col_id)
                            
                            # Special formatting for failed vehicles
                            if not vehicle.processing_success and col_id == "processing_error":
                                # Make error message more prominent
                                value = f"❌ {value}" if value else "❌ Unknown Error"
                            elif not vehicle.processing_success and col_id == "vin":
                                # Add error indicator to VIN column for failed vehicles
                                value = f"⚠️ {value}" if value else "⚠️ No VIN"
                            
                            row_data.append(str(value) if value is not None else "")
                        else:
                            row_data.append("")
                    
                    logger.info(f"🔧 DEBUG: Row data prepared for vehicle {i+1}, inserting into treeview")
                    
                    # Insert row
                    iid = self.tree.insert("", tk.END, values=row_data)
                    self.data_map[iid] = vehicle
                    
                    logger.info(f"🔧 DEBUG: Vehicle {i+1} inserted with iid={iid}, applying color coding")
                    
                    # Apply color coding for different types of vehicles
                    self._apply_row_color_coding(iid, vehicle)
                    
                    # Count vehicle types for debugging
                    if vehicle.processing_success:
                        successful_count += 1
                    else:
                        failed_count += 1
                        if "Missing VIN" in vehicle.processing_error:
                            missing_vin_count += 1
                        elif "Empty" in vehicle.processing_error:
                            empty_row_count += 1
                        elif "Invalid VIN" in vehicle.processing_error:
                            invalid_vin_count += 1
                    
                    logger.info(f"🔧 DEBUG: Vehicle {i+1} processed successfully")
                    
                except Exception as ve:
                    logger.error(f"🔧 DEBUG: Error adding vehicle {i+1} to tree: {ve}")
                    logger.error(f"🔧 DEBUG: Vehicle VIN: {getattr(vehicle, 'vin', 'Unknown')}")
                    logger.error(f"🔧 DEBUG: Vehicle success: {getattr(vehicle, 'processing_success', 'Unknown')}")
                    import traceback
                    logger.error(f"🔧 DEBUG: Vehicle error traceback: {traceback.format_exc()}")
            
            # Log summary
            logger.info(f"🔧 DEBUG: populate_data() vehicle processing complete:")
            logger.info(f"   Successful vehicles: {successful_count}")
            logger.info(f"   Failed vehicles: {failed_count}")
            logger.info(f"     - Missing VIN: {missing_vin_count}")
            logger.info(f"     - Empty rows: {empty_row_count}")
            logger.info(f"     - Invalid VIN: {invalid_vin_count}")
            
            # Apply current filter if any
            logger.info(f"🔧 DEBUG: Applying filters")
            self._apply_filter()
            logger.info(f"🔧 DEBUG: Filters applied successfully")
            
            logger.info(f"🔧 DEBUG: populate_data() completed successfully")
            
        except Exception as e:
            logger.error(f"🔧 DEBUG: Exception in populate_data(): {e}")
            logger.error(f"🔧 DEBUG: populate_data exception type: {type(e).__name__}")
            import traceback
            logger.error(f"🔧 DEBUG: populate_data traceback: {traceback.format_exc()}")
            raise  # Re-raise to see the full error chain
    
    def refresh(self):
        """Refresh the display."""
        # Clear any filter
        self.search_var.set("")
        
        # Repopulate data
        self.populate_data()
    
    def copy_selection(self):
        """Copy selected rows to clipboard in tab-delimited format."""
        # Get selected items
        selection = self.tree.selection()
        
        if not selection:
            return
        
        # Prepare clipboard text
        clipboard_lines = []
        
        # Add header row
        headers = [self.all_columns_map.get(col, col) for col in self.visible_columns]
        clipboard_lines.append("\t".join(headers))
        
        # Add selected rows
        for iid in selection:
            values = [self.tree.set(iid, col) for col in self.visible_columns]
            clipboard_lines.append("\t".join(values))
        
        # Join lines and copy to clipboard
        clipboard_text = "\n".join(clipboard_lines)
        self.clipboard_clear()
        self.clipboard_append(clipboard_text)
    
    def on_resize(self):
        """Handle resize events."""
        # Adjust column widths if needed
        pass

    def _export_for_excel(self):
        """Export data optimized for Excel analysis with preserved order and quality indicators."""
        import os
        import csv
        import datetime
        
        # Get currently displayed data in display order (preserves filtering and sorting)
        displayed_vehicles = []
        for iid in self.tree.get_children():
            if iid in self.data_map:
                displayed_vehicles.append(self.data_map[iid])
        
        if not displayed_vehicles:
            messagebox.showwarning("No Data", "No data to export.")
            return
        
        # Auto-generate filename with timestamp
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"fleet_analysis_{timestamp}.csv"
        
        # Ask user for save location with suggested filename
        from tkinter import filedialog
        filepath = filedialog.asksaveasfilename(
            title="Export Fleet Analysis for Excel",
            initialname=filename,
            defaultextension=".csv",
            filetypes=[
                ("CSV Files (Excel Compatible)", "*.csv"),
                ("All Files", "*.*")
            ]
        )
        
        if not filepath:
            return
        
        try:
            # Define Excel-optimized column order with data quality indicators
            excel_columns = [
                "VIN", "Processing Status", "Data Quality", "Year", "Make", "Model", 
                "FuelTypePrimary", "BodyClass", "GVWR",
                "MPG Combined", "MPG City", "MPG Highway", 
                "CO2 emissions", "Asset ID", "Department", "Location",
                "Odometer", "Annual Mileage", "Processing Error"
            ]
            
            # Add any additional custom fields that exist
            sample_row = displayed_vehicles[0].to_row_dict()
            for key in sample_row.keys():
                if key not in excel_columns and key not in ["co2A", "rangeA", "Assumed Vehicle (Text)", "Assumed Vehicle (ID)"]:
                    excel_columns.append(key)
            
            # Write to CSV with Excel-friendly formatting
            with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:  # UTF-8 BOM for Excel
                writer = csv.writer(f)
                
                # Write header row
                writer.writerow(excel_columns)
                
                # Write data rows in preserved order
                for vehicle in displayed_vehicles:
                    row_dict = vehicle.to_row_dict()
                    row = [row_dict.get(col, "") for col in excel_columns]
                    writer.writerow(row)
            
            # Show success message with helpful info
            success_msg = f"✅ Fleet analysis exported successfully!\n\n"
            success_msg += f"📁 File: {os.path.basename(filepath)}\n"
            success_msg += f"📊 Vehicles: {len(displayed_vehicles)}\n"
            success_msg += f"📋 Columns: {len(excel_columns)}\n\n"
            success_msg += "💡 Tips for Excel analysis:\n"
            success_msg += "• Data Quality column shows confidence (0-100%)\n"
            success_msg += "• Processing Status shows Success/Failed\n"
            success_msg += "• Sort by Data Quality to prioritize reliable data\n"
            success_msg += "• Filter by Processing Status to focus on successful matches"
            
            messagebox.showinfo("Export Complete", success_msg)
            
        except Exception as e:
            logger.error(f"Error exporting for Excel: {e}")
            messagebox.showerror(
                "Export Failed",
                f"Failed to export data:\n{str(e)}"
            )

    def _get_vehicle_value(self, vehicle: FleetVehicle, col_id: str) -> Any:
        """
        Get a field value from a vehicle object by column ID.
        
        Args:
            vehicle: FleetVehicle object
            col_id: Column identifier
            
        Returns:
            Field value or empty string if not found
        """
        # Get the row dictionary which has all the properly formatted data
        row_dict = vehicle.to_row_dict()
        
        # Try to get the value directly by column ID
        if col_id in row_dict:
            return row_dict[col_id]
        
        # If not found, try some common mappings
        column_mappings = {
            "vin": "VIN",
            "year": "Year", 
            "make": "Make",
            "model": "Model",
            "fuel_type": "FuelTypePrimary",
            "body_class": "BodyClass",
            "mpg_combined": "MPG Combined",
            "co2_emissions": "CO2 emissions",
            "commercial_category": "Commercial Category",
            "gvwr_lbs": "GVWR (lbs)",
            "data_quality": "Data Quality",
            "processing_status": "Processing Status",
            "annual_mileage": "Annual Mileage",
            "asset_id": "Asset ID",
            "department": "Department",
            "location": "Location",
            "odometer": "Odometer"
        }
        
        # Try the mapping
        mapped_key = column_mappings.get(col_id.lower(), col_id)
        if mapped_key in row_dict:
            return row_dict[mapped_key]
        
        # Still not found, try the vehicle's get_field method as last resort
        try:
            return vehicle.get_field(col_id)
        except:
            return ""