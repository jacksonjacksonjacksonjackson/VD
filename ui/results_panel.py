"""
results_panel.py

Panel for displaying and interacting with processed vehicle data in the
Fleet Electrification Analyzer.
"""

import logging
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List, Any, Optional, Callable, Tuple

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
        self.data_map = {}  # Maps treeview IIDs to data objects
        
        # Create UI components
        self._create_toolbar()
        self._create_treeview()
        
        # Populate data
        self.populate_data()
        
        # Create context menu
        self._create_context_menu()
    
    def _create_toolbar(self):
        """Create the toolbar with search and actions."""
        # Create toolbar frame
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        # Left side - search
        search_frame = ttk.Frame(toolbar)
        search_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT, padx=(0, 5))
        
        search_entry = ttk.Entry(
            search_frame, 
            textvariable=self.search_var,
            width=30
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
            values=["all", "VIN", "Make", "Model", "Year"],
            width=10,
            state="readonly"
        )
        search_fields.pack(side=tk.LEFT, padx=(0, 5))
        SimpleTooltip(search_fields, "Select fields to search in")
        
        # Bind search fields to apply filter
        search_fields.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())
        
        # Right side - actions
        actions_frame = ttk.Frame(toolbar)
        actions_frame.pack(side=tk.RIGHT)
        
        # Clear filter button
        clear_btn = ttk.Button(
            actions_frame,
            text="Clear Filter",
            command=self._clear_filter
        )
        clear_btn.pack(side=tk.RIGHT, padx=5)
        SimpleTooltip(clear_btn, "Clear search filter")
        
        # Refresh button
        refresh_btn = ttk.Button(
            actions_frame,
            text="Refresh",
            command=self.refresh
        )
        refresh_btn.pack(side=tk.RIGHT, padx=5)
        SimpleTooltip(refresh_btn, "Refresh data display")
        
        # Export button
        export_btn = ttk.Button(
            actions_frame,
            text="Export",
            command=self._export_data
        )
        export_btn.pack(side=tk.RIGHT, padx=5)
        SimpleTooltip(export_btn, "Export displayed data")
    
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
        
        # Configure row height and alternating colors
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)
        
        # Bind selection event
        self.tree.bind("<<TreeviewSelect>>", self._on_selection_change)
        
        # Bind double-click for details
        self.tree.bind("<Double-1>", self._show_details)
    
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
        """Apply search filter to the data."""
        search_text = self.search_var.get().lower().strip()
        search_field = self.search_fields_var.get()
        
        # Clear existing items
        for i in self.tree.get_children():
            self.tree.delete(i)
        
        # If no search text, show all data
        if not search_text:
            self.populate_data()
            return
        
        # Filter data
        if search_field == "all":
            # Search in all fields
            for vehicle in self.data:
                row_dict = vehicle.to_row_dict()
                
                # Convert values to strings for searching
                row_values = [str(v).lower() for v in row_dict.values()]
                
                # If search text appears in any value, include this row
                if any(search_text in val for val in row_values):
                    values = [row_dict.get(col, "") for col in self.visible_columns]
                    iid = self.tree.insert("", tk.END, values=values)
                    self.data_map[iid] = vehicle
        else:
            # Search in specific field
            for vehicle in self.data:
                row_dict = vehicle.to_row_dict()
                
                # Get value for the search field
                field_value = str(row_dict.get(search_field, "")).lower()
                
                # If search text appears in the field value, include this row
                if search_text in field_value:
                    values = [row_dict.get(col, "") for col in self.visible_columns]
                    iid = self.tree.insert("", tk.END, values=values)
                    self.data_map[iid] = vehicle
    
    def _clear_filter(self):
        """Clear the search filter."""
        self.search_var.set("")
        self.search_fields_var.set("all")
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
        Set the data to display.
        
        Args:
            data: List of FleetVehicle objects
        """
        self.data = data or []
        
        # Rebuild column map
        self._build_column_map()
        
        # Refresh the display
        self.refresh()
    
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
        """Populate the treeview with data."""
        # Clear existing items and data map
        self.tree.delete(*self.tree.get_children())
        self.data_map.clear()
        
        # Check if we have data
        if not self.data:
            return
        
        # Add each vehicle to the treeview
        for vehicle in self.data:
            row_dict = vehicle.to_row_dict()
            values = [row_dict.get(col, "") for col in self.visible_columns]
            
            # Insert into treeview
            iid = self.tree.insert("", tk.END, values=values)
            
            # Map iid to vehicle object
            self.data_map[iid] = vehicle
    
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