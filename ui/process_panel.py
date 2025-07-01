"""
process_panel.py

Panel for processing vehicle data in the Fleet Electrification Analyzer.
Handles input/output file selection, processing options, and log display.
"""

import os
import csv
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Any, Optional, Callable
import datetime
import json

# Import drag-and-drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DRAG_DROP_AVAILABLE = True
except ImportError:
    DRAG_DROP_AVAILABLE = False
    logging.warning("tkinterdnd2 not available. Drag-and-drop functionality disabled.")

from settings import (
    MAX_THREADS,
    PRIMARY_HEX_1,
    PRIMARY_HEX_2,
    PRIMARY_HEX_3,
    SECONDARY_HEX_1
)
from utils import SimpleTooltip, ScrollableFrame
from data.processor import CsvFileValidator, FileValidationResult

# Set up module logger
logger = logging.getLogger(__name__)

class ProcessPanel(ttk.Frame):
    """
    Panel for processing vehicle data.
    Handles input/output file selection, processing options, and log display.
    """
    
    def __init__(self, parent, on_process=None, on_stop=None, on_log=None):
        """
        Initialize the process panel.
        
        Args:
            parent: Parent widget
            on_process: Callback when processing starts
            on_stop: Callback when processing is stopped
            on_log: Callback when a log message is added
        """
        super().__init__(parent)
        
        # Store callbacks
        self.on_process_callback = on_process
        self.on_stop_callback = on_stop
        self.on_log_callback = on_log
        
        # Initialize variables
        self.input_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.max_threads_var = tk.IntVar(value=MAX_THREADS)
        self.skip_existing_var = tk.BooleanVar(value=True)
        self.auto_filename_var = tk.BooleanVar(value=True)
        
        # File validation state
        self.current_validation: Optional[FileValidationResult] = None
        
        # Processing history for reprocessing
        self.processing_history: List[Dict[str, Any]] = []
        self._load_processing_history()
        
        # Enhanced progress tracking
        self.current_progress = {"current": 0, "total": 0, "stage": ""}
        
        # Create UI components
        self._create_input_section()
        self._create_preview_section()
        self._create_options_section()
        self._create_reprocessing_section()
        self._create_actions_section()
        self._create_log_section()
        
        # Set default output file
        self._update_default_output()
        
        # Trace auto filename setting
        self.auto_filename_var.trace_add("write", lambda *args: self._update_default_output())
    
    def _create_input_section(self):
        """Create the input/output file selection section with drag-and-drop support."""
        # Create frame
        input_frame = ttk.LabelFrame(self, text="Input/Output")
        input_frame.pack(fill=tk.X, padx=10, pady=(10, 5), ipady=5)
        
        # Input file row with drag-and-drop zone
        input_container = ttk.Frame(input_frame)
        input_container.grid(row=0, column=0, columnspan=4, sticky=tk.W+tk.E, padx=5, pady=5)
        
        ttk.Label(input_container, text="Input CSV:").pack(side=tk.LEFT, padx=(0, 5))
        
        # Create drag-and-drop zone
        self.drop_zone = tk.Frame(
            input_container,
            bg="#f0f8ff",
            relief=tk.SUNKEN,
            bd=2,
            height=60
        )
        self.drop_zone.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        # Drop zone label
        self.drop_zone_label = tk.Label(
            self.drop_zone,
            text="📁 Drag & Drop CSV file here\nor click Browse to select",
            bg="#f0f8ff",
            fg="#666666",
            font=("Segoe UI", 9),
            justify=tk.CENTER
        )
        self.drop_zone_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        # Enable drag-and-drop if available
        if DRAG_DROP_AVAILABLE:
            self._setup_drag_drop()
        
        # Browse button
        input_browse_btn = ttk.Button(
            input_container, 
            text="Browse...", 
            command=self._browse_input
        )
        input_browse_btn.pack(side=tk.LEFT, padx=5)
        SimpleTooltip(input_browse_btn, "Select CSV file with VINs")
        
        # Download sample button
        sample_btn = ttk.Button(
            input_container,
            text="Sample CSV",
            command=self._download_sample_csv
        )
        sample_btn.pack(side=tk.LEFT, padx=5)
        SimpleTooltip(sample_btn, "Download a sample CSV template")
        
        # Input file display
        input_entry = ttk.Entry(
            input_frame, 
            textvariable=self.input_file_var, 
            width=60
        )
        input_entry.grid(row=1, column=0, columnspan=4, sticky=tk.W+tk.E, padx=5, pady=5)
        SimpleTooltip(input_entry, "Path to input CSV file with VINs")
        
        # Output file row with auto-generation option
        output_container = ttk.Frame(input_frame)
        output_container.grid(row=2, column=0, columnspan=4, sticky=tk.W+tk.E, padx=5, pady=5)
        
        ttk.Label(output_container, text="Output CSV:").pack(side=tk.LEFT, padx=(0, 5))
        
        # Auto-generate filename checkbox
        auto_check = ttk.Checkbutton(
            output_container,
            text="Auto-generate filename",
            variable=self.auto_filename_var
        )
        auto_check.pack(side=tk.LEFT, padx=(10, 5))
        SimpleTooltip(auto_check, "Automatically generate output filename with timestamp")
        
        output_browse_btn = ttk.Button(
            output_container, 
            text="Browse...", 
            command=self._browse_output
        )
        output_browse_btn.pack(side=tk.RIGHT, padx=5)
        SimpleTooltip(output_browse_btn, "Select output CSV location")
        
        # Output file display
        output_entry = ttk.Entry(
            input_frame, 
            textvariable=self.output_file_var, 
            width=60
        )
        output_entry.grid(row=3, column=0, columnspan=4, sticky=tk.W+tk.E, padx=5, pady=5)
        SimpleTooltip(output_entry, "Path to save processed results")
        
        # Configure grid to expand properly
        input_frame.columnconfigure(0, weight=1)
        
        # Bind event to update default output when input changes
        self.input_file_var.trace_add("write", lambda *args: self._update_default_output())
        self.input_file_var.trace_add("write", lambda *args: self._validate_input_file())
    
    def _setup_drag_drop(self):
        """Set up drag-and-drop functionality for the drop zone."""
        if not DRAG_DROP_AVAILABLE:
            return
        
        # Register drop zone for file drops
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind('<<Drop>>', self._handle_file_drop)
        
        # Visual feedback for drag events
        self.drop_zone.dnd_bind('<<DragEnter>>', self._drag_enter)
        self.drop_zone.dnd_bind('<<DragLeave>>', self._drag_leave)
        
        # Make drop zone clickable as well
        self.drop_zone.bind("<Button-1>", lambda e: self._browse_input())
        self.drop_zone_label.bind("<Button-1>", lambda e: self._browse_input())
    
    def _handle_file_drop(self, event):
        """Handle file drop event."""
        # Get the dropped file path
        files = event.data.split()
        if files:
            filepath = files[0].replace("{", "").replace("}", "")
            if filepath.lower().endswith('.csv'):
                self.input_file_var.set(filepath)
                self._update_drop_zone_display(filepath)
                self.add_log(f"File dropped: {os.path.basename(filepath)}")
            else:
                messagebox.showerror("Invalid File", "Please drop a CSV file.")
        
        # Reset drop zone visual state
        self._drag_leave(event)
    
    def _drag_enter(self, event):
        """Visual feedback when drag enters the drop zone."""
        self.drop_zone.config(bg="#e6f3ff", relief=tk.RAISED)
        self.drop_zone_label.config(
            text="🎯 Drop CSV file here",
            bg="#e6f3ff",
            fg="#0066cc",
            font=("Segoe UI", 10, "bold")
        )
    
    def _drag_leave(self, event):
        """Reset visual state when drag leaves the drop zone."""
        self.drop_zone.config(bg="#f0f8ff", relief=tk.SUNKEN)
        if self.input_file_var.get().strip():
            filename = os.path.basename(self.input_file_var.get())
            self.drop_zone_label.config(
                text=f"✅ {filename}",
                bg="#f0f8ff",
                fg="#006600",
                font=("Segoe UI", 9)
            )
        else:
            self.drop_zone_label.config(
                text="📁 Drag & Drop CSV file here\nor click Browse to select",
                bg="#f0f8ff",
                fg="#666666",
                font=("Segoe UI", 9)
            )
    
    def _update_drop_zone_display(self, filepath):
        """Update drop zone display with selected file."""
        filename = os.path.basename(filepath)
        self.drop_zone_label.config(
            text=f"✅ {filename}",
            bg="#f0f8ff",
            fg="#006600",
            font=("Segoe UI", 9)
        )
    
    def _create_preview_section(self):
        """Create the file preview and validation section."""
        # Create collapsible frame
        self.preview_frame = ttk.LabelFrame(self, text="File Preview & Validation")
        # Initially hidden
        
        # Validation status row
        status_frame = ttk.Frame(self.preview_frame)
        status_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.status_label = ttk.Label(status_frame, text="No file selected", foreground="gray")
        self.status_label.pack(side=tk.LEFT)
        
        # File info frame  
        info_frame = ttk.LabelFrame(self.preview_frame, text="File Information")
        info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Info labels
        self.info_text = tk.Text(info_frame, height=4, wrap=tk.WORD, state=tk.DISABLED, bg="#f5f5f5")
        self.info_text.pack(fill=tk.X, padx=5, pady=5)
        
        # Column mapping frame
        mapping_frame = ttk.LabelFrame(self.preview_frame, text="Column Mapping")
        mapping_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.mapping_text = tk.Text(mapping_frame, height=3, wrap=tk.WORD, state=tk.DISABLED, bg="#f5f5f5")
        self.mapping_text.pack(fill=tk.X, padx=5, pady=5)
        
        # Data preview frame
        data_frame = ttk.LabelFrame(self.preview_frame, text="Data Preview (First 5 Rows)")
        data_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create treeview for data preview
        self.preview_tree = ttk.Treeview(data_frame, height=6)
        preview_scrollbar = ttk.Scrollbar(data_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=preview_scrollbar.set)
        
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def _create_options_section(self):
        """Create the processing options section."""
        # Create frame
        options_frame = ttk.LabelFrame(self, text="Processing Options")
        options_frame.pack(fill=tk.X, padx=10, pady=5, ipady=5)
        
        # Max threads
        ttk.Label(options_frame, text="Max Threads:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        threads_spinbox = ttk.Spinbox(
            options_frame,
            from_=1,
            to=32,
            textvariable=self.max_threads_var,
            width=5
        )
        threads_spinbox.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        SimpleTooltip(threads_spinbox, "Maximum number of parallel processing threads")
        
        # Skip existing checkbox
        skip_checkbox = ttk.Checkbutton(
            options_frame,
            text="Skip existing VINs in output file",
            variable=self.skip_existing_var
        )
        skip_checkbox.grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        SimpleTooltip(skip_checkbox, "Skip processing VINs that already exist in the output file")
    
    def _create_reprocessing_section(self):
        """Create the reprocessing section with history display."""
        # Create frame
        reprocessing_frame = ttk.LabelFrame(self, text="Quick Reprocessing")
        reprocessing_frame.pack(fill=tk.X, padx=10, pady=5, ipady=5)
        
        # Top row with reprocess button and history info
        top_row = ttk.Frame(reprocessing_frame)
        top_row.pack(fill=tk.X, padx=5, pady=5)
        
        # Reprocessing button
        self.reprocess_btn = ttk.Button(
            top_row,
            text="🔄 Reprocess Last File",
            command=self._reprocess_file,
            style="Accent.TButton"
        )
        self.reprocess_btn.pack(side=tk.LEFT, padx=(0, 10))
        SimpleTooltip(self.reprocess_btn, "Reprocess the last processed file with a new timestamp")
        
        # History dropdown
        ttk.Label(top_row, text="Recent:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.history_var = tk.StringVar()
        self.history_combo = ttk.Combobox(
            top_row,
            textvariable=self.history_var,
            width=40,
            state="readonly"
        )
        self.history_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.history_combo.bind("<<ComboboxSelected>>", self._on_history_selection)
        SimpleTooltip(self.history_combo, "Select from recent processing jobs")
        
        # Reprocess selected button
        reprocess_selected_btn = ttk.Button(
            top_row,
            text="Process",
            command=self._reprocess_selected
        )
        reprocess_selected_btn.pack(side=tk.RIGHT, padx=5)
        SimpleTooltip(reprocess_selected_btn, "Reprocess selected file")
        
        # Update history display
        self._update_history_display()
        
        # Update reprocess button state
        self._update_reprocess_button_state()
    
    def _update_history_display(self):
        """Update the history dropdown with recent processing jobs."""
        if not hasattr(self, 'history_combo'):
            return
        
        history_items = []
        for job in reversed(self.processing_history[-5:]):  # Show last 5 jobs
            timestamp = job.get('timestamp', '')
            filename = job.get('input_filename', 'Unknown')
            
            # Format timestamp for display
            try:
                dt = datetime.datetime.fromisoformat(timestamp)
                time_str = dt.strftime("%m/%d %H:%M")
            except:
                time_str = "Unknown"
            
            display_text = f"{time_str} - {filename}"
            history_items.append(display_text)
        
        self.history_combo['values'] = history_items
        
        # Select most recent if available
        if history_items:
            self.history_combo.set(history_items[0])
    
    def _update_reprocess_button_state(self):
        """Update the reprocess button enabled state based on history."""
        if hasattr(self, 'reprocess_btn'):
            if self.processing_history:
                self.reprocess_btn.config(state=tk.NORMAL)
            else:
                self.reprocess_btn.config(state=tk.DISABLED)
    
    def _on_history_selection(self, event):
        """Handle history selection from dropdown."""
        selection = self.history_combo.get()
        if selection:
            # Find the corresponding job
            selected_index = self.history_combo['values'].index(selection)
            reversed_history = list(reversed(self.processing_history[-5:]))
            
            if selected_index < len(reversed_history):
                job = reversed_history[selected_index]
                
                # Preview the job details in log
                self.add_log(f"Selected: {job.get('input_filename', 'Unknown')} "
                           f"({job.get('max_threads', 'Unknown')} threads)")
    
    def _reprocess_selected(self):
        """Reprocess the selected file from history."""
        selection = self.history_combo.get()
        if not selection:
            messagebox.showinfo("No Selection", "Please select a file from the recent history.")
            return
        
        # Find the corresponding job
        try:
            selected_index = self.history_combo['values'].index(selection)
            reversed_history = list(reversed(self.processing_history[-5:]))
            
            if selected_index < len(reversed_history):
                job = reversed_history[selected_index]
                self._reprocess_job(job)
            else:
                messagebox.showerror("Error", "Could not find the selected job.")
        except Exception as e:
            messagebox.showerror("Error", f"Error reprocessing selected file: {e}")
    
    def _reprocess_job(self, job):
        """Reprocess a specific job."""
        # Check if input file still exists
        input_file = job.get("input_file", "")
        if not os.path.exists(input_file):
            messagebox.showerror(
                "File Not Found", 
                f"The input file no longer exists:\n{input_file}\n\n"
                "Please select the file manually."
            )
            return
        
        # Set the parameters
        self.input_file_var.set(input_file)
        self.max_threads_var.set(job.get("max_threads", MAX_THREADS))
        self.skip_existing_var.set(job.get("skip_existing", True))
        
        # Auto-generate new output filename with fresh timestamp
        self._update_default_output()
        
        # Log the reprocessing
        self.add_log(f"🔄 Reprocessing: {os.path.basename(input_file)}")
        
        # Start processing immediately
        self._start_processing()
    
    def _create_actions_section(self):
        """Create the actions section with control buttons."""
        # Create frame
        actions_frame = ttk.Frame(self)
        actions_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Start button
        self.start_button = ttk.Button(
            actions_frame,
            text="Start Processing",
            command=self._start_processing,
            style="Accent.TButton"
        )
        self.start_button.pack(side=tk.LEFT, padx=5, pady=5)
        SimpleTooltip(self.start_button, "Begin processing VINs")
        
        # Stop button (initially disabled)
        self.stop_button = ttk.Button(
            actions_frame,
            text="Stop",
            command=self._stop_processing,
            state=tk.DISABLED
        )
        self.stop_button.pack(side=tk.LEFT, padx=5, pady=5)
        SimpleTooltip(self.stop_button, "Stop processing (can't be resumed)")
        
        # Clear log button
        clear_log_btn = ttk.Button(
            actions_frame,
            text="Clear Log",
            command=self._clear_log
        )
        clear_log_btn.pack(side=tk.RIGHT, padx=5, pady=5)
        SimpleTooltip(clear_log_btn, "Clear the log display")
    
    def _create_log_section(self):
        """Create the log display section."""
        # Create frame
        log_frame = ttk.LabelFrame(self, text="Processing Log")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 10), ipady=5)
        
        # Create text widget with scrollbar
        self.log_text = tk.Text(
            log_frame,
            wrap=tk.WORD,
            width=80,
            height=20,
            bg="#f5f5f5",
            state=tk.DISABLED
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Configure tags for different log levels
        self.log_text.tag_configure("info", foreground="black")
        self.log_text.tag_configure("warning", foreground="orange")
        self.log_text.tag_configure("error", foreground="red")
        self.log_text.tag_configure("timestamp", foreground="gray")
        
        # Add scrollbar
        log_scrollbar = ttk.Scrollbar(log_frame)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Connect scrollbar to text widget
        self.log_text.config(yscrollcommand=log_scrollbar.set)
        log_scrollbar.config(command=self.log_text.yview)
        
        # Add progress bar
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            self,
            variable=self.progress_var,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        # Create right-click menu for log
        self.log_menu = tk.Menu(self, tearoff=0)
        self.log_menu.add_command(label="Copy", command=self.copy_selection)
        self.log_menu.add_command(label="Select All", command=self._select_all_log)
        self.log_menu.add_separator()
        self.log_menu.add_command(label="Clear Log", command=self._clear_log)
        
        # Bind right-click to show menu
        self.log_text.bind("<Button-3>", self._show_log_menu)
    
    def _browse_input(self):
        """Open file dialog to select input CSV."""
        filepath = filedialog.askopenfilename(
            title="Select Input CSV",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        
        if filepath:
            self.input_file_var.set(filepath)
            # Validation will be triggered by the trace
    
    def _browse_output(self):
        """Open file dialog to select output CSV location."""
        filepath = filedialog.asksaveasfilename(
            title="Select Output CSV",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        
        if filepath:
            self.output_file_var.set(filepath)
    
    def _update_default_output(self):
        """Update default output path based on input path with auto-generation and timestamps."""
        input_path = self.input_file_var.get().strip()
        
        if input_path and os.path.exists(input_path) and self.auto_filename_var.get():
            # Generate output filename with timestamp
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_dir = os.path.dirname(input_path)
            
            # Add timestamp for uniqueness
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"{base_name}_fleet_analysis_{timestamp}.csv"
            output_path = os.path.join(output_dir, output_filename)
            
            # Set the output path
            self.output_file_var.set(output_path)
            
            # Update drop zone display if file is selected
            if hasattr(self, 'drop_zone_label'):
                self._update_drop_zone_display(input_path)
    
    def _validate_input_file(self):
        """Validate the selected input file and update preview."""
        input_path = self.input_file_var.get().strip()
        
        if not input_path:
            self._hide_preview()
            return
        
        if not os.path.exists(input_path):
            self._update_validation_status("File not found", "error")
            self._hide_preview()
            return
        
        try:
            # Validate the file
            validator = CsvFileValidator(input_path)
            self.current_validation = validator.validate_and_preview()
            
            # Ensure all required attributes exist
            if not hasattr(self.current_validation, 'valid'):
                raise ValueError("Invalid validation result: missing 'valid' attribute")
            
            # Update the UI with validation results
            self._update_preview_display()
            
        except Exception as e:
            logger.error(f"Error validating file: {e}")
            self._update_validation_status(f"Error validating file: {str(e)}", "error")
            self._hide_preview()
            
            # Clear current validation to prevent further errors
            self.current_validation = None
    
    def _update_validation_status(self, message: str, level: str = "info"):
        """Update the validation status label."""
        colors = {
            "info": "black",
            "warning": "orange", 
            "error": "red",
            "success": "green"
        }
        
        self.status_label.config(text=message, foreground=colors.get(level, "black"))
    
    def _hide_preview(self):
        """Hide the preview section."""
        self.preview_frame.pack_forget()
        self.current_validation = None
    
    def _show_preview(self):
        """Show the preview section."""
        # Use a safer positioning method instead of assuming child widget order
        try:
            # Check if preview_frame exists and is not already packed
            if not hasattr(self, 'preview_frame'):
                logger.warning("Preview frame not found, skipping preview display")
                return
            
            # Check if already visible to avoid re-packing
            if self.preview_frame.winfo_viewable():
                return
            
            # Try to place after the input section if it exists
            input_widgets = []
            if hasattr(self, 'master') and self.master:
                try:
                    input_widgets = [child for child in self.master.children.values() 
                                   if isinstance(child, ttk.LabelFrame) and 
                                   hasattr(child, 'cget') and 
                                   child.cget('text') == 'Input/Output']
                except Exception as e:
                    logger.debug(f"Error finding input widgets: {e}")
            
            if input_widgets:
                self.preview_frame.pack(fill=tk.X, padx=10, pady=5, after=input_widgets[0])
            else:
                # Fallback: just pack normally at the end
                self.preview_frame.pack(fill=tk.X, padx=10, pady=5)
                
        except Exception as e:
            logger.debug(f"Error positioning preview frame: {e}")
            # Ultimate fallback: pack without positioning
            try:
                self.preview_frame.pack(fill=tk.X, padx=10, pady=5)
            except Exception as e2:
                logger.error(f"Failed to show preview frame: {e2}")
    
    def _update_preview_display(self):
        """Update the preview display with validation results."""
        if not self.current_validation:
            self._hide_preview()
            return
        
        # Show the preview section
        self._show_preview()
        
        # Update validation status
        if self.current_validation.valid:
            status_msg = f"✅ Valid CSV - {self.current_validation.valid_vins} valid VINs found"
            if self.current_validation.invalid_vins > 0:
                status_msg += f" ({self.current_validation.invalid_vins} invalid)"
            self._update_validation_status(status_msg, "success")
        else:
            self._update_validation_status(f"❌ {self.current_validation.error_message}", "error")
        
        # Update file information
        self._update_info_display()
        
        # Update column mapping
        self._update_mapping_display()
        
        # Update data preview
        self._update_data_preview()
    
    def _update_info_display(self):
        """Update the file information display."""
        if not self.current_validation:
            return
        
        info_lines = []
        info_lines.append(f"📄 File: {os.path.basename(self.input_file_var.get())}")
        info_lines.append(f"📊 Rows: {self.current_validation.total_rows}")
        info_lines.append(f"📑 Columns: {self.current_validation.column_count}")
        info_lines.append(f"🔤 Encoding: {self.current_validation.detected_encoding}")
        
        if self.current_validation.valid:
            info_lines.append(f"✅ Valid VINs: {self.current_validation.valid_vins}")
            if self.current_validation.invalid_vins > 0:
                info_lines.append(f"❌ Invalid VINs: {self.current_validation.invalid_vins}")
        
        # Add warnings
        for warning in self.current_validation.warning_messages:
            info_lines.append(f"⚠️ {warning}")
        
        # Update the text widget
        self.info_text.config(state=tk.NORMAL)
        self.info_text.delete(1.0, tk.END)
        self.info_text.insert(1.0, "\n".join(info_lines))
        self.info_text.config(state=tk.DISABLED)
    
    def _update_mapping_display(self):
        """Update the column mapping display."""
        if not self.current_validation:
            return
        
        mapping_lines = []
        
        if self.current_validation.vin_column:
            mapping_lines.append(f"🔑 VIN Column: '{self.current_validation.vin_column}'")
        
        # Show fleet management fields that were auto-mapped
        if self.current_validation.mapped_columns:
            mapping_lines.append(f"📊 Fleet Management Fields Detected:")
            for original_col, standard_field in self.current_validation.mapped_columns.items():
                mapping_lines.append(f"   • '{original_col}' → {standard_field}")
        
        # Show custom columns that will be preserved
        if self.current_validation.unmapped_columns:
            unmapped_preview = ", ".join(self.current_validation.unmapped_columns[:3])
            if len(self.current_validation.unmapped_columns) > 3:
                unmapped_preview += f" ... (+{len(self.current_validation.unmapped_columns) - 3} more)"
            mapping_lines.append(f"📋 Custom Columns (preserved as-is): {unmapped_preview}")
        
        if not self.current_validation.additional_columns:
            mapping_lines.append("📋 Additional Columns: None")
        
        # Show summary of fleet management capabilities
        if self.current_validation.fleet_management_fields:
            field_summary = ", ".join(self.current_validation.fleet_management_fields[:5])
            if len(self.current_validation.fleet_management_fields) > 5:
                field_summary += f" ... (+{len(self.current_validation.fleet_management_fields) - 5} more)"
            mapping_lines.append(f"🎯 Auto-mapped Fields: {field_summary}")
        
        # Show sample VINs if available
        if self.current_validation.sample_valid_vins:
            sample_vins = ", ".join(self.current_validation.sample_valid_vins[:3])
            mapping_lines.append(f"✅ Sample Valid VINs: {sample_vins}")
        
        if self.current_validation.sample_invalid_vins:
            sample_invalid = "; ".join(self.current_validation.sample_invalid_vins[:2])
            mapping_lines.append(f"❌ Sample Invalid VINs: {sample_invalid}")
        
        # Update the text widget
        self.mapping_text.config(state=tk.NORMAL)
        self.mapping_text.delete(1.0, tk.END)
        self.mapping_text.insert(1.0, "\n".join(mapping_lines))
        self.mapping_text.config(state=tk.DISABLED)
    
    def _update_data_preview(self):
        """Update the data preview table."""
        if not self.current_validation or not self.current_validation.sample_rows:
            # Clear the treeview
            for item in self.preview_tree.get_children():
                self.preview_tree.delete(item)
            return
        
        # Clear existing items
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        
        # Get first 5 rows for preview
        sample_rows = self.current_validation.sample_rows[:5]
        if not sample_rows:
            return
        
        # Set up columns
        columns = self.current_validation.columns
        self.preview_tree["columns"] = columns
        self.preview_tree["show"] = "headings"
        
        # Configure column headers and widths
        for col in columns:
            self.preview_tree.heading(col, text=col)
            # Adjust column width based on content
            max_width = max(len(col) * 8, 100)  # Minimum 100px
            self.preview_tree.column(col, width=max_width, minwidth=80)
        
        # Add sample data
        for row in sample_rows:
            values = [row.get(col, "") for col in columns]
            # Truncate long values for display
            display_values = [str(val)[:50] + "..." if len(str(val)) > 50 else str(val) for val in values]
            self.preview_tree.insert("", tk.END, values=display_values)
    
    def _start_processing(self):
        """Start the vehicle data processing with enhanced tracking."""
        # Store start time for progress estimates
        import time
        self._start_time = time.time()
        
        # Get file paths
        input_file = self.input_file_var.get().strip()
        output_file = self.output_file_var.get().strip()
        
        # Validate inputs
        if not input_file:
            messagebox.showerror("Error", "Please select an input CSV file.")
            return
        
        if not output_file:
            messagebox.showerror("Error", "Please specify an output CSV file path.")
            return
        
        # Check file validation
        if not self.current_validation or not self.current_validation.valid:
            if not self.current_validation:
                # Force validation
                self._validate_input_file()
            
            if not self.current_validation or not self.current_validation.valid:
                error_msg = "The selected CSV file has validation errors:\n\n"
                if self.current_validation and self.current_validation.error_message:
                    error_msg += self.current_validation.error_message
                else:
                    error_msg += "Unknown validation error"
                
                error_msg += "\n\nPlease fix the file issues or select a different file."
                messagebox.showerror("File Validation Error", error_msg)
                return
        
        # Show warning for files with invalid VINs
        if self.current_validation.invalid_vins > 0:
            invalid_percent = (self.current_validation.invalid_vins / self.current_validation.total_rows) * 100
            warning_msg = f"Your file contains {self.current_validation.invalid_vins} invalid VINs "
            warning_msg += f"({invalid_percent:.1f}% of total).\n\n"
            warning_msg += "These VINs will appear in results with error messages.\n\n"
            warning_msg += "Do you want to continue processing?"
            
            if not messagebox.askyesno("Invalid VINs Detected", warning_msg):
                return
        
        # Add to processing history before starting
        self._add_to_processing_history(
            input_file, 
            output_file, 
            self.max_threads_var.get(),
            self.skip_existing_var.get()
        )
        
        # Show enhanced processing info
        if self.current_validation.total_rows > 500:
            info_msg = f"Processing {self.current_validation.total_rows} VINs with {self.max_threads_var.get()} threads.\n"
            info_msg += "This may take several minutes. Progress will be shown below.\n\n"
            info_msg += "💡 Tip: You can process other files while this runs using the Reprocess button!"
            messagebox.showinfo("Processing Started", info_msg)
        
        # Call the processing callback
        if self.on_process_callback:
            # Bundle processing options into a dictionary
            options = {
                "max_threads": self.max_threads_var.get(),
                "skip_existing": self.skip_existing_var.get()
            }
            # Add detailed logging before callback
            self.add_log(f"🔧 DEBUG: Calling processing callback with args:", "info")
            self.add_log(f"    input_file: {input_file}", "info")
            self.add_log(f"    output_file: {output_file}", "info")
            self.add_log(f"    options: {options}", "info")
            
            try:
                self.on_process_callback(input_file, output_file, options)
                self.add_log(f"✅ DEBUG: Processing callback completed successfully", "info")
            except Exception as e:
                self.add_log(f"❌ DEBUG: Processing callback failed: {e}", "error")
                messagebox.showerror("Processing Error", f"Failed to start processing: {e}")
                return
        
        # Update UI state
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        
        # Enhanced logging
        self.add_log(f"🚀 Starting processing: {self.current_validation.valid_vins} valid VINs detected")
        self.add_log(f"📂 Input: {os.path.basename(input_file)}")
        self.add_log(f"💾 Output: {os.path.basename(output_file)}")
        self.add_log(f"⚙️ Threads: {self.max_threads_var.get()}")
        
        if self.current_validation.invalid_vins > 0:
            self.add_log(f"⚠️ Warning: {self.current_validation.invalid_vins} invalid VINs will be included with errors", "warning")
    
    def _stop_processing(self):
        """Stop the processing pipeline."""
        # Add log entry
        self.add_log("Stopping processing...", level="warning")
        
        # Call the callback
        if self.on_stop_callback:
            self.on_stop_callback()
    
    def _clear_log(self):
        """Clear the log display."""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def _show_log_menu(self, event):
        """Show the context menu for the log text widget."""
        self.log_menu.tk_popup(event.x_root, event.y_root)
    
    def _select_all_log(self):
        """Select all text in the log widget."""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.tag_add(tk.SEL, "1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.log_text.focus_set()
    
    def set_input_file(self, filepath):
        """
        Set the input file path.
        
        Args:
            filepath: Path to the input file
        """
        self.input_file_var.set(filepath)
    
    def set_max_threads(self, max_threads):
        """
        Set the maximum number of processing threads.
        
        Args:
            max_threads: Maximum number of threads
        """
        self.max_threads_var.set(max_threads)
    
    def add_log(self, message, level="info"):
        """
        Add a message to the log display.
        
        Args:
            message: Message to add
            level: Log level ("info", "warning", "error")
        """
        # Get timestamp
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        
        # Enable text widget for editing
        self.log_text.config(state=tk.NORMAL)
        
        # Insert timestamp and message with appropriate tags
        self.log_text.insert(tk.END, f"[{timestamp}] ", "timestamp")
        self.log_text.insert(tk.END, f"{message}\n", level)
        
        # Scroll to end
        self.log_text.see(tk.END)
        
        # Disable text widget again
        self.log_text.config(state=tk.DISABLED)
        
        # Call the log callback if provided
        if self.on_log_callback:
            self.on_log_callback(message)
    
    def update_progress(self, current, total, stage="Processing"):
        """
        Update the progress bar with enhanced status messages.
        
        Args:
            current: Current progress value
            total: Maximum progress value
            stage: Current processing stage
        """
        if total <= 0:
            return
        
        # Store current progress
        self.current_progress = {"current": current, "total": total, "stage": stage}
        
        # Calculate percentage
        percentage = (current / total) * 100
        
        # Update progress bar
        self.progress_bar.config(maximum=total)
        self.progress_var.set(current)
        
        # Create enhanced status message
        status_msg = f"{stage}: {current}/{total} ({percentage:.1f}%)"
        
        # Add time estimates for longer jobs
        if total > 100 and current > 0:
            elapsed_ratio = current / total
            if elapsed_ratio > 0.1:  # Only show estimates after 10% completion
                import time
                if not hasattr(self, '_start_time'):
                    self._start_time = time.time()
                
                elapsed_time = time.time() - self._start_time
                estimated_total = elapsed_time / elapsed_ratio
                remaining_time = estimated_total - elapsed_time
                
                if remaining_time > 60:
                    remaining_str = f"{int(remaining_time // 60)}m {int(remaining_time % 60)}s"
                else:
                    remaining_str = f"{int(remaining_time)}s"
                
                status_msg += f" | ETA: {remaining_str}"
        
        # Update progress text in log
        if self.on_log_callback:
            self.on_log_callback(status_msg)
    
    def processing_complete(self):
        """Called when processing is complete with enhanced messaging."""
        # Reset start time for next run
        if hasattr(self, '_start_time'):
            delattr(self, '_start_time')
        
        # Update UI
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        
        # Add completion log with summary
        progress = self.current_progress
        completion_msg = f"Processing complete! Processed {progress.get('total', 0)} VINs"
        
        if progress.get('total', 0) > 0:
            success_rate = (progress.get('current', 0) / progress.get('total', 1)) * 100
            completion_msg += f" ({success_rate:.1f}% success rate)"
        
        self.add_log(completion_msg, "info")
        
        # Show completion notification for large jobs
        if progress.get('total', 0) > 100:
            messagebox.showinfo(
                "Processing Complete",
                f"Successfully processed {progress.get('total', 0)} VINs!\n\n"
                f"Results saved to:\n{self.output_file_var.get()}"
            )
    
    def processing_stopped(self):
        """Called when processing is stopped."""
        # Update UI
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        
        # Add log entry
        self.add_log("Processing stopped.", level="warning")
    
    def copy_selection(self):
        """Copy selected text to clipboard."""
        try:
            # Get selected text
            selected = self.log_text.get(tk.SEL_FIRST, tk.SEL_LAST)
            
            # Copy to clipboard
            self.clipboard_clear()
            self.clipboard_append(selected)
        except tk.TclError:
            # No selection
            pass
    
    def _download_sample_csv(self):
        """Download a sample CSV template for users."""
        # Create sample data with real VIN examples and essential fleet management columns
        sample_data = [
            ["VIN", "Asset ID", "Department", "Odometer"],
            ["1HGBH41JXMN109186", "TRUCK001", "Public Works", "45000"],
            ["1FTFW1ET5DKE55321", "VAN002", "Parks & Recreation", "32000"],
            ["1FAHP2D86CG123456", "CAR003", "Administration", "28000"],
            ["1C4RJFAG4FC654321", "SUV004", "Fire Department", "67000"],
            ["3C6TRVAG9GE123789", "TRUCK005", "Public Works", "52000"],
        ]
        
        # Ask user where to save the sample file
        filename = filedialog.asksaveasfilename(
            title="Save Sample CSV Template",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile="fleet_sample_template.csv"
        )
        
        if filename:
            try:
                # Write sample CSV file
                with open(filename, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerows(sample_data)
                
                # Show success message
                messagebox.showinfo(
                    "Sample CSV Created",
                    f"Sample CSV template saved to:\n{filename}\n\n"
                    f"This template includes:\n"
                    f"• Example VINs (column required)\n"
                    f"• Asset ID (optional)\n"
                    f"• Department (optional)\n"
                    f"• Odometer reading (optional)\n\n"
                    f"You can add, remove, or modify columns as needed.\n"
                    f"Only the VIN column is required for processing.\n\n"
                    f"Note: Annual mileage will be calculated from odometer\n"
                    f"readings and vehicle age from VIN decoding."
                )
                
                # Optionally set as input file
                if messagebox.askyesno(
                    "Use Sample File",
                    "Would you like to use this sample file as your input file?"
                ):
                    self.set_input_file(filename)
                
            except Exception as e:
                messagebox.showerror(
                    "Error Creating Sample",
                    f"Failed to create sample CSV file:\n{e}"
                )

    def refresh(self):
        """Refresh the panel."""
        # Nothing to refresh specifically
        pass

    def _reprocess_file(self):
        """Reprocess the last processed file with one-click."""
        if not self.processing_history:
            messagebox.showinfo("No History", "No previous processing jobs found.")
            return
        
        # Get the most recent processing job
        last_job = self.processing_history[-1]
        
        # Set the parameters
        self.input_file_var.set(last_job.get("input_file", ""))
        self.max_threads_var.set(last_job.get("max_threads", MAX_THREADS))
        self.skip_existing_var.set(last_job.get("skip_existing", True))
        
        # Auto-generate new output filename with fresh timestamp
        self._update_default_output()
        
        # Log the reprocessing
        self.add_log(f"Reprocessing: {os.path.basename(last_job.get('input_file', ''))}")
        
        # Start processing immediately
        self._start_processing()
    
    def _load_processing_history(self):
        """Load processing history from file."""
        try:
            history_file = os.path.join(os.path.dirname(__file__), "..", "data", "processing_history.json")
            if os.path.exists(history_file):
                with open(history_file, 'r') as f:
                    self.processing_history = json.load(f)
                # Keep only last 10 jobs
                self.processing_history = self.processing_history[-10:]
        except Exception as e:
            logger.warning(f"Could not load processing history: {e}")
            self.processing_history = []
    
    def _save_processing_history(self):
        """Save processing history to file."""
        try:
            history_file = os.path.join(os.path.dirname(__file__), "..", "data", "processing_history.json")
            os.makedirs(os.path.dirname(history_file), exist_ok=True)
            
            with open(history_file, 'w') as f:
                json.dump(self.processing_history, f, indent=2)
        except Exception as e:
            logger.warning(f"Could not save processing history: {e}")
    
    def _add_to_processing_history(self, input_file: str, output_file: str, max_threads: int, skip_existing: bool):
        """Add a processing job to history."""
        job = {
            "timestamp": datetime.datetime.now().isoformat(),
            "input_file": input_file,
            "output_file": output_file,
            "max_threads": max_threads,
            "skip_existing": skip_existing,
            "input_filename": os.path.basename(input_file)
        }
        
        self.processing_history.append(job)
        
        # Keep only last 10 jobs
        self.processing_history = self.processing_history[-10:]
        
        # Save to file
        self._save_processing_history()