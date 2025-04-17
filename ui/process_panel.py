"""
process_panel.py

Panel for processing vehicle data in the Fleet Electrification Analyzer.
Handles input/output file selection, processing options, and log display.
"""

import os
import logging
import tkinter as tk
from tkinter import ttk, filedialog
from typing import Dict, List, Any, Optional, Callable

from settings import (
    MAX_THREADS,
    PRIMARY_HEX_1,
    PRIMARY_HEX_2,
    PRIMARY_HEX_3,
    SECONDARY_HEX_1
)
from utils import SimpleTooltip, ScrollableFrame

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
        
        # Create UI components
        self._create_input_section()
        self._create_options_section()
        self._create_actions_section()
        self._create_log_section()
        
        # Set default output file
        self._update_default_output()
    
    def _create_input_section(self):
        """Create the input/output file selection section."""
        # Create frame
        input_frame = ttk.LabelFrame(self, text="Input/Output")
        input_frame.pack(fill=tk.X, padx=10, pady=(10, 5), ipady=5)
        
        # Input file row
        ttk.Label(input_frame, text="Input CSV:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        input_browse_btn = ttk.Button(
            input_frame, 
            text="Browse...", 
            command=self._browse_input
        )
        input_browse_btn.grid(row=0, column=1, padx=5, pady=5)
        SimpleTooltip(input_browse_btn, "Select CSV file with VINs")
        
        input_entry = ttk.Entry(
            input_frame, 
            textvariable=self.input_file_var, 
            width=40
        )
        input_entry.grid(row=0, column=2, sticky=tk.W+tk.E, padx=5, pady=5)
        SimpleTooltip(input_entry, "Path to input CSV file with VINs")
        
        # Output file row
        ttk.Label(input_frame, text="Output CSV:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        
        output_browse_btn = ttk.Button(
            input_frame, 
            text="Browse...", 
            command=self._browse_output
        )
        output_browse_btn.grid(row=1, column=1, padx=5, pady=5)
        SimpleTooltip(output_browse_btn, "Select output CSV location")
        
        output_entry = ttk.Entry(
            input_frame, 
            textvariable=self.output_file_var, 
            width=40
        )
        output_entry.grid(row=1, column=2, sticky=tk.W+tk.E, padx=5, pady=5)
        SimpleTooltip(output_entry, "Path to save processed results")
        
        # Configure grid to expand properly
        input_frame.columnconfigure(2, weight=1)
        
        # Bind event to update default output when input changes
        self.input_file_var.trace_add("write", lambda *args: self._update_default_output())
    
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
        """Update default output path based on input path."""
        input_path = self.input_file_var.get().strip()
        if input_path and not self.output_file_var.get().strip():
            # Generate default output path
            input_dir = os.path.dirname(input_path)
            input_name = os.path.basename(input_path)
            base_name, _ = os.path.splitext(input_name)
            
            output_name = f"{base_name}_results.csv"
            output_path = os.path.join(input_dir, output_name)
            
            self.output_file_var.set(output_path)
    
    def _start_processing(self):
        """Start the processing pipeline."""
        # Get input and output paths
        input_path = self.input_file_var.get().strip()
        output_path = self.output_file_var.get().strip()
        
        # Validate paths
        if not input_path:
            self.add_log("Error: Input file path is required.", level="error")
            return
        
        if not os.path.exists(input_path):
            self.add_log(f"Error: Input file does not exist: {input_path}", level="error")
            return
        
        if not output_path:
            self.add_log("Error: Output file path is required.", level="error")
            return
        
        # Get options
        options = {
            "max_threads": self.max_threads_var.get(),
            "skip_existing": self.skip_existing_var.get()
        }
        
        # Update UI
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.progress_var.set(0)
        
        # Add log entry
        self.add_log(f"Starting processing with {options['max_threads']} threads...")
        
        # Call the callback
        if self.on_process_callback:
            self.on_process_callback(input_path, output_path, options)
    
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
        import datetime
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
    
    def update_progress(self, current, total):
        """
        Update the progress bar.
        
        Args:
            current: Current progress value
            total: Maximum progress value
        """
        if total <= 0:
            return
        
        # Calculate percentage
        percentage = (current / total) * 100
        
        # Update progress bar
        self.progress_bar.config(maximum=total)
        self.progress_var.set(current)
        
        # Update progress text in status if callback is available
        if self.on_log_callback:
            self.on_log_callback(f"Processing: {current}/{total} ({percentage:.1f}%)")
    
    def processing_complete(self):
        """Called when processing is complete."""
        # Update UI
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        
        # Add log entry
        self.add_log("Processing complete.")
    
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
    
    def refresh(self):
        """Refresh the panel."""
        # Nothing to refresh specifically
        pass