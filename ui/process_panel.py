"""
process_panel.py

Panel for processing vehicle data in the Fleet Electrification Analyzer.
Redesigned as a clear step-based flow: Upload → Review → Process.
"""

import os
import csv
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Any, Optional
import datetime
import json

# Import drag-and-drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DRAG_DROP_AVAILABLE = True
except ImportError:
    DRAG_DROP_AVAILABLE = False
    logging.warning("tkinterdnd2 not available. Drag-and-drop functionality disabled.")

from settings import MAX_THREADS
from utils import SimpleTooltip
from data.processor import CsvFileValidator, FileValidationResult
from ui.theme import Colors, Fonts, Spacing

# Set up module logger
logger = logging.getLogger(__name__)


class ProcessPanel(ttk.Frame):
    """
    Panel for processing vehicle data.
    Organized as a step-based flow: Upload → Review → Process.
    """

    def __init__(self, parent, on_process=None, on_stop=None, on_log=None):
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
        self._advanced_visible = False

        # File validation state
        self.current_validation: Optional[FileValidationResult] = None

        # Processing history for reprocessing
        self.processing_history: List[Dict[str, Any]] = []
        self._load_processing_history()

        # Enhanced progress tracking
        self.current_progress = {"current": 0, "total": 0, "stage": ""}

        # Create UI — step-based layout
        self._create_step1_upload()
        self._create_step2_review()
        self._create_step3_process()
        self._create_log_section()

        # Trace input changes to trigger validation and output path generation
        self.input_file_var.trace_add("write", lambda *args: self._update_default_output())
        self.input_file_var.trace_add("write", lambda *args: self._validate_input_file())

    # ── Step 1: Upload ──────────────────────────────────────────────────

    def _create_step1_upload(self):
        """Step 1: File upload with drag-and-drop zone."""
        frame = ttk.LabelFrame(self, text="Step 1 — Select CSV File")
        frame.pack(fill=tk.X, padx=Spacing.MARGIN_ELEMENT, pady=(Spacing.MARGIN_ELEMENT, Spacing.SM))

        # Main row: drop zone + buttons
        row = ttk.Frame(frame)
        row.pack(fill=tk.X, padx=Spacing.SM, pady=Spacing.SM)

        # Drag-and-drop zone
        self.drop_zone = tk.Frame(
            row, bg="#f0f8ff", relief=tk.SUNKEN, bd=2, height=56
        )
        self.drop_zone.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, Spacing.SM))

        self.drop_zone_label = tk.Label(
            self.drop_zone,
            text="Drag & Drop CSV file here, or click Browse",
            bg="#f0f8ff", fg="#666666",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY),
            justify=tk.CENTER
        )
        self.drop_zone_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        if DRAG_DROP_AVAILABLE:
            self._setup_drag_drop()

        # Buttons column
        btn_col = ttk.Frame(row)
        btn_col.pack(side=tk.RIGHT)

        browse_btn = ttk.Button(btn_col, text="Browse...", command=self._browse_input)
        browse_btn.pack(fill=tk.X, pady=(0, Spacing.XS))
        SimpleTooltip(browse_btn, "Select CSV file with VINs")

        sample_btn = ttk.Button(btn_col, text="Sample CSV", command=self._download_sample_csv)
        sample_btn.pack(fill=tk.X)
        SimpleTooltip(sample_btn, "Download a sample CSV template")

        # Recent files row (compact — only shown when history exists)
        if self.processing_history:
            recent_row = ttk.Frame(frame)
            recent_row.pack(fill=tk.X, padx=Spacing.SM, pady=(0, Spacing.SM))

            ttk.Label(
                recent_row, text="Recent:",
                font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL),
                foreground=Colors.TEXT_TERTIARY
            ).pack(side=tk.LEFT, padx=(0, Spacing.XS))

            self.history_var = tk.StringVar()
            self.history_combo = ttk.Combobox(
                recent_row,
                textvariable=self.history_var,
                width=45,
                state="readonly"
            )
            self.history_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, Spacing.SM))
            self.history_combo.bind("<<ComboboxSelected>>", self._on_history_selection)
            self._update_history_display()

            reprocess_btn = ttk.Button(
                recent_row, text="Load",
                command=self._reprocess_selected,
                style="Accent.TButton"
            )
            reprocess_btn.pack(side=tk.RIGHT)
            SimpleTooltip(reprocess_btn, "Load selected file for reprocessing")

    # ── Step 2: Review ──────────────────────────────────────────────────

    def _create_step2_review(self):
        """Step 2: Compact validation summary — hidden until file selected."""
        self.review_frame = ttk.LabelFrame(self, text="Step 2 — Review")
        # Initially hidden — shown after file selection & validation

        # Status line: icon + summary text
        self.review_status = ttk.Label(
            self.review_frame,
            text="Select a file above to see validation results.",
            foreground=Colors.TEXT_TERTIARY,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY)
        )
        self.review_status.pack(fill=tk.X, padx=Spacing.SM, pady=(Spacing.SM, Spacing.XS))

        # Detail line: columns detected, mapping info
        self.review_detail = ttk.Label(
            self.review_frame,
            text="",
            foreground=Colors.TEXT_SECONDARY,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL),
            wraplength=700
        )
        self.review_detail.pack(fill=tk.X, padx=Spacing.SM, pady=(0, Spacing.XS))

        # Data preview (compact treeview — 4 rows)
        self.preview_tree_frame = ttk.Frame(self.review_frame)
        # Initially hidden — shown when valid file has sample rows

        self.preview_tree = ttk.Treeview(self.preview_tree_frame, height=4)
        preview_scroll = ttk.Scrollbar(
            self.preview_tree_frame, orient=tk.HORIZONTAL,
            command=self.preview_tree.xview
        )
        self.preview_tree.configure(xscrollcommand=preview_scroll.set)
        self.preview_tree.pack(fill=tk.X, expand=True)
        preview_scroll.pack(fill=tk.X)

    # ── Step 3: Process ─────────────────────────────────────────────────

    def _create_step3_process(self):
        """Step 3: Start processing — prominent button with collapsible options."""
        frame = ttk.Frame(self)
        frame.pack(fill=tk.X, padx=Spacing.MARGIN_ELEMENT, pady=Spacing.SM)
        self._step3_widget = frame  # Reference for review frame insertion

        # Primary action row
        action_row = ttk.Frame(frame)
        action_row.pack(fill=tk.X)

        self.start_button = ttk.Button(
            action_row,
            text="Start Processing",
            command=self._start_processing,
            style="Primary.TButton"
        )
        self.start_button.pack(side=tk.LEFT, padx=(0, Spacing.SM))
        SimpleTooltip(self.start_button,
                      "Decode VINs and enrich with fuel economy data")

        self.stop_button = ttk.Button(
            action_row,
            text="Stop",
            command=self._stop_processing,
            style="Danger.TButton",
            state=tk.DISABLED
        )
        self.stop_button.pack(side=tk.LEFT, padx=(0, Spacing.SM))
        SimpleTooltip(self.stop_button, "Stop processing (progress will be lost)")

        # Advanced options toggle (right side)
        self._adv_toggle_btn = ttk.Button(
            action_row,
            text="Advanced Options",
            command=self._toggle_advanced,
            style="Secondary.TButton"
        )
        self._adv_toggle_btn.pack(side=tk.RIGHT)

        clear_log_btn = ttk.Button(
            action_row,
            text="Clear Log",
            command=self._clear_log,
            style="Secondary.TButton"
        )
        clear_log_btn.pack(side=tk.RIGHT, padx=(0, Spacing.SM))
        SimpleTooltip(clear_log_btn, "Clear the processing log")

        # Advanced options panel (collapsed by default)
        self._adv_frame = ttk.LabelFrame(frame, text="Advanced Options")
        # Not packed initially

        adv_inner = ttk.Frame(self._adv_frame)
        adv_inner.pack(fill=tk.X, padx=Spacing.SM, pady=Spacing.SM)

        ttk.Label(adv_inner, text="Max Threads:").pack(side=tk.LEFT, padx=(0, Spacing.XS))
        threads_spin = ttk.Spinbox(
            adv_inner, from_=1, to=32,
            textvariable=self.max_threads_var, width=4
        )
        threads_spin.pack(side=tk.LEFT, padx=(0, Spacing.MD))
        SimpleTooltip(threads_spin, "Number of parallel API requests (higher = faster, but may hit rate limits)")

        skip_cb = ttk.Checkbutton(
            adv_inner,
            text="Skip VINs already in output file",
            variable=self.skip_existing_var
        )
        skip_cb.pack(side=tk.LEFT, padx=(0, Spacing.MD))
        SimpleTooltip(skip_cb, "When reprocessing, skip VINs that already appear in the output CSV")

        # Output path override
        ttk.Label(adv_inner, text="Output:").pack(side=tk.LEFT, padx=(Spacing.MD, Spacing.XS))
        output_entry = ttk.Entry(adv_inner, textvariable=self.output_file_var, width=30)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, Spacing.XS))

        output_browse = ttk.Button(adv_inner, text="...", command=self._browse_output, width=3)
        output_browse.pack(side=tk.LEFT)
        SimpleTooltip(output_browse, "Choose custom output file location")

        return frame

    # ── Processing Log ──────────────────────────────────────────────────

    def _create_log_section(self):
        """Processing log with progress bar."""
        log_frame = ttk.LabelFrame(self, text="Processing Log")
        log_frame.pack(fill=tk.BOTH, expand=True,
                       padx=Spacing.MARGIN_ELEMENT,
                       pady=(Spacing.SM, Spacing.MARGIN_ELEMENT))

        self.log_text = tk.Text(
            log_frame,
            wrap=tk.WORD,
            width=80, height=18,
            bg=Colors.SURFACE,
            fg=Colors.TEXT_PRIMARY,
            font=(Fonts.FAMILY_MONO, Fonts.SIZE_SMALL),
            state=tk.DISABLED,
            relief=tk.FLAT
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Color tags
        self.log_text.tag_configure("info", foreground=Colors.TEXT_PRIMARY)
        self.log_text.tag_configure("warning", foreground=Colors.WARNING,
                                    font=(Fonts.FAMILY_MONO, Fonts.SIZE_SMALL, "bold"))
        self.log_text.tag_configure("error", foreground=Colors.ERROR,
                                    font=(Fonts.FAMILY_MONO, Fonts.SIZE_SMALL, "bold"))
        self.log_text.tag_configure("success", foreground=Colors.SUCCESS,
                                    font=(Fonts.FAMILY_MONO, Fonts.SIZE_SMALL, "bold"))
        self.log_text.tag_configure("timestamp", foreground=Colors.TEXT_TERTIARY)

        log_scroll = ttk.Scrollbar(log_frame)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=log_scroll.set)
        log_scroll.config(command=self.log_text.yview)

        # Progress bar
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            self, variable=self.progress_var,
            mode='determinate', style="TProgressbar"
        )
        self.progress_bar.pack(fill=tk.X, padx=Spacing.MARGIN_ELEMENT,
                               pady=(0, Spacing.MARGIN_ELEMENT))

        # Right-click menu for log
        self.log_menu = tk.Menu(self, tearoff=0)
        self.log_menu.add_command(label="Copy", command=self.copy_selection)
        self.log_menu.add_command(label="Select All", command=self._select_all_log)
        self.log_menu.add_separator()
        self.log_menu.add_command(label="Clear Log", command=self._clear_log)
        self.log_text.bind("<Button-3>", self._show_log_menu)

    # ── Advanced Options Toggle ─────────────────────────────────────────

    def _toggle_advanced(self):
        """Show/hide the advanced options panel."""
        if self._advanced_visible:
            self._adv_frame.pack_forget()
            self._adv_toggle_btn.config(text="Advanced Options")
            self._advanced_visible = False
        else:
            self._adv_frame.pack(fill=tk.X, pady=(Spacing.SM, 0))
            self._adv_toggle_btn.config(text="Hide Options")
            self._advanced_visible = True

    # ── Drag-and-Drop ───────────────────────────────────────────────────

    def _setup_drag_drop(self):
        if not DRAG_DROP_AVAILABLE:
            return
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind('<<Drop>>', self._handle_file_drop)
        self.drop_zone.dnd_bind('<<DragEnter>>', self._drag_enter)
        self.drop_zone.dnd_bind('<<DragLeave>>', self._drag_leave)
        self.drop_zone.bind("<Button-1>", lambda e: self._browse_input())
        self.drop_zone_label.bind("<Button-1>", lambda e: self._browse_input())

    def _handle_file_drop(self, event):
        files = event.data.split()
        if files:
            filepath = files[0].replace("{", "").replace("}", "")
            if filepath.lower().endswith('.csv'):
                self.input_file_var.set(filepath)
                self._update_drop_zone_display(filepath)
                self.add_log(f"File loaded: {os.path.basename(filepath)}")
            else:
                messagebox.showerror("Invalid File", "Please drop a CSV file.")
        self._drag_leave(event)

    def _drag_enter(self, event):
        self.drop_zone.config(bg="#e6f3ff", relief=tk.RAISED)
        self.drop_zone_label.config(
            text="Drop CSV file here",
            bg="#e6f3ff", fg="#0066cc",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY, "bold")
        )

    def _drag_leave(self, event):
        self.drop_zone.config(bg="#f0f8ff", relief=tk.SUNKEN)
        if self.input_file_var.get().strip():
            self._update_drop_zone_display(self.input_file_var.get())
        else:
            self.drop_zone_label.config(
                text="Drag & Drop CSV file here, or click Browse",
                bg="#f0f8ff", fg="#666666",
                font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY)
            )

    def _update_drop_zone_display(self, filepath):
        filename = os.path.basename(filepath)
        self.drop_zone_label.config(
            text=f"{filename}",
            bg="#f0f8ff", fg="#006600",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY, "bold")
        )

    # ── File Browsing ───────────────────────────────────────────────────

    def _browse_input(self):
        filepath = filedialog.askopenfilename(
            title="Select Input CSV",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if filepath:
            self.input_file_var.set(filepath)

    def _browse_output(self):
        filepath = filedialog.asksaveasfilename(
            title="Select Output CSV",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if filepath:
            self.output_file_var.set(filepath)

    # ── Auto Output Path ────────────────────────────────────────────────

    def _update_default_output(self):
        input_path = self.input_file_var.get().strip()
        if input_path and os.path.exists(input_path):
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_dir = os.path.dirname(input_path)
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(output_dir, f"{base_name}_fleet_analysis_{timestamp}.csv")
            self.output_file_var.set(output_path)
            if hasattr(self, 'drop_zone_label'):
                self._update_drop_zone_display(input_path)

    # ── Validation & Review ─────────────────────────────────────────────

    def _validate_input_file(self):
        input_path = self.input_file_var.get().strip()

        if not input_path:
            self._hide_review()
            return

        if not os.path.exists(input_path):
            self._update_review("File not found", "", "error")
            return

        try:
            validator = CsvFileValidator(input_path)
            self.current_validation = validator.validate_and_preview()

            if not hasattr(self.current_validation, 'valid'):
                raise ValueError("Invalid validation result")

            self._update_review_from_validation()

        except Exception as e:
            logger.error(f"Error validating file: {e}")
            self._update_review(f"Error: {e}", "", "error")
            self.current_validation = None

    def _update_review(self, status_text: str, detail_text: str, level: str = "info"):
        """Update review section with status and detail text."""
        colors = {
            "info": Colors.TEXT_PRIMARY,
            "success": Colors.SUCCESS,
            "warning": Colors.WARNING,
            "error": Colors.ERROR
        }
        self.review_status.config(text=status_text, foreground=colors.get(level, Colors.TEXT_PRIMARY))
        self.review_detail.config(text=detail_text)
        self._show_review()

    def _update_review_from_validation(self):
        """Populate review section from current_validation."""
        v = self.current_validation
        if not v:
            self._hide_review()
            return

        # Status line
        if v.valid:
            status = f"{v.valid_vins} valid VINs found"
            if v.invalid_vins > 0:
                status += f"  ({v.invalid_vins} invalid)"
            level = "success"
        else:
            status = v.error_message or "Validation failed"
            level = "error"

        # Detail line: file info + column mapping summary
        parts = []
        parts.append(f"{v.total_rows} rows")
        parts.append(f"{v.column_count} columns")
        parts.append(f"Encoding: {v.detected_encoding}")

        if v.vin_column:
            parts.append(f"VIN column: '{v.vin_column}'")

        if v.mapped_columns:
            mapped_names = list(v.mapped_columns.values())[:4]
            mapped_str = ", ".join(mapped_names)
            if len(v.mapped_columns) > 4:
                mapped_str += f" (+{len(v.mapped_columns) - 4} more)"
            parts.append(f"Fleet fields: {mapped_str}")

        detail = "  |  ".join(parts)

        self._update_review(status, detail, level)

        # Update data preview
        self._update_data_preview()

    def _show_review(self):
        if not self.review_frame.winfo_viewable():
            self.review_frame.pack(
                fill=tk.X, padx=Spacing.MARGIN_ELEMENT, pady=(0, Spacing.SM),
                before=self._step3_widget
            )

    def _hide_review(self):
        self.review_frame.pack_forget()
        self.preview_tree_frame.pack_forget()
        self.current_validation = None

    def _update_data_preview(self):
        """Show a compact data preview table."""
        v = self.current_validation
        if not v or not v.sample_rows:
            self.preview_tree_frame.pack_forget()
            return

        # Clear
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)

        sample = v.sample_rows[:4]
        columns = v.columns
        self.preview_tree["columns"] = columns
        self.preview_tree["show"] = "headings"

        for col in columns:
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, width=max(len(col) * 8, 80), minwidth=60)

        for row in sample:
            values = [str(row.get(col, ""))[:40] for col in columns]
            self.preview_tree.insert("", tk.END, values=values)

        self.preview_tree_frame.pack(fill=tk.X, padx=Spacing.SM, pady=(0, Spacing.SM))

    # ── Processing ──────────────────────────────────────────────────────

    def _start_processing(self):
        import time
        self._start_time = time.time()

        input_file = self.input_file_var.get().strip()
        output_file = self.output_file_var.get().strip()

        if not input_file:
            messagebox.showerror("Error", "Please select an input CSV file.")
            return

        if not output_file:
            messagebox.showerror("Error", "Please specify an output CSV file path.")
            return

        # Ensure validation
        if not self.current_validation or not self.current_validation.valid:
            if not self.current_validation:
                self._validate_input_file()

            if not self.current_validation or not self.current_validation.valid:
                error_msg = "The selected CSV file has validation errors:\n\n"
                if self.current_validation and self.current_validation.error_message:
                    error_msg += self.current_validation.error_message
                else:
                    error_msg += "Unknown validation error"
                error_msg += "\n\nPlease fix the issues or select a different file."
                messagebox.showerror("Validation Error", error_msg)
                return

        # Warn about invalid VINs
        if self.current_validation.invalid_vins > 0:
            pct = (self.current_validation.invalid_vins / self.current_validation.total_rows) * 100
            msg = (f"Your file contains {self.current_validation.invalid_vins} invalid VINs "
                   f"({pct:.1f}% of total).\n\n"
                   "These will appear in results with error messages.\n\n"
                   "Continue processing?")
            if not messagebox.askyesno("Invalid VINs Detected", msg):
                return

        # Record to history
        self._add_to_processing_history(
            input_file, output_file,
            self.max_threads_var.get(),
            self.skip_existing_var.get()
        )

        # Call processing callback
        if self.on_process_callback:
            options = {
                "max_threads": self.max_threads_var.get(),
                "skip_existing": self.skip_existing_var.get(),
                "cached_validation": self.current_validation,
            }
            try:
                self.on_process_callback(input_file, output_file, options)
            except Exception as e:
                self.add_log(f"Failed to start processing: {e}", "error")
                messagebox.showerror("Processing Error", f"Failed to start processing: {e}")
                return

        # Update UI
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        self.add_log(f"Processing {self.current_validation.valid_vins} VINs from {os.path.basename(input_file)}")
        self.add_log(f"Output: {os.path.basename(output_file)}")
        if self.current_validation.invalid_vins > 0:
            self.add_log(
                f"{self.current_validation.invalid_vins} invalid VINs will be included with errors",
                "warning"
            )

    def _stop_processing(self):
        self.add_log("Stopping processing...", level="warning")
        if self.on_stop_callback:
            self.on_stop_callback()

    # ── Reprocessing History ────────────────────────────────────────────

    def _update_history_display(self):
        if not hasattr(self, 'history_combo'):
            return
        items = []
        for job in reversed(self.processing_history[-5:]):
            ts = job.get('timestamp', '')
            fname = job.get('input_filename', 'Unknown')
            try:
                dt = datetime.datetime.fromisoformat(ts)
                time_str = dt.strftime("%m/%d %H:%M")
            except Exception:
                time_str = "Unknown"
            items.append(f"{time_str} — {fname}")
        self.history_combo['values'] = items
        if items:
            self.history_combo.set(items[0])

    def _on_history_selection(self, event):
        selection = self.history_combo.get()
        if not selection:
            return
        idx = list(self.history_combo['values']).index(selection)
        rev = list(reversed(self.processing_history[-5:]))
        if idx < len(rev):
            job = rev[idx]
            input_file = job.get('input_file', '')
            if os.path.exists(input_file):
                self.input_file_var.set(input_file)
                self.max_threads_var.set(job.get('max_threads', MAX_THREADS))
                self.skip_existing_var.set(job.get('skip_existing', True))
                self.add_log(f"Loaded: {os.path.basename(input_file)}")
            else:
                messagebox.showwarning("File Not Found",
                                       f"The file no longer exists:\n{input_file}")

    def _reprocess_selected(self):
        """Load the selected history file (does not auto-start)."""
        if not hasattr(self, 'history_combo'):
            return
        selection = self.history_combo.get()
        if not selection:
            messagebox.showinfo("No Selection", "Please select a file from the recent list.")
            return
        # Trigger the same logic as selecting from dropdown
        self._on_history_selection(None)

    def _reprocess_file(self):
        """Reprocess the last processed file."""
        if not self.processing_history:
            messagebox.showinfo("No History", "No previous processing jobs found.")
            return
        last_job = self.processing_history[-1]
        input_file = last_job.get("input_file", "")
        if not os.path.exists(input_file):
            messagebox.showerror("File Not Found",
                                 f"The input file no longer exists:\n{input_file}")
            return
        self.input_file_var.set(input_file)
        self.max_threads_var.set(last_job.get("max_threads", MAX_THREADS))
        self.skip_existing_var.set(last_job.get("skip_existing", True))
        self._update_default_output()
        self.add_log(f"Reprocessing: {os.path.basename(input_file)}")
        self._start_processing()

    # ── Log Helpers ─────────────────────────────────────────────────────

    def add_log(self, message, level="info"):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] ", "timestamp")
        self.log_text.insert(tk.END, f"{message}\n", level)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        if self.on_log_callback:
            self.on_log_callback(message)

    def _clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _show_log_menu(self, event):
        self.log_menu.tk_popup(event.x_root, event.y_root)

    def _select_all_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.tag_add(tk.SEL, "1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.log_text.focus_set()

    def copy_selection(self):
        try:
            selected = self.log_text.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.clipboard_clear()
            self.clipboard_append(selected)
        except tk.TclError:
            pass

    # ── Progress ────────────────────────────────────────────────────────

    def update_progress(self, current, total, stage="Processing"):
        if total <= 0:
            return
        self.current_progress = {"current": current, "total": total, "stage": stage}
        percentage = (current / total) * 100
        self.progress_bar.config(maximum=total)
        self.progress_var.set(current)

        status_msg = f"{stage}: {current}/{total} ({percentage:.1f}%)"

        if total > 100 and current > 0:
            ratio = current / total
            if ratio > 0.1:
                import time
                if not hasattr(self, '_start_time'):
                    self._start_time = time.time()
                elapsed = time.time() - self._start_time
                remaining = (elapsed / ratio) - elapsed
                if remaining > 60:
                    remaining_str = f"{int(remaining // 60)}m {int(remaining % 60)}s"
                else:
                    remaining_str = f"{int(remaining)}s"
                status_msg += f" | ETA: {remaining_str}"

        if self.on_log_callback:
            self.on_log_callback(status_msg)

    def processing_complete(self):
        if hasattr(self, '_start_time'):
            delattr(self, '_start_time')
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

        progress = self.current_progress
        total = progress.get('total', 0)
        msg = f"Processing complete! {total} VINs processed"
        self.add_log(msg, "success")

        if total > 100:
            messagebox.showinfo(
                "Processing Complete",
                f"Processed {total} VINs.\n\nResults saved to:\n{self.output_file_var.get()}"
            )

    def processing_stopped(self):
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.add_log("Processing stopped.", level="warning")

    # ── Public Setters ──────────────────────────────────────────────────

    def set_input_file(self, filepath):
        self.input_file_var.set(filepath)

    def set_max_threads(self, max_threads):
        self.max_threads_var.set(max_threads)

    def refresh(self):
        pass

    # ── Sample CSV ──────────────────────────────────────────────────────

    def _download_sample_csv(self):
        sample_data = [
            ["VIN", "Asset ID", "Department", "Odometer"],
            ["1HGBH41JXMN109186", "TRUCK001", "Public Works", "45000"],
            ["1FTFW1ET5DKE55321", "VAN002", "Parks & Recreation", "32000"],
            ["1FAHP2D86CG123456", "CAR003", "Administration", "28000"],
            ["1C4RJFAG4FC654321", "SUV004", "Fire Department", "67000"],
            ["3C6TRVAG9GE123789", "TRUCK005", "Public Works", "52000"],
        ]

        filename = filedialog.asksaveasfilename(
            title="Save Sample CSV Template",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile="fleet_sample_template.csv"
        )

        if filename:
            try:
                with open(filename, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerows(sample_data)

                messagebox.showinfo(
                    "Sample CSV Created",
                    f"Saved to: {filename}\n\n"
                    "Only the VIN column is required.\n"
                    "Asset ID, Department, and Odometer are optional fleet fields."
                )

                if messagebox.askyesno("Use Sample File",
                                       "Use this as your input file?"):
                    self.set_input_file(filename)

            except Exception as e:
                messagebox.showerror("Error", f"Failed to create sample CSV:\n{e}")

    # ── Processing History Persistence ──────────────────────────────────

    def _load_processing_history(self):
        try:
            history_file = os.path.join(os.path.dirname(__file__), "..", "data", "processing_history.json")
            if os.path.exists(history_file):
                with open(history_file, 'r') as f:
                    self.processing_history = json.load(f)
                self.processing_history = self.processing_history[-10:]
        except Exception as e:
            logger.warning(f"Could not load processing history: {e}")
            self.processing_history = []

    def _save_processing_history(self):
        try:
            history_file = os.path.join(os.path.dirname(__file__), "..", "data", "processing_history.json")
            os.makedirs(os.path.dirname(history_file), exist_ok=True)
            with open(history_file, 'w') as f:
                json.dump(self.processing_history, f, indent=2)
        except Exception as e:
            logger.warning(f"Could not save processing history: {e}")

    def _add_to_processing_history(self, input_file, output_file, max_threads, skip_existing):
        job = {
            "timestamp": datetime.datetime.now().isoformat(),
            "input_file": input_file,
            "output_file": output_file,
            "max_threads": max_threads,
            "skip_existing": skip_existing,
            "input_filename": os.path.basename(input_file)
        }
        self.processing_history.append(job)
        self.processing_history = self.processing_history[-10:]
        self._save_processing_history()
        if hasattr(self, 'history_combo'):
            self._update_history_display()
