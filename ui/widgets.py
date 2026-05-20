"""
ui/widgets.py

Reusable Tkinter UI widgets and dialog helpers for the Fleet Electrification Analyzer.
Extracted from utils.py to keep UI code in the ui/ package.
"""

import os
import datetime
import tkinter as tk
from tkinter import ttk
from typing import List, Tuple, Optional


###############################################################################
# Tooltips
###############################################################################

class SimpleTooltip:
    """
    Creates a tooltip for a Tkinter widget.

    Example:
        button = tk.Button(root, text="Help")
        SimpleTooltip(button, "Click for help")
    """

    def __init__(self, widget, text, delay=500, fg="#000000", bg="#FFFFEA",
                 padx=5, pady=3, font=None):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.fg = fg
        self.bg = bg
        self.padx = padx
        self.pady = pady
        self.font = font or ("tahoma", "8", "normal")

        self.tipwindow = None
        self.after_id = None

        self.widget.bind("<Enter>", self.schedule_show)
        self.widget.bind("<Leave>", self.hide)
        self.widget.bind("<ButtonPress>", self.hide)

    def schedule_show(self, event=None):
        """Schedule the tooltip to appear after the delay."""
        self.cancel_scheduled()
        self.after_id = self.widget.after(self.delay, self.show)

    def cancel_scheduled(self):
        """Cancel any scheduled tooltip showing."""
        if self.after_id:
            self.widget.after_cancel(self.after_id)
            self.after_id = None

    def show(self):
        """Create and display the tooltip."""
        if self.tipwindow or not self.text:
            return

        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 10

        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")

        label = tk.Label(
            tw, text=self.text, justify=tk.LEFT,
            background=self.bg, foreground=self.fg,
            relief="solid", borderwidth=1,
            font=self.font
        )
        label.pack(ipadx=self.padx, ipady=self.pady)

    def hide(self, event=None):
        """Destroy the tooltip window."""
        self.cancel_scheduled()
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None


###############################################################################
# Status Bar
###############################################################################

class StatusBar(ttk.Frame):
    """
    Status bar widget with multiple sections.

    Example:
        status_bar = StatusBar(root)
        status_bar.pack(side="bottom", fill="x")
        status_bar.set("Ready")
        status_bar.set("Processing...", section="process")
    """

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.sections = {}

        self.sections["main"] = ttk.Label(
            self, relief="sunken", anchor="w", padding=(5, 2)
        )
        self.sections["main"].pack(side="left", fill="x", expand=True)

        self.set("Ready")

    def add_section(self, name, width=None, side="right"):
        """Add a new section to the status bar."""
        if name in self.sections:
            return

        self.sections[name] = ttk.Label(
            self, relief="sunken", anchor="w", padding=(5, 2), width=width
        )
        self.sections[name].pack(side=side, fill="y", padx=(1, 0))

        self.set("", section=name)

    def set(self, text, section="main"):
        """Set the text for a section."""
        if section in self.sections:
            self.sections[section].config(text=text)


###############################################################################
# Progress Dialog
###############################################################################

class ProgressDialog(tk.Toplevel):
    """
    Modal dialog with a progress bar and cancel button.

    Example:
        progress = ProgressDialog(root, "Processing Files", "Please wait...")
        for i in range(100):
            if progress.cancelled:
                break
            progress.update(i)
        progress.destroy()
    """

    def __init__(self, parent, title, message, maximum=100, cancelable=True):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.cancelled = False
        self.maximum = maximum

        width = 350
        height = 150
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        x = parent_x + (parent_width // 2) - (width // 2)
        y = parent_y + (parent_height // 2) - (height // 2)

        self.geometry(f"{width}x{height}+{x}+{y}")

        self.message_var = tk.StringVar(value=message)
        self.message_label = ttk.Label(
            self, textvariable=self.message_var, wraplength=width-20
        )
        self.message_label.pack(padx=10, pady=(10, 5))

        self.progress_var = tk.DoubleVar(value=0)
        self.progressbar = ttk.Progressbar(
            self, orient=tk.HORIZONTAL, length=300,
            mode='determinate', variable=self.progress_var,
            maximum=maximum
        )
        self.progressbar.pack(padx=10, pady=5)

        self.status_var = tk.StringVar(value="Starting...")
        self.status_label = ttk.Label(
            self, textvariable=self.status_var
        )
        self.status_label.pack(padx=10, pady=5)

        if cancelable:
            self.cancel_button = ttk.Button(
                self, text="Cancel", command=self.cancel
            )
            self.cancel_button.pack(pady=10)

        self.update_idletasks()

    def update(self, value, message=None, status=None):
        """Update the progress and optionally the message."""
        self.progress_var.set(value)
        if message is not None:
            self.message_var.set(message)

        if status is not None:
            percent = int((value / self.maximum) * 100)
            self.status_var.set(f"{status} ({percent}%)")
        else:
            percent = int((value / self.maximum) * 100)
            self.status_var.set(f"{percent}%")

        self.update_idletasks()

    def cancel(self):
        """Mark as cancelled and update UI."""
        self.cancelled = True
        self.status_var.set("Cancelling...")
        if hasattr(self, 'cancel_button'):
            self.cancel_button.config(state="disabled")


###############################################################################
# Scrollable Frame
###############################################################################

class ScrollableFrame(ttk.Frame):
    """
    A frame with scrollbars that can contain other widgets.

    Example:
        scrollable = ScrollableFrame(root)
        scrollable.pack(fill="both", expand=True)

        for i in range(50):
            ttk.Label(scrollable.scrollable_frame, text=f"Row {i}").pack()
    """

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.canvas = tk.Canvas(self)

        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=self.hsb.set)

        self.vsb.pack(side="right", fill="y")
        self.hsb.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw"
        )

        self.canvas.bind("<Configure>", self._configure_canvas_window)

        self.scrollable_frame.bind("<Enter>", self._bind_mousewheel)
        self.scrollable_frame.bind("<Leave>", self._unbind_mousewheel)

    def _configure_canvas_window(self, event):
        """Adjust the width of the canvas window when canvas size changes."""
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _bind_mousewheel(self, event):
        """Bind mouse wheel to scroll vertically."""
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _unbind_mousewheel(self, event):
        """Unbind mouse wheel when mouse leaves widget."""
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        """Handle mouse wheel events."""
        if event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")


###############################################################################
# Error Communication
###############################################################################

class ErrorCommunicator:
    """Enhanced error communication with user-friendly messages and suggested fixes."""

    ERROR_CATEGORIES = {
        "vin_format": {
            "title": "VIN Format Error",
            "icon": "\u274c",
            "color": "#d32f2f"
        },
        "file_access": {
            "title": "File Access Error",
            "icon": "\U0001f4c1",
            "color": "#f57c00"
        },
        "api_error": {
            "title": "Data Lookup Error",
            "icon": "\U0001f310",
            "color": "#1976d2"
        },
        "processing": {
            "title": "Processing Error",
            "icon": "\u2699\ufe0f",
            "color": "#7b1fa2"
        },
        "validation": {
            "title": "Data Validation Error",
            "icon": "\u26a0\ufe0f",
            "color": "#f57c00"
        }
    }

    @classmethod
    def show_error_dialog(cls, parent, category: str, message: str, details: str = "",
                         suggested_fixes: List[str] = None, context_help: str = ""):
        """Show an enhanced error dialog with category-specific styling and helpful suggestions."""
        category_info = cls.ERROR_CATEGORIES.get(category, cls.ERROR_CATEGORIES["processing"])

        dialog = tk.Toplevel(parent)
        dialog.title(category_info["title"])
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()

        width = 500
        height = 400
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        x = parent_x + (parent_width // 2) - (width // 2)
        y = parent_y + (parent_height // 2) - (height // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")

        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))

        icon_label = tk.Label(
            header_frame,
            text=category_info["icon"],
            font=("Segoe UI", 24),
            fg=category_info["color"]
        )
        icon_label.pack(side=tk.LEFT, padx=(0, 10))

        title_label = tk.Label(
            header_frame,
            text=category_info["title"],
            font=("Segoe UI", 14, "bold"),
            fg=category_info["color"]
        )
        title_label.pack(side=tk.LEFT, anchor=tk.W)

        message_label = tk.Label(
            main_frame,
            text=message,
            font=("Segoe UI", 11),
            wraplength=450,
            justify=tk.LEFT
        )
        message_label.pack(fill=tk.X, pady=(0, 10))

        if details:
            details_frame = ttk.LabelFrame(main_frame, text="Details")
            details_frame.pack(fill=tk.X, pady=(0, 10))

            details_text = tk.Text(
                details_frame,
                height=4,
                wrap=tk.WORD,
                bg="#f5f5f5",
                font=("Consolas", 9)
            )
            details_text.pack(fill=tk.X, padx=10, pady=10)
            details_text.insert(1.0, details)
            details_text.config(state=tk.DISABLED)

        if suggested_fixes:
            fixes_frame = ttk.LabelFrame(main_frame, text="\U0001f4a1 Suggested Solutions")
            fixes_frame.pack(fill=tk.X, pady=(0, 10))

            for i, fix in enumerate(suggested_fixes, 1):
                fix_label = tk.Label(
                    fixes_frame,
                    text=f"{i}. {fix}",
                    font=("Segoe UI", 10),
                    wraplength=450,
                    justify=tk.LEFT,
                    anchor=tk.W
                )
                fix_label.pack(fill=tk.X, padx=10, pady=2)

        if context_help:
            help_frame = ttk.LabelFrame(main_frame, text="\u2139\ufe0f Additional Help")
            help_frame.pack(fill=tk.X, pady=(0, 15))

            help_label = tk.Label(
                help_frame,
                text=context_help,
                font=("Segoe UI", 9),
                wraplength=450,
                justify=tk.LEFT,
                fg="#666666"
            )
            help_label.pack(fill=tk.X, padx=10, pady=10)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(
            button_frame,
            text="OK",
            command=dialog.destroy
        ).pack(side=tk.RIGHT, padx=(10, 0))

        if context_help:
            ttk.Button(
                button_frame,
                text="Copy Details",
                command=lambda: cls._copy_error_details(category, message, details)
            ).pack(side=tk.RIGHT)

    @classmethod
    def _copy_error_details(cls, category: str, message: str, details: str):
        """Copy error details to clipboard for support."""
        error_report = f"""Fleet Electrification Analyzer - Error Report
Generated: {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

Category: {category}
Message: {message}

Details:
{details}

Please include this information when reporting issues.
"""
        root = tk._default_root
        if root:
            root.clipboard_clear()
            root.clipboard_append(error_report)

    @classmethod
    def get_vin_error_message(cls, vin: str, error_type: str) -> Tuple[str, List[str]]:
        """Get user-friendly VIN error message with suggested fixes."""
        if error_type == "length":
            message = f"The VIN '{vin}' has {len(vin)} characters, but VINs must be exactly 17 characters long."
            fixes = [
                "Check for missing characters at the beginning or end",
                "Remove any spaces or special characters",
                "Verify the VIN from the vehicle documentation"
            ]

        elif error_type == "invalid_chars":
            message = f"The VIN '{vin}' contains invalid characters. VINs can only contain letters and numbers (except I, O, and Q)."
            fixes = [
                "Replace any I, O, Q characters with 1, 0, or other valid characters",
                "Remove spaces, hyphens, or other special characters",
                "Check if characters were misread (e.g., 8 vs B, 5 vs S)"
            ]

        elif error_type == "placeholder":
            message = f"The VIN '{vin}' appears to be a placeholder or template rather than a real VIN."
            fixes = [
                "Replace with the actual VIN from the vehicle",
                "Check vehicle registration or insurance documents",
                "Look for the VIN on the dashboard or driver's side door frame"
            ]

        else:
            message = f"The VIN '{vin}' is not valid."
            fixes = [
                "Verify the VIN is exactly 17 characters",
                "Check for invalid characters (I, O, Q are not allowed)",
                "Ensure it's a real VIN, not a placeholder"
            ]

        return message, fixes

    @classmethod
    def get_file_error_message(cls, filepath: str, error_type: str) -> Tuple[str, List[str]]:
        """Get user-friendly file error message with suggested fixes."""
        filename = os.path.basename(filepath)

        if error_type == "not_found":
            message = f"The file '{filename}' could not be found."
            fixes = [
                "Check if the file was moved or deleted",
                "Verify the file path is correct",
                "Use the Browse button to select the file again"
            ]

        elif error_type == "permission":
            message = f"Permission denied when trying to access '{filename}'."
            fixes = [
                "Check if the file is open in another program (like Excel)",
                "Run the application as administrator",
                "Verify you have read/write permissions for this location"
            ]

        elif error_type == "format":
            message = f"The file '{filename}' is not in the expected CSV format."
            fixes = [
                "Save the file as CSV format (.csv extension)",
                "Check if the file uses the correct delimiter (comma)",
                "Use the 'Download Sample CSV' button to see the expected format"
            ]

        elif error_type == "encoding":
            message = f"The file '{filename}' has encoding issues and cannot be read properly."
            fixes = [
                "Save the file with UTF-8 encoding",
                "Open in Excel and 'Save As' CSV (UTF-8)",
                "Remove any special characters that might cause encoding issues"
            ]

        else:
            message = f"An error occurred while processing the file '{filename}'."
            fixes = [
                "Check if the file is corrupted",
                "Try opening the file in Excel to verify it's readable",
                "Use a different file or recreate the CSV"
            ]

        return message, fixes


###############################################################################
# Context Help
###############################################################################

class ContextHelp:
    """Context-sensitive help system for user guidance."""

    HELP_TOPICS = {
        "vin_format": """
VIN (Vehicle Identification Number) Format:
\u2022 Must be exactly 17 characters long
\u2022 Contains letters and numbers only
\u2022 Cannot contain I, O, or Q (to avoid confusion with 1, 0)
\u2022 Each position has a specific meaning for vehicle details

Common VIN locations on vehicles:
\u2022 Dashboard (visible through windshield)
\u2022 Driver's side door frame
\u2022 Vehicle registration documents
\u2022 Insurance cards
        """,

        "csv_format": """
CSV File Format Requirements:
\u2022 Must have a 'VIN' or 'VINs' column header
\u2022 One VIN per row
\u2022 Additional columns are preserved (Asset ID, Department, etc.)
\u2022 Use comma as delimiter
\u2022 UTF-8 encoding recommended

Example format:
VIN,Asset ID,Department
1HGBH41JXMN109186,TRUCK001,Public Works
1FTFW1ET5DKE55321,VAN002,Parks & Recreation
        """,

        "processing_options": """
Processing Options:
\u2022 Max Threads: Number of parallel processing threads (1-32)
  - Higher values = faster processing
  - Lower values = less system resource usage

\u2022 Skip Existing VINs: Skip VINs already in output file
  - Useful for resuming interrupted processing
  - Prevents duplicate entries

\u2022 Auto-generate filename: Creates timestamped output files
  - Format: filename_fleet_analysis_YYYYMMDD_HHMMSS.csv
  - Prevents accidental overwrites
        """,

        "data_quality": """
Data Quality Indicators:
\u2022 High (80-100%): Complete vehicle information available
\u2022 Medium (50-80%): Some details missing but usable
\u2022 Low (0-50%): Limited information found
\u2022 Failed: VIN invalid or no data found

Quality factors:
\u2022 VIN validity and format
\u2022 Successful API data retrieval
\u2022 Commercial vehicle data completeness
\u2022 Cross-source data consistency
        """
    }

    @classmethod
    def show_help_dialog(cls, parent, topic: str, title: str = "Help"):
        """Show context-sensitive help dialog."""
        help_text = cls.HELP_TOPICS.get(topic, "Help information not available for this topic.")

        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.resizable(True, True)
        dialog.transient(parent)

        width = 600
        height = 400
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        x = parent_x + (parent_width // 2) - (width // 2)
        y = parent_y + (parent_height // 2) - (height // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")

        content_frame = ttk.Frame(dialog)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        text_widget = tk.Text(
            content_frame,
            wrap=tk.WORD,
            font=("Segoe UI", 10),
            bg="#f8f9fa",
            relief=tk.FLAT,
            padx=10,
            pady=10
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert(1.0, help_text.strip())
        text_widget.config(state=tk.DISABLED)

        scrollbar = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.configure(yscrollcommand=scrollbar.set)

        button_frame = ttk.Frame(content_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(
            button_frame,
            text="Close",
            command=dialog.destroy
        ).pack(side=tk.RIGHT)
