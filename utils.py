"""
utils.py

Utility functions and helpers for the Fleet Electrification Analyzer.
Contains reusable tools, logging setup, and helper classes.
"""

import os
import time
import json
import logging
import datetime
import threading
from typing import Dict, Any, List, Optional, Callable, Union, Tuple
from pathlib import Path

import tkinter as tk
from tkinter import ttk

from settings import LOG_FORMAT, LOG_LEVEL, LOG_DATE_FORMAT, DEFAULT_LOG_FILE, CACHE_EXPIRY

###############################################################################
# Logging Setup
###############################################################################

def setup_logging(log_file: Optional[str] = None, console: bool = True, level: int = LOG_LEVEL) -> logging.Logger:
    """
    Configure application logging with file and/or console handlers.
    
    Args:
        log_file: Path to log file. If None, uses the default path.
        console: Whether to also log to console.
        
    Returns:
        Logger instance for the application.
    """
    logger = logging.getLogger("fleet_analyzer")
    logger.setLevel(LOG_LEVEL)
    formatter = logging.Formatter(LOG_FORMAT, datefmt=LOG_DATE_FORMAT)
    
    # Clear existing handlers to avoid duplicates
    if logger.handlers:
        logger.handlers.clear()
    
    # File handler
    if log_file is None:
        log_file = DEFAULT_LOG_FILE
        
    # Ensure log directory exists
    log_dir = os.path.dirname(log_file)
    os.makedirs(log_dir, exist_ok=True)
    
    file_handler = logging.FileHandler(log_file)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # Console handler (optional)
    if console:
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    
    return logger

# Create a default logger instance
logger = setup_logging()

###############################################################################
# Time and Date Utilities
###############################################################################

def timestamp() -> str:
    """
    Returns a string timestamp of the current date and time.
    Format: YYYY-MM-DD HH:MM:SS
    """
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def format_timestamp(dt: Optional[datetime.datetime] = None) -> str:
    """
    Format a datetime object as a string timestamp.
    If no datetime is provided, uses the current time.
    """
    if dt is None:
        dt = datetime.datetime.now()
    return dt.strftime("%Y-%m-%d %H:%M:%S")

def elapsed_time(start_time: float) -> str:
    """
    Calculate and format the elapsed time since start_time.
    
    Args:
        start_time: Starting time from time.time()
        
    Returns:
        Formatted string of elapsed time (e.g., "5.2 seconds" or "2 minutes 15 seconds")
    """
    elapsed = time.time() - start_time
    if elapsed < 60:
        return f"{elapsed:.1f} seconds"
    else:
        minutes = int(elapsed // 60)
        seconds = int(elapsed % 60)
        return f"{minutes} minutes {seconds} seconds"

###############################################################################
# File Handling Utilities
###############################################################################

def ensure_dir(directory: Union[str, Path]) -> Path:
    """
    Ensure a directory exists, creating it if necessary.
    
    Args:
        directory: Path to the directory
        
    Returns:
        Path object of the directory
    """
    path = Path(directory)
    path.mkdir(parents=True, exist_ok=True)
    return path

def safe_filename(name: str) -> str:
    """
    Convert a string to a safe filename by removing invalid characters.
    
    Args:
        name: Original filename
        
    Returns:
        Safe filename string
    """
    # Replace invalid filename characters
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '_')
    
    # Handle special cases
    if name.startswith('.'):
        name = '_' + name
    
    # Limit length to avoid path length issues on some systems
    if len(name) > 200:
        name = name[:197] + '...'
        
    return name

def load_json_file(file_path: Union[str, Path], default: Any = None) -> Any:
    """
    Load and parse a JSON file.
    
    Args:
        file_path: Path to the JSON file
        default: Value to return if file doesn't exist or is invalid
        
    Returns:
        Parsed JSON data or default value
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, IOError) as e:
        logger.warning(f"Failed to load JSON file {file_path}: {e}")
        return {} if default is None else default

def save_json_file(data: Any, file_path: Union[str, Path], indent: int = 2) -> bool:
    """
    Save data to a JSON file.
    
    Args:
        data: Data to save
        file_path: Path to save the file
        indent: JSON indentation level
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Ensure directory exists
        ensure_dir(Path(file_path).parent)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=indent)
        return True
    except IOError as e:
        logger.error(f"Failed to save JSON file {file_path}: {e}")
        return False

###############################################################################
# String and Data Processing Utilities
###############################################################################

def case_insensitive_equal(a: str, b: str) -> bool:
    """
    Compare two strings case-insensitively after stripping whitespace.
    
    Args:
        a: First string
        b: Second string
        
    Returns:
        True if strings are equal (ignoring case and whitespace)
    """
    return a.strip().lower() == b.strip().lower()

def normalize_vehicle_model(model: str) -> str:
    """
    Normalize a vehicle model name for comparison.
    Removes special characters and standardizes common variations.
    
    Args:
        model: Vehicle model name
        
    Returns:
        Normalized model name
    """
    # Convert to lowercase and strip whitespace
    normalized = model.strip().lower()
    
    # Remove special characters but preserve numbers
    normalized = ''.join(c for c in normalized if c.isalnum() or c.isspace())
    
    # Standardize spacing
    normalized = ' '.join(normalized.split())
    
    # Handle special cases like F-150 -> f150
    normalized = normalized.replace('f 150', 'f150')
    normalized = normalized.replace('f 250', 'f250')
    normalized = normalized.replace('f 350', 'f350')
    normalized = normalized.replace('e 150', 'e150')
    normalized = normalized.replace('e 250', 'e250')
    normalized = normalized.replace('e 350', 'e350')
    
    return normalized

def safe_cast(value: Any, to_type: Callable, default: Any = None) -> Any:
    """
    Safely cast a value to a specified type.
    
    Args:
        value: Value to cast
        to_type: Type function (int, float, str, etc.)
        default: Value to return if casting fails
        
    Returns:
        Converted value or default
    """
    try:
        return to_type(value)
    except (ValueError, TypeError):
        return default

def format_number(value: Optional[Union[int, float]], precision: int = 1) -> str:
    """
    Format a number with thousand separators and fixed precision.
    
    Args:
        value: Number to format
        precision: Decimal precision
        
    Returns:
        Formatted string representation
    """
    if value is None:
        return ""
    
    try:
        if isinstance(value, int) or value.is_integer():
            return f"{int(value):,}"
        else:
            return f"{value:,.{precision}f}"
    except (ValueError, AttributeError):
        return str(value)

###############################################################################
# Thread-Safe Collections
###############################################################################

class SafeDict:
    """Thread-safe dictionary implementation using a lock."""
    
    def __init__(self, initial_data: Optional[Dict] = None):
        self._dict = initial_data or {}
        self._lock = threading.RLock()
    
    def get(self, key: Any, default: Any = None) -> Any:
        with self._lock:
            return self._dict.get(key, default)
    
    def set(self, key: Any, value: Any) -> None:
        with self._lock:
            self._dict[key] = value
    
    def delete(self, key: Any) -> None:
        with self._lock:
            if key in self._dict:
                del self._dict[key]
    
    def contains(self, key: Any) -> bool:
        with self._lock:
            return key in self._dict
    
    def clear(self) -> None:
        with self._lock:
            self._dict.clear()
    
    def items(self) -> List[Tuple[Any, Any]]:
        with self._lock:
            return list(self._dict.items())
    
    def keys(self) -> List[Any]:
        with self._lock:
            return list(self._dict.keys())
    
    def values(self) -> List[Any]:
        with self._lock:
            return list(self._dict.values())
    
    def to_dict(self) -> Dict:
        with self._lock:
            return dict(self._dict)
    
    def update(self, other: Dict) -> None:
        with self._lock:
            self._dict.update(other)
    
    def size(self) -> int:
        with self._lock:
            return len(self._dict)

###############################################################################
# Caching
###############################################################################

class Cache:
    """
    Thread-safe caching system with expiration.
    """
    
    def __init__(self, expiry_seconds: int = CACHE_EXPIRY):
        self._cache = SafeDict()
        self._expiry = expiry_seconds
    
    def get(self, key: str) -> Optional[Any]:
        """
        Get a value from the cache if it exists and is not expired.
        
        Args:
            key: Cache key
            
        Returns:
            Cached value or None if not found or expired
        """
        entry = self._cache.get(key)
        if entry is None:
            return None
        
        timestamp, value = entry
        if time.time() - timestamp > self._expiry:
            # Expired
            self._cache.delete(key)
            return None
        
        return value
    
    def set(self, key: str, value: Any) -> None:
        """
        Store a value in the cache with current timestamp.
        
        Args:
            key: Cache key
            value: Value to store
        """
        self._cache.set(key, (time.time(), value))
    
    def delete(self, key: str) -> None:
        """
        Remove a value from the cache.
        
        Args:
            key: Cache key to remove
        """
        self._cache.delete(key)
    
    def clear(self) -> None:
        """Clear all entries from the cache."""
        self._cache.clear()
    
    def size(self) -> int:
        """Get the number of entries in the cache."""
        return self._cache.size()
    
    def prune(self) -> int:
        """
        Remove all expired entries from the cache.
        
        Returns:
            Number of entries removed
        """
        keys_to_delete = []
        now = time.time()
        
        for key, entry in self._cache.items():
            timestamp, _ = entry
            if now - timestamp > self._expiry:
                keys_to_delete.append(key)
        
        for key in keys_to_delete:
            self._cache.delete(key)
        
        return len(keys_to_delete)

###############################################################################
# UI Widgets and Helpers
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
        """
        Initialize tooltip with widget and text.
        
        Args:
            widget: The widget to attach the tooltip to
            text: Tooltip text
            delay: Delay in ms before showing tooltip
            fg: Text color
            bg: Background color
            padx: Horizontal padding
            pady: Vertical padding
            font: Font to use (None for default)
        """
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
        
        # Position tooltip below and slightly to the right of the widget
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 10
        
        # Create tooltip window
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)  # Remove window decorations
        tw.wm_geometry(f"+{x}+{y}")
        
        # Create tooltip content
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
        """
        Initialize the status bar.
        
        Args:
            master: Parent widget
            **kwargs: Additional keyword arguments for ttk.Frame
        """
        super().__init__(master, **kwargs)
        
        self.sections = {}
        
        # Main status section (left-aligned)
        self.sections["main"] = ttk.Label(
            self, relief="sunken", anchor="w", padding=(5, 2)
        )
        self.sections["main"].pack(side="left", fill="x", expand=True)
        
        # Set default message
        self.set("Ready")
    
    def add_section(self, name, width=None, side="right"):
        """
        Add a new section to the status bar.
        
        Args:
            name: Section identifier
            width: Width in characters (None for auto)
            side: Side to place section ("left" or "right")
        """
        if name in self.sections:
            return
        
        self.sections[name] = ttk.Label(
            self, relief="sunken", anchor="w", padding=(5, 2), width=width
        )
        self.sections[name].pack(side=side, fill="y", padx=(1, 0))
        
        # Set default text
        self.set("", section=name)
    
    def set(self, text, section="main"):
        """
        Set the text for a section.
        
        Args:
            text: Text to display
            section: Section identifier
        """
        if section in self.sections:
            self.sections[section].config(text=text)

class ProgressDialog(tk.Toplevel):
    """
    Modal dialog with a progress bar and cancel button.
    
    Example:
        progress = ProgressDialog(root, "Processing Files", "Please wait...")
        for i in range(100):
            if progress.cancelled:
                break
            progress.update(i)
            # Do work...
        progress.destroy()
    """
    
    def __init__(self, parent, title, message, maximum=100, cancelable=True):
        """
        Initialize the progress dialog.
        
        Args:
            parent: Parent window
            title: Dialog title
            message: Message to display
            maximum: Maximum progress value
            cancelable: Whether the operation can be cancelled
        """
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        self.cancelled = False
        self.maximum = maximum
        
        # Calculate position (center on parent)
        width = 350
        height = 150
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        x = parent_x + (parent_width // 2) - (width // 2)
        y = parent_y + (parent_height // 2) - (height // 2)
        
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # Message label
        self.message_var = tk.StringVar(value=message)
        self.message_label = ttk.Label(
            self, textvariable=self.message_var, wraplength=width-20
        )
        self.message_label.pack(padx=10, pady=(10, 5))
        
        # Progress bar
        self.progress_var = tk.DoubleVar(value=0)
        self.progressbar = ttk.Progressbar(
            self, orient=tk.HORIZONTAL, length=300,
            mode='determinate', variable=self.progress_var,
            maximum=maximum
        )
        self.progressbar.pack(padx=10, pady=5)
        
        # Status label
        self.status_var = tk.StringVar(value="Starting...")
        self.status_label = ttk.Label(
            self, textvariable=self.status_var
        )
        self.status_label.pack(padx=10, pady=5)
        
        # Cancel button (if cancelable)
        if cancelable:
            self.cancel_button = ttk.Button(
                self, text="Cancel", command=self.cancel
            )
            self.cancel_button.pack(pady=10)
        
        # Update the UI
        self.update_idletasks()
    
    def update(self, value, message=None, status=None):
        """
        Update the progress and optionally the message.
        
        Args:
            value: Current progress value
            message: New message (or None to keep current)
            status: Status text (or None to keep current)
        """
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

class ScrollableFrame(ttk.Frame):
    """
    A frame with scrollbars that can contain other widgets.
    
    Example:
        scrollable = ScrollableFrame(root)
        scrollable.pack(fill="both", expand=True)
        
        # Add widgets to the scrollable area
        for i in range(50):
            ttk.Label(scrollable.scrollable_frame, text=f"Row {i}").pack()
    """
    
    def __init__(self, master, **kwargs):
        """
        Initialize the scrollable frame.
        
        Args:
            master: Parent widget
            **kwargs: Additional keyword arguments for ttk.Frame
        """
        super().__init__(master, **kwargs)
        
        # Create a canvas widget
        self.canvas = tk.Canvas(self)
        
        # Create vertical scrollbar
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        
        # Create horizontal scrollbar
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=self.hsb.set)
        
        # Layout scrollbars and canvas
        self.vsb.pack(side="right", fill="y")
        self.hsb.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Create the scrollable frame inside the canvas
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        # Bind frame size changes to update canvas scrollregion
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # Create window in canvas to contain the frame
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw"
        )
        
        # Adjust window size when canvas size changes
        self.canvas.bind("<Configure>", self._configure_canvas_window)
        
        # Bind mouse wheel to scroll
        self.scrollable_frame.bind("<Enter>", self._bind_mousewheel)
        self.scrollable_frame.bind("<Leave>", self._unbind_mousewheel)
    
    def _configure_canvas_window(self, event):
        """Adjust the width of the canvas window when canvas size changes."""
        self.canvas.itemconfig(
            self.canvas_window, width=event.width
        )
    
    def _bind_mousewheel(self, event):
        """Bind mouse wheel to scroll vertically."""
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        # Linux support
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
            # Scroll up
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            # Scroll down
            self.canvas.yview_scroll(1, "units")

###############################################################################
# Data Validation
###############################################################################

def validate_vin(vin: str) -> bool:
    """
    Check if a VIN is valid according to basic rules.
    
    Args:
        vin: Vehicle Identification Number to validate
        
    Returns:
        True if VIN passes basic validation
    """
    # Basic validation (more comprehensive validation would include checksum)
    if not vin:
        return False
    
    # Remove spaces and convert to uppercase
    vin = vin.replace(" ", "").upper()
    
    # Check length (should be 17 characters for modern vehicles)
    if len(vin) != 17:
        return False
    
    # Check for invalid characters
    valid_chars = set("0123456789ABCDEFGHJKLMNPRSTUVWXYZ")
    if not all(c in valid_chars for c in vin):
        return False
    
    return True

def validate_year(year: str) -> bool:
    """
    Check if a year string is valid.
    
    Args:
        year: Year string to validate
        
    Returns:
        True if year is valid
    """
    try:
        year_int = int(year)
        current_year = datetime.datetime.now().year
        return 1900 <= year_int <= current_year + 1
    except ValueError:
        return False