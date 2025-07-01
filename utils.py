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
    Enhanced for commercial vehicle matching (F-150 vs F150, etc.).
    
    Args:
        model: Vehicle model name
        
    Returns:
        Normalized model name
    """
    import re
    
    # Convert to lowercase and strip whitespace
    normalized = model.strip().lower()
    
    # Enhanced Ford F-series normalization (F-150, F-250, etc. -> f150, f250)
    # Handles F-150, F 150, F150 all to f150
    normalized = re.sub(r'f[-\s]*(\d{3})', r'f\1', normalized)
    
    # Enhanced Ford E-series normalization (E-250, E-350, etc. -> e250, e350)  
    # Handles E-250, E 250, E250 all to e250
    normalized = re.sub(r'e[-\s]*(\d{3})', r'e\1', normalized)
    
    # Ford Super Duty normalization (remove "super duty" text)
    normalized = re.sub(r'\bsuper\s*duty\b', '', normalized)
    
    # Chevy/GMC commercial vehicle patterns
    # Silverado 1500 -> silverado1500, Sierra 2500 -> sierra2500
    normalized = re.sub(r'(silverado|sierra)[-\s]*(\d{4})', r'\1\2', normalized)
    
    # Ram truck patterns (Ram 1500 -> ram1500)
    normalized = re.sub(r'ram[-\s]*(\d{4})', r'ram\1', normalized)
    
    # Express/Savana van patterns (Express 2500 -> express2500)
    normalized = re.sub(r'(express|savana)[-\s]*(\d{4})', r'\1\2', normalized)
    
    # Remove special characters but preserve numbers
    normalized = ''.join(c for c in normalized if c.isalnum() or c.isspace())
    
    # Standardize spacing
    normalized = ' '.join(normalized.split())
    
    # Additional legacy F-series handling for any remaining cases
    normalized = normalized.replace('f 150', 'f150')
    normalized = normalized.replace('f 250', 'f250') 
    normalized = normalized.replace('f 350', 'f350')
    normalized = normalized.replace('f 450', 'f450')
    normalized = normalized.replace('f 550', 'f550')
    
    # Additional legacy E-series handling
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
    # Use detailed validation and return only True/False for backward compatibility
    is_valid, _ = validate_vin_detailed(vin)
    return is_valid

def validate_vin_detailed(vin: str) -> Tuple[bool, str]:
    """
    Check if a VIN is valid and provide detailed error information.
    
    Args:
        vin: Vehicle Identification Number to validate
        
    Returns:
        Tuple of (is_valid, error_message)
        If valid: (True, "")
        If invalid: (False, "User-friendly error description")
    """
    # Check if VIN is provided
    if not vin or not vin.strip():
        return False, "VIN is required - please provide a Vehicle Identification Number"
    
    # Clean the VIN
    original_vin = vin
    vin = vin.replace(" ", "").replace("-", "").upper().strip()
    
    # Check length (should be 17 characters for modern vehicles)
    if len(vin) < 17:
        return False, f"VIN is too short - found {len(vin)} characters, need 17 (example: 1HGBH41JXMN109186)"
    elif len(vin) > 17:
        return False, f"VIN is too long - found {len(vin)} characters, need exactly 17"
    
    # Check for invalid characters (VINs exclude I, O, Q to avoid confusion with 1, 0)
    valid_chars = set("0123456789ABCDEFGHJKLMNPRSTUVWXYZ")
    invalid_chars = [c for c in vin if c not in valid_chars]
    
    if invalid_chars:
        invalid_list = ", ".join(set(invalid_chars))
        return False, f"VIN contains invalid characters: {invalid_list} (VINs cannot contain I, O, or Q)"
    
    # Check for common mistakes
    if vin.count('0') > 8:  # Too many zeros might indicate placeholder
        return False, "VIN appears to contain placeholder zeros - please verify the actual VIN"
    
    if vin == "1" * 17 or vin == "0" * 17:  # Obvious test data
        return False, "VIN appears to be test data - please provide the actual vehicle VIN"
    
    return True, ""

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

###############################################################################
# Commercial Vehicle Detection - Step 12 Enhancement
###############################################################################

def extract_gvwr_pounds(gvwr_text: str) -> Tuple[float, str]:
    """
    Extract numeric GVWR value from various text formats.
    
    Args:
        gvwr_text: Raw GVWR text from NHTSA API
        
    Returns:
        Tuple of (gvwr_pounds, commercial_category)
        
    Examples:
        "Class 1: 6,000 lb (2,722 kg)" -> (6000.0, "Light Duty")
        "8500" -> (8500.0, "Light Duty")
        "19,500 lb" -> (19500.0, "Medium Duty")
    """
    import re
    
    if not gvwr_text:
        return 0.0, ""
    
    # Patterns to extract GVWR in pounds
    patterns = [
        r'(\d{1,2},?\d{3})\s*(?:lb|lbs|pounds?)',  # "6,000 lb" or "6000 lb"
        r'(\d{1,2},?\d{3})\s*(?:\(|$)',  # "6,000 (" or "6000" at end
        r':\s*(\d{1,2},?\d{3})',  # ": 6,000"
        r'(\d{4,6})'  # Raw numbers like "6000"
    ]
    
    gvwr_pounds = 0.0
    for pattern in patterns:
        match = re.search(pattern, gvwr_text.replace(',', ''))
        if match:
            try:
                gvwr_pounds = float(match.group(1).replace(',', ''))
                break
            except ValueError:
                continue
    
    # Classify commercial category based on GVWR
    if gvwr_pounds > 0:
        if gvwr_pounds <= 8500:
            commercial_category = "Light Duty"
        elif gvwr_pounds <= 19500:
            commercial_category = "Medium Duty" 
        elif gvwr_pounds <= 33000:
            commercial_category = "Heavy Duty"
        else:
            commercial_category = "Extra Heavy Duty"
    else:
        commercial_category = ""
    
    return gvwr_pounds, commercial_category

def detect_commercial_vehicle(make: str, model: str, body_class: str, 
                            gvwr_pounds: float = 0.0) -> bool:
    """
    Determine if a vehicle is likely a commercial vehicle.
    
    Args:
        make: Vehicle manufacturer
        model: Vehicle model
        body_class: Vehicle body class
        gvwr_pounds: Gross vehicle weight rating in pounds
        
    Returns:
        True if vehicle appears to be commercial
    """
    # Commercial indicators by body class
    commercial_body_classes = [
        "truck", "van", "bus", "chassis cab", "cutaway", "pickup",
        "cargo van", "passenger van", "step van", "box truck",
        "utility", "commercial", "work truck", "crew cab"
    ]
    
    # Commercial vehicle models (common fleet vehicles)
    commercial_models = [
        # Ford commercial lineup
        "transit", "e-series", "f-150", "f-250", "f-350", "f-450", "f-550",
        "f150", "f250", "f350", "f450", "f550", "econoline",
        
        # GM commercial lineup  
        "express", "savana", "silverado", "sierra", 
        
        # Chrysler/RAM commercial lineup
        "promaster", "ram 1500", "ram 2500", "ram 3500", "ram1500", "ram2500", "ram3500",
        
        # Other commercial vehicles
        "sprinter", "metris", "nv200", "nv1500", "nv2500", "nv3500",
        "city express", "connect"
    ]
    
    # Check body class
    body_lower = body_class.lower()
    for indicator in commercial_body_classes:
        if indicator in body_lower:
            return True
    
    # Check model name
    model_lower = model.lower()
    for indicator in commercial_models:
        if indicator in model_lower:
            return True
    
    # Check GVWR (vehicles over 8,500 lbs are typically commercial)
    if gvwr_pounds > 8500:
        return True
    
    # Check make-specific patterns
    make_lower = make.lower()
    if make_lower in ["isuzu", "mack", "peterbilt", "kenworth", "freightliner", "volvo trucks"]:
        return True  # These are primarily commercial vehicle manufacturers
    
    return False

def detect_diesel_engine(fuel_type_primary: str = "", fuel_type_secondary: str = "", 
                        engine_type: str = "", model: str = "") -> bool:
    """
    Detect if a vehicle uses diesel fuel.
    
    Args:
        fuel_type_primary: Primary fuel type from NHTSA
        fuel_type_secondary: Secondary fuel type from NHTSA
        engine_type: Engine type/configuration from NHTSA
        model: Vehicle model name
        
    Returns:
        True if vehicle appears to use diesel
    """
    diesel_indicators = [
        "diesel", "biodiesel", "b20", "b100", 
        "compression ignition", "ci", "cng/diesel",
        "diesel electric", "hybrid diesel"
    ]
    
    # Check all fuel and engine fields
    fields_to_check = [
        fuel_type_primary.lower(),
        fuel_type_secondary.lower(), 
        engine_type.lower()
    ]
    
    for field in fields_to_check:
        for indicator in diesel_indicators:
            if indicator in field:
                return True
    
    # Check model name for diesel indicators
    model_lower = model.lower()
    diesel_model_indicators = [
        "duramax", "powerstroke", "cummins", "ecodiesel", "bluetec"
    ]
    
    for indicator in diesel_model_indicators:
        if indicator in model_lower:
            return True
    
    return False

def classify_commercial_category(gvwr_pounds: float, body_class: str = "") -> str:
    """
    Classify commercial category based on GVWR and body class.
    
    Args:
        gvwr_pounds: Gross vehicle weight rating in pounds
        body_class: Vehicle body class (optional)
        
    Returns:
        Commercial category string
    """
    # GVWR-based classification (US DOT standards)
    if gvwr_pounds <= 0:
        return ""
    elif gvwr_pounds <= 8500:
        return "Light Duty"
    elif gvwr_pounds <= 19500:
        return "Medium Duty"
    elif gvwr_pounds <= 33000:
        return "Heavy Duty"
    else:
        return "Extra Heavy Duty"

def get_commercial_summary(commercial_category: str, is_diesel: bool, 
                          gvwr_pounds: float, is_commercial: bool) -> str:
    """
    Generate a summary string of commercial vehicle characteristics.
    
    Args:
        commercial_category: Light/Medium/Heavy Duty classification
        is_diesel: Whether vehicle uses diesel
        gvwr_pounds: Gross vehicle weight in pounds
        is_commercial: Whether classified as commercial vehicle
        
    Returns:
        Human-readable commercial summary string
    """
    if not is_commercial:
        return "Passenger Vehicle"
    
    parts = []
    
    if commercial_category:
        parts.append(commercial_category)
    
    if is_diesel:
        parts.append("Diesel")
    
    if gvwr_pounds > 0:
        parts.append(f"GVWR: {gvwr_pounds:,.0f} lb")
    
    return " | ".join(parts) if parts else "Commercial Vehicle"

def extract_engine_power(engine_hp: str, engine_kw: str) -> Tuple[str, str]:
    """
    Clean and format engine power values.
    
    Args:
        engine_hp: Horsepower string from NHTSA
        engine_kw: Kilowatt string from NHTSA
        
    Returns:
        Tuple of (formatted_hp, formatted_kw)
    """
    import re
    
    # Clean HP value
    hp_clean = ""
    if engine_hp:
        # Extract numeric value from HP string
        hp_match = re.search(r'(\d+(?:\.\d+)?)', str(engine_hp))
        if hp_match:
            hp_clean = f"{hp_match.group(1)} HP"
    
    # Clean KW value  
    kw_clean = ""
    if engine_kw:
        # Extract numeric value from KW string
        kw_match = re.search(r'(\d+(?:\.\d+)?)', str(engine_kw))
        if kw_match:
            kw_clean = f"{kw_match.group(1)} kW"
    
    return hp_clean, kw_clean

# Enhanced error communication utilities
class ErrorCommunicator:
    """Enhanced error communication with user-friendly messages and suggested fixes."""
    
    ERROR_CATEGORIES = {
        "vin_format": {
            "title": "VIN Format Error",
            "icon": "❌",
            "color": "#d32f2f"
        },
        "file_access": {
            "title": "File Access Error", 
            "icon": "📁",
            "color": "#f57c00"
        },
        "api_error": {
            "title": "Data Lookup Error",
            "icon": "🌐", 
            "color": "#1976d2"
        },
        "processing": {
            "title": "Processing Error",
            "icon": "⚙️",
            "color": "#7b1fa2"
        },
        "validation": {
            "title": "Data Validation Error",
            "icon": "⚠️",
            "color": "#f57c00"
        }
    }
    
    @classmethod
    def show_error_dialog(cls, parent, category: str, message: str, details: str = "", 
                         suggested_fixes: List[str] = None, context_help: str = ""):
        """
        Show an enhanced error dialog with category-specific styling and helpful suggestions.
        
        Args:
            parent: Parent window
            category: Error category from ERROR_CATEGORIES
            message: Main error message
            details: Additional details
            suggested_fixes: List of suggested solutions
            context_help: Context-sensitive help information
        """
        category_info = cls.ERROR_CATEGORIES.get(category, cls.ERROR_CATEGORIES["processing"])
        
        # Create custom dialog
        dialog = tk.Toplevel(parent)
        dialog.title(category_info["title"])
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()
        
        # Calculate position (center on parent)
        width = 500
        height = 400
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        x = parent_x + (parent_width // 2) - (width // 2)
        y = parent_y + (parent_height // 2) - (height // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        
        # Main frame
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header with icon and title
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
        
        # Main message
        message_label = tk.Label(
            main_frame,
            text=message,
            font=("Segoe UI", 11),
            wraplength=450,
            justify=tk.LEFT
        )
        message_label.pack(fill=tk.X, pady=(0, 10))
        
        # Details section
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
        
        # Suggested fixes
        if suggested_fixes:
            fixes_frame = ttk.LabelFrame(main_frame, text="💡 Suggested Solutions")
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
        
        # Context help
        if context_help:
            help_frame = ttk.LabelFrame(main_frame, text="ℹ️ Additional Help")
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
        
        # Buttons
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
        import datetime
        
        error_report = f"""Fleet Electrification Analyzer - Error Report
Generated: {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

Category: {category}
Message: {message}

Details:
{details}

Please include this information when reporting issues.
"""
        
        # Copy to clipboard
        root = tk._default_root
        if root:
            root.clipboard_clear()
            root.clipboard_append(error_report)
    
    @classmethod
    def get_vin_error_message(cls, vin: str, error_type: str) -> Tuple[str, List[str]]:
        """
        Get user-friendly VIN error message with suggested fixes.
        
        Args:
            vin: The problematic VIN
            error_type: Type of VIN error
            
        Returns:
            Tuple of (message, suggested_fixes)
        """
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
        """
        Get user-friendly file error message with suggested fixes.
        
        Args:
            filepath: The problematic file path
            error_type: Type of file error
            
        Returns:
            Tuple of (message, suggested_fixes)
        """
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

class ContextHelp:
    """Context-sensitive help system for user guidance."""
    
    HELP_TOPICS = {
        "vin_format": """
VIN (Vehicle Identification Number) Format:
• Must be exactly 17 characters long
• Contains letters and numbers only
• Cannot contain I, O, or Q (to avoid confusion with 1, 0)
• Each position has a specific meaning for vehicle details

Common VIN locations on vehicles:
• Dashboard (visible through windshield)
• Driver's side door frame
• Vehicle registration documents
• Insurance cards
        """,
        
        "csv_format": """
CSV File Format Requirements:
• Must have a 'VIN' or 'VINs' column header
• One VIN per row
• Additional columns are preserved (Asset ID, Department, etc.)
• Use comma as delimiter
• UTF-8 encoding recommended

Example format:
VIN,Asset ID,Department
1HGBH41JXMN109186,TRUCK001,Public Works
1FTFW1ET5DKE55321,VAN002,Parks & Recreation
        """,
        
        "processing_options": """
Processing Options:
• Max Threads: Number of parallel processing threads (1-32)
  - Higher values = faster processing
  - Lower values = less system resource usage
  
• Skip Existing VINs: Skip VINs already in output file
  - Useful for resuming interrupted processing
  - Prevents duplicate entries

• Auto-generate filename: Creates timestamped output files
  - Format: filename_fleet_analysis_YYYYMMDD_HHMMSS.csv
  - Prevents accidental overwrites
        """,
        
        "data_quality": """
Data Quality Indicators:
• High (80-100%): Complete vehicle information available
• Medium (50-80%): Some details missing but usable
• Low (0-50%): Limited information found
• Failed: VIN invalid or no data found

Quality factors:
• VIN validity and format
• Successful API data retrieval
• Commercial vehicle data completeness
• Cross-source data consistency
        """
    }
    
    @classmethod
    def show_help_dialog(cls, parent, topic: str, title: str = "Help"):
        """Show context-sensitive help dialog."""
        help_text = cls.HELP_TOPICS.get(topic, "Help information not available for this topic.")
        
        # Create help dialog
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.resizable(True, True)
        dialog.transient(parent)
        
        # Size and position
        width = 600
        height = 400
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        x = parent_x + (parent_width // 2) - (width // 2)
        y = parent_y + (parent_height // 2) - (height // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        
        # Content frame
        content_frame = ttk.Frame(dialog)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Help text
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
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        # Close button
        button_frame = ttk.Frame(content_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(
            button_frame,
            text="Close",
            command=dialog.destroy
        ).pack(side=tk.RIGHT)