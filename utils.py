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
    Thread-safe caching system with expiration and optional disk persistence.

    When a `file_path` is provided, the cache loads from disk on init and
    can be saved back with `save_to_disk()`.  Values are stored as JSON —
    dataclass objects that implement `to_dict()` are serialized with a
    `_cache_type` tag so they can be reconstructed on load.
    """

    # Registry of types that can be round-tripped through JSON
    _TYPE_REGISTRY: Dict[str, Any] = {}

    @classmethod
    def register_type(cls, type_class: Any) -> None:
        """Register a dataclass type for cache serialization."""
        cls._TYPE_REGISTRY[type_class.__name__] = type_class

    def __init__(self, expiry_seconds: int = CACHE_EXPIRY,
                 file_path: Optional[Union[str, Path]] = None):
        self._cache = SafeDict()
        self._expiry = expiry_seconds
        self._file_path = Path(file_path) if file_path else None
        self._dirty = False  # Track whether in-memory state differs from disk

        # Load from disk if a file was provided
        if self._file_path:
            self._load_from_disk()

    # ── public API (unchanged interface) ──────────────────────────

    def get(self, key: str) -> Optional[Any]:
        """Get a value from the cache if it exists and is not expired."""
        entry = self._cache.get(key)
        if entry is None:
            return None

        timestamp, value = entry
        if time.time() - timestamp > self._expiry:
            self._cache.delete(key)
            self._dirty = True
            return None

        return value

    def set(self, key: str, value: Any) -> None:
        """Store a value in the cache with current timestamp."""
        self._cache.set(key, (time.time(), value))
        self._dirty = True

    def delete(self, key: str) -> None:
        """Remove a value from the cache."""
        self._cache.delete(key)
        self._dirty = True

    def clear(self) -> None:
        """Clear all entries from the cache."""
        self._cache.clear()
        self._dirty = True

    def size(self) -> int:
        """Get the number of entries in the cache."""
        return self._cache.size()

    def prune(self) -> int:
        """Remove all expired entries. Returns count removed."""
        keys_to_delete = []
        now = time.time()

        for key, entry in self._cache.items():
            timestamp, _ = entry
            if now - timestamp > self._expiry:
                keys_to_delete.append(key)

        for key in keys_to_delete:
            self._cache.delete(key)

        if keys_to_delete:
            self._dirty = True
        return len(keys_to_delete)

    # ── disk persistence ──────────────────────────────────────────

    def save_to_disk(self) -> bool:
        """
        Persist the current cache to the JSON file specified at init.
        Returns True on success, False on failure or if no file was configured.
        """
        if not self._file_path:
            return False
        if not self._dirty:
            return True  # Nothing changed — skip the write

        try:
            # Prune expired entries before saving
            self.prune()

            serializable = {}
            for key, entry in self._cache.items():
                timestamp, value = entry
                serializable[key] = {
                    "t": timestamp,
                    "v": self._serialize_value(value)
                }

            # Ensure parent directory exists
            self._file_path.parent.mkdir(parents=True, exist_ok=True)

            # Atomic-ish write: write to temp file then rename
            tmp_path = self._file_path.with_suffix(".tmp")
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(serializable, f, separators=(",", ":"))
            tmp_path.replace(self._file_path)

            self._dirty = False
            logger.info(f"Cache saved: {self.size()} entries to {self._file_path.name}")
            return True

        except Exception as e:
            logger.error(f"Failed to save cache to {self._file_path}: {e}")
            return False

    def _load_from_disk(self) -> None:
        """Load cache entries from the JSON file."""
        if not self._file_path or not self._file_path.exists():
            return

        try:
            with open(self._file_path, "r", encoding="utf-8") as f:
                raw = json.load(f)

            now = time.time()
            loaded = 0
            expired = 0

            for key, entry in raw.items():
                timestamp = entry.get("t", 0)
                if now - timestamp > self._expiry:
                    expired += 1
                    continue  # Skip expired entries on load
                value = self._deserialize_value(entry.get("v"))
                self._cache.set(key, (timestamp, value))
                loaded += 1

            logger.info(f"Cache loaded: {loaded} entries from {self._file_path.name}"
                        f"{f' ({expired} expired, skipped)' if expired else ''}")

        except (json.JSONDecodeError, IOError, KeyError) as e:
            logger.warning(f"Could not load cache from {self._file_path}: {e}")

    # ── serialization helpers ─────────────────────────────────────

    @classmethod
    def _serialize_value(cls, value: Any) -> Any:
        """Convert a value to a JSON-safe representation."""
        if value is None:
            return None
        # Dataclass objects with to_dict()
        if hasattr(value, "to_dict") and hasattr(value, "__dataclass_fields__"):
            return {"_cache_type": type(value).__name__, "data": value.to_dict()}
        # Lists — recurse
        if isinstance(value, list):
            return [cls._serialize_value(item) for item in value]
        # Dicts — recurse on values
        if isinstance(value, dict):
            return {k: cls._serialize_value(v) for k, v in value.items()}
        # Primitives (str, int, float, bool) are already JSON-safe
        return value

    @classmethod
    def _deserialize_value(cls, raw: Any) -> Any:
        """Reconstruct a value from its JSON representation."""
        if raw is None:
            return None
        if isinstance(raw, dict):
            type_name = raw.get("_cache_type")
            if type_name and type_name in cls._TYPE_REGISTRY:
                type_class = cls._TYPE_REGISTRY[type_name]
                return type_class.from_dict(raw["data"])
            # Regular dict — recurse
            return {k: cls._deserialize_value(v) for k, v in raw.items()}
        if isinstance(raw, list):
            return [cls._deserialize_value(item) for item in raw]
        return raw

###############################################################################
# UI Widgets and Helpers — canonical location is now ui/widgets.py
# Re-exported here for backward compatibility with existing imports.
###############################################################################

from ui.widgets import (  # noqa: F401, E402
    SimpleTooltip,
    StatusBar,
    ProgressDialog,
    ScrollableFrame,
    ErrorCommunicator,
    ContextHelp,
)

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