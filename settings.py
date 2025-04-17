"""
settings.py

Centralized configuration for the Fleet Electrification Analyzer application.
This module contains all configuration parameters, styling options, and constants
used throughout the application.
"""
import os
import json
from pathlib import Path
import logging
from typing import Dict, Any

###############################################################################
# Application Information
###############################################################################
APP_NAME = "Fleet Electrification Analyzer"
APP_VERSION = "3.0.0"
APP_DESCRIPTION = "A tool for analyzing fleet vehicles and planning electrification strategies"
APP_AUTHOR = "Fleet Analytics"

###############################################################################
# File Paths & Directories
###############################################################################
# Application directories
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
CACHE_DIR = DATA_DIR / "cache"
EXPORT_DIR = DATA_DIR / "exports"
LOG_DIR = DATA_DIR / "logs"
TEMP_DIR = DATA_DIR / "temp"

# Ensure directories exist
for directory in [DATA_DIR, CACHE_DIR, EXPORT_DIR, LOG_DIR, TEMP_DIR]:
    directory.mkdir(exist_ok=True, parents=True)

# Default file paths
DEFAULT_CONFIG_FILE = DATA_DIR / "config.json"
DEFAULT_CACHE_FILE = CACHE_DIR / "api_cache.json"
DEFAULT_LOG_FILE = LOG_DIR / "app.log"

###############################################################################
# Logging Configuration
###############################################################################
LOG_LEVEL = logging.INFO
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

###############################################################################
# API Configuration
###############################################################################
# VIN Decoder API (NHTSA)
NHTSA_BASE_URL = "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/"
NHTSA_BATCH_URL = "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVINValuesBatch/"

# Fuel Economy API
FUELECONOMY_BASE_URL = "https://www.fueleconomy.gov/ws/rest/vehicle"
FUELECONOMY_MENU_URL = "https://www.fueleconomy.gov/ws/rest/vehicle/menu"

# API Request Parameters
API_TIMEOUT = 10  # seconds
MAX_RETRIES = 3
RETRY_DELAY = 1.0  # seconds
RATE_LIMIT_DELAY = 0.1  # seconds between API calls

###############################################################################
# Threading Configuration
###############################################################################
MAX_THREADS = 10  # Default maximum threads for parallel processing
MAX_QUEUE_SIZE = 1000  # Maximum size of processing queue

###############################################################################
# Vehicle Matching Parameters
###############################################################################
# Criteria weights for vehicle matching confidence
MATCHING_WEIGHTS = {
    "exact_vin": 100,
    "year_make_model": 80,
    "year_make": 60,
    "make_model": 50,
    "engine_match": 20,
    "displacement_match": 15,
    "cylinders_match": 10,
    "fuel_type_match": 10,
    "drive_match": 5,
    "transmission_match": 5
}

# Minimum confidence threshold for accepting a match
MIN_MATCH_CONFIDENCE = 60

###############################################################################
# User Interface Configuration
###############################################################################
# Window settings
DEFAULT_WINDOW_SIZE = "1400x900"
MIN_WINDOW_SIZE = "800x600"

# Colors
PRIMARY_HEX_1 = "#3C465A"     # Charcoal
PRIMARY_HEX_2 = "#FBFCFF"     # White
PRIMARY_HEX_3 = "#5B7553"     # Reseda green
SECONDARY_HEX_1 = "#C45911"   # Deep orange
SECONDARY_HEX_2 = "#D0CCD0"   # Light grey

# Other UI colors
SUCCESS_COLOR = "#28a745"     # Green for success messages
WARNING_COLOR = "#ffc107"     # Yellow for warnings
ERROR_COLOR = "#dc3545"       # Red for errors
INFO_COLOR = "#17a2b8"        # Blue for information

# Fonts
HEADING1_FONT = ("Franklin Gothic Book", 16, "bold")
HEADING2_FONT = ("Franklin Gothic Book", 14, "bold")
HEADING3_FONT = ("Franklin Gothic Book", 12, "bold")
BODY_FONT = ("Calibri", 11, "")
TABLE_FONT = ("Calibri", 10, "")
SMALL_FONT = ("Calibri", 9, "")

###############################################################################
# Data Analysis Configuration
###############################################################################
# Chart types available for visualization
CHART_TYPES = [
    "Body Class Distribution",
    "MPG Distribution",
    "CO2 Emissions Distribution",
    "CO2 vs MPG Correlation",
    "CO2 Comparison (Primary vs Alt)",
    "EV Range Distribution",
    "Make Frequency",
    "Model Distribution",
    "Fuel Type Distribution",
    "Annual Cost Comparison",
    "Fleet Age Distribution",
    "Electrification Potential"
]

# Default values for analysis calculations
DEFAULT_GAS_PRICE = 3.50          # $/gallon
DEFAULT_ELECTRICITY_PRICE = 0.13  # $/kWh
DEFAULT_EV_EFFICIENCY = 0.30      # kWh/mile
DEFAULT_ANNUAL_MILEAGE = 12000    # miles/year
DEFAULT_VEHICLE_LIFESPAN = 12     # years
DEFAULT_BATTERY_DEGRADATION = 2.0 # %/year
DEFAULT_ICE_MAINTENANCE = 0.10    # $/mile
DEFAULT_EV_MAINTENANCE = 0.06     # $/mile

###############################################################################
# Field Mappings & Conversions
###############################################################################
# Table column mapping (user-friendly names)
COLUMN_NAME_MAP = {
    "VIN": "VIN",
    "Year": "Model Year",
    "Make": "Make",
    "Model": "Model",
    "FuelTypePrimary": "Fuel Type (Primary)",
    "fuelType2": "Fuel Type (Secondary)",
    "BodyClass": "Body Class",
    "GVWR": "Gross Vehicle Wt",
    "MPG City": "MPG (City)",
    "MPG Highway": "MPG (Highway)",
    "MPG Combined": "MPG (Combined)",
    "CO2 emissions": "CO2 (Primary)",
    "co2A": "CO2 (Alt)",
    "rangeA": "Alt Range",
    "Odometer": "Odometer",
    "Annual Mileage": "Annual Mileage",
    "Asset ID": "Asset ID",
    "Department": "Department",
    "Location": "Location"
}

# Define field categories for UI organization
FIELD_CATEGORIES = {
    "Essential": [
        "VIN", "Year", "Make", "Model", "FuelTypePrimary", "BodyClass", 
        "GVWR", "MPG City", "MPG Highway", "MPG Combined", "CO2 emissions"
    ],
    "Alternative Fuel": [
        "fuelType2", "co2A", "rangeA", "phevBlended", "phevCity", "phevComb"
    ],
    "Technical": [
        "cylinders", "displ", "drive", "trany", "fuelCost08"
    ],
    "Fleet Management": [
        "Odometer", "Annual Mileage", "Asset ID", "Department", "Location"
    ],
    "Advanced": [
        "feScore", "ghgScore", "ghgScoreA", "combinedCD", "startStop", "VClass"
    ]
}

# Initial visible columns in the results table (can be customized by user)
DEFAULT_VISIBLE_COLUMNS = [
    "VIN", "Year", "Make", "Model", "FuelTypePrimary", "BodyClass", 
    "MPG Combined", "CO2 emissions", "Annual Mileage", "Asset ID"
]

###############################################################################
# Cache Configuration
###############################################################################
CACHE_ENABLED = True
CACHE_EXPIRY = 604800  # 7 days in seconds
MAX_CACHE_SIZE = 100000  # Maximum number of items in cache

# Default structure for the API cache
DEFAULT_API_CACHE = {
    "vin": {},
    "makes": {},
    "models": {},
    "options": {},
    "vehicledetails": {},
}

###############################################################################
# Export Configuration
###############################################################################
# Available export formats
EXPORT_FORMATS = {
    "CSV": ".csv",
    "Excel": ".xlsx",
    "PDF": ".pdf",
    "JSON": ".json",
    "HTML": ".html"
}

# Default export settings
DEFAULT_EXPORT_FORMAT = "Excel"
DEFAULT_EXPORT_FOLDER = str(EXPORT_DIR)

###############################################################################
# User Settings Management
###############################################################################
def load_user_settings() -> Dict[str, Any]:
    """
    Load user settings from the config file.
    Returns default settings if file doesn't exist or is invalid.
    """
    try:
        if os.path.exists(DEFAULT_CONFIG_FILE):
            with open(DEFAULT_CONFIG_FILE, 'r') as f:
                return json.load(f)
    except (json.JSONDecodeError, IOError) as e:
        logging.warning(f"Failed to load user settings: {e}")
    
    # Return default settings
    return {
        "window_size": DEFAULT_WINDOW_SIZE,
        "max_threads": MAX_THREADS,
        "visible_columns": DEFAULT_VISIBLE_COLUMNS,
        "gas_price": DEFAULT_GAS_PRICE,
        "electricity_price": DEFAULT_ELECTRICITY_PRICE,
        "ev_efficiency": DEFAULT_EV_EFFICIENCY,
        "annual_mileage": DEFAULT_ANNUAL_MILEAGE,
        "export_format": DEFAULT_EXPORT_FORMAT,
        "export_folder": DEFAULT_EXPORT_FOLDER
    }

def save_user_settings(settings: Dict[str, Any]) -> bool:
    """
    Save user settings to the config file.
    Returns True if successful, False otherwise.
    """
    try:
        with open(DEFAULT_CONFIG_FILE, 'w') as f:
            json.dump(settings, f, indent=4)
        return True
    except IOError as e:
        logging.error(f"Failed to save user settings: {e}")
        return False

# Load user settings on module import
USER_SETTINGS = load_user_settings()

###############################################################################
# All Fuel Economy API Fields
###############################################################################
# Complete list of all fields from Fuel Economy API for advanced users
ALL_FUEL_ECONOMY_FIELDS = [
    "atvtype",
    "barrels08",
    "barrelsA08",
    "charge120",
    "charge240",
    "charge240b",
    "c240Dscr",
    "c240bDscr",
    "city08",
    "city08U",
    "cityA08",
    "cityA08U",
    "cityCD",
    "cityE",
    "cityUF",
    "co2",
    "co2A",
    "co2TailpipeAGpm",
    "co2TailpipeGpm",
    "comb08",
    "comb08U",
    "combA08",
    "combA08U",
    "combE",
    "combinedCD",
    "combinedUF",
    "cylinders",
    "displ",
    "drive",
    "feScore",
    "fuelCost08",
    "fuelCostA08",
    "fuelType",
    "fuelType1",
    "fuelType2",
    "ghgScore",
    "ghgScoreA",
    "guzzler",
    "highway08",
    "highway08U",
    "highwayA08",
    "highwayA08U",
    "highwayCD",
    "highwayE",
    "highwayUF",
    "hlv",
    "hpv",
    "id",
    "lv2",
    "lv4",
    "make",
    "mfrCode",
    "model",
    "mpgData",
    "phevBlended",
    "phevCity",
    "phevHwy",
    "phevComb",
    "pv2",
    "pv4",
    "rangeA",
    "rangeCityA",
    "rangeHwyA",
    "sCharger",
    "tCharger",
    "trans_dscr",
    "trany",
    "UCity",
    "UCityA",
    "UHighway",
    "UHighwayA",
    "VClass",
    "year",
    "youSaveSpend",
    "createdOn",
    "modifiedOn",
    "startStop",
]