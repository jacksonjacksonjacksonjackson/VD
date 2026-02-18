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
APP_VERSION = "3.0.3"
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
# Threading Configuration (Optimized for Commercial Vehicle Processing)
###############################################################################
MAX_THREADS = 8   # Reduced for stability with web scraping
MAX_QUEUE_SIZE = 500  # Reduced queue size for better memory management

# Advanced threading settings
THREADING_CONFIG = {
    "core_threads": 4,              # Core processing threads (non-scraping)
    "scraping_threads": 2,          # Dedicated scraping threads
    "io_threads": 2,                # File I/O and database operations
    "batch_size": 50,               # VINs processed per batch
    "scraping_batch_size": 5,       # VINs scraped per batch
    "timeout_seconds": 300,         # Total processing timeout
    "memory_limit_mb": 512,         # Memory usage limit
}

###############################################################################
# Commercial Vehicle Web Scraping Configuration
###############################################################################
# Scraping behavior settings
SCRAPING_CONFIG = {
    # Browser simulation
    "user_agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "request_timeout": 12,           # Increased timeout for stability
    "selenium_timeout": 20,          # Increased Selenium timeout
    
    # Rate limiting (production-friendly)
    "rate_limit_delay": 1.5,         # Conservative delay between requests
    "max_retries": 2,                # Reduced retries for faster processing
    "backoff_factor": 1.5,           # Exponential backoff multiplier
    
    # Caching strategy
    "cache_expiry_hours": 336,       # Extended cache (2 weeks) for stability
    "cache_cleanup_interval": 24,    # Hours between cache cleanup
    "max_cache_size_mb": 100,        # Cache size limit
    
    # Quality and performance
    "confidence_threshold": 0.65,    # Slightly lower for more data
    "enable_scraping": True,         # Global scraping enable/disable
    "use_selenium": False,           # Prefer requests for better performance
    "max_concurrent_scrapes": 2,     # Reduced for politeness
    "scraping_batch_size": 5,        # Process in small batches
    
    # Fallback behavior
    "enable_estimates": True,        # Allow estimated specs when scraping fails
    "estimation_confidence": 0.5,    # Confidence for estimated values
    "graceful_degradation": True,    # Continue processing if scraping fails
}

###############################################################################
# Commercial Vehicle Help Documentation
###############################################################################
# Contextual help text for commercial vehicle features
COMMERCIAL_VEHICLE_HELP = {
    "Commercial Category": "Classification based on GVWR: Light Duty (≤8,500 lbs), Medium Duty (8,501-19,500 lbs), Heavy Duty (19,501-33,000 lbs), Extra Heavy Duty (>33,000 lbs)",
    "GVWR (lbs)": "Gross Vehicle Weight Rating - maximum allowable total weight of vehicle including cargo, passengers, and fluids",
    "payload_capacity_lbs": "Maximum weight of cargo that can be carried (GVWR minus curb weight)",
    "towing_capacity_lbs": "Maximum weight the vehicle can safely tow behind it",
    "duty_cycle": "Operational classification: Urban Delivery, Regional Haul, Long Haul, Construction, Emergency Services, etc.",
    "electrification_suitability": "Assessment of how suitable this vehicle is for electrification: High, Medium, Low based on operational patterns",
    "gcwr_lbs": "Gross Combined Weight Rating - maximum total weight of vehicle plus trailer",
    "front_gawr_lbs": "Front Gross Axle Weight Rating - maximum weight on front axle",
    "rear_gawr_lbs": "Rear Gross Axle Weight Rating - maximum weight on rear axle",
    "engine_torque_lb_ft": "Engine torque output in pound-feet - indicates pulling power",
    "vocation": "Specific vehicle application or use case (delivery, construction, passenger transport, etc.)",
    "cab_configuration": "Cab style: Regular (single row), Extended (partial second row), Crew (full second row)",
    "bed_length": "Truck bed length classification: Short, Standard, Long",
    "axle_configuration": "Drive configuration: 4x2 (2WD), 4x4 (4WD), 6x4 (6-wheel with 4 driven), etc.",
    "wheelbase_inches": "Distance between front and rear axle centers - affects ride quality and cargo space",
    "overall_length_inches": "Total vehicle length from front bumper to rear",
    "cargo_length_inches": "Usable cargo area length",
    "fuel_tank_capacity_gal": "Total fuel tank capacity in gallons",
    "def_tank_capacity_gal": "Diesel Exhaust Fluid tank capacity for emissions control systems"
}

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
    "Electrification Potential",
    "Emissions Reduction",
    "ROI Analysis",
    "Charging Infrastructure",
    "Emissions Inventory",
    "Emissions Trends",
    "Emissions by Department",
    "Emissions by Vehicle Type",
    "Fleet Cash Flow",
    "Replacement Priority",
    "Scenario Comparison",
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

# Phase 9A: Enhanced TCO parameters
DEFAULT_FUEL_ESCALATION_RATE = 3.0    # % annual increase in fuel/electricity prices
DEFAULT_INCENTIVE_AMOUNT = 0.0        # $ federal/state incentive (subtracted from EV price)
DEFAULT_INFRASTRUCTURE_COST_PER_VEHICLE = 0.0  # $ charging infrastructure amortized per vehicle
DEFAULT_RESIDUAL_VALUE_ICE_PCT = 15.0  # % of purchase price at end of analysis period
DEFAULT_RESIDUAL_VALUE_EV_PCT = 20.0   # % of purchase price at end of analysis period (battery has value)

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
    "Location": "Location",
    
    # Enhanced commercial vehicle fields for Step 12
    "Commercial Category": "Commercial Category",
    "GVWR (lbs)": "GVWR (pounds)",
    "Engine HP": "Engine Horsepower",
    "Engine Type": "Engine Configuration",
    "FuelTypeSecondary": "Fuel Type (Secondary)",
    "Vehicle Class": "Vehicle Class",
    "Series": "Series",
    "Trim": "Trim Level",
    "Is Diesel": "Diesel Engine",
    "Is Commercial": "Commercial Vehicle",
    "Commercial Summary": "Commercial Summary",

    # Match quality
    "Match Confidence": "Match Confidence",
    "Fuel Type Mismatch": "Fuel Type Mismatch",

    # ACF compliance & electrification timeline
    "ACF Category": "ACF Compliance",
    "ACF Detail": "ACF Detail",
    "Proposed EV Year": "Proposed EV Year",

    # EV equivalent matching (Phase 9B)
    "EV Equivalent": "EV Equivalent",
    "EV MSRP Range": "EV MSRP Range",
    "EV EPA Range": "EV EPA Range",
    "EV Fit Score": "EV Match Score",

    # Enhanced Commercial Vehicle Specifications (from web scraping)
    "payload_capacity_lbs": "Payload Capacity (lbs)",
    "towing_capacity_lbs": "Towing Capacity (lbs)",
    "gcwr_lbs": "GCWR (lbs)",
    "front_gawr_lbs": "Front GAWR (lbs)",
    "rear_gawr_lbs": "Rear GAWR (lbs)",
    "engine_torque_lb_ft": "Engine Torque (lb-ft)",
    "engine_torque_rpm": "Torque RPM",
    "max_hp_rpm": "Max HP RPM",
    "engine_manufacturer": "Engine Manufacturer",
    "engine_model": "Engine Model",
    "fuel_tank_capacity_gal": "Fuel Tank (gal)",
    "def_tank_capacity_gal": "DEF Tank (gal)",
    "wheelbase_inches": "Wheelbase (in)",
    "overall_length_inches": "Length (in)",
    "overall_width_inches": "Width (in)",
    "overall_height_inches": "Height (in)",
    "cargo_length_inches": "Cargo Length (in)",
    "cargo_width_inches": "Cargo Width (in)",
    "cargo_height_inches": "Cargo Height (in)",
    "cab_configuration": "Cab Configuration",
    "bed_length": "Bed Length",
    "axle_configuration": "Axle Configuration",
    "axle_ratio": "Axle Ratio",
    "duty_cycle": "Duty Cycle",
    "vocation": "Vocation",
    "electrification_suitability": "EV Suitability",
    "data_source": "Data Source",
    "data_confidence": "Data Confidence"
}

# Define field categories for UI organization
FIELD_CATEGORIES = {
    "Essential": [
        "VIN", "Year", "Make", "Model", "FuelTypePrimary", "BodyClass", 
        "GVWR", "MPG City", "MPG Highway", "MPG Combined", "CO2 emissions"
    ],
    "Commercial Vehicle": [
        "Commercial Category", "GVWR (lbs)", "Engine HP", "Engine Type", 
        "Is Diesel", "Is Commercial", "Commercial Summary", "Vehicle Class",
        "payload_capacity_lbs", "towing_capacity_lbs", "duty_cycle", "vocation",
        "electrification_suitability"
    ],
    "Commercial Specifications": [
        "gcwr_lbs", "front_gawr_lbs", "rear_gawr_lbs", "engine_torque_lb_ft",
        "engine_torque_rpm", "max_hp_rpm", "engine_manufacturer", "engine_model",
        "fuel_tank_capacity_gal", "def_tank_capacity_gal"
    ],
    "Vehicle Dimensions": [
        "wheelbase_inches", "overall_length_inches", "overall_width_inches", 
        "overall_height_inches", "cargo_length_inches", "cargo_width_inches",
        "cargo_height_inches", "cab_configuration", "bed_length"
    ],
    "Configuration": [
        "axle_configuration", "axle_ratio", "data_source", "data_confidence"
    ],
    "Vehicle Details": [
        "FuelTypeSecondary", "Series", "Trim", "cylinders", "displ", "drive", "trany"
    ],
    "Alternative Fuel": [
        "fuelType2", "co2A", "rangeA", "phevBlended", "phevCity", "phevComb"
    ],
    "Fleet Management": [
        "Odometer", "Annual Mileage", "Asset ID", "Department", "Location"
    ],
    "Analysis": [
        "Data Quality", "Processing Status", "feScore", "ghgScore", "ghgScoreA"
    ],
    "Advanced": [
        "combinedCD", "startStop", "VClass", "fuelCost08", "Processing Error"
    ]
}

# Initial visible columns in the results table (optimized for commercial fleet analysis)
DEFAULT_VISIBLE_COLUMNS = [
    # Essential identification
    "VIN", "Year", "Make", "Model",

    # Key vehicle specs
    "FuelTypePrimary", "BodyClass",

    # Fuel economy — the core value of the app
    "MPG Combined", "MPG City", "MPG Highway",
    "CO2 emissions",

    # Commercial classification
    "Commercial Category", "GVWR (lbs)",

    # Fleet management
    "Asset ID", "Department",

    # ACF compliance & electrification
    "ACF Category", "ACF Detail", "Proposed EV Year",

    # EV replacement recommendations
    "EV Equivalent", "EV MSRP Range",

    # Data quality & matching
    "Match Confidence", "Data Quality", "Processing Status"
]

# Additional data column mappings for common fleet management field names
# Maps various input column names to standardized field names
ADDITIONAL_DATA_MAPPINGS = {
    # Asset/Vehicle identification
    "asset_id": ["asset_id", "asset id", "asset_number", "asset number", "asset#", "unit_id", 
                 "unit id", "unit_number", "unit number", "unit#", "vehicle_id", "vehicle id", 
                 "vehicle_number", "vehicle number", "fleet_id", "fleet id", "fleet_number", 
                 "tag_number", "tag number", "equipment_id", "equipment id"],
    
    # Mileage/Odometer
    "odometer": ["odometer", "odo", "mileage", "miles", "current_mileage", "current mileage",
                 "meter_reading", "meter reading", "total_miles", "total miles"],
    
    "annual_mileage": ["annual_mileage", "annual mileage", "yearly_mileage", "yearly mileage",
                       "miles_per_year", "miles per year", "average_annual_miles", 
                       "average annual miles", "estimated_annual_miles", "est_annual_miles"],
    
    # Department/Organization
    "department": ["department", "dept", "division", "section", "unit", "group", "team",
                   "organization", "org", "cost_center", "cost center", "business_unit",
                   "business unit", "program", "service_area", "service area"],
    
    # Location/Site
    "location": ["location", "site", "facility", "station", "base", "yard", "depot",
                 "garage", "address", "city", "region", "district", "zone", "area"],
    
    # Driver/Operator
    "driver": ["driver", "operator", "assigned_to", "assigned to", "user", "employee",
               "operator_name", "operator name", "driver_name", "driver name"],
    
    # Dates
    "acquisition_date": ["acquisition_date", "acquisition date", "purchase_date", 
                         "purchase date", "date_acquired", "date acquired", "in_service_date",
                         "in service date", "start_date", "start date"],
    
    "retire_date": ["retire_date", "retire date", "retirement_date", "retirement date",
                    "end_date", "end date", "disposal_date", "disposal date"],
    
    # Financial
    "purchase_price": ["purchase_price", "purchase price", "cost", "price", "value",
                       "acquisition_cost", "acquisition cost", "original_cost", "original cost"],
    
    "fuel_card": ["fuel_card", "fuel card", "card_number", "card number", "fuel_id", "fuel id"],
    
    # Vehicle specifications
    "license_plate": ["license_plate", "license plate", "plate", "plate_number", 
                      "plate number", "tag", "registration", "reg_number", "reg number"],
    
    "vin_last_8": ["vin_last_8", "vin last 8", "last_8", "last 8", "partial_vin", "partial vin"],
    
    # Usage/Classification  
    "vehicle_type": ["vehicle_type", "vehicle type", "type", "class", "category", 
                     "classification", "use_type", "use type"],
    
    "fuel_type": ["fuel_type", "fuel type", "fuel", "primary_fuel", "primary fuel"],
    
    # Maintenance
    "last_service": ["last_service", "last service", "last_maintenance", "last maintenance",
                     "service_date", "service date", "maintenance_date", "maintenance date"],
    
    "next_service": ["next_service", "next service", "next_maintenance", "next maintenance",
                     "due_date", "due date", "service_due", "service due"],
    
    # Custom fields
    "notes": ["notes", "comments", "remarks", "description", "memo", "additional_info",
              "additional info", "special_instructions", "special instructions"],
    
    "status": ["status", "condition", "state", "active", "inactive", "available", "unavailable"]
}

###############################################################################
# Electrification Timeline
###############################################################################
# Final year of the fleet electrification timeline (inclusive).
# For CA state & local government fleets, 2040 is a common planning horizon.
DEFAULT_ELECTRIFICATION_END_YEAR = 2040

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