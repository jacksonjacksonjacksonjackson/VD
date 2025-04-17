"""
models.py

Data models and structures for the Fleet Electrification Analyzer.
Defines dataclasses for vehicles, API responses, and analysis results.
"""

import re
import json
import datetime
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Any, Union, Set

from settings import ALL_FUEL_ECONOMY_FIELDS

###############################################################################
# Base Models
###############################################################################

@dataclass
class BaseModel:
    """Base class for data models with common utilities."""
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert model to dictionary."""
        return asdict(self)
    
    def to_json(self, indent: int = 2) -> str:
        """Convert model to JSON string."""
        return json.dumps(self.to_dict(), indent=indent)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'BaseModel':
        """Create model instance from dictionary."""
        # Filter the dictionary to include only fields defined in the class
        fields = {f.name for f in cls.__dataclass_fields__.values()}
        filtered_data = {k: v for k, v in data.items() if k in fields}
        return cls(**filtered_data)

###############################################################################
# Vehicle Models
###############################################################################

@dataclass
class VehicleIdentification(BaseModel):
    """Vehicle identification from VIN decoding."""
    
    vin: str
    year: str = ""
    make: str = ""
    model: str = ""
    fuel_type: str = ""
    body_class: str = ""
    gvwr: str = ""
    engine_displacement: str = ""
    engine_cylinders: str = ""
    drive_type: str = ""
    transmission: str = ""
    
    @property
    def valid(self) -> bool:
        """Check if the basic identification fields are populated."""
        return bool(self.vin and self.year and self.make and self.model)
    
    @property
    def display_name(self) -> str:
        """Get a human-readable display name for the vehicle."""
        return f"{self.year} {self.make} {self.model}".strip()

@dataclass
class FuelEconomyData(BaseModel):
    """Fuel economy data from FuelEconomy.gov API."""
    
    # Key efficiency metrics
    city_mpg: float = 0.0
    highway_mpg: float = 0.0
    combined_mpg: float = 0.0
    
    # CO2 emissions
    co2_primary: float = 0.0  # g/mile
    co2_alt: float = 0.0  # g/mile for alternative fuel
    
    # Alternative fuel data
    alt_fuel_type: str = ""
    alt_range: float = 0.0  # miles
    
    # Other useful fields
    fuel_cost_primary: float = 0.0  # $ annual
    fuel_cost_alt: float = 0.0  # $ annual
    
    # Raw data from API
    raw_data: Dict[str, Any] = field(default_factory=dict)
    
    def __post_init__(self):
        """Process raw_data if provided but other fields are empty."""
        if self.raw_data and not self.combined_mpg:
            self._extract_from_raw()
    
    def _extract_from_raw(self):
        """Extract structured fields from raw API data."""
        # MPG data
        self.city_mpg = float(self.raw_data.get('city08', 0) or 0)
        self.highway_mpg = float(self.raw_data.get('highway08', 0) or 0)
        self.combined_mpg = float(self.raw_data.get('comb08', 0) or 0)
        
        # CO2 data
        self.co2_primary = float(self.raw_data.get('co2TailpipeGpm', 0) or 0)
        self.co2_alt = float(self.raw_data.get('co2TailpipeAGpm', 0) or 0)
        
        # Alternative fuel
        self.alt_fuel_type = self.raw_data.get('fuelType2', '')
        self.alt_range = float(self.raw_data.get('rangeA', 0) or 0)
        
        # Fuel costs
        self.fuel_cost_primary = float(self.raw_data.get('fuelCost08', 0) or 0)
        self.fuel_cost_alt = float(self.raw_data.get('fuelCostA08', 0) or 0)

@dataclass
class FleetVehicle(BaseModel):
    """
    Comprehensive vehicle model combining identification, fuel economy,
    and fleet-specific data.
    """
    
    # Vehicle identification
    vin: str
    vehicle_id: VehicleIdentification = field(default_factory=VehicleIdentification)
    
    # Fuel economy data
    fuel_economy: FuelEconomyData = field(default_factory=FuelEconomyData)
    
    # Fleet management data
    asset_id: str = ""
    department: str = ""
    location: str = ""
    odometer: float = 0.0
    annual_mileage: float = 0.0
    acquisition_date: Optional[datetime.date] = None
    retire_date: Optional[datetime.date] = None
    
    # Custom fields for user data
    custom_fields: Dict[str, Any] = field(default_factory=dict)
    
    # Analysis and matching metadata
    match_confidence: float = 0.0
    assumed_vehicle_id: str = ""
    assumed_vehicle_text: str = ""
    
    @property
    def age(self) -> float:
        """Calculate vehicle age in years based on model year."""
        if not self.vehicle_id.year:
            return 0.0
        
        try:
            year = int(self.vehicle_id.year)
            current_year = datetime.datetime.now().year
            return current_year - year
        except ValueError:
            return 0.0
    
    @property
    def display_name(self) -> str:
        """Get a human-readable display name for the vehicle."""
        return self.vehicle_id.display_name
    
    def get_field(self, field_name: str) -> Any:
        """
        Get a field value by name, checking all nested objects.
        
        Args:
            field_name: Name of the field to retrieve
            
        Returns:
            Field value or empty string if not found
        """
        # Check direct attributes first
        if hasattr(self, field_name):
            return getattr(self, field_name)
        
        # Check vehicle_id
        if hasattr(self.vehicle_id, field_name):
            return getattr(self.vehicle_id, field_name)
        
        # Check fuel_economy
        if hasattr(self.fuel_economy, field_name):
            return getattr(self.fuel_economy, field_name)
        
        # Check fuel_economy.raw_data
        if field_name in self.fuel_economy.raw_data:
            return self.fuel_economy.raw_data[field_name]
        
        # Check custom fields
        if field_name in self.custom_fields:
            return self.custom_fields[field_name]
        
        # Not found
        return ""
    
    def set_field(self, field_name: str, value: Any) -> bool:
        """
        Set a field value by name, determining the appropriate object.
        
        Args:
            field_name: Name of the field to set
            value: Value to set
            
        Returns:
            True if field was set, False if not found
        """
        # Check direct attributes
        if hasattr(self, field_name):
            setattr(self, field_name, value)
            return True
        
        # Check vehicle_id
        if hasattr(self.vehicle_id, field_name):
            setattr(self.vehicle_id, field_name, value)
            return True
        
        # Check fuel_economy
        if hasattr(self.fuel_economy, field_name):
            setattr(self.fuel_economy, field_name, value)
            return True
        
        # Check fuel_economy.raw_data for FuelEconomy.gov fields
        if field_name in ALL_FUEL_ECONOMY_FIELDS:
            self.fuel_economy.raw_data[field_name] = value
            return True
        
        # Use custom fields as fallback
        self.custom_fields[field_name] = value
        return True
    
    def to_row_dict(self) -> Dict[str, Any]:
        """
        Convert to a flattened dictionary for table display or export.
        
        Returns:
            Dictionary with all fields flattened to a single level
        """
        result = {
            "VIN": self.vin,
            "Year": self.vehicle_id.year,
            "Make": self.vehicle_id.make,
            "Model": self.vehicle_id.model,
            "FuelTypePrimary": self.vehicle_id.fuel_type,
            "BodyClass": self.vehicle_id.body_class,
            "GVWR": self.vehicle_id.gvwr,
            "MPG City": str(self.fuel_economy.city_mpg),
            "MPG Highway": str(self.fuel_economy.highway_mpg),
            "MPG Combined": str(self.fuel_economy.combined_mpg),
            "CO2 emissions": str(self.fuel_economy.co2_primary),
            "co2A": str(self.fuel_economy.co2_alt),
            "rangeA": str(self.fuel_economy.alt_range),
            "Odometer": str(self.odometer),
            "Annual Mileage": str(self.annual_mileage),
            "Asset ID": self.asset_id,
            "Department": self.department,
            "Location": self.location,
            "Assumed Vehicle (Text)": self.assumed_vehicle_text,
            "Assumed Vehicle (ID)": self.assumed_vehicle_id
        }
        
        # Add all raw data fields from fuel economy
        for key, value in self.fuel_economy.raw_data.items():
            if key not in result:
                result[key] = str(value)
        
        # Add custom fields
        for key, value in self.custom_fields.items():
            if key not in result:
                result[key] = str(value)
        
        return result

###############################################################################
# Fleet Models
###############################################################################

@dataclass
class Fleet(BaseModel):
    """Collection of vehicles with fleet-level metadata and analysis."""
    
    name: str
    vehicles: List[FleetVehicle] = field(default_factory=list)
    creation_date: datetime.datetime = field(default_factory=datetime.datetime.now)
    last_modified: datetime.datetime = field(default_factory=datetime.datetime.now)
    notes: str = ""
    
    @property
    def size(self) -> int:
        """Get the number of vehicles in the fleet."""
        return len(self.vehicles)
    
    @property
    def makes(self) -> Dict[str, int]:
        """Get count of vehicles by make."""
        makes = {}
        for vehicle in self.vehicles:
            make = vehicle.vehicle_id.make
            if make:
                makes[make] = makes.get(make, 0) + 1
        return makes
    
    @property
    def models(self) -> Dict[str, int]:
        """Get count of vehicles by model."""
        models = {}
        for vehicle in self.vehicles:
            model = vehicle.vehicle_id.model
            if model:
                models[model] = models.get(model, 0) + 1
        return models
    
    @property
    def fuel_types(self) -> Dict[str, int]:
        """Get count of vehicles by fuel type."""
        fuel_types = {}
        for vehicle in self.vehicles:
            fuel_type = vehicle.vehicle_id.fuel_type
            if fuel_type:
                fuel_types[fuel_type] = fuel_types.get(fuel_type, 0) + 1
        return fuel_types
    
    @property
    def body_classes(self) -> Dict[str, int]:
        """Get count of vehicles by body class."""
        body_classes = {}
        for vehicle in self.vehicles:
            body_class = vehicle.vehicle_id.body_class
            if body_class:
                body_classes[body_class] = body_classes.get(body_class, 0) + 1
        return body_classes
    
    @property
    def departments(self) -> Dict[str, int]:
        """Get count of vehicles by department."""
        departments = {}
        for vehicle in self.vehicles:
            dept = vehicle.department
            if dept:
                departments[dept] = departments.get(dept, 0) + 1
        return departments
    
    @property
    def avg_mpg(self) -> float:
        """Calculate average combined MPG for the fleet."""
        mpg_values = [v.fuel_economy.combined_mpg for v in self.vehicles 
                      if v.fuel_economy.combined_mpg > 0]
        if not mpg_values:
            return 0.0
        return sum(mpg_values) / len(mpg_values)
    
    @property
    def avg_co2(self) -> float:
        """Calculate average CO2 emissions (g/mile) for the fleet."""
        co2_values = [v.fuel_economy.co2_primary for v in self.vehicles 
                      if v.fuel_economy.co2_primary > 0]
        if not co2_values:
            return 0.0
        return sum(co2_values) / len(co2_values)
    
    @property
    def avg_age(self) -> float:
        """Calculate average age of vehicles in the fleet."""
        ages = [v.age for v in self.vehicles if v.age > 0]
        if not ages:
            return 0.0
        return sum(ages) / len(ages)
    
    @property
    def total_annual_mileage(self) -> float:
        """Calculate total annual mileage for the fleet."""
        return sum(v.annual_mileage for v in self.vehicles if v.annual_mileage > 0)
    
    def add_vehicle(self, vehicle: FleetVehicle) -> None:
        """Add a vehicle to the fleet."""
        self.vehicles.append(vehicle)
        self.last_modified = datetime.datetime.now()
    
    def remove_vehicle(self, vin: str) -> bool:
        """
        Remove a vehicle from the fleet by VIN.
        
        Args:
            vin: VIN of vehicle to remove
            
        Returns:
            True if vehicle was found and removed
        """
        for i, vehicle in enumerate(self.vehicles):
            if vehicle.vin == vin:
                self.vehicles.pop(i)
                self.last_modified = datetime.datetime.now()
                return True
        return False
    
    def get_vehicle(self, vin: str) -> Optional[FleetVehicle]:
        """
        Get a vehicle by VIN.
        
        Args:
            vin: VIN to search for
            
        Returns:
            FleetVehicle or None if not found
        """
        for vehicle in self.vehicles:
            if vehicle.vin == vin:
                return vehicle
        return None
    
    def filter_vehicles(self, **criteria) -> List[FleetVehicle]:
        """
        Filter vehicles based on field values.
        
        Args:
            **criteria: Field name/value pairs to match
            
        Returns:
            List of matching vehicles
        """
        results = []
        
        for vehicle in self.vehicles:
            match = True
            
            for field, value in criteria.items():
                vehicle_value = vehicle.get_field(field)
                
                # Handle string matching (case-insensitive)
                if isinstance(value, str) and isinstance(vehicle_value, str):
                    if value.lower() not in vehicle_value.lower():
                        match = False
                        break
                
                # Handle numeric ranges
                elif isinstance(value, tuple) and len(value) == 2:
                    min_val, max_val = value
                    try:
                        num_value = float(vehicle_value)
                        if not (min_val <= num_value <= max_val):
                            match = False
                            break
                    except (ValueError, TypeError):
                        match = False
                        break
                
                # Handle exact matches
                elif vehicle_value != value:
                    match = False
                    break
            
            if match:
                results.append(vehicle)
        
        return results

###############################################################################
# Analysis Models
###############################################################################

@dataclass
class ElectrificationAnalysis(BaseModel):
    """Analysis of fleet electrification potential and benefits."""
    
    # Fleet reference
    fleet_name: str
    
    # Analysis parameters
    gas_price: float = 3.50  # $/gallon
    electricity_price: float = 0.13  # $/kWh
    ev_efficiency: float = 0.30  # kWh/mile
    analysis_period: int = 10  # years
    discount_rate: float = 5.0  # %
    
    # Results
    co2_savings: float = 0.0  # tons
    fuel_cost_savings: float = 0.0  # $
    maintenance_savings: float = 0.0  # $
    total_savings: float = 0.0  # $
    payback_period: float = 0.0  # years
    
    # Vehicle specific results
    vehicle_results: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    
    # Prioritized electrification list
    prioritized_vehicles: List[str] = field(default_factory=list)

@dataclass
class ChargingAnalysis(BaseModel):
    """Analysis of charging infrastructure needs."""
    
    # Fleet reference
    fleet_name: str
    
    # Analysis parameters
    daily_usage_pattern: str = "standard"  # standard, extended, 24-hour
    charging_window: tuple = (18, 6)  # hours (start, end)
    peak_charging_time: int = 4  # hours
    
    # Results
    level2_chargers_needed: int = 0
    dcfc_chargers_needed: int = 0
    max_power_required: float = 0.0  # kW
    recommended_layout: Dict[str, Any] = field(default_factory=dict)
    estimated_installation_cost: float = 0.0  # $

@dataclass
class EmissionsInventory(BaseModel):
    """Greenhouse gas emissions inventory for a fleet."""
    
    # Fleet reference
    fleet_name: str
    
    # Inventory year
    inventory_year: int = field(default_factory=lambda: datetime.datetime.now().year)
    
    # Emissions by category (metric tons CO2e)
    total_emissions: float = 0.0
    by_department: Dict[str, float] = field(default_factory=dict)
    by_vehicle_type: Dict[str, float] = field(default_factory=dict)
    by_fuel_type: Dict[str, float] = field(default_factory=dict)
    
    # Historical data
    historical_data: Dict[int, float] = field(default_factory=dict)
    
    # Reduction targets
    baseline_year: int = 0
    target_year: int = 0
    reduction_target: float = 0.0  # %
    projected_emissions: Dict[int, float] = field(default_factory=dict)

###############################################################################
# API Response Models
###############################################################################

@dataclass
class VinDecoderResponse(BaseModel):
    """Response from VIN decoder API."""
    
    success: bool = False
    error_message: str = ""
    vin: str = ""
    data: Dict[str, Any] = field(default_factory=dict)
    
    def to_vehicle_id(self) -> VehicleIdentification:
        """Convert API response to VehicleIdentification object."""
        if not self.success or not self.data:
            return VehicleIdentification(vin=self.vin)
        
        return VehicleIdentification(
            vin=self.vin,
            year=self.data.get("ModelYear", ""),
            make=self.data.get("Make", ""),
            model=self.data.get("Model", ""),
            fuel_type=self.data.get("FuelTypePrimary", ""),
            body_class=self.data.get("BodyClass", ""),
            gvwr=self.data.get("GVWR", ""),
            engine_displacement=self.data.get("DisplacementL", ""),
            engine_cylinders=self.data.get("EngineCylinders", ""),
            drive_type=self.data.get("DriveType", ""),
            transmission=self.data.get("TransmissionStyle", "")
        )

@dataclass
class FuelEconomyResponse(BaseModel):
    """Response from FuelEconomy.gov API."""
    
    success: bool = False
    error_message: str = ""
    vehicle_id: str = ""
    data: Dict[str, Any] = field(default_factory=dict)
    
    def to_fuel_economy(self) -> FuelEconomyData:
        """Convert API response to FuelEconomyData object."""
        if not self.success or not self.data:
            return FuelEconomyData()
        
        return FuelEconomyData(raw_data=self.data)