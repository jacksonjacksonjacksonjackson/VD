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
import logging

from settings import ALL_FUEL_ECONOMY_FIELDS

# Import CommercialVehicleSpecs after models are defined to avoid circular imports
# We'll use TYPE_CHECKING to avoid runtime import issues
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from commercial_vehicle_scraper import CommercialVehicleSpecs

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
    
    # Enhanced commercial vehicle fields for Step 12
    gvwr_raw: str = ""  # Raw GVWR value from NHTSA
    gvwr_pounds: float = 0.0  # Parsed GVWR in pounds
    commercial_category: str = ""  # Light/Medium/Heavy Duty
    engine_power_hp: str = ""  # Horsepower
    engine_power_kw: str = ""  # Kilowatts
    engine_type: str = ""  # Engine type/configuration
    fuel_type_secondary: str = ""  # Secondary fuel type
    vehicle_class: str = ""  # NHTSA vehicle class
    plant_country: str = ""  # Manufacturing country
    series: str = ""  # Vehicle series
    trim: str = ""  # Trim level
    is_diesel: bool = False  # Diesel detection flag
    is_commercial: bool = False  # Commercial vehicle flag
    
    def __post_init__(self):
        """Process data after initialization."""
        self._process_gvwr()
        self._classify_commercial()
        self._detect_diesel()
    
    def _process_gvwr(self):
        """Extract numeric GVWR value and classify commercial category."""
        if not self.gvwr and not self.gvwr_raw:
            return
            
        # Use gvwr_raw if available, otherwise use gvwr
        gvwr_text = self.gvwr_raw or self.gvwr
        
        # Extract numeric value from GVWR string
        import re
        # Look for patterns like "Class 1: 6,000 lb (2,722 kg)" or "6000" or "6,000 lb"
        patterns = [
            r'(\d{1,2},?\d{3})\s*(?:lb|lbs|pounds?)',  # "6,000 lb" or "6000 lb"
            r'(\d{1,2},?\d{3})\s*(?:\(|$)',  # "6,000 (" or "6000" at end
            r':\s*(\d{1,2},?\d{3})',  # ": 6,000"
            r'(\d{4,6})'  # Raw numbers like "6000"
        ]
        
        for pattern in patterns:
            match = re.search(pattern, gvwr_text.replace(',', ''))
            if match:
                try:
                    self.gvwr_pounds = float(match.group(1).replace(',', ''))
                    break
                except ValueError:
                    continue
        
        # Classify commercial category based on GVWR
        if self.gvwr_pounds > 0:
            if self.gvwr_pounds <= 8500:
                self.commercial_category = "Light Duty"
            elif self.gvwr_pounds <= 19500:
                self.commercial_category = "Medium Duty" 
            elif self.gvwr_pounds <= 33000:
                self.commercial_category = "Heavy Duty"
            else:
                self.commercial_category = "Extra Heavy Duty"
    
    def _classify_commercial(self):
        """Determine if this is a commercial vehicle."""
        # Commercial indicators
        commercial_body_classes = [
            "truck", "van", "bus", "chassis cab", "cutaway", "pickup",
            "cargo van", "passenger van", "step van", "box truck"
        ]
        
        commercial_models = [
            "transit", "e-series", "express", "savana", "sprinter",
            "promaster", "f-150", "f-250", "f-350", "f-450", "f-550",
            "silverado", "sierra", "ram 1500", "ram 2500", "ram 3500"
        ]
        
        # Check body class
        body_lower = self.body_class.lower()
        for indicator in commercial_body_classes:
            if indicator in body_lower:
                self.is_commercial = True
                return
        
        # Check model name
        model_lower = self.model.lower()
        for indicator in commercial_models:
            if indicator in model_lower:
                self.is_commercial = True
                return
        
        # Check GVWR (vehicles over 8,500 lbs are typically commercial)
        if self.gvwr_pounds > 8500:
            self.is_commercial = True
    
    def _detect_diesel(self):
        """Detect if vehicle uses diesel fuel."""
        diesel_indicators = [
            "diesel", "biodiesel", "b20", "b100", 
            "compression ignition", "ci"
        ]
        
        # Check primary fuel type
        fuel_lower = self.fuel_type.lower()
        for indicator in diesel_indicators:
            if indicator in fuel_lower:
                self.is_diesel = True
                return
        
        # Check secondary fuel type
        fuel2_lower = self.fuel_type_secondary.lower()
        for indicator in diesel_indicators:
            if indicator in fuel2_lower:
                self.is_diesel = True
                return
        
        # Check engine type
        engine_lower = self.engine_type.lower()
        for indicator in diesel_indicators:
            if indicator in engine_lower:
                self.is_diesel = True
                return
    
    @property
    def valid(self) -> bool:
        """Check if the basic identification fields are populated."""
        return bool(self.vin and self.year and self.make and self.model)
    
    @property
    def display_name(self) -> str:
        """Get a human-readable display name for the vehicle."""
        return f"{self.year} {self.make} {self.model}".strip()
    
    @property
    def commercial_summary(self) -> str:
        """Get a summary of commercial vehicle characteristics."""
        parts = []
        
        if self.commercial_category:
            parts.append(self.commercial_category)
        
        if self.is_diesel:
            parts.append("Diesel")
        
        if self.gvwr_pounds > 0:
            parts.append(f"GVWR: {self.gvwr_pounds:,.0f} lb")
        
        return " | ".join(parts) if parts else "Passenger Vehicle"

@dataclass
class FuelEconomyData(BaseModel):
    """Fuel economy data from FuelEconomy.gov API."""

    # Key efficiency metrics
    city_mpg: float = 0.0
    highway_mpg: float = 0.0
    combined_mpg: float = 0.0

    # MPG data provenance — who provided this MPG value.
    # Values: "" (unknown/API), "FuelEconomy.gov", "Fuelly (community)",
    #         "EPA SmartWay", "Ford Commercial", "EPA Class Estimate", "Manual Override"
    mpg_source: str = ""

    # Whether the MPG value is an estimate rather than a measured/published figure.
    # True for EPA class-average fallbacks.  False for API or scraper data.
    mpg_is_estimate: bool = False

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

        # Tag source when MPG comes from FuelEconomy.gov API
        if self.combined_mpg > 0 and not self.mpg_source:
            self.mpg_source = "FuelEconomy.gov"
            self.mpg_is_estimate = False
        
        # CO2 data
        self.co2_primary = float(self.raw_data.get('co2TailpipeGpm', 0) or 0)
        self.co2_alt = float(self.raw_data.get('co2TailpipeAGpm', 0) or 0)
        
        # Alternative fuel
        self.alt_fuel_type = self.raw_data.get('fuelType2', '')
        
        # Handle alternative range with safe conversion for formats like "361/479"
        range_value = self.raw_data.get('rangeA', 0) or 0
        try:
            # If it's already a number, use it directly
            if isinstance(range_value, (int, float)):
                self.alt_range = float(range_value)
            else:
                # Handle string formats like "361/479" - extract first numeric value
                range_str = str(range_value).strip()
                if range_str:
                    # Look for numeric values like "361" from "361/479"
                    match = re.search(r'(\d+\.?\d*)', range_str)
                    if match:
                        self.alt_range = float(match.group(1))
                    else:
                        self.alt_range = 0.0
                else:
                    self.alt_range = 0.0
        except (ValueError, TypeError, AttributeError):
            # If conversion fails, default to 0
            self.alt_range = 0.0
        
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
    vehicle_id: VehicleIdentification = field(default_factory=lambda: VehicleIdentification(vin=""))
    
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
    
    # Input order tracking for preserving CSV order in results
    input_order_index: int = 0
    
    # Processing error tracking for debugging failed VINs
    processing_error: str = ""
    processing_success: bool = True
    data_quality_score: float = 0.0
    
    # Enhanced quality tracking for Step 13
    quality_breakdown: Dict[str, float] = field(default_factory=dict)
    consistency_score: float = 0.0
    confidence_factors: Dict[str, Any] = field(default_factory=dict)
    last_quality_check: Optional[datetime.datetime] = None
    
    # Commercial vehicle specifications from web scraping
    commercial_specs: Optional['CommercialVehicleSpecs'] = None
    
    def calculate_detailed_quality(self) -> Dict[str, Any]:
        """
        Calculate detailed quality metrics for this vehicle.
        
        Returns:
            Dictionary with comprehensive quality analysis
        """
        from datetime import datetime
        
        # Update last quality check timestamp
        self.last_quality_check = datetime.now()
        
        # Core data completeness (0-35 points)
        core_score = 0.0
        if self.vehicle_id.year: core_score += 8
        if self.vehicle_id.make: core_score += 8  
        if self.vehicle_id.model: core_score += 8
        if self.vehicle_id.fuel_type: core_score += 6
        if self.vehicle_id.body_class: core_score += 5
        
        # Fuel economy data (0-25 points)
        fuel_score = 0.0
        if self.fuel_economy.combined_mpg > 0: fuel_score += 12
        if self.fuel_economy.city_mpg > 0: fuel_score += 6
        if self.fuel_economy.highway_mpg > 0: fuel_score += 6
        if self.fuel_economy.co2_primary > 0: fuel_score += 1
        
        # Commercial vehicle data (0-15 points)
        commercial_score = 0.0
        if self.vehicle_id.gvwr_pounds > 0: commercial_score += 4
        if self.vehicle_id.commercial_category: commercial_score += 3
        if self.vehicle_id.engine_power_hp: commercial_score += 2
        if self.vehicle_id.engine_type: commercial_score += 2
        if self.vehicle_id.vehicle_class: commercial_score += 2
        if self.vehicle_id.series or self.vehicle_id.trim: commercial_score += 2
        
        # Technical details (0-10 points)
        technical_score = 0.0
        if self.vehicle_id.engine_displacement: technical_score += 3
        if self.vehicle_id.engine_cylinders: technical_score += 2
        if self.vehicle_id.transmission: technical_score += 2
        if self.vehicle_id.drive_type: technical_score += 2
        if self.vehicle_id.fuel_type_secondary: technical_score += 1
        
        # Match confidence (0-10 points)
        confidence_score = self.match_confidence / 10.0
        
        # Calculate consistency bonus (0-5 points)
        consistency_bonus = self._calculate_data_consistency()
        self.consistency_score = consistency_bonus
        
        # Update quality breakdown
        self.quality_breakdown = {
            "core_data": core_score,
            "fuel_economy": fuel_score,
            "commercial_data": commercial_score,
            "technical_details": technical_score,
            "match_confidence": confidence_score,
            "consistency_bonus": consistency_bonus
        }
        
        # Calculate total score
        total_score = sum(self.quality_breakdown.values())
        self.data_quality_score = min(total_score, 100.0)
        
        # Update confidence factors
        self.confidence_factors = {
            "vin_validation": self._assess_vin_confidence(),
            "api_matching": self._assess_api_confidence(),
            "data_completeness": self._assess_completeness_confidence(),
            "commercial_classification": self._assess_commercial_confidence()
        }
        
        return {
            "total_score": self.data_quality_score,
            "breakdown": self.quality_breakdown,
            "consistency_score": self.consistency_score,
            "confidence_factors": self.confidence_factors,
            "assessment_time": self.last_quality_check.isoformat() if self.last_quality_check else None
        }
    
    def _calculate_data_consistency(self) -> float:
        """Calculate consistency bonus based on data relationships."""
        bonus = 0.0
        
        # VIN year consistency
        if self.vehicle_id.year and len(self.vin) == 17:
            try:
                year_int = int(self.vehicle_id.year)
                if year_int >= 2010:
                    bonus += 1.0
                elif year_int >= 1980:
                    bonus += 0.5
            except ValueError:
                pass
        
        # Commercial/diesel consistency
        if self.vehicle_id.is_commercial and self.vehicle_id.is_diesel:
            bonus += 1.0
        elif not self.vehicle_id.is_commercial and not self.vehicle_id.is_diesel:
            bonus += 0.5
        
        # GVWR/body class consistency
        if self.vehicle_id.gvwr_pounds > 0 and self.vehicle_id.body_class:
            body_lower = self.vehicle_id.body_class.lower()
            if (self.vehicle_id.gvwr_pounds <= 8500 and 
                any(term in body_lower for term in ['sedan', 'coupe', 'hatchback', 'suv', 'wagon'])):
                bonus += 1.0
            elif (self.vehicle_id.gvwr_pounds > 19500 and 
                  any(term in body_lower for term in ['truck', 'bus', 'commercial', 'chassis'])):
                bonus += 1.0
        
        # MPG/weight consistency
        if (self.fuel_economy.combined_mpg > 0 and self.vehicle_id.gvwr_pounds > 0):
            if (self.vehicle_id.gvwr_pounds > 8500 and self.fuel_economy.combined_mpg < 25) or \
               (self.vehicle_id.gvwr_pounds <= 6000 and self.fuel_economy.combined_mpg > 20):
                bonus += 0.5
        
        return min(bonus, 5.0)
    
    def _assess_vin_confidence(self) -> Dict[str, Any]:
        """Assess confidence in VIN validation."""
        return {
            "valid_format": len(self.vin) == 17 and self.vin.isalnum(),
            "year_consistency": bool(self.vehicle_id.year),
            "checksum_valid": True  # Simplified - real VIN checksum validation is complex
        }
    
    def _assess_api_confidence(self) -> Dict[str, Any]:
        """Assess confidence in API data matching."""
        return {
            "match_confidence": self.match_confidence,
            "confidence_level": "High" if self.match_confidence >= 80 else 
                              "Medium" if self.match_confidence >= 60 else "Low",
            "assumed_match": bool(self.assumed_vehicle_text),
            "nhtsa_success": self.processing_success
        }
    
    def _assess_completeness_confidence(self) -> Dict[str, Any]:
        """Assess confidence based on data completeness."""
        total_fields = 20  # Core fields we expect
        populated_fields = sum([
            bool(self.vehicle_id.year),
            bool(self.vehicle_id.make),
            bool(self.vehicle_id.model),
            bool(self.vehicle_id.fuel_type),
            bool(self.vehicle_id.body_class),
            bool(self.vehicle_id.engine_displacement),
            bool(self.vehicle_id.transmission),
            bool(self.fuel_economy.combined_mpg > 0),
            bool(self.fuel_economy.city_mpg > 0),
            bool(self.fuel_economy.highway_mpg > 0),
            bool(self.fuel_economy.co2_primary > 0),
            bool(self.vehicle_id.gvwr_pounds > 0),
            bool(self.vehicle_id.commercial_category),
            bool(self.vehicle_id.engine_power_hp),
            bool(self.vehicle_id.engine_type),
            bool(self.vehicle_id.vehicle_class),
            bool(self.vehicle_id.drive_type),
            bool(self.vehicle_id.engine_cylinders),
            bool(self.vehicle_id.fuel_type_secondary),
            bool(self.vehicle_id.series or self.vehicle_id.trim)
        ])
        
        completeness_percent = (populated_fields / total_fields) * 100
        
        return {
            "populated_fields": populated_fields,
            "total_fields": total_fields,
            "completeness_percent": completeness_percent,
            "completeness_level": "High" if completeness_percent >= 75 else
                                 "Medium" if completeness_percent >= 50 else "Low"
        }
    
    def _assess_commercial_confidence(self) -> Dict[str, Any]:
        """Assess confidence in commercial vehicle classification."""
        commercial_indicators = 0
        
        # Count classification indicators
        if self.vehicle_id.gvwr_pounds > 8500: commercial_indicators += 1
        if "truck" in self.vehicle_id.body_class.lower(): commercial_indicators += 1
        if "van" in self.vehicle_id.body_class.lower(): commercial_indicators += 1
        if any(term in self.vehicle_id.model.lower() for term in ['f-150', 'f150', 'transit', 'silverado']): 
            commercial_indicators += 1
        if self.vehicle_id.is_diesel: commercial_indicators += 1
        
        return {
            "classification_indicators": commercial_indicators,
            "is_commercial": self.vehicle_id.is_commercial,
            "commercial_category": self.vehicle_id.commercial_category,
            "confidence_level": "High" if commercial_indicators >= 3 else
                              "Medium" if commercial_indicators >= 2 else "Low"
        }
    
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
    
    @staticmethod
    def _fmt_mpg(value: float) -> str:
        """Format MPG value for display."""
        if value <= 0:
            return ""
        return f"{value:.1f}" if value != int(value) else str(int(value))

    @staticmethod
    def _fmt_number(value: float, decimals: int = 0) -> str:
        """Format a number with thousand separators."""
        if value <= 0:
            return ""
        if decimals == 0:
            return f"{int(value):,}"
        return f"{value:,.{decimals}f}"

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
            "MPG City": self._fmt_mpg(self.fuel_economy.city_mpg),
            "MPG Highway": self._fmt_mpg(self.fuel_economy.highway_mpg),
            "MPG Combined": self._fmt_mpg(self.fuel_economy.combined_mpg),
            "MPG Source": self.fuel_economy.mpg_source,
            "MPG Estimated": "Yes" if self.fuel_economy.mpg_is_estimate else ("" if not self.fuel_economy.mpg_source else "No"),
            "CO2 emissions": self._fmt_number(self.fuel_economy.co2_primary, 1) if self.fuel_economy.co2_primary > 0 else "",
            "co2A": self._fmt_number(self.fuel_economy.co2_alt, 1) if self.fuel_economy.co2_alt > 0 else "",
            "rangeA": self._fmt_number(self.fuel_economy.alt_range) if self.fuel_economy.alt_range > 0 else "",
            "Odometer": self._fmt_number(self.odometer),
            "Annual Mileage": self._fmt_number(self.annual_mileage),
            "Asset ID": self.asset_id,
            "Department": self.department,
            "Location": self.location,
            "Assumed Vehicle (Text)": self.assumed_vehicle_text,
            "Assumed Vehicle (ID)": self.assumed_vehicle_id,

            # Quality, matching, and status
            "Match Confidence": f"{self.match_confidence:.0f}%" if self.match_confidence > 0 else "",
            "Data Quality": f"{self.data_quality_score:.0f}%",
            "Processing Status": "Success" if self.processing_success else "Failed",
            "Processing Error": self.processing_error,

            # Commercial vehicle fields
            "Commercial Category": self.vehicle_id.commercial_category,
            "GVWR (lbs)": self._fmt_number(self.vehicle_id.gvwr_pounds),
            "Engine HP": self.vehicle_id.engine_power_hp,
            "Engine Type": self.vehicle_id.engine_type,
            "FuelTypeSecondary": self.vehicle_id.fuel_type_secondary,
            "Vehicle Class": self.vehicle_id.vehicle_class,
            "Series": self.vehicle_id.series,
            "Trim": self.vehicle_id.trim,
            "Is Diesel": "Yes" if self.vehicle_id.is_diesel else "No",
            "Is Commercial": "Yes" if self.vehicle_id.is_commercial else "No",
            "Commercial Summary": self.vehicle_id.commercial_summary,

            # ACF compliance & electrification timeline (populated by processor)
            "ACF Category": self.custom_fields.get("ACF Category", ""),
            "ACF Detail": self.custom_fields.get("ACF Detail", ""),
            "ACF Relevance": self.custom_fields.get("ACF Relevance", ""),
            "Proposed EV Year": self.custom_fields.get("Proposed EV Year", ""),
            "EV Year Reason": self.custom_fields.get("EV Year Reason", ""),

            # Fuel type mismatch indicator (diesel VIN matched to gasoline MPG)
            "Fuel Type Mismatch": self.custom_fields.get("Fuel Type Mismatch", ""),
        }
        
        # Add commercial vehicle specifications from web scraping
        if self.commercial_specs:
            commercial_data = {
                "payload_capacity_lbs": str(self.commercial_specs.payload_capacity_lbs or ""),
                "towing_capacity_lbs": str(self.commercial_specs.towing_capacity_lbs or ""),
                "gcwr_lbs": str(self.commercial_specs.gcwr_lbs or ""),
                "front_gawr_lbs": str(self.commercial_specs.front_gawr_lbs or ""),
                "rear_gawr_lbs": str(self.commercial_specs.rear_gawr_lbs or ""),
                "engine_torque_lb_ft": str(self.commercial_specs.engine_torque_lb_ft or ""),
                "engine_torque_rpm": str(self.commercial_specs.engine_torque_rpm or ""),
                "max_hp_rpm": str(self.commercial_specs.max_hp_rpm or ""),
                "engine_manufacturer": self.commercial_specs.engine_manufacturer,
                "engine_model": self.commercial_specs.engine_model,
                "fuel_tank_capacity_gal": str(self.commercial_specs.fuel_tank_capacity_gal or ""),
                "def_tank_capacity_gal": str(self.commercial_specs.def_tank_capacity_gal or ""),
                "wheelbase_inches": str(self.commercial_specs.wheelbase_inches or ""),
                "overall_length_inches": str(self.commercial_specs.overall_length_inches or ""),
                "overall_width_inches": str(self.commercial_specs.overall_width_inches or ""),
                "overall_height_inches": str(self.commercial_specs.overall_height_inches or ""),
                "cargo_length_inches": str(self.commercial_specs.cargo_length_inches or ""),
                "cargo_width_inches": str(self.commercial_specs.cargo_width_inches or ""),
                "cargo_height_inches": str(self.commercial_specs.cargo_height_inches or ""),
                "cab_configuration": self.commercial_specs.cab_configuration,
                "bed_length": self.commercial_specs.bed_length,
                "axle_configuration": self.commercial_specs.axle_configuration,
                "axle_ratio": self.commercial_specs.axle_ratio,
                "duty_cycle": self.commercial_specs.duty_cycle,
                "vocation": self.commercial_specs.vocation,
                "electrification_suitability": self.commercial_specs.electrification_suitability,
                "data_source": self.commercial_specs.data_source,
                "data_confidence": f"{self.commercial_specs.data_confidence:.0%}" if self.commercial_specs.data_confidence else "",
            }
            result.update(commercial_data)
        
        # Add all raw data fields from fuel economy
        for key, value in self.fuel_economy.raw_data.items():
            if key not in result:
                result[key] = str(value)
        
        # Add custom fields (exclude internal implementation details)
        for key, value in self.custom_fields.items():
            if key not in result and not key.startswith("_"):
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
    fleet_type: str = "hpf"   # "hpf", "non_hpf", or "state_agency" — controls ACF deadlines
    max_vehicles_per_year: int = 0  # 0 = no cap (auto even-spread); >0 = greedy-fill cap
    
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

    # Fleet-wide aggregate year-by-year cash flows (Phase 9A)
    fleet_cash_flows: List[Dict[str, Any]] = field(default_factory=list)

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

    # Extended results (Phase 28)
    hourly_load_kw: List[float] = field(default_factory=lambda: [0.0] * 24)
    facility_breakdown: List[Dict[str, Any]] = field(default_factory=list)
    daily_energy_kwh: float = 0.0
    charging_hours: float = 0.0

    # Utility rate fields (populated when rate section is applied)
    state_code: str = ""
    electricity_price_per_kwh: float = 0.0
    demand_charge_per_kw_month: float = 0.0
    estimated_annual_charging_cost: float = 0.0

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

    # Whether historical/projected data is synthetic (fabricated for illustration)
    is_synthetic: bool = False

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
            transmission=self.data.get("TransmissionStyle", ""),
            gvwr_raw=self.data.get("GVWR", ""),
            engine_power_hp=self.data.get("EngineHP", ""),
            engine_power_kw=self.data.get("EngineKW", ""),
            engine_type=self.data.get("EngineConfiguration", ""),
            fuel_type_secondary=self.data.get("FuelTypeSecondary", ""),
            vehicle_class=self.data.get("VehicleType", ""),
            plant_country=self.data.get("PlantCountry", ""),
            series=self.data.get("Series", ""),
            trim=self.data.get("Trim", "")
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

###############################################################################
# Data Quality Analysis - Step 13 Enhancement
###############################################################################

@dataclass
class DataQualityAnalysis(BaseModel):
    """
    Comprehensive data quality analysis for fleet-level quality tracking.
    Enhanced for Step 13 with trending and confidence metrics.
    """
    
    # Fleet reference
    fleet_name: str
    analysis_date: datetime.datetime = field(default_factory=datetime.datetime.now)
    
    # Overall quality metrics
    total_vehicles: int = 0
    avg_quality_score: float = 0.0
    quality_distribution: Dict[str, int] = field(default_factory=dict)  # High/Medium/Low counts
    
    # Quality breakdown by category
    avg_core_data_score: float = 0.0
    avg_fuel_economy_score: float = 0.0
    avg_commercial_data_score: float = 0.0
    avg_technical_score: float = 0.0
    avg_match_confidence: float = 0.0
    avg_consistency_score: float = 0.0
    
    # Commercial vehicle analysis
    commercial_vehicle_count: int = 0
    commercial_quality_avg: float = 0.0
    diesel_vehicle_count: int = 0
    gvwr_data_completeness: float = 0.0
    
    # Data completeness analysis
    completeness_by_field: Dict[str, float] = field(default_factory=dict)
    missing_critical_fields: List[str] = field(default_factory=list)
    
    # Confidence analysis
    high_confidence_count: int = 0  # >= 80% quality
    medium_confidence_count: int = 0  # 50-79% quality
    low_confidence_count: int = 0  # < 50% quality
    
    # Trend tracking
    quality_trends: Dict[str, List[float]] = field(default_factory=dict)
    historical_averages: Dict[str, float] = field(default_factory=dict)
    
    # Problem identification
    common_issues: List[str] = field(default_factory=list)
    improvement_recommendations: List[str] = field(default_factory=list)
    
    @classmethod
    def analyze_fleet(cls, vehicles: List[FleetVehicle], 
                     fleet_name: str = "Fleet Analysis") -> 'DataQualityAnalysis':
        """
        Perform comprehensive quality analysis on a fleet.
        
        Args:
            vehicles: List of vehicles to analyze
            fleet_name: Name of the fleet
            
        Returns:
            DataQualityAnalysis object with complete analysis
        """
        analysis = cls(fleet_name=fleet_name, total_vehicles=len(vehicles))
        
        if not vehicles:
            return analysis
        
        # Calculate detailed quality for all vehicles
        quality_scores = []
        core_scores = []
        fuel_scores = []
        commercial_scores = []
        technical_scores = []
        confidence_scores = []
        consistency_scores = []
        
        high_count = medium_count = low_count = 0
        commercial_count = diesel_count = 0
        commercial_quality_scores = []
        
        # Field completeness tracking
        field_counts = {}
        total_vehicles = len(vehicles)
        
        for vehicle in vehicles:
            # Calculate detailed quality if not already done
            if not vehicle.quality_breakdown:
                vehicle.calculate_detailed_quality()
            
            # Collect scores
            quality_scores.append(vehicle.data_quality_score)
            core_scores.append(vehicle.quality_breakdown.get("core_data", 0))
            fuel_scores.append(vehicle.quality_breakdown.get("fuel_economy", 0))
            commercial_scores.append(vehicle.quality_breakdown.get("commercial_data", 0))
            technical_scores.append(vehicle.quality_breakdown.get("technical_details", 0))
            confidence_scores.append(vehicle.quality_breakdown.get("match_confidence", 0))
            consistency_scores.append(vehicle.consistency_score)
            
            # Quality distribution
            if vehicle.data_quality_score >= 80:
                high_count += 1
            elif vehicle.data_quality_score >= 50:
                medium_count += 1
            else:
                low_count += 1
            
            # Commercial vehicle analysis
            if vehicle.vehicle_id.is_commercial:
                commercial_count += 1
                commercial_quality_scores.append(vehicle.data_quality_score)
            
            if vehicle.vehicle_id.is_diesel:
                diesel_count += 1
            
            # Field completeness analysis
            analysis._track_field_completeness(vehicle, field_counts)
        
        # Calculate averages
        analysis.avg_quality_score = sum(quality_scores) / len(quality_scores)
        analysis.avg_core_data_score = sum(core_scores) / len(core_scores)
        analysis.avg_fuel_economy_score = sum(fuel_scores) / len(fuel_scores)
        analysis.avg_commercial_data_score = sum(commercial_scores) / len(commercial_scores)
        analysis.avg_technical_score = sum(technical_scores) / len(technical_scores)
        analysis.avg_match_confidence = sum(confidence_scores) / len(confidence_scores)
        analysis.avg_consistency_score = sum(consistency_scores) / len(consistency_scores)
        
        # Quality distribution
        analysis.quality_distribution = {
            "High (80%+)": high_count,
            "Medium (50-79%)": medium_count,
            "Low (<50%)": low_count
        }
        
        analysis.high_confidence_count = high_count
        analysis.medium_confidence_count = medium_count
        analysis.low_confidence_count = low_count
        
        # Commercial vehicle metrics
        analysis.commercial_vehicle_count = commercial_count
        analysis.diesel_vehicle_count = diesel_count
        
        if commercial_quality_scores:
            analysis.commercial_quality_avg = sum(commercial_quality_scores) / len(commercial_quality_scores)
        
        # GVWR data completeness
        gvwr_complete = sum(1 for v in vehicles if v.vehicle_id.gvwr_pounds > 0)
        analysis.gvwr_data_completeness = (gvwr_complete / total_vehicles) * 100
        
        # Field completeness percentages
        analysis.completeness_by_field = {
            field: (count / total_vehicles) * 100 
            for field, count in field_counts.items()
        }
        
        # Identify issues and recommendations
        analysis._identify_issues_and_recommendations()
        
        return analysis
    
    def _track_field_completeness(self, vehicle: FleetVehicle, field_counts: Dict[str, int]):
        """Track field completeness for analysis."""
        fields_to_track = {
            "year": bool(vehicle.vehicle_id.year),
            "make": bool(vehicle.vehicle_id.make),
            "model": bool(vehicle.vehicle_id.model),
            "fuel_type": bool(vehicle.vehicle_id.fuel_type),
            "body_class": bool(vehicle.vehicle_id.body_class),
            "combined_mpg": vehicle.fuel_economy.combined_mpg > 0,
            "city_mpg": vehicle.fuel_economy.city_mpg > 0,
            "highway_mpg": vehicle.fuel_economy.highway_mpg > 0,
            "co2_emissions": vehicle.fuel_economy.co2_primary > 0,
            "gvwr": vehicle.vehicle_id.gvwr_pounds > 0,
            "commercial_category": bool(vehicle.vehicle_id.commercial_category),
            "engine_hp": bool(vehicle.vehicle_id.engine_power_hp),
            "engine_type": bool(vehicle.vehicle_id.engine_type),
            "engine_displacement": bool(vehicle.vehicle_id.engine_displacement),
            "transmission": bool(vehicle.vehicle_id.transmission),
            "drive_type": bool(vehicle.vehicle_id.drive_type)
        }
        
        for field, is_present in fields_to_track.items():
            if is_present:
                field_counts[field] = field_counts.get(field, 0) + 1
    
    def _identify_issues_and_recommendations(self):
        """Identify common issues and generate improvement recommendations."""
        issues = []
        recommendations = []
        
        # Low overall quality
        if self.avg_quality_score < 60:
            issues.append(f"Low average data quality: {self.avg_quality_score:.1f}%")
            recommendations.append("Review VIN data sources and API response handling")
        
        # Low commercial data completeness
        if self.avg_commercial_data_score < 40:
            issues.append("Limited commercial vehicle data completeness")
            recommendations.append("Verify NHTSA API extraction for GVWR and engine specifications")
        
        # Poor fuel economy data
        if self.avg_fuel_economy_score < 50:
            issues.append("Low fuel economy data completeness")
            recommendations.append("Check FuelEconomy.gov API matching algorithms")
        
        # High percentage of low-confidence vehicles
        low_confidence_pct = (self.low_confidence_count / self.total_vehicles) * 100
        if low_confidence_pct > 25:
            issues.append(f"High percentage of low-confidence vehicles: {low_confidence_pct:.1f}%")
            recommendations.append("Review VIN validation and API response handling")
        
        # Low GVWR completeness for commercial analysis
        if self.gvwr_data_completeness < 70:
            issues.append(f"Low GVWR data completeness: {self.gvwr_data_completeness:.1f}%")
            recommendations.append("Improve GVWR extraction from NHTSA responses")
        
        # Identify missing critical fields
        critical_fields = ["year", "make", "model", "combined_mpg"]
        missing_critical = []
        
        for field in critical_fields:
            completeness = self.completeness_by_field.get(field, 0)
            if completeness < 80:
                missing_critical.append(f"{field} ({completeness:.1f}% complete)")
        
        if missing_critical:
            issues.append("Critical fields have low completeness")
            self.missing_critical_fields = missing_critical
            recommendations.append("Focus on improving data extraction for critical vehicle identification fields")
        
        self.common_issues = issues
        self.improvement_recommendations = recommendations
    
    def get_quality_summary(self) -> Dict[str, Any]:
        """
        Get a summary of quality analysis results.
        
        Returns:
            Dictionary with key quality metrics and insights
        """
        return {
            "fleet_name": self.fleet_name,
            "analysis_date": self.analysis_date.isoformat(),
            "total_vehicles": self.total_vehicles,
            "overall_quality": {
                "average_score": round(self.avg_quality_score, 1),
                "distribution": self.quality_distribution,
                "grade": self._get_quality_grade()
            },
            "commercial_vehicles": {
                "count": self.commercial_vehicle_count,
                "diesel_count": self.diesel_vehicle_count,
                "avg_quality": round(self.commercial_quality_avg, 1),
                "gvwr_completeness": round(self.gvwr_data_completeness, 1)
            },
            "data_completeness": {
                "top_fields": self._get_top_complete_fields(),
                "improvement_needed": self.missing_critical_fields
            },
            "issues_and_recommendations": {
                "issues": self.common_issues,
                "recommendations": self.improvement_recommendations
            }
        }
    
    def _get_quality_grade(self) -> str:
        """Get a letter grade for overall fleet data quality."""
        if self.avg_quality_score >= 90:
            return "A"
        elif self.avg_quality_score >= 80:
            return "B"
        elif self.avg_quality_score >= 70:
            return "C"
        elif self.avg_quality_score >= 60:
            return "D"
        else:
            return "F"
    
    def _get_top_complete_fields(self) -> List[str]:
        """Get the top 5 most complete fields."""
        sorted_fields = sorted(
            self.completeness_by_field.items(),
            key=lambda x: x[1],
            reverse=True
        )
        return [f"{field} ({pct:.1f}%)" for field, pct in sorted_fields[:5]]


###############################################################################
# Presentation Profile
###############################################################################

@dataclass
class PresentationProfile(BaseModel):
    """Per-fleet presentation profile: client info, consultant contacts, content."""

    # Client info
    client_name: str = ""
    meeting_date: str = ""               # e.g. "March 5th, 2026"
    presentation_type: str = "Kickoff"   # Kickoff / Update 1 / Update 2 / Final

    # Presenter (the analyst building the deck)
    presenter_name: str = ""
    presenter_title: str = ""
    presenter_company: str = ""

    # Partner 1 contact (e.g. utility partner)
    partner_1_name: str = ""
    partner_1_title: str = ""
    partner_1_org: str = ""
    partner_1_email: str = ""

    # Partner 2 contact (optional)
    partner_2_name: str = ""
    partner_2_title: str = ""
    partner_2_org: str = ""
    partner_2_email: str = ""

    # Editable consulting content (one item per element)
    agenda_items: List[str] = field(default_factory=list)
    data_needs_items: List[str] = field(default_factory=list)
    next_steps_items: List[str] = field(default_factory=list)

    # Slide selection & order
    included_slides: List[str] = field(default_factory=list)   # ordered list of slide IDs
    optional_slides: List[str] = field(default_factory=list)   # extra slides to append

    # Template override (None = use DEFAULT_TEMPLATE_PATH)
    template_path: Optional[str] = None