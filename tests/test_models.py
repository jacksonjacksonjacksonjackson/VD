"""Tests for data models: GVWR parsing, commercial classification, diesel detection, quality scoring."""

import pytest
from data.models import VehicleIdentification, FuelEconomyData, FleetVehicle
from tests.conftest import make_vehicle_id, make_fleet_vehicle, make_fuel_economy


# =========================================================================
# GVWR Parsing (_process_gvwr)
# =========================================================================

class TestGvwrParsing:
    """Tests for VehicleIdentification._process_gvwr()."""

    def test_nhtsa_class_format(self):
        """'Class 2: 6,001 - 10,000 lb (2,722 - 4,536 kg)' -> 10000 (upper bound)"""
        vid = make_vehicle_id(gvwr="Class 2: 6,001 - 10,000 lb (2,722 - 4,536 kg)")
        assert vid.gvwr_pounds == 10000.0

    def test_simple_pounds(self):
        vid = make_vehicle_id(gvwr="8500 lbs")
        assert vid.gvwr_pounds == 8500.0

    def test_with_comma(self):
        vid = make_vehicle_id(gvwr="14,500 lb")
        assert vid.gvwr_pounds == 14500.0

    def test_raw_number(self):
        vid = make_vehicle_id(gvwr_raw="19500", gvwr="")
        assert vid.gvwr_pounds == 19500.0

    def test_empty_gvwr(self):
        vid = make_vehicle_id(gvwr="", gvwr_raw="")
        assert vid.gvwr_pounds == 0.0

    # Commercial category classification based on GVWR
    def test_light_duty_classification(self):
        vid = make_vehicle_id(gvwr="6000 lb")
        assert vid.commercial_category == "Light Duty"

    def test_medium_duty_classification(self):
        vid = make_vehicle_id(gvwr="14500 lb")
        assert vid.commercial_category == "Medium Duty"

    def test_heavy_duty_classification(self):
        vid = make_vehicle_id(gvwr="26000 lb")
        assert vid.commercial_category == "Heavy Duty"

    def test_extra_heavy_duty_classification(self):
        vid = make_vehicle_id(gvwr="40000 lb")
        assert vid.commercial_category == "Extra Heavy Duty"

    def test_boundary_light_medium(self):
        """8500 is the last light-duty value."""
        vid = make_vehicle_id(gvwr="8500 lb")
        assert vid.commercial_category == "Light Duty"

    def test_boundary_medium_starts(self):
        """8501 is medium-duty."""
        vid = make_vehicle_id(gvwr="8501 lb")
        assert vid.commercial_category == "Medium Duty"


# =========================================================================
# Commercial Vehicle Detection
# =========================================================================

class TestCommercialDetection:
    """Tests for VehicleIdentification._classify_commercial()."""

    def test_pickup_body_class(self):
        vid = make_vehicle_id(body_class="Pickup")
        assert vid.is_commercial is True

    def test_sedan_not_commercial(self):
        vid = make_vehicle_id(body_class="Sedan/Saloon", model="Civic", gvwr="3500 lb")
        assert vid.is_commercial is False

    def test_transit_model(self):
        vid = make_vehicle_id(model="Transit 350", body_class="Van", gvwr="9000 lb")
        assert vid.is_commercial is True

    def test_heavy_gvwr_triggers_commercial(self):
        vid = make_vehicle_id(body_class="Incomplete", model="Custom", gvwr="12000 lb")
        assert vid.is_commercial is True


# =========================================================================
# Diesel Detection
# =========================================================================

class TestDieselDetection:
    """Tests for VehicleIdentification._detect_diesel()."""

    def test_diesel_fuel_type(self):
        vid = make_vehicle_id(fuel_type="Diesel")
        assert vid.is_diesel is True

    def test_biodiesel_fuel_type(self):
        vid = make_vehicle_id(fuel_type="Biodiesel (B20)")
        assert vid.is_diesel is True

    def test_gasoline_not_diesel(self):
        vid = make_vehicle_id(fuel_type="Gasoline")
        assert vid.is_diesel is False

    def test_electric_not_diesel(self):
        vid = make_vehicle_id(fuel_type="Battery Electric Vehicle (BEV)")
        assert vid.is_diesel is False

    def test_diesel_secondary_fuel(self):
        vid = make_vehicle_id(fuel_type="Gasoline", fuel_type_secondary="Diesel")
        assert vid.is_diesel is True


# =========================================================================
# Quality Scoring
# =========================================================================

class TestQualityScoring:
    """Tests for FleetVehicle.calculate_detailed_quality()."""

    def test_full_data_high_score(self):
        """Vehicle with all data fields should score well."""
        v = make_fleet_vehicle(match_confidence=95.0)
        result = v.calculate_detailed_quality()
        assert result["total_score"] >= 50  # Core + fuel + confidence

    def test_minimal_data_low_score(self):
        """Vehicle with only VIN should score low."""
        v = FleetVehicle(
            vin="1HGBH41JXMN109186",
            vehicle_id=VehicleIdentification(vin="1HGBH41JXMN109186"),
            fuel_economy=FuelEconomyData(),
        )
        result = v.calculate_detailed_quality()
        assert result["total_score"] < 30

    def test_score_capped_at_100(self):
        """Score should never exceed 100."""
        v = make_fleet_vehicle(match_confidence=100.0)
        result = v.calculate_detailed_quality()
        assert result["total_score"] <= 100.0

    def test_breakdown_keys_present(self):
        v = make_fleet_vehicle()
        result = v.calculate_detailed_quality()
        for key in ("core_data", "fuel_economy", "commercial_data",
                     "technical_details", "match_confidence", "consistency_bonus"):
            assert key in result["breakdown"]

    def test_fuel_economy_points(self):
        """Having all MPG values should contribute fuel_economy points."""
        v = make_fleet_vehicle(
            fuel_overrides=dict(combined_mpg=22, city_mpg=18, highway_mpg=26, co2_primary=400)
        )
        v.calculate_detailed_quality()
        assert v.quality_breakdown["fuel_economy"] == 25.0  # 12 + 6 + 6 + 1

    def test_no_fuel_economy_zero_points(self):
        v = make_fleet_vehicle(
            fuel_overrides=dict(combined_mpg=0, city_mpg=0, highway_mpg=0, co2_primary=0)
        )
        v.calculate_detailed_quality()
        assert v.quality_breakdown["fuel_economy"] == 0.0
