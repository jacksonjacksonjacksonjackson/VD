"""
Shared test fixtures for the Fleet Electrification Analyzer test suite.
"""

import sys
import os
import datetime
import pytest

# Ensure the project root is on sys.path so imports work
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from data.models import VehicleIdentification, FuelEconomyData, FleetVehicle


# ---------------------------------------------------------------------------
# Factory helpers — build dataclass objects with sensible defaults
# ---------------------------------------------------------------------------

def make_vehicle_id(**overrides) -> VehicleIdentification:
    """Create a VehicleIdentification with reasonable defaults."""
    defaults = dict(
        vin="1HGBH41JXMN109186",
        year="2020",
        make="Ford",
        model="F-150",
        fuel_type="Gasoline",
        body_class="Pickup",
        gvwr="Class 2: 6,001 - 10,000 lb (2,722 - 4,536 kg)",
    )
    defaults.update(overrides)
    return VehicleIdentification(**defaults)


def make_fuel_economy(**overrides) -> FuelEconomyData:
    """Create a FuelEconomyData with reasonable defaults."""
    defaults = dict(
        city_mpg=20.0,
        highway_mpg=26.0,
        combined_mpg=22.0,
        co2_primary=404.0,
    )
    defaults.update(overrides)
    return FuelEconomyData(**defaults)


def make_fleet_vehicle(**overrides) -> FleetVehicle:
    """Create a FleetVehicle with reasonable defaults.

    Pass ``vid_overrides`` and ``fuel_overrides`` dicts to customise the
    nested VehicleIdentification / FuelEconomyData without constructing
    them manually.
    """
    vid_kw = overrides.pop("vid_overrides", {})
    fuel_kw = overrides.pop("fuel_overrides", {})

    defaults = dict(
        vin="1HGBH41JXMN109186",
        vehicle_id=make_vehicle_id(**vid_kw),
        fuel_economy=make_fuel_economy(**fuel_kw),
        department="General Services",
        odometer=45000.0,
        annual_mileage=12000.0,
        match_confidence=85.0,
        processing_success=True,
    )
    defaults.update(overrides)
    return FleetVehicle(**defaults)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def sample_vehicle():
    """A typical light-duty gasoline pickup."""
    return make_fleet_vehicle()


@pytest.fixture
def diesel_vehicle():
    """A medium-duty diesel truck."""
    return make_fleet_vehicle(
        vid_overrides=dict(
            vin="1FDUF5HT4FEB12345",
            year="2018",
            make="Ford",
            model="F-550",
            fuel_type="Diesel",
            body_class="Chassis Cab",
            gvwr="Class 5: 16,001 - 19,500 lb",
        ),
        fuel_overrides=dict(
            combined_mpg=12.0,
            city_mpg=10.0,
            highway_mpg=14.0,
            co2_primary=741.0,
        ),
        department="Public Works",
        odometer=92000.0,
        annual_mileage=18000.0,
    )


@pytest.fixture
def electric_vehicle():
    """A battery-electric light-duty vehicle."""
    return make_fleet_vehicle(
        vid_overrides=dict(
            vin="5YJ3E1EA1LF123456",
            year="2022",
            make="Tesla",
            model="Model 3",
            fuel_type="Battery Electric Vehicle (BEV)",
            body_class="Sedan/Saloon",
            gvwr="Class 1: 0 - 6,000 lb",
        ),
        fuel_overrides=dict(
            combined_mpg=0.0,
            city_mpg=0.0,
            highway_mpg=0.0,
            co2_primary=0.0,
        ),
        department="Administration",
        odometer=15000.0,
        annual_mileage=10000.0,
    )


@pytest.fixture
def emergency_vehicle():
    """A police pursuit vehicle."""
    return make_fleet_vehicle(
        vid_overrides=dict(
            vin="1FM5K8AR6MNA12345",
            year="2021",
            make="Ford",
            model="Police Interceptor Utility",
            fuel_type="Gasoline",
            body_class="Sport Utility Vehicle (SUV)/Multi-Purpose Vehicle (MPV)",
            gvwr="Class 2: 6,001 - 10,000 lb",
            trim="PPV",
        ),
        department="Police Department",
        odometer=68000.0,
        annual_mileage=25000.0,
    )


@pytest.fixture
def sample_fleet(sample_vehicle, diesel_vehicle, electric_vehicle):
    """A small mixed fleet for integration tests."""
    from data.models import Fleet
    fleet = Fleet(name="Test Fleet")
    fleet.vehicles = [sample_vehicle, diesel_vehicle, electric_vehicle]
    return fleet
