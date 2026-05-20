"""Tests for TCO, ROI, and emissions calculations in analysis/calculations.py."""

import pytest
from analysis.calculations import (
    calculate_annual_fuel_cost,
    calculate_annual_ev_cost,
    calculate_annual_co2_emissions,
    calculate_emissions_reduction,
    calculate_electrification_savings,
    calculate_ev_roi,
    create_emissions_inventory,
)
from tests.conftest import make_fleet_vehicle, make_fuel_economy


# =========================================================================
# Annual Fuel Cost
# =========================================================================

class TestAnnualFuelCost:

    def test_basic_calculation(self, sample_vehicle):
        """12,000 miles / 22 MPG * $3.50 = ~$1909."""
        cost = calculate_annual_fuel_cost(sample_vehicle, gas_price=3.50)
        expected = 12000 / 22.0 * 3.50
        assert abs(cost - expected) < 0.01

    def test_zero_mpg_returns_zero(self):
        v = make_fleet_vehicle(fuel_overrides=dict(combined_mpg=0))
        cost = calculate_annual_fuel_cost(v, gas_price=3.50)
        assert cost == 0.0

    def test_higher_gas_price_increases_cost(self, sample_vehicle):
        cost_low = calculate_annual_fuel_cost(sample_vehicle, gas_price=3.00)
        cost_high = calculate_annual_fuel_cost(sample_vehicle, gas_price=5.00)
        assert cost_high > cost_low


# =========================================================================
# Annual EV Cost
# =========================================================================

class TestAnnualEvCost:

    def test_basic_calculation(self, sample_vehicle):
        """12,000 miles * 0.30 kWh/mi * $0.13/kWh = $468."""
        cost = calculate_annual_ev_cost(sample_vehicle, electricity_price=0.13, ev_efficiency=0.30)
        expected = 12000 * 0.30 * 0.13
        assert abs(cost - expected) < 0.01


# =========================================================================
# Annual CO2 Emissions
# =========================================================================

class TestAnnualCo2Emissions:

    def test_with_co2_data(self, sample_vehicle):
        """CO2 per mile (404 g/mi) * annual mileage / 1e6 -> metric tons."""
        emissions = calculate_annual_co2_emissions(sample_vehicle)
        expected = 404.0 * 12000.0 / 1_000_000
        assert abs(emissions - expected) < 0.001

    def test_zero_emissions_when_no_data(self):
        v = make_fleet_vehicle(fuel_overrides=dict(co2_primary=0, combined_mpg=0))
        emissions = calculate_annual_co2_emissions(v)
        assert emissions == 0.0

    def test_estimates_from_mpg_when_no_co2(self):
        """When co2_primary is 0 but MPG exists, should estimate from MPG."""
        v = make_fleet_vehicle(fuel_overrides=dict(co2_primary=0, combined_mpg=25))
        emissions = calculate_annual_co2_emissions(v)
        assert emissions > 0  # Should estimate ~4.27 metric tons


# =========================================================================
# Emissions Reduction
# =========================================================================

class TestEmissionsReduction:

    def test_positive_reduction(self, sample_vehicle):
        """ICE -> EV should reduce emissions."""
        reduction = calculate_emissions_reduction(sample_vehicle)
        assert reduction > 0

    def test_ev_has_near_zero_reduction(self, electric_vehicle):
        """EV already has zero tailpipe; reduction is small or zero."""
        reduction = calculate_emissions_reduction(electric_vehicle)
        assert reduction <= 0.01


# =========================================================================
# TCO / ROI (Fix 32 regression test)
# =========================================================================

class TestCalculateEvRoi:

    def test_ice_tco_includes_operating_costs(self, sample_vehicle):
        """ICE TCO = purchase + fuel + maintenance over analysis period."""
        result = calculate_ev_roi(
            vehicle=sample_vehicle,
            ev_purchase_price=45000,
            ice_purchase_price=35000,
            gas_price=3.50,
            electricity_price=0.13,
            ev_efficiency=0.30,
            analysis_years=10,
            ice_maintenance=0.10,
            ev_maintenance=0.06,
        )
        # ICE TCO should be > purchase price (operating costs added)
        assert result["ice_tco"] > 35000

    def test_ev_tco_includes_operating_costs(self, sample_vehicle):
        """EV TCO = purchase + electricity + maintenance over period."""
        result = calculate_ev_roi(
            vehicle=sample_vehicle,
            ev_purchase_price=45000,
            ice_purchase_price=35000,
            gas_price=3.50,
            electricity_price=0.13,
            ev_efficiency=0.30,
            analysis_years=10,
        )
        # EV TCO should be > purchase price
        assert result["ev_tco"] > 45000

    def test_tco_savings_is_difference(self, sample_vehicle):
        result = calculate_ev_roi(
            vehicle=sample_vehicle,
            ev_purchase_price=45000,
            ice_purchase_price=35000,
        )
        assert abs(result["tco_savings"] - (result["ice_tco"] - result["ev_tco"])) < 0.02

    def test_payback_period_positive(self, sample_vehicle):
        result = calculate_ev_roi(
            vehicle=sample_vehicle,
            ev_purchase_price=45000,
            ice_purchase_price=35000,
        )
        assert result["payback_years"] > 0

    def test_zero_savings_infinite_payback(self):
        """If EV costs more to operate, payback should be infinite."""
        v = make_fleet_vehicle(
            fuel_overrides=dict(combined_mpg=50),  # Very efficient ICE
            annual_mileage=5000,  # Low usage
        )
        result = calculate_ev_roi(
            vehicle=v,
            ev_purchase_price=60000,
            ice_purchase_price=25000,
            gas_price=2.00,
            electricity_price=0.30,  # Expensive electricity
            ev_efficiency=0.50,  # Poor EV efficiency
        )
        # With very efficient ICE and expensive electricity, savings might be <= 0
        if result["annual_savings"] <= 0:
            assert result["payback_years"] == float("inf")


# =========================================================================
# Emissions Inventory (Fix 34 regression test)
# =========================================================================

class TestEmissionsInventory:

    def test_synthetic_flag_set(self, sample_fleet):
        """create_emissions_inventory() should set is_synthetic=True."""
        inventory = create_emissions_inventory(sample_fleet)
        assert inventory.is_synthetic is True

    def test_total_emissions_positive(self, sample_fleet):
        inventory = create_emissions_inventory(sample_fleet)
        assert inventory.total_emissions > 0

    def test_historical_data_present(self, sample_fleet):
        inventory = create_emissions_inventory(sample_fleet)
        assert len(inventory.historical_data) > 0

    def test_projected_emissions_present(self, sample_fleet):
        inventory = create_emissions_inventory(sample_fleet)
        assert len(inventory.projected_emissions) > 0
