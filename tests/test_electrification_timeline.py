"""Tests for electrification timeline scoring and year assignment."""

import datetime
import pytest
from unittest.mock import patch

from analysis.electrification_timeline import (
    _score_vehicle,
    assign_electrification_years,
    WEIGHT_AGE,
    WEIGHT_MILEAGE,
    WEIGHT_ANNUAL_USAGE,
    MAX_AGE_YEARS,
    MAX_ODOMETER,
    MAX_ANNUAL_MILEAGE,
    ACF_BOOST,
)
from tests.conftest import make_fleet_vehicle


# =========================================================================
# _score_vehicle() — individual vehicle priority scoring
# =========================================================================

class TestScoreVehicle:

    # ── Age component ────────────────────────────────────────────────

    def test_max_age_scores_full_weight(self):
        """Vehicle at MAX_AGE_YEARS should get the full age weight."""
        current_year = datetime.datetime.now().year
        v = make_fleet_vehicle(
            vid_overrides=dict(year=str(current_year - MAX_AGE_YEARS)),
            odometer=0.0,
            annual_mileage=0.0,
        )
        v.custom_fields["_acf_code"] = "B"
        score = _score_vehicle(v, "B")
        # Age component should be WEIGHT_AGE * 1.0, plus ACF boost
        expected_min = WEIGHT_AGE * 1.0
        assert score >= expected_min

    def test_zero_age_scores_zero_age_component(self):
        """Vehicle with current model year should get zero from age."""
        current_year = datetime.datetime.now().year
        v = make_fleet_vehicle(
            vid_overrides=dict(year=str(current_year)),
            odometer=0.0,
            annual_mileage=0.0,
        )
        v.custom_fields["_acf_code"] = "C"
        score = _score_vehicle(v, "C")
        # Only ACF boost (halved because no data present)
        assert score == ACF_BOOST["C"] * 0.5

    # ── Mileage component ────────────────────────────────────────────

    def test_high_odometer_increases_score(self):
        """Vehicle with MAX_ODOMETER should get full mileage weight."""
        v = make_fleet_vehicle(
            vid_overrides=dict(year="2020"),
            odometer=MAX_ODOMETER,
            annual_mileage=0.0,
        )
        v.custom_fields["_acf_code"] = "B"
        score_high = _score_vehicle(v, "B")

        v2 = make_fleet_vehicle(
            vid_overrides=dict(year="2020"),
            odometer=0.0,
            annual_mileage=0.0,
        )
        v2.custom_fields["_acf_code"] = "B"
        score_low = _score_vehicle(v2, "B")
        assert score_high > score_low

    # ── Annual usage component ───────────────────────────────────────

    def test_high_annual_mileage_increases_score(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(year="2020"),
            odometer=0.0,
            annual_mileage=MAX_ANNUAL_MILEAGE,
        )
        v.custom_fields["_acf_code"] = "B"
        score = _score_vehicle(v, "B")
        # Should include usage contribution
        assert score > ACF_BOOST["B"]

    # ── ACF boost ────────────────────────────────────────────────────

    def test_category_b_gets_highest_boost(self):
        """ACF-subject (B) vehicles get 0.30 boost, the highest."""
        v = make_fleet_vehicle(
            vid_overrides=dict(year="2018"),
            odometer=50000.0,
            annual_mileage=12000.0,
        )
        v.custom_fields["_acf_code"] = "B"
        score_b = _score_vehicle(v, "B")

        v2 = make_fleet_vehicle(
            vid_overrides=dict(year="2018"),
            odometer=50000.0,
            annual_mileage=12000.0,
        )
        v2.custom_fields["_acf_code"] = "C"
        score_c = _score_vehicle(v2, "C")
        assert score_b > score_c

    def test_no_acf_boost_for_unknown_code(self):
        """Unknown ACF code should add zero boost."""
        v = make_fleet_vehicle(
            vid_overrides=dict(year="2018"),
            odometer=50000.0,
            annual_mileage=12000.0,
        )
        v.custom_fields["_acf_code"] = "X"
        score = _score_vehicle(v, "X")
        # Same vehicle with code B should score higher
        v2 = make_fleet_vehicle(
            vid_overrides=dict(year="2018"),
            odometer=50000.0,
            annual_mileage=12000.0,
        )
        v2.custom_fields["_acf_code"] = "B"
        score_b = _score_vehicle(v2, "B")
        assert score_b > score

    # ── Data-completeness penalty (Fix 20 regression) ────────────────

    def test_no_data_halves_acf_boost(self):
        """Vehicle with zero age, zero odo, zero annual mileage gets halved boost."""
        current_year = datetime.datetime.now().year
        v = make_fleet_vehicle(
            vid_overrides=dict(year=str(current_year)),
            odometer=0.0,
            annual_mileage=0.0,
        )
        v.custom_fields["_acf_code"] = "B"
        score = _score_vehicle(v, "B")
        assert score == ACF_BOOST["B"] * 0.5

    def test_one_metric_present_gives_full_boost(self):
        """If at least one metric is present, full ACF boost applies."""
        v = make_fleet_vehicle(
            vid_overrides=dict(year="2018"),  # has age
            odometer=0.0,
            annual_mileage=0.0,
        )
        v.custom_fields["_acf_code"] = "B"
        score = _score_vehicle(v, "B")
        # Score should include full boost (0.30) + age component
        assert score > ACF_BOOST["B"] * 0.5

    # ── Score range ──────────────────────────────────────────────────

    def test_score_non_negative(self):
        """Score should never be negative."""
        v = make_fleet_vehicle(odometer=0.0, annual_mileage=0.0)
        v.custom_fields["_acf_code"] = "B"
        score = _score_vehicle(v, "B")
        assert score >= 0.0

    def test_score_upper_bound_reasonable(self):
        """Max score should be roughly WEIGHT_AGE + WEIGHT_MILEAGE + WEIGHT_ANNUAL_USAGE + max boost."""
        current_year = datetime.datetime.now().year
        v = make_fleet_vehicle(
            vid_overrides=dict(year=str(current_year - 20)),
            odometer=300000.0,
            annual_mileage=30000.0,
        )
        v.custom_fields["_acf_code"] = "B"
        score = _score_vehicle(v, "B")
        theoretical_max = WEIGHT_AGE + WEIGHT_MILEAGE + WEIGHT_ANNUAL_USAGE + ACF_BOOST["B"]
        assert score <= theoretical_max + 0.01


# =========================================================================
# assign_electrification_years() — fleet-wide year assignment
# =========================================================================

class TestAssignElectrificationYears:

    # ── ZEV handling ─────────────────────────────────────────────────

    def test_zev_gets_na(self):
        """ZEVs should get 'N/A' — already electric."""
        v = make_fleet_vehicle(
            vid_overrides=dict(fuel_type="Battery Electric Vehicle (BEV)"),
        )
        v.custom_fields["_acf_code"] = "ZEV"
        assign_electrification_years([v])
        assert v.custom_fields["Proposed EV Year"] == "N/A"

    # ── Failed vehicle handling ──────────────────────────────────────

    def test_failed_vehicle_gets_na(self):
        v = make_fleet_vehicle(processing_success=False)
        v.custom_fields["_acf_code"] = "B"
        assign_electrification_years([v])
        assert v.custom_fields["Proposed EV Year"] == "N/A"

    # ── Missing ACF code ─────────────────────────────────────────────

    def test_missing_acf_code_gets_na(self):
        v = make_fleet_vehicle()
        # No _acf_code set
        assign_electrification_years([v])
        assert v.custom_fields["Proposed EV Year"] == "N/A"

    # ── Light-duty exempt (Category A) ───────────────────────────────

    def test_category_a_gets_exempt(self):
        """Light-duty exempt vehicles should show 'Exempt'."""
        v = make_fleet_vehicle(
            vid_overrides=dict(gvwr="6000 lb", fuel_type="Gasoline"),
        )
        v.custom_fields["_acf_code"] = "A"
        assign_electrification_years([v])
        assert v.custom_fields["Proposed EV Year"] == "Exempt"

    # ── Schedulable vehicles ─────────────────────────────────────────

    def test_category_b_gets_year(self):
        """ACF-subject vehicles should get a numeric year."""
        v = make_fleet_vehicle(
            vid_overrides=dict(
                gvwr="14000 lb",
                fuel_type="Diesel",
                body_class="Truck",
            ),
            odometer=80000.0,
            annual_mileage=15000.0,
        )
        v.custom_fields["_acf_code"] = "B"
        assign_electrification_years([v], end_year=2040)
        year_str = v.custom_fields["Proposed EV Year"]
        year = int(year_str)
        current_year = datetime.datetime.now().year
        assert current_year < year <= 2040

    def test_category_c_gets_year(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(body_class="Dump Truck", gvwr="33000 lb", fuel_type="Diesel"),
        )
        v.custom_fields["_acf_code"] = "C"
        assign_electrification_years([v], end_year=2040)
        year_str = v.custom_fields["Proposed EV Year"]
        year = int(year_str)
        assert year > datetime.datetime.now().year

    def test_category_d_gets_year(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(body_class="Ambulance", gvwr="14000 lb", fuel_type="Gasoline"),
            department="EMS",
        )
        v.custom_fields["_acf_code"] = "D"
        assign_electrification_years([v], end_year=2040)
        year_str = v.custom_fields["Proposed EV Year"]
        year = int(year_str)
        assert year > datetime.datetime.now().year

    # ── Ordering: higher-priority vehicles assigned earlier years ─────

    def test_higher_score_gets_earlier_year(self):
        """An old high-mileage vehicle should be scheduled before a new low-mileage one."""
        old = make_fleet_vehicle(
            vid_overrides=dict(year="2005", gvwr="14000 lb", fuel_type="Diesel"),
            odometer=180000.0,
            annual_mileage=20000.0,
        )
        old.custom_fields["_acf_code"] = "B"

        new = make_fleet_vehicle(
            vid_overrides=dict(year="2023", gvwr="14000 lb", fuel_type="Diesel"),
            odometer=10000.0,
            annual_mileage=5000.0,
        )
        new.custom_fields["_acf_code"] = "B"

        assign_electrification_years([old, new], end_year=2040)
        old_year = int(old.custom_fields["Proposed EV Year"])
        new_year = int(new.custom_fields["Proposed EV Year"])
        assert old_year <= new_year

    # ── Budget smoothing ─────────────────────────────────────────────

    def test_budget_smoothing_distributes_evenly(self):
        """Multiple vehicles should be spread across available years."""
        vehicles = []
        for i in range(10):
            v = make_fleet_vehicle(
                vid_overrides=dict(
                    year=str(2010 + i),
                    gvwr="14000 lb",
                    fuel_type="Diesel",
                ),
                odometer=50000.0 + i * 10000,
                annual_mileage=12000.0,
            )
            v.custom_fields["_acf_code"] = "B"
            vehicles.append(v)

        assign_electrification_years(vehicles, end_year=2040)

        years = [int(v.custom_fields["Proposed EV Year"]) for v in vehicles]
        unique_years = set(years)
        # With 10 vehicles and ~15 available years, should use at least 2 different years
        assert len(unique_years) >= 2

    # ── Edge case: end_year <= current_year ───────────────────────────

    def test_past_end_year_assigns_na(self):
        """If end_year is in the past, all vehicles get N/A."""
        v = make_fleet_vehicle()
        v.custom_fields["_acf_code"] = "B"
        assign_electrification_years([v], end_year=2020)
        assert v.custom_fields["Proposed EV Year"] == "N/A"

    # ── Mixed fleet ──────────────────────────────────────────────────

    def test_mixed_fleet_triage(self):
        """A mixed fleet should have ZEV->N/A, A->Exempt, B/C/D->year."""
        zev = make_fleet_vehicle(
            vid_overrides=dict(fuel_type="Battery Electric Vehicle (BEV)"),
        )
        zev.custom_fields["_acf_code"] = "ZEV"

        light = make_fleet_vehicle(
            vid_overrides=dict(gvwr="6000 lb", fuel_type="Gasoline"),
        )
        light.custom_fields["_acf_code"] = "A"

        truck = make_fleet_vehicle(
            vid_overrides=dict(gvwr="14000 lb", fuel_type="Diesel"),
            odometer=80000.0,
        )
        truck.custom_fields["_acf_code"] = "B"

        failed = make_fleet_vehicle(processing_success=False)
        failed.custom_fields["_acf_code"] = "B"

        assign_electrification_years([zev, light, truck, failed], end_year=2040)
        assert zev.custom_fields["Proposed EV Year"] == "N/A"
        assert light.custom_fields["Proposed EV Year"] == "Exempt"
        assert failed.custom_fields["Proposed EV Year"] == "N/A"
        truck_year = int(truck.custom_fields["Proposed EV Year"])
        assert truck_year > datetime.datetime.now().year

    # ── Empty fleet ──────────────────────────────────────────────────

    def test_empty_fleet(self):
        """Should handle empty list without error."""
        assign_electrification_years([], end_year=2040)
        # No assertion needed — just should not raise
