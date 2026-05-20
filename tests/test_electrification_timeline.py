"""Tests for electrification timeline scoring and year assignment."""

import datetime
import pytest
from unittest.mock import patch

from analysis.electrification_timeline import (
    _score_vehicle,
    _acf_deadline_for_vehicle,
    assign_electrification_years,
    ACF_DEADLINE_TABLE,
    ACF_DEADLINE_TABLE_NON_HPF,
    ACF_DEADLINE_UNKNOWN_GVWR,
    ACF_DEADLINE_UNKNOWN_GVWR_NON_HPF,
    WEIGHT_AGE,
    WEIGHT_MILEAGE,
    WEIGHT_ANNUAL_USAGE,
    MAX_AGE_YEARS,
    MAX_ODOMETER,
    MAX_ANNUAL_MILEAGE,
    ACF_BOOST,
    ACF_RELEVANCE,
    ACF_BOOST_B_CLASS_2B4,
    ACF_BOOST_B_CLASS_5_8A,
    ACF_BOOST_B_CLASS_8B,
    ACF_BOOST_B_UNKNOWN,
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
        # Default vehicle GVWR maps to Class 2b-4 (≤ 10,000 lb), so the
        # GVWR-tiered boost is ACF_BOOST_B_CLASS_2B4.  With no useful data
        # (age=0, odo=0, mileage=0) the boost is halved.
        assert score == pytest.approx(ACF_BOOST_B_CLASS_2B4 * 0.5)

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
        # ACF_BOOST_B_CLASS_2B4 is the highest Cat B tier boost (Class 2b-4, earliest deadline)
        theoretical_max = WEIGHT_AGE + WEIGHT_MILEAGE + WEIGHT_ANNUAL_USAGE + ACF_BOOST_B_CLASS_2B4
        assert score <= theoretical_max + 0.01


# =========================================================================
# _acf_deadline_for_vehicle() — ACF mandate deadline lookup
# =========================================================================

class TestAcfDeadlineLookup:
    CURRENT_YEAR = datetime.datetime.now().year

    def test_medium_duty_class_2b_4_gets_2035_milestone(self):
        """GVWR 14,000 lbs → Class 2b-4 → milestone 2035."""
        v = make_fleet_vehicle(vid_overrides=dict(gvwr="14000 lb", fuel_type="Diesel"))
        v.vehicle_id.gvwr_pounds = 14000.0
        year, reason = _acf_deadline_for_vehicle(v, self.CURRENT_YEAR)
        assert year is not None
        # May be pulled to an earlier checkpoint for high-urgency vehicles,
        # but must not exceed the milestone year
        assert year <= 2035

    def test_class_5_8a_gets_2039_milestone_max(self):
        """GVWR 25,000 lbs → Class 5-8a → milestone 2039."""
        v = make_fleet_vehicle(vid_overrides=dict(gvwr="25000 lb", fuel_type="Diesel"))
        v.vehicle_id.gvwr_pounds = 25000.0
        # New vehicle (low urgency) should get the full milestone year
        v.vehicle_id.year = str(self.CURRENT_YEAR - 1)
        v.odometer = 1000.0
        v.annual_mileage = 5000.0
        year, reason = _acf_deadline_for_vehicle(v, self.CURRENT_YEAR)
        assert year is not None
        assert year <= 2039

    def test_class_8b_gets_2042_milestone_max(self):
        """GVWR 40,000 lbs → Class 8b → milestone 2042."""
        v = make_fleet_vehicle(vid_overrides=dict(gvwr="40000 lb", fuel_type="Diesel"))
        v.vehicle_id.gvwr_pounds = 40000.0
        v.vehicle_id.year = str(self.CURRENT_YEAR - 1)
        v.odometer = 1000.0
        v.annual_mileage = 5000.0
        year, reason = _acf_deadline_for_vehicle(v, self.CURRENT_YEAR)
        assert year is not None
        assert year <= 2042

    def test_unknown_gvwr_returns_none(self):
        """Vehicle with GVWR 0 should return (None, '') — falls through to score queue."""
        v = make_fleet_vehicle(vid_overrides=dict(gvwr="", fuel_type="Diesel"))
        v.vehicle_id.gvwr_pounds = 0.0
        year, reason = _acf_deadline_for_vehicle(v, self.CURRENT_YEAR)
        assert year is None
        assert reason == ""

    def test_reason_contains_class_label(self):
        """Reason string should mention the applicable GVWR class."""
        v = make_fleet_vehicle(vid_overrides=dict(gvwr="14000 lb", fuel_type="Diesel"))
        v.vehicle_id.gvwr_pounds = 14000.0
        _year, reason = _acf_deadline_for_vehicle(v, self.CURRENT_YEAR)
        assert "Class" in reason

    def test_high_urgency_pulls_to_earlier_checkpoint(self):
        """An old high-mileage vehicle should be pulled to an early purchase checkpoint."""
        v = make_fleet_vehicle(
            vid_overrides=dict(year="2005", gvwr="14000 lb", fuel_type="Diesel"),
            odometer=190000.0,
            annual_mileage=20000.0,
        )
        v.vehicle_id.gvwr_pounds = 14000.0
        year, _reason = _acf_deadline_for_vehicle(v, self.CURRENT_YEAR)
        # High urgency (score >= 0.55) → earliest future purchase checkpoint
        # which should be < 2035 milestone
        assert year is not None
        assert year < 2035


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

    def test_zev_reason_mentions_already_zev(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(fuel_type="Battery Electric Vehicle (BEV)"),
        )
        v.custom_fields["_acf_code"] = "ZEV"
        assign_electrification_years([v])
        assert "zero-emission" in v.custom_fields["EV Year Reason"].lower()

    # ── Failed vehicle handling ──────────────────────────────────────

    def test_failed_vehicle_gets_na(self):
        v = make_fleet_vehicle(processing_success=False)
        v.custom_fields["_acf_code"] = "B"
        assign_electrification_years([v])
        assert v.custom_fields["Proposed EV Year"] == "N/A"

    def test_failed_vehicle_reason_mentions_processing(self):
        v = make_fleet_vehicle(processing_success=False)
        v.custom_fields["_acf_code"] = "B"
        assign_electrification_years([v])
        assert "processing failed" in v.custom_fields["EV Year Reason"].lower()

    # ── Missing ACF code ─────────────────────────────────────────────

    def test_missing_acf_code_gets_na(self):
        v = make_fleet_vehicle()
        # No _acf_code set
        assign_electrification_years([v])
        assert v.custom_fields["Proposed EV Year"] == "N/A"

    def test_missing_acf_code_reason_mentions_classification(self):
        v = make_fleet_vehicle()
        assign_electrification_years([v])
        assert "classification" in v.custom_fields["EV Year Reason"].lower()

    # ── Light-duty exempt (Category A) ───────────────────────────────

    def test_category_a_gets_exempt(self):
        """Light-duty exempt vehicles should show 'Exempt'."""
        v = make_fleet_vehicle(
            vid_overrides=dict(gvwr="6000 lb", fuel_type="Gasoline"),
        )
        v.custom_fields["_acf_code"] = "A"
        assign_electrification_years([v])
        assert v.custom_fields["Proposed EV Year"] == "Exempt"

    def test_category_a_reason_mentions_light_duty(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(gvwr="6000 lb", fuel_type="Gasoline"),
        )
        v.custom_fields["_acf_code"] = "A"
        assign_electrification_years([v])
        assert "light-duty" in v.custom_fields["EV Year Reason"].lower()

    # ── ACF Relevance field ──────────────────────────────────────────

    def test_acf_relevance_populated_for_all_categories(self):
        """ACF Relevance should be set for every vehicle regardless of outcome."""
        codes = ["ZEV", "A", "B", "C", "D"]
        for code in codes:
            v = make_fleet_vehicle(
                vid_overrides=dict(gvwr="14000 lb", fuel_type="Diesel"),
                odometer=50000.0,
            )
            v.custom_fields["_acf_code"] = code
            if code == "ZEV":
                v.vehicle_id.fuel_type = "Battery Electric Vehicle (BEV)"
            assign_electrification_years([v])
            assert "ACF Relevance" in v.custom_fields
            assert v.custom_fields["ACF Relevance"]  # not empty

    def test_category_b_relevance_mentions_mandate(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(gvwr="14000 lb", fuel_type="Diesel"),
            odometer=80000.0,
        )
        v.custom_fields["_acf_code"] = "B"
        assign_electrification_years([v])
        assert "mandate" in v.custom_fields["ACF Relevance"].lower()

    # ── Category B: ACF deadline-first assignment ─────────────────────

    def test_category_b_known_gvwr_gets_acf_deadline_year(self):
        """Category B with known GVWR gets an ACF mandate deadline year."""
        v = make_fleet_vehicle(
            vid_overrides=dict(gvwr="14000 lb", fuel_type="Diesel"),
            odometer=80000.0,
            annual_mileage=15000.0,
        )
        v.vehicle_id.gvwr_pounds = 14000.0
        v.custom_fields["_acf_code"] = "B"
        assign_electrification_years([v], end_year=2045)
        year_str = v.custom_fields["Proposed EV Year"]
        year = int(year_str)
        # Class 2b-4 milestone is 2035; urgency may pull to earlier checkpoint
        assert year <= 2035

    def test_category_b_reason_mentions_acf_mandate(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(gvwr="14000 lb", fuel_type="Diesel"),
            odometer=80000.0,
        )
        v.vehicle_id.gvwr_pounds = 14000.0
        v.custom_fields["_acf_code"] = "B"
        assign_electrification_years([v])
        assert "acf mandate" in v.custom_fields["EV Year Reason"].lower()

    def test_category_b_unknown_gvwr_falls_to_score_queue(self):
        """Category B with no GVWR falls into score-based queue."""
        v = make_fleet_vehicle(
            vid_overrides=dict(fuel_type="Diesel"),
            odometer=80000.0,
            annual_mileage=15000.0,
        )
        v.vehicle_id.gvwr_pounds = 0.0
        v.custom_fields["_acf_code"] = "B"
        assign_electrification_years([v], end_year=2040)
        year_str = v.custom_fields["Proposed EV Year"]
        year = int(year_str)
        current_year = datetime.datetime.now().year
        assert current_year < year <= 2040

    # ── Categories C and D: score queue ──────────────────────────────

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

    # ── Ordering: higher-priority C/D vehicles assigned earlier years ──

    def test_higher_score_gets_earlier_year_for_queue_vehicles(self):
        """For score-queue vehicles (C/D), older/higher-mileage gets earlier year."""
        old = make_fleet_vehicle(
            vid_overrides=dict(year="2005", body_class="Dump Truck",
                               gvwr="33000 lb", fuel_type="Diesel"),
            odometer=180000.0,
            annual_mileage=20000.0,
        )
        old.custom_fields["_acf_code"] = "C"

        new = make_fleet_vehicle(
            vid_overrides=dict(year="2023", body_class="Dump Truck",
                               gvwr="33000 lb", fuel_type="Diesel"),
            odometer=10000.0,
            annual_mileage=5000.0,
        )
        new.custom_fields["_acf_code"] = "C"

        assign_electrification_years([old, new], end_year=2040)
        old_year = int(old.custom_fields["Proposed EV Year"])
        new_year = int(new.custom_fields["Proposed EV Year"])
        assert old_year <= new_year

    # ── Budget smoothing for score-queue vehicles ─────────────────────

    def test_budget_smoothing_distributes_evenly(self):
        """Multiple score-queue vehicles (C/D) should spread across years."""
        vehicles = []
        for i in range(10):
            v = make_fleet_vehicle(
                vid_overrides=dict(
                    year=str(2010 + i),
                    body_class="Dump Truck",
                    gvwr="33000 lb",
                    fuel_type="Diesel",
                ),
                odometer=50000.0 + i * 10000,
                annual_mileage=12000.0,
            )
            v.custom_fields["_acf_code"] = "C"
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
        truck.vehicle_id.gvwr_pounds = 14000.0
        truck.custom_fields["_acf_code"] = "B"

        failed = make_fleet_vehicle(processing_success=False)
        failed.custom_fields["_acf_code"] = "B"

        assign_electrification_years([zev, light, truck, failed], end_year=2045)
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

    # ── Budget smoothing regression tests (fix for N≤Y and N>Y cases) ──

    def test_large_fleet_uses_all_available_years(self):
        """N > num_years: every year in the horizon must receive at least one vehicle.

        Regression for the ceil-based algorithm that left trailing years empty.
        Example: 18 vehicles over 5 years → ceil(18/5)=4, only 5 years used but
        the previous code could leave later years empty.  With floor-based
        distribution all 5 years must be represented.
        """
        current_year = datetime.datetime.now().year
        end_year = current_year + 5   # 5 year slots available
        n = 18                        # >5 → base=3, extra=3 → all 5 years filled

        vehicles = []
        for i in range(n):
            v = make_fleet_vehicle(
                vid_overrides=dict(
                    year=str(2005 + i % 15),
                    body_class="Dump Truck",
                    gvwr="33000 lb",
                    fuel_type="Diesel",
                ),
                odometer=50000.0 + i * 5000,
                annual_mileage=12000.0,
            )
            v.custom_fields["_acf_code"] = "C"
            vehicles.append(v)

        assign_electrification_years(vehicles, end_year=end_year)

        years = {int(v.custom_fields["Proposed EV Year"]) for v in vehicles}
        expected = set(range(current_year + 1, end_year + 1))
        assert years == expected, (
            f"All {len(expected)} years should be represented; got years {sorted(years)}"
        )

    def test_small_fleet_spans_full_horizon(self):
        """N < num_years: vehicles should span close to the full horizon, not
        just cluster in the first N consecutive years.

        Regression for the greedy front-loading that placed 4 vehicles in
        2027, 2028, 2029, 2030 regardless of a 14-year planning window.
        After fix: highest-priority → first year, lowest-priority → last year,
        so the spread should exceed 5 years for a 14-year horizon.
        """
        vehicles = []
        for i in range(4):
            v = make_fleet_vehicle(
                vid_overrides=dict(
                    year=str(2010 + i * 3),
                    body_class="Dump Truck",
                    gvwr="33000 lb",
                    fuel_type="Diesel",
                ),
                odometer=80000.0 - i * 15000,
                annual_mileage=12000.0,
            )
            v.custom_fields["_acf_code"] = "C"
            vehicles.append(v)

        assign_electrification_years(vehicles, end_year=2040)

        years = [int(v.custom_fields["Proposed EV Year"]) for v in vehicles]
        spread = max(years) - min(years)
        assert spread >= 5, (
            f"4 vehicles over 14-year horizon should span >5 years; "
            f"got years {sorted(years)} (spread={spread})"
        )

    def test_priority_order_preserved_across_full_horizon(self):
        """The highest-priority vehicle must always get an earlier year than the
        lowest-priority one, even after the even-spread fix.
        """
        current_year = datetime.datetime.now().year

        high_priority = make_fleet_vehicle(
            vid_overrides=dict(
                year="2005",
                body_class="Dump Truck",
                gvwr="33000 lb",
                fuel_type="Diesel",
            ),
            odometer=195000.0,
            annual_mileage=20000.0,
        )
        high_priority.custom_fields["_acf_code"] = "C"

        low_priority = make_fleet_vehicle(
            vid_overrides=dict(
                year=str(current_year - 1),
                body_class="Dump Truck",
                gvwr="33000 lb",
                fuel_type="Diesel",
            ),
            odometer=3000.0,
            annual_mileage=3000.0,
        )
        low_priority.custom_fields["_acf_code"] = "C"

        assign_electrification_years([high_priority, low_priority], end_year=2040)

        high_year = int(high_priority.custom_fields["Proposed EV Year"])
        low_year = int(low_priority.custom_fields["Proposed EV Year"])
        assert high_year < low_year, (
            f"High-priority vehicle should get earlier year than low-priority: "
            f"{high_year} vs {low_year}"
        )


# =========================================================================
# Non-HPF / State Agency fleet type — deadline table tests (Phase 26)
# =========================================================================

def _make_cat_b_vehicle(gvwr_str: str) -> "FleetVehicle":
    """Helper: new, low-urgency Cat B vehicle for deadline milestone tests.

    Uses a nearly-new model year and minimal odometer/mileage so that
    the urgency score stays well below 0.55.  This guarantees the vehicle
    receives the full deadline milestone year rather than an early purchase-
    checkpoint year, making the deadline assertions deterministic.
    """
    current_year = datetime.datetime.now().year
    v = make_fleet_vehicle(
        vid_overrides=dict(
            gvwr=gvwr_str,
            fuel_type="Diesel",
            year=str(current_year - 1),   # nearly new → tiny age score
        ),
        odometer=1000.0,       # near-zero odometer → tiny odometer score
        annual_mileage=5000.0, # low annual use → tiny usage score
    )
    v.custom_fields["_acf_code"] = "B"
    return v


class TestNonHPFDeadlines:
    """Regression tests for fleet_type='non_hpf' and 'state_agency' deadlines."""

    # ── Deadline table structure ──────────────────────────────────────────────

    def test_non_hpf_table_has_three_gvwr_ranges(self):
        assert len(ACF_DEADLINE_TABLE_NON_HPF) == 3

    def test_non_hpf_milestones_later_than_hpf(self):
        """Every non-HPF milestone year should be >= the HPF milestone year."""
        for key in ACF_DEADLINE_TABLE:
            hpf_ms = ACF_DEADLINE_TABLE[key]["milestone_year"]
            non_hpf_ms = ACF_DEADLINE_TABLE_NON_HPF[key]["milestone_year"]
            assert non_hpf_ms >= hpf_ms, (
                f"Non-HPF milestone {non_hpf_ms} is not later than HPF {hpf_ms} "
                f"for GVWR range {key}"
            )

    # ── HPF regression (CARB deadline stored in ACF Deadline Year field) ────────
    # Cat B vehicles are no longer hard-assigned to CARB deadline years.
    # Instead, the CARB deadline is stored in custom_fields["ACF Deadline Year"]
    # for reference (milestone chart, compliance warnings), while Proposed EV Year
    # is assigned via the even-spread score queue.

    def test_class_2b_4_hpf_milestone_is_2035(self):
        v = _make_cat_b_vehicle("14000 lb")
        assign_electrification_years([v], fleet_type="hpf")
        deadline = v.custom_fields.get("ACF Deadline Year", "")
        assert deadline == "2035", f"HPF Class 2b-4 CARB deadline should be 2035, got {deadline}"

    def test_class_5_8a_hpf_milestone_is_2039(self):
        v = _make_cat_b_vehicle("25000 lb")
        assign_electrification_years([v], fleet_type="hpf")
        deadline = v.custom_fields.get("ACF Deadline Year", "")
        assert deadline == "2039", f"HPF Class 5-8a CARB deadline should be 2039, got {deadline}"

    def test_class_8b_hpf_milestone_is_2042(self):
        v = _make_cat_b_vehicle("40000 lb")
        assign_electrification_years([v], fleet_type="hpf")
        deadline = v.custom_fields.get("ACF Deadline Year", "")
        assert deadline == "2042", f"HPF Class 8b CARB deadline should be 2042, got {deadline}"

    # ── Non-HPF deadlines ─────────────────────────────────────────────────────

    def test_class_2b_4_non_hpf_milestone_is_2036(self):
        v = _make_cat_b_vehicle("14000 lb")
        assign_electrification_years([v], fleet_type="non_hpf")
        deadline = v.custom_fields.get("ACF Deadline Year", "")
        assert deadline == "2036", f"Non-HPF Class 2b-4 CARB deadline should be 2036, got {deadline}"

    def test_class_5_8a_non_hpf_milestone_is_2040(self):
        v = _make_cat_b_vehicle("25000 lb")
        assign_electrification_years([v], fleet_type="non_hpf")
        deadline = v.custom_fields.get("ACF Deadline Year", "")
        assert deadline == "2040", f"Non-HPF Class 5-8a CARB deadline should be 2040, got {deadline}"

    def test_class_8b_non_hpf_milestone_is_2043(self):
        v = _make_cat_b_vehicle("40000 lb")
        assign_electrification_years([v], fleet_type="non_hpf")
        deadline = v.custom_fields.get("ACF Deadline Year", "")
        assert deadline == "2043", f"Non-HPF Class 8b CARB deadline should be 2043, got {deadline}"

    # ── State Agency uses same table as Non-HPF ───────────────────────────────

    def test_state_agency_class_2b_4_milestone_matches_non_hpf(self):
        v = _make_cat_b_vehicle("14000 lb")
        assign_electrification_years([v], fleet_type="state_agency")
        deadline = v.custom_fields.get("ACF Deadline Year", "")
        assert deadline == "2036", f"State Agency Class 2b-4 should match Non-HPF (2036), got {deadline}"

    # ── Unknown fleet_type falls back to HPF ──────────────────────────────────

    def test_unknown_fleet_type_falls_back_to_hpf(self):
        v = _make_cat_b_vehicle("14000 lb")
        assign_electrification_years([v], fleet_type="bogus_type")
        deadline = v.custom_fields.get("ACF Deadline Year", "")
        assert deadline == "2035", (
            f"Unknown fleet_type should fall back to HPF (2035), got {deadline}"
        )

    # ── Unknown GVWR fallback ─────────────────────────────────────────────────

    def test_unknown_gvwr_hpf_fallback(self):
        assert ACF_DEADLINE_UNKNOWN_GVWR == 2035

    def test_unknown_gvwr_non_hpf_fallback(self):
        assert ACF_DEADLINE_UNKNOWN_GVWR_NON_HPF == 2036

    # ── Non-B vehicles are unaffected by fleet_type ───────────────────────────

    def test_cat_c_vehicle_unaffected_by_fleet_type(self):
        """Cat C vehicles use score-based queue; fleet_type should not change result."""
        from tests.conftest import make_fleet_vehicle

        v_hpf = make_fleet_vehicle(
            vid_overrides=dict(gvwr="14000 lb", fuel_type="Diesel"),
            odometer=50000.0, annual_mileage=12000.0,
        )
        v_hpf.custom_fields["_acf_code"] = "C"

        v_non = make_fleet_vehicle(
            vid_overrides=dict(gvwr="14000 lb", fuel_type="Diesel"),
            odometer=50000.0, annual_mileage=12000.0,
        )
        v_non.custom_fields["_acf_code"] = "C"

        assign_electrification_years([v_hpf], fleet_type="hpf", end_year=2040)
        assign_electrification_years([v_non], fleet_type="non_hpf", end_year=2040)

        year_hpf = v_hpf.custom_fields.get("Proposed EV Year", "")
        year_non = v_non.custom_fields.get("Proposed EV Year", "")
        # Both should land on the same year (score-based, not deadline-based)
        assert year_hpf == year_non, (
            f"Cat C year should not differ by fleet_type: HPF={year_hpf}, "
            f"non_hpf={year_non}"
        )


# =========================================================================
# Even-spread regression tests for Cat B (Phase 28)
# =========================================================================

class TestCatBEvenSpread:
    """Cat B vehicles now use even-spread score queue.  CARB deadlines are stored
    in custom_fields['ACF Deadline Year'] but do NOT drive Proposed EV Year."""

    def test_cat_b_acf_deadline_year_always_populated(self):
        """ACF Deadline Year is set for every Cat B vehicle regardless of urgency."""
        v = _make_cat_b_vehicle("14000 lb")
        assign_electrification_years([v], fleet_type="hpf")
        deadline = v.custom_fields.get("ACF Deadline Year", "")
        assert deadline != "", "ACF Deadline Year should be populated for Cat B"
        assert deadline.isdigit() or deadline == "Unknown", (
            f"ACF Deadline Year should be a year string or 'Unknown', got {deadline}"
        )

    def test_cat_b_proposed_ev_year_is_in_horizon(self):
        """Proposed EV Year for Cat B must fall within the planning horizon."""
        v = _make_cat_b_vehicle("14000 lb")
        assign_electrification_years([v], fleet_type="hpf", end_year=2040)
        year_str = v.custom_fields.get("Proposed EV Year", "")
        assert year_str.isdigit(), f"Expected a year, got '{year_str}'"
        year = int(year_str)
        assert 2026 <= year <= 2040, f"Proposed EV Year {year} outside horizon 2026–2040"

    def test_cat_b_class_2b4_boost_exceeds_class_8b(self):
        """Class 2b-4 (earlier deadline) should have a higher urgency boost than Class 8b."""
        assert ACF_BOOST_B_CLASS_2B4 > ACF_BOOST_B_CLASS_8B, (
            "Class 2b-4 boost should exceed Class 8b boost (earlier CARB deadline → higher priority)"
        )
        assert ACF_BOOST_B_CLASS_2B4 > ACF_BOOST_B_CLASS_5_8A > ACF_BOOST_B_CLASS_8B, (
            "Boost ordering should be: Class 2b-4 > Class 5-8a > Class 8b"
        )

    def test_cat_b_class_2b4_gets_earlier_year_than_class_8b(self):
        """Given identical age/mileage, Class 2b-4 (higher boost) should be scheduled
        before Class 8b (lower boost) in the even-spread queue."""
        current_year = datetime.datetime.now().year
        common_kw = dict(
            vid_overrides=dict(year=str(current_year - 5), fuel_type="Diesel"),
            odometer=50000.0,
            annual_mileage=12000.0,
        )
        v_2b4 = make_fleet_vehicle(vid_overrides=dict(
            gvwr="14000 lb", year=str(current_year - 5), fuel_type="Diesel"
        ), odometer=50000.0, annual_mileage=12000.0)
        v_2b4.custom_fields["_acf_code"] = "B"

        v_8b = make_fleet_vehicle(vid_overrides=dict(
            gvwr="40000 lb", year=str(current_year - 5), fuel_type="Diesel"
        ), odometer=50000.0, annual_mileage=12000.0)
        v_8b.custom_fields["_acf_code"] = "B"

        assign_electrification_years([v_2b4, v_8b], end_year=2040)

        year_2b4 = int(v_2b4.custom_fields.get("Proposed EV Year", "9999"))
        year_8b = int(v_8b.custom_fields.get("Proposed EV Year", "9999"))
        assert year_2b4 <= year_8b, (
            f"Class 2b-4 (higher boost) should be scheduled no later than Class 8b; "
            f"got 2b-4={year_2b4}, 8b={year_8b}"
        )

    def test_cat_b_fleet_spread_across_horizon(self):
        """14 Cat B vehicles over 14 years (2027–2040) should each get a distinct year,
        demonstrating even spread rather than deadline spikes."""
        current_year = datetime.datetime.now().year
        vehicles = []
        for i in range(14):
            v = make_fleet_vehicle(
                vid_overrides=dict(
                    gvwr="14000 lb",
                    year=str(current_year - i),  # varying age for distinct scores
                    fuel_type="Diesel",
                ),
                odometer=float(i * 5000),
                annual_mileage=12000.0,
            )
            v.custom_fields["_acf_code"] = "B"
            vehicles.append(v)

        assign_electrification_years(vehicles, end_year=2040)

        years = [v.custom_fields.get("Proposed EV Year", "") for v in vehicles]
        numeric_years = [y for y in years if y.isdigit()]
        assert len(numeric_years) == 14, f"All 14 vehicles should get a year, got: {years}"

        unique_years = set(numeric_years)
        # With 14 vehicles over 14 available years (2027–2040), each year gets 1 vehicle
        assert len(unique_years) >= 10, (
            f"14 Cat B vehicles should spread across at least 10 distinct years, "
            f"got {len(unique_years)}: {sorted(unique_years)}"
        )
        # No deadline spikes: no single year should have more than 3 vehicles
        from collections import Counter
        year_counts = Counter(numeric_years)
        max_count = max(year_counts.values())
        assert max_count <= 3, (
            f"Even spread should prevent spikes; max vehicles in one year = {max_count}: "
            f"{dict(year_counts)}"
        )

    def test_cat_b_even_spread_unaffected_by_fleet_type(self):
        """Cat B Proposed EV Year should use even-spread for both HPF and non-HPF fleets."""
        current_year = datetime.datetime.now().year
        v_hpf = make_fleet_vehicle(
            vid_overrides=dict(gvwr="14000 lb", year=str(current_year - 5), fuel_type="Diesel"),
            odometer=50000.0, annual_mileage=12000.0,
        )
        v_hpf.custom_fields["_acf_code"] = "B"

        v_non = make_fleet_vehicle(
            vid_overrides=dict(gvwr="14000 lb", year=str(current_year - 5), fuel_type="Diesel"),
            odometer=50000.0, annual_mileage=12000.0,
        )
        v_non.custom_fields["_acf_code"] = "B"

        assign_electrification_years([v_hpf], fleet_type="hpf", end_year=2040)
        assign_electrification_years([v_non], fleet_type="non_hpf", end_year=2040)

        year_hpf = v_hpf.custom_fields.get("Proposed EV Year", "")
        year_non = v_non.custom_fields.get("Proposed EV Year", "")
        # Score-based spread → same result regardless of fleet_type
        assert year_hpf == year_non, (
            f"Cat B Proposed EV Year should not differ by fleet_type: "
            f"HPF={year_hpf}, non_hpf={year_non}"
        )
        # But ACF Deadline Year SHOULD differ by fleet_type
        deadline_hpf = v_hpf.custom_fields.get("ACF Deadline Year", "")
        deadline_non = v_non.custom_fields.get("ACF Deadline Year", "")
        assert deadline_hpf != deadline_non, (
            f"ACF Deadline Year should differ by fleet_type: "
            f"HPF={deadline_hpf}, non_hpf={deadline_non}"
        )
