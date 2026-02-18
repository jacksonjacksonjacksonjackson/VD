"""Tests for CARB ACF compliance classification in analysis/acf_compliance.py."""

import pytest
from analysis.acf_compliance import classify_acf_vehicle
from tests.conftest import make_fleet_vehicle


# =========================================================================
# ZEV Detection
# =========================================================================

class TestZevClassification:

    def test_bev_classified_as_zev(self, electric_vehicle):
        code, label, detail = classify_acf_vehicle(electric_vehicle)
        assert code == "ZEV"
        assert "zero" in label.lower() or "zev" in label.lower()

    def test_fuel_cell_classified_as_zev(self):
        v = make_fleet_vehicle(vid_overrides=dict(fuel_type="Hydrogen Fuel Cell"))
        code, _, _ = classify_acf_vehicle(v)
        assert code == "ZEV"


# =========================================================================
# Category A — Light-Duty Exempt
# =========================================================================

class TestLightDutyExempt:

    def test_sedan_is_exempt(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(
                body_class="Sedan/Saloon",
                gvwr="4500 lb",
                fuel_type="Gasoline",
                model="Camry",
                make="Toyota",
            )
        )
        code, label, _ = classify_acf_vehicle(v)
        assert code == "A"
        assert "exempt" in label.lower()

    def test_light_suv_is_exempt(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(
                body_class="Sport Utility Vehicle (SUV)",
                gvwr="6000 lb",
                fuel_type="Gasoline",
                model="Explorer",
            )
        )
        code, _, _ = classify_acf_vehicle(v)
        assert code == "A"

    def test_boundary_8500_is_light_duty(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(gvwr="8500 lb", fuel_type="Gasoline")
        )
        code, _, _ = classify_acf_vehicle(v)
        assert code == "A"


# =========================================================================
# Category B — Subject to ACF
# =========================================================================

class TestSubjectToAcf:

    def test_medium_duty_truck(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(
                body_class="Truck",
                gvwr="14000 lb",
                fuel_type="Gasoline",
                model="F-450",
            ),
            department="Public Works",
        )
        code, label, _ = classify_acf_vehicle(v)
        assert code == "B"
        assert "subject" in label.lower() or "acf" in label.lower()

    def test_heavy_duty_truck(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(
                body_class="Cargo Van",
                gvwr="26000 lb",
                fuel_type="Diesel",
                model="Box Truck",
            ),
        )
        code, _, _ = classify_acf_vehicle(v)
        assert code == "B"


# =========================================================================
# Category C — Exempt Body Type
# =========================================================================

class TestExemptBodyType:

    def test_dump_truck_exempt(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(
                body_class="Dump Truck",
                gvwr="33000 lb",
                fuel_type="Diesel",
            ),
        )
        code, label, _ = classify_acf_vehicle(v)
        assert code == "C"
        assert "body" in label.lower()

    def test_concrete_mixer_exempt(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(
                body_class="Concrete Mixer",
                gvwr="40000 lb",
                fuel_type="Diesel",
            ),
        )
        code, _, _ = classify_acf_vehicle(v)
        assert code == "C"

    def test_crane_exempt(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(
                body_class="Crane",
                gvwr="50000 lb",
                fuel_type="Diesel",
            ),
        )
        code, _, _ = classify_acf_vehicle(v)
        assert code == "C"


# =========================================================================
# Category D — Emergency Vehicle
# =========================================================================

class TestEmergencyVehicle:

    def test_ppv_trim_detected(self, emergency_vehicle):
        """PPV (Police Pursuit Vehicle) is a strong signal."""
        code, label, _ = classify_acf_vehicle(emergency_vehicle)
        assert code == "D"
        assert "emergency" in label.lower()

    def test_ambulance_body_class(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(
                body_class="Ambulance",
                gvwr="14000 lb",
                fuel_type="Gasoline",
            ),
            department="EMS",
        )
        code, _, _ = classify_acf_vehicle(v)
        assert code == "D"

    def test_fire_apparatus_make(self):
        v = make_fleet_vehicle(
            vid_overrides=dict(
                make="Pierce",
                model="Velocity",
                body_class="Fire Apparatus",
                gvwr="40000 lb",
                fuel_type="Diesel",
            ),
            department="Fire Department",
        )
        code, _, _ = classify_acf_vehicle(v)
        assert code == "D"

    # ── False positive prevention (Fix 22 regression) ─────────────

    def test_crossfire_not_emergency(self):
        """'Crossfire' should not match 'fire'."""
        v = make_fleet_vehicle(
            vid_overrides=dict(
                model="Crossfire",
                make="Chrysler",
                body_class="Coupe",
                gvwr="3500 lb",
                fuel_type="Gasoline",
            ),
            department="General Services",
        )
        code, _, _ = classify_acf_vehicle(v)
        assert code != "D"

    def test_nissan_patrol_not_emergency(self):
        """Nissan 'Patrol' civilian SUV — weak signal without emergency department."""
        v = make_fleet_vehicle(
            vid_overrides=dict(
                model="Patrol",
                make="Nissan",
                body_class="Sport Utility Vehicle (SUV)",
                gvwr="7000 lb",
                fuel_type="Gasoline",
            ),
            department="Fleet Services",
        )
        code, _, _ = classify_acf_vehicle(v)
        assert code != "D"

    def test_weak_keyword_with_emergency_dept_triggers(self):
        """'Police' model in a Police Department should trigger D.

        Note: GVWR must exceed 8,500 lbs or the vehicle will be classified
        as Category A (light-duty exempt) before emergency detection runs.
        """
        v = make_fleet_vehicle(
            vid_overrides=dict(
                model="Police Interceptor",
                make="Ford",
                body_class="Sedan/Saloon",
                gvwr="10000 lb",
                fuel_type="Gasoline",
            ),
            department="Police Department",
        )
        code, _, _ = classify_acf_vehicle(v)
        assert code == "D"
