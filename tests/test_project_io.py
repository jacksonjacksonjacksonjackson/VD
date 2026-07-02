"""
Tests for data/project_io.py — project save/load (.fea files).

Covers round-trip serialization of Fleet objects including nested
VehicleIdentification, FuelEconomyData, custom_fields (ACF codes,
EV year assignments), datetime fields, and scenario_results.
"""

import datetime
import json
import os
import sys
import tempfile

import pytest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from tests.conftest import make_fleet_vehicle
from data.models import Fleet
from data.project_io import load_project, save_project


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_fleet(num_vehicles: int = 3, fleet_type: str = "hpf") -> Fleet:
    vehicles = []
    for i in range(num_vehicles):
        vin = f"1HGBH41JXM{i:07d}"
        v = make_fleet_vehicle(
            vin=vin,
            vid_overrides={"vin": vin, "year": str(2015 + i), "make": "Ford", "model": "F-150"},
            fuel_overrides={"combined_mpg": 20.0 + i, "co2_primary": 400.0},
            custom_fields={
                "_acf_code": "B",
                "ACF Category": "Category B — Mandate-Subject",
                "Proposed EV Year": str(2028 + i),
                "_ev_year_overridden": False,
            },
            acquisition_date=datetime.date(2015 + i, 6, 1),
            odometer=float(50_000 + i * 10_000),
        )
        vehicles.append(v)
    return Fleet(
        name="Test Fleet",
        vehicles=vehicles,
        fleet_type=fleet_type,
        max_vehicles_per_year=5,
    )


# ---------------------------------------------------------------------------
# Round-trip tests
# ---------------------------------------------------------------------------

class TestRoundTrip:
    def test_basic_round_trip(self, tmp_path):
        fleet = _make_fleet(3)
        path = str(tmp_path / "test.fea")
        save_project(path, fleet)
        loaded_fleet, scenario_results = load_project(path)

        assert len(loaded_fleet.vehicles) == 3
        assert loaded_fleet.name == fleet.name
        assert loaded_fleet.fleet_type == fleet.fleet_type
        assert loaded_fleet.max_vehicles_per_year == fleet.max_vehicles_per_year
        assert scenario_results is None

    def test_vehicle_fields_preserved(self, tmp_path):
        fleet = _make_fleet(1)
        original = fleet.vehicles[0]
        path = str(tmp_path / "test.fea")
        save_project(path, fleet)
        loaded, _ = load_project(path)
        v = loaded.vehicles[0]

        assert v.vin == original.vin
        assert v.vehicle_id.year == original.vehicle_id.year
        assert v.vehicle_id.make == original.vehicle_id.make
        assert v.vehicle_id.model == original.vehicle_id.model
        assert v.fuel_economy.combined_mpg == original.fuel_economy.combined_mpg
        assert v.odometer == original.odometer

    def test_acf_codes_preserved(self, tmp_path):
        fleet = _make_fleet(2)
        path = str(tmp_path / "test.fea")
        save_project(path, fleet)
        loaded, _ = load_project(path)

        for orig, v in zip(fleet.vehicles, loaded.vehicles):
            assert v.custom_fields.get("_acf_code") == orig.custom_fields["_acf_code"]
            assert v.custom_fields.get("Proposed EV Year") == orig.custom_fields["Proposed EV Year"]

    def test_datetime_date_preserved(self, tmp_path):
        fleet = _make_fleet(1)
        fleet.vehicles[0].acquisition_date = datetime.date(2019, 3, 15)
        path = str(tmp_path / "test.fea")
        save_project(path, fleet)
        loaded, _ = load_project(path)
        assert loaded.vehicles[0].acquisition_date == datetime.date(2019, 3, 15)

    def test_last_quality_check_datetime_preserved(self, tmp_path):
        fleet = _make_fleet(1)
        dt = datetime.datetime(2025, 1, 20, 14, 30, 0)
        fleet.vehicles[0].last_quality_check = dt
        path = str(tmp_path / "test.fea")
        save_project(path, fleet)
        loaded, _ = load_project(path)
        assert loaded.vehicles[0].last_quality_check == dt

    def test_none_dates_preserved(self, tmp_path):
        fleet = _make_fleet(1)
        fleet.vehicles[0].acquisition_date = None
        fleet.vehicles[0].last_quality_check = None
        path = str(tmp_path / "test.fea")
        save_project(path, fleet)
        loaded, _ = load_project(path)
        assert loaded.vehicles[0].acquisition_date is None
        assert loaded.vehicles[0].last_quality_check is None

    def test_scenario_results_round_trip(self, tmp_path):
        fleet = _make_fleet(2)
        fake_scenario_results = {
            "scenarios": [{"name": "Aggressive 2030", "vehicles_per_year": {2026: 1}}],
            "all_years": [2026, 2027, 2028],
            "best_roi": "Aggressive 2030",
        }
        path = str(tmp_path / "test.fea")
        save_project(path, fleet, scenario_results=fake_scenario_results)
        _, scenario_results = load_project(path)

        assert scenario_results is not None
        assert scenario_results["best_roi"] == "Aggressive 2030"
        assert scenario_results["all_years"] == [2026, 2027, 2028]

    def test_fleet_type_options(self, tmp_path):
        for fleet_type in ("hpf", "non_hpf", "state_agency"):
            fleet = _make_fleet(1, fleet_type=fleet_type)
            path = str(tmp_path / f"test_{fleet_type}.fea")
            save_project(path, fleet)
            loaded, _ = load_project(path)
            assert loaded.fleet_type == fleet_type

    def test_empty_fleet(self, tmp_path):
        fleet = Fleet(name="Empty")
        path = str(tmp_path / "empty.fea")
        save_project(path, fleet)
        loaded, _ = load_project(path)
        assert loaded.vehicles == []
        assert loaded.name == "Empty"

    def test_file_is_valid_json(self, tmp_path):
        fleet = _make_fleet(1)
        path = str(tmp_path / "test.fea")
        save_project(path, fleet)
        with open(path) as f:
            data = json.load(f)
        assert data["version"] == 1
        assert "fleet" in data
        assert len(data["fleet"]["vehicles"]) == 1

    def test_large_fleet(self, tmp_path):
        fleet = _make_fleet(50)
        path = str(tmp_path / "large.fea")
        save_project(path, fleet)
        loaded, _ = load_project(path)
        assert len(loaded.vehicles) == 50

    def test_custom_fields_unicode(self, tmp_path):
        fleet = _make_fleet(1)
        fleet.vehicles[0].custom_fields["Department"] = "Public Works — Sanitation"
        path = str(tmp_path / "unicode.fea")
        save_project(path, fleet)
        loaded, _ = load_project(path)
        assert loaded.vehicles[0].custom_fields["Department"] == "Public Works — Sanitation"

    def test_ev_year_override_flag_preserved(self, tmp_path):
        fleet = _make_fleet(1)
        fleet.vehicles[0].custom_fields["_ev_year_overridden"] = True
        path = str(tmp_path / "override.fea")
        save_project(path, fleet)
        loaded, _ = load_project(path)
        assert loaded.vehicles[0].custom_fields.get("_ev_year_overridden") is True


# ---------------------------------------------------------------------------
# Error handling
# ---------------------------------------------------------------------------

class TestErrorHandling:
    def test_missing_file_raises(self, tmp_path):
        with pytest.raises((FileNotFoundError, IOError)):
            load_project(str(tmp_path / "nonexistent.fea"))

    def test_corrupt_json_raises(self, tmp_path):
        path = str(tmp_path / "corrupt.fea")
        with open(path, "w") as f:
            f.write("not valid json {{{")
        with pytest.raises(json.JSONDecodeError):
            load_project(path)

    def test_future_version_raises(self, tmp_path):
        path = str(tmp_path / "future.fea")
        with open(path, "w") as f:
            json.dump({"version": 999, "fleet": {"vehicles": [], "name": "x"}}, f)
        with pytest.raises(ValueError, match="newer than this app"):
            load_project(path)
