"""
tests/test_vehicle_database.py

Unit tests for VehicleDatabaseManager — the Tier 3 SQLite MPG reference lookup.

Coverage targets:
  - CRUD: add, update (allow-list enforcement), delete
  - 4-tier lookup fallback (exact → any-fuel → null-year → GVWR range)
  - NULL-year UNIQUE constraint handling (explicit pre-delete)
  - Search filters

Each test uses an isolated in-memory-style SQLite file via pytest's tmp_path fixture
so tests never touch the real vehicle_database.db.

Run: python -m pytest tests/test_vehicle_database.py -v
"""

import sys
import os
import pytest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from pathlib import Path
from data.vehicle_database import VehicleDatabaseManager


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _db(tmp_path: Path) -> VehicleDatabaseManager:
    """Open a fresh VehicleDatabaseManager at a temp path."""
    return VehicleDatabaseManager(tmp_path / "test_vehicles.db")


def _make_vehicle_stub(make="ford", model="f-250", year=2020, fuel_type="diesel",
                       gvwr_pounds=None):
    """Return a minimal duck-typed vehicle object accepted by lookup_mpg()."""
    class _VID:
        pass

    class _FE:
        pass

    class _V:
        pass

    vid = _VID()
    vid.make = make
    vid.model = model
    vid.year = str(year) if year else None
    vid.fuel_type = fuel_type
    vid.gvwr_pounds = gvwr_pounds or 0

    v = _V()
    v.vehicle_id = vid
    v.fuel_economy = _FE()
    return v


# ---------------------------------------------------------------------------
# Group 1: Basic CRUD
# ---------------------------------------------------------------------------

class TestCrud:

    def test_add_and_retrieve_vehicle(self, tmp_path):
        """Adding a vehicle must make it appear in get_all_ice_vehicles()."""
        db = _db(tmp_path)
        row_id = db.add_ice_vehicle("Ford", "F-250", mpg_combined=15.0, year=2020,
                                    fuel_type="Diesel", source="test")
        assert row_id > 0

        all_rows = db.get_all_ice_vehicles()
        assert len(all_rows) == 1
        row = all_rows[0]
        assert row["make"] == "ford"          # normalised to lowercase
        assert row["mpg_combined"] == 15.0
        assert row["year"] == 2020
        db.close()

    def test_delete_vehicle(self, tmp_path):
        """Deleting a row by id must remove it; second delete returns False."""
        db = _db(tmp_path)
        row_id = db.add_ice_vehicle("Ford", "F-350", mpg_combined=13.0, year=2019)
        assert db.delete_ice_vehicle(row_id) is True
        assert db.get_all_ice_vehicles() == []
        assert db.delete_ice_vehicle(row_id) is False   # already gone
        db.close()

    def test_update_vehicle_allowed_column(self, tmp_path):
        """update_ice_vehicle must accept columns in _ALLOWED_UPDATE_COLS."""
        db = _db(tmp_path)
        row_id = db.add_ice_vehicle("Ford", "Transit", mpg_combined=18.0, year=2021)
        result = db.update_ice_vehicle(row_id, mpg_combined=19.5)
        assert result is True
        row = db.get_all_ice_vehicles()[0]
        assert row["mpg_combined"] == 19.5
        db.close()

    def test_update_vehicle_disallowed_column(self, tmp_path):
        """update_ice_vehicle must silently reject columns outside the allow-list
        (e.g., 'id') and return False when no valid columns remain."""
        db = _db(tmp_path)
        row_id = db.add_ice_vehicle("Ford", "Ranger", mpg_combined=22.0, year=2022)
        result = db.update_ice_vehicle(row_id, id=999)    # 'id' not in allow-list
        assert result is False
        # Original row untouched
        row = db.get_all_ice_vehicles()[0]
        assert row["id"] == row_id
        db.close()

    def test_update_nonexistent_row(self, tmp_path):
        """Updating a row that does not exist must return False."""
        db = _db(tmp_path)
        result = db.update_ice_vehicle(9999, mpg_combined=10.0)
        assert result is False
        db.close()


# ---------------------------------------------------------------------------
# Group 2: 4-tier Lookup Fallback
# ---------------------------------------------------------------------------

class TestLookupFallback:

    def test_tier1_exact_year_make_model_fuel(self, tmp_path):
        """Tier 1: exact year + make + model + fuel_type returns the right MPG."""
        db = _db(tmp_path)
        db.add_ice_vehicle("Ford", "F-250", mpg_combined=15.0, year=2020,
                           fuel_type="Diesel")
        # Add a different fuel row to ensure fuel_type is matched, not confused
        db.add_ice_vehicle("Ford", "F-250", mpg_combined=20.0, year=2020,
                           fuel_type="Gasoline")

        v = _make_vehicle_stub(make="ford", model="f-250", year=2020, fuel_type="diesel")
        result = db.lookup_mpg(v)
        assert result is not None
        assert result["combined"] == 15.0
        db.close()

    def test_tier2_any_fuel_type(self, tmp_path):
        """Tier 2: if no exact fuel_type match, fall back to any fuel for same year/make/model."""
        db = _db(tmp_path)
        db.add_ice_vehicle("Ford", "F-250", mpg_combined=15.0, year=2020,
                           fuel_type="Diesel")

        # Lookup with gasoline — Tier 1 misses, Tier 2 should match
        v = _make_vehicle_stub(make="ford", model="f-250", year=2020, fuel_type="gasoline")
        result = db.lookup_mpg(v)
        assert result is not None
        assert result["combined"] == 15.0
        db.close()

    def test_tier3_null_year_catchall(self, tmp_path):
        """Tier 3: a NULL-year entry covers any year when Tiers 1–2 fail."""
        db = _db(tmp_path)
        db.add_ice_vehicle("Ford", "F-250", mpg_combined=14.0, year=None,
                           fuel_type="Diesel")

        # No exact-year row → should fall through to null-year catch-all
        v = _make_vehicle_stub(make="ford", model="f-250", year=2018, fuel_type="diesel")
        result = db.lookup_mpg(v)
        assert result is not None
        assert result["combined"] == 14.0
        db.close()

    def test_tier3_not_used_when_tier1_matches(self, tmp_path):
        """Tier 1 exact match must take precedence over Tier 3 null-year entry."""
        db = _db(tmp_path)
        db.add_ice_vehicle("Ford", "F-250", mpg_combined=14.0, year=None,
                           fuel_type="Diesel")
        db.add_ice_vehicle("Ford", "F-250", mpg_combined=16.0, year=2020,
                           fuel_type="Diesel")

        v = _make_vehicle_stub(make="ford", model="f-250", year=2020, fuel_type="diesel")
        result = db.lookup_mpg(v)
        assert result is not None
        assert result["combined"] == 16.0   # exact match wins
        db.close()

    def test_tier4_gvwr_range(self, tmp_path):
        """Tier 4: GVWR-range overlap returns a match when Tiers 1–3 all fail."""
        db = _db(tmp_path)
        db.add_ice_vehicle("Peterbilt", "567", mpg_combined=6.5, year=None,
                           fuel_type="Diesel",
                           gvwr_lbs_min=33001, gvwr_lbs_max=60000)

        # No year row and no null-year row → should fall through to GVWR range
        v = _make_vehicle_stub(make="peterbilt", model="567", year=2019,
                               fuel_type="diesel", gvwr_pounds=45000)
        result = db.lookup_mpg(v)
        assert result is not None
        assert result["combined"] == 6.5
        db.close()

    def test_tier4_gvwr_outside_range_returns_none(self, tmp_path):
        """A GVWR outside the stored min/max must NOT match.

        Uses a year=2019 row so Tiers 1–3 all fail for a year=2021 lookup, leaving
        only the Tier 4 GVWR range check to decide. 70000 lbs is outside 33001–60000.
        """
        db = _db(tmp_path)
        # Specific year — ensures Tier 3 (null-year) does not fire
        db.add_ice_vehicle("Peterbilt", "567", mpg_combined=6.5, year=2019,
                           fuel_type="Diesel",
                           gvwr_lbs_min=33001, gvwr_lbs_max=60000)

        # Different year lookup → Tiers 1/2 fail; no null-year row → Tier 3 fails;
        # GVWR 70000 outside 33001–60000 → Tier 4 fails → None
        v = _make_vehicle_stub(make="peterbilt", model="567", year=2021,
                               fuel_type="diesel", gvwr_pounds=70000)
        result = db.lookup_mpg(v)
        assert result is None
        db.close()

    def test_lookup_returns_none_when_no_match(self, tmp_path):
        """lookup_mpg must return None when no tier matches."""
        db = _db(tmp_path)
        v = _make_vehicle_stub(make="notamake", model="notamodel", year=2020)
        result = db.lookup_mpg(v)
        assert result is None
        db.close()

    def test_lookup_returns_none_for_empty_make(self, tmp_path):
        """lookup_mpg must return None immediately when make is blank."""
        db = _db(tmp_path)
        v = _make_vehicle_stub(make="", model="f-250", year=2020)
        result = db.lookup_mpg(v)
        assert result is None
        db.close()


# ---------------------------------------------------------------------------
# Group 3: NULL-year UNIQUE Constraint Handling
# ---------------------------------------------------------------------------

class TestNullYearUnique:

    def test_add_null_year_vehicle(self, tmp_path):
        """Inserting a NULL-year entry must succeed and appear in the database."""
        db = _db(tmp_path)
        row_id = db.add_ice_vehicle("Ford", "F-550", mpg_combined=12.0, year=None,
                                    fuel_type="Diesel")
        assert row_id > 0
        rows = db.get_all_ice_vehicles()
        assert len(rows) == 1
        assert rows[0]["year"] is None
        db.close()

    def test_add_duplicate_null_year_replaces_existing(self, tmp_path):
        """Inserting a second NULL-year row for the same make/model/fuel must
        replace the first (via the explicit pre-delete), leaving exactly one row."""
        db = _db(tmp_path)
        db.add_ice_vehicle("Ford", "F-550", mpg_combined=12.0, year=None,
                           fuel_type="Diesel")
        db.add_ice_vehicle("Ford", "F-550", mpg_combined=13.5, year=None,
                           fuel_type="Diesel")

        rows = db.get_all_ice_vehicles()
        assert len(rows) == 1, (
            "Duplicate NULL-year row should replace the original, not create a second."
        )
        assert rows[0]["mpg_combined"] == 13.5
        db.close()

    def test_null_year_and_specific_year_coexist(self, tmp_path):
        """A NULL-year catch-all and a specific-year row must coexist without collision."""
        db = _db(tmp_path)
        db.add_ice_vehicle("Ford", "F-550", mpg_combined=12.0, year=None,
                           fuel_type="Diesel")
        db.add_ice_vehicle("Ford", "F-550", mpg_combined=14.0, year=2021,
                           fuel_type="Diesel")

        rows = db.get_all_ice_vehicles()
        assert len(rows) == 2
        db.close()


# ---------------------------------------------------------------------------
# Group 4: Search
# ---------------------------------------------------------------------------

class TestSearch:

    def test_search_by_make(self, tmp_path):
        """search_ice_vehicles(make=...) must return only rows for that make."""
        db = _db(tmp_path)
        db.add_ice_vehicle("Ford",     "F-250", mpg_combined=15.0, year=2020)
        db.add_ice_vehicle("Ford",     "F-350", mpg_combined=13.0, year=2020)
        db.add_ice_vehicle("Chevrolet","Silverado", mpg_combined=16.0, year=2020)

        results = db.search_ice_vehicles(make="Ford")
        assert len(results) == 2
        assert all(r["make"] == "ford" for r in results)
        db.close()

    def test_search_by_make_and_model(self, tmp_path):
        """search_ice_vehicles(make=..., model=...) must apply both filters."""
        db = _db(tmp_path)
        db.add_ice_vehicle("Ford", "F-250", mpg_combined=15.0, year=2020)
        db.add_ice_vehicle("Ford", "F-250", mpg_combined=14.0, year=2019)
        db.add_ice_vehicle("Ford", "F-350", mpg_combined=13.0, year=2020)

        results = db.search_ice_vehicles(make="Ford", model="F-250")
        assert len(results) == 2
        # normalize_vehicle_model("F-250") → "f250" (strips hyphens, lowercases)
        assert all(r["model"] == "f250" for r in results)
        db.close()

    def test_search_by_year(self, tmp_path):
        """search_ice_vehicles(year=...) must return only rows for that year."""
        db = _db(tmp_path)
        db.add_ice_vehicle("Ford", "F-250", mpg_combined=15.0, year=2020)
        db.add_ice_vehicle("Ford", "F-250", mpg_combined=14.0, year=2019)

        results = db.search_ice_vehicles(year=2020)
        assert len(results) == 1
        assert results[0]["year"] == 2020
        db.close()

    def test_search_no_filters_returns_all(self, tmp_path):
        """search_ice_vehicles() with no args returns all rows."""
        db = _db(tmp_path)
        db.add_ice_vehicle("Ford",      "F-250", mpg_combined=15.0, year=2020)
        db.add_ice_vehicle("Chevrolet", "Silverado", mpg_combined=16.0, year=2021)

        results = db.search_ice_vehicles()
        assert len(results) == 2
        db.close()

    def test_search_no_match_returns_empty_list(self, tmp_path):
        """search_ice_vehicles for a non-existent make must return []."""
        db = _db(tmp_path)
        db.add_ice_vehicle("Ford", "F-250", mpg_combined=15.0, year=2020)

        results = db.search_ice_vehicles(make="Toyota")
        assert results == []
        db.close()
