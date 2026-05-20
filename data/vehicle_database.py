"""
data/vehicle_database.py

SQLite-backed vehicle MPG reference database for the Fleet Electrification Analyzer.

Provides a persistent lookup table for analyst-sourced MPG data, filling the gap
between FuelEconomy.gov / commercial scraper coverage and the EPA class-average
fallback. Entries are keyed by make/model/year/fuel_type and are matched with
decreasing specificity so a single entry can cover an entire model line.

Schema covers both ICE vehicles (active) and EV vehicles (stub for future use).

Phase 16 of the Fleet Electrification Analyzer improvement track.
"""

import sqlite3
import threading
import logging
import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any

from utils import normalize_vehicle_model

logger = logging.getLogger(__name__)

# Allowed column names for update_ice_vehicle to prevent SQL injection
_ALLOWED_UPDATE_COLS = {
    "year", "make", "model", "fuel_type", "body_class",
    "gvwr_lbs_min", "gvwr_lbs_max",
    "mpg_combined", "mpg_city", "mpg_highway",
    "notes", "source",
}

_CREATE_ICE_TABLE = """
CREATE TABLE IF NOT EXISTS ice_vehicles (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    year          INTEGER,
    make          TEXT    NOT NULL,
    model         TEXT    NOT NULL,
    fuel_type     TEXT,
    body_class    TEXT,
    gvwr_lbs_min  INTEGER,
    gvwr_lbs_max  INTEGER,
    mpg_combined  REAL,
    mpg_city      REAL,
    mpg_highway   REAL,
    notes         TEXT,
    source        TEXT,
    added_by      TEXT,
    added_date    TEXT,
    updated_date  TEXT,
    UNIQUE(year, make, model, fuel_type)
);
"""

_CREATE_EV_TABLE = """
CREATE TABLE IF NOT EXISTS ev_vehicles (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    ev_model           TEXT    NOT NULL,
    ev_make            TEXT    NOT NULL,
    msrp_low           REAL,
    msrp_high          REAL,
    battery_kwh        REAL,
    epa_range_miles    INTEGER,
    body_class         TEXT,
    gvwr_min           INTEGER,
    gvwr_max           INTEGER,
    towing_lbs         INTEGER,
    payload_lbs        INTEGER,
    cargo_cu_ft        REAL,
    availability_year  INTEGER,
    ice_msrp_low       REAL,
    ice_msrp_high      REAL,
    notes              TEXT,
    match_keywords     TEXT,
    added_date         TEXT,
    updated_date       TEXT
);
"""


###############################################################################
# Manager class
###############################################################################

class VehicleDatabaseManager:
    """
    Manages the SQLite vehicle reference database.

    Thread-safety model:
    - A single connection is opened with check_same_thread=False.
    - Read operations (lookup, search, get_all) need no lock — SQLite allows
      concurrent reads.
    - Write operations (add, update, delete) acquire self._lock to serialise
      all mutations on the shared connection.
    """

    def __init__(self, db_path: Path) -> None:
        """
        Open (or create) the database at db_path and ensure tables exist.

        Args:
            db_path: Filesystem path to the SQLite database file.
        """
        self._db_path = db_path
        self._lock = threading.Lock()

        try:
            self._conn = sqlite3.connect(str(db_path), check_same_thread=False)
            self._conn.row_factory = sqlite3.Row   # rows accessible by column name
            self._conn.execute("PRAGMA journal_mode=WAL;")  # better concurrent access
            self._create_tables()
            logger.info(f"Vehicle database opened at {db_path}")
        except Exception as e:
            logger.error(f"Failed to open vehicle database at {db_path}: {e}")
            raise

    # -------------------------------------------------------------------------
    # Schema
    # -------------------------------------------------------------------------

    def _create_tables(self) -> None:
        """Create tables if they do not already exist."""
        self._conn.executescript(_CREATE_ICE_TABLE + _CREATE_EV_TABLE)
        self._conn.commit()

    # -------------------------------------------------------------------------
    # Lookup (read-only, used by processing pipeline)
    # -------------------------------------------------------------------------

    def lookup_mpg(self, vehicle) -> Optional[Dict[str, Any]]:
        """
        Find an MPG entry for a vehicle using four-tier specificity fallback.

        Tiers (first match wins):
          1. Exact year + make + model + fuel_type
          2. Exact year + make + model (any fuel_type)
          3. NULL-year catch-all: make + model (any year)
          4. GVWR-range match: make + model + GVWR within stored min/max

        Args:
            vehicle: FleetVehicle instance (needs .vehicle_id attributes).

        Returns:
            Dict with keys combined/city/highway/source/notes, or None.
        """
        vid = vehicle.vehicle_id

        make  = (vid.make  or "").strip().lower()
        model = normalize_vehicle_model(vid.model or "")

        try:
            year = int(vid.year) if vid.year else None
        except (ValueError, TypeError):
            year = None

        fuel_type = (vid.fuel_type or "").strip().lower() or None
        gvwr      = vid.gvwr_pounds or 0

        if not make or not model:
            return None

        try:
            # Tier 1 — exact year + make + model + fuel_type
            if year and fuel_type:
                row = self._fetchone(
                    "SELECT * FROM ice_vehicles "
                    "WHERE year=? AND make=? AND model=? AND fuel_type=? "
                    "AND mpg_combined > 0 LIMIT 1",
                    (year, make, model, fuel_type),
                )
                if row:
                    return self._row_to_result(row)

            # Tier 2 — exact year + make + model (any fuel_type)
            if year:
                row = self._fetchone(
                    "SELECT * FROM ice_vehicles "
                    "WHERE year=? AND make=? AND model=? "
                    "AND mpg_combined > 0 "
                    "ORDER BY fuel_type NULLS LAST LIMIT 1",
                    (year, make, model),
                )
                if row:
                    return self._row_to_result(row)

            # Tier 3 — NULL-year model-level catch-all
            row = self._fetchone(
                "SELECT * FROM ice_vehicles "
                "WHERE year IS NULL AND make=? AND model=? "
                "AND mpg_combined > 0 LIMIT 1",
                (make, model),
            )
            if row:
                return self._row_to_result(row)

            # Tier 4 — GVWR-range overlap (only if GVWR is known)
            if gvwr > 0:
                row = self._fetchone(
                    "SELECT * FROM ice_vehicles "
                    "WHERE make=? AND model=? "
                    "AND (gvwr_lbs_min IS NULL OR gvwr_lbs_min <= ?) "
                    "AND (gvwr_lbs_max IS NULL OR gvwr_lbs_max >= ?) "
                    "AND mpg_combined > 0 LIMIT 1",
                    (make, model, gvwr, gvwr),
                )
                if row:
                    return self._row_to_result(row)

        except Exception as e:
            logger.warning(f"Vehicle DB lookup error for {vid.year} {vid.make} {vid.model}: {e}")

        return None

    def _fetchone(self, sql: str, params: tuple) -> Optional[sqlite3.Row]:
        """Execute a SELECT and return the first row, or None."""
        cursor = self._conn.execute(sql, params)
        return cursor.fetchone()

    def _row_to_result(self, row: sqlite3.Row) -> Dict[str, Any]:
        """Convert a sqlite3.Row from ice_vehicles to the lookup result dict."""
        src = row["source"] or "analyst"
        return {
            "combined": row["mpg_combined"] or 0.0,
            "city":     row["mpg_city"]     or 0.0,
            "highway":  row["mpg_highway"]  or 0.0,
            "source":   f"User Database ({src})",
            "notes":    row["notes"] or "",
        }

    # -------------------------------------------------------------------------
    # CRUD — ICE vehicles
    # -------------------------------------------------------------------------

    def add_ice_vehicle(
        self,
        make: str,
        model: str,
        mpg_combined: float,
        year: Optional[int] = None,
        fuel_type: Optional[str] = None,
        body_class: Optional[str] = None,
        mpg_city: float = 0.0,
        mpg_highway: float = 0.0,
        notes: str = "",
        source: str = "analyst",
        gvwr_lbs_min: Optional[int] = None,
        gvwr_lbs_max: Optional[int] = None,
    ) -> int:
        """
        Add or replace an ICE vehicle MPG entry.

        Returns the row id of the inserted/replaced row.

        Note on NULL-year uniqueness: SQLite treats NULL != NULL for UNIQUE
        constraints, so two NULL-year rows with the same make/model/fuel_type
        would both be accepted. We handle this by explicitly deleting any
        existing NULL-year match before inserting.
        """
        make      = make.strip().lower()
        model     = normalize_vehicle_model(model)
        fuel_type = (fuel_type or "").strip().lower() or None
        now       = datetime.datetime.now().isoformat()

        with self._lock:
            # Handle NULL-year duplicate collision (see note above)
            if year is None:
                self._conn.execute(
                    "DELETE FROM ice_vehicles "
                    "WHERE year IS NULL AND make=? AND model=? AND fuel_type IS ?",
                    (make, model, fuel_type),
                )

            cursor = self._conn.execute(
                """
                INSERT OR REPLACE INTO ice_vehicles
                    (year, make, model, fuel_type, body_class,
                     gvwr_lbs_min, gvwr_lbs_max,
                     mpg_combined, mpg_city, mpg_highway,
                     notes, source, added_by, added_date, updated_date)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    year, make, model, fuel_type, body_class,
                    gvwr_lbs_min, gvwr_lbs_max,
                    mpg_combined, mpg_city, mpg_highway,
                    notes, source, "analyst", now, now,
                ),
            )
            self._conn.commit()

        logger.info(
            f"DB: saved {year or 'any-year'} {make} {model} — "
            f"{mpg_combined} mpg combined (source: {source})"
        )
        return cursor.lastrowid

    def update_ice_vehicle(self, row_id: int, **kwargs) -> bool:
        """
        Update specific columns on an existing ice_vehicles row.

        Only columns in _ALLOWED_UPDATE_COLS are accepted; others are silently
        ignored to prevent SQL injection via kwarg names.

        Returns True if a row was updated, False otherwise.
        """
        safe_kwargs = {k: v for k, v in kwargs.items() if k in _ALLOWED_UPDATE_COLS}
        if not safe_kwargs:
            logger.warning("update_ice_vehicle called with no valid columns")
            return False

        # Normalise make/model if being updated
        if "make" in safe_kwargs:
            safe_kwargs["make"] = safe_kwargs["make"].strip().lower()
        if "model" in safe_kwargs:
            safe_kwargs["model"] = normalize_vehicle_model(safe_kwargs["model"])
        if "fuel_type" in safe_kwargs:
            safe_kwargs["fuel_type"] = (safe_kwargs["fuel_type"] or "").strip().lower() or None

        safe_kwargs["updated_date"] = datetime.datetime.now().isoformat()

        set_clause = ", ".join(f"{col}=?" for col in safe_kwargs)
        values     = list(safe_kwargs.values()) + [row_id]

        with self._lock:
            cursor = self._conn.execute(
                f"UPDATE ice_vehicles SET {set_clause} WHERE id=?", values
            )
            self._conn.commit()

        return cursor.rowcount > 0

    def delete_ice_vehicle(self, row_id: int) -> bool:
        """
        Delete an ice_vehicles row by id.

        Returns True if a row was deleted, False otherwise.
        """
        with self._lock:
            cursor = self._conn.execute(
                "DELETE FROM ice_vehicles WHERE id=?", (row_id,)
            )
            self._conn.commit()
        return cursor.rowcount > 0

    # -------------------------------------------------------------------------
    # Read helpers
    # -------------------------------------------------------------------------

    def search_ice_vehicles(
        self,
        make: Optional[str] = None,
        model: Optional[str] = None,
        year: Optional[int] = None,
        fuel_type: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """
        Search ice_vehicles with optional filters.  All supplied filters are
        ANDed together.  Returns a list of row dicts ordered by make, model, year.
        """
        conditions: List[str] = []
        params: List[Any]     = []

        if make:
            conditions.append("make=?")
            params.append(make.strip().lower())
        if model:
            conditions.append("model=?")
            params.append(normalize_vehicle_model(model))
        if year is not None:
            conditions.append("year=?")
            params.append(year)
        if fuel_type:
            conditions.append("fuel_type=?")
            params.append(fuel_type.strip().lower())

        where = f"WHERE {' AND '.join(conditions)}" if conditions else ""
        sql   = f"SELECT * FROM ice_vehicles {where} ORDER BY make, model, year"

        cursor = self._conn.execute(sql, params)
        return [dict(row) for row in cursor.fetchall()]

    def get_all_ice_vehicles(self) -> List[Dict[str, Any]]:
        """Return all ice_vehicles rows ordered by make, model, year."""
        cursor = self._conn.execute(
            "SELECT * FROM ice_vehicles ORDER BY make, model, year"
        )
        return [dict(row) for row in cursor.fetchall()]

    def get_all_ev_vehicles(self) -> List[Dict[str, Any]]:
        """
        Return all ev_vehicles rows.  Empty in Phase 16 — the table exists
        for future use and the hardcoded ev_database.py is still active.
        """
        cursor = self._conn.execute(
            "SELECT * FROM ev_vehicles ORDER BY ev_make, ev_model"
        )
        return [dict(row) for row in cursor.fetchall()]

    # -------------------------------------------------------------------------
    # Lifecycle
    # -------------------------------------------------------------------------

    def close(self) -> None:
        """Close the database connection."""
        try:
            self._conn.close()
            logger.info("Vehicle database connection closed")
        except Exception as e:
            logger.warning(f"Error closing vehicle database: {e}")
