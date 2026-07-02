"""
project_io.py

Save and load Fleet Electrification Analyzer project files (.fea).

A .fea file is a JSON document containing the full Fleet state — all
FleetVehicle records (including VehicleIdentification, FuelEconomyData,
ACF codes, EV year assignments, and custom_fields) plus optional
scenario_results from the last analysis run.

Loading a project reconstructs the Fleet in memory and bypasses the
VIN-processing pipeline entirely.
"""

import copy
import datetime
import json
import logging
from typing import Any, Dict, Optional, Tuple

from data.models import Fleet, FleetVehicle, FuelEconomyData, VehicleIdentification

logger = logging.getLogger(__name__)

PROJECT_VERSION = 1
FILE_EXTENSION = ".fea"


# ---------------------------------------------------------------------------
# JSON encoder / decoder helpers
# ---------------------------------------------------------------------------

class _DatetimeEncoder(json.JSONEncoder):
    """Serialize datetime.date and datetime.datetime as tagged dicts."""

    def default(self, obj: Any) -> Any:
        if isinstance(obj, datetime.datetime):
            return {"__type__": "datetime", "value": obj.isoformat()}
        if isinstance(obj, datetime.date):
            return {"__type__": "date", "value": obj.isoformat()}
        return super().default(obj)


def _datetime_hook(obj: Dict) -> Any:
    """Reconstruct datetime objects from tagged dicts produced by _DatetimeEncoder."""
    t = obj.get("__type__")
    if t == "datetime":
        return datetime.datetime.fromisoformat(obj["value"])
    if t == "date":
        return datetime.date.fromisoformat(obj["value"])
    return obj


# ---------------------------------------------------------------------------
# Serialization
# ---------------------------------------------------------------------------

def _serialize_vehicle(v: FleetVehicle) -> Dict:
    """Convert a FleetVehicle to a plain JSON-serializable dict.

    commercial_specs is intentionally excluded — it holds scraped specs that
    can be re-fetched and is not a plain dataclass on all code paths.
    """
    import dataclasses

    v_copy = copy.copy(v)
    v_copy.commercial_specs = None  # exclude — not reliably serializable
    return dataclasses.asdict(v_copy)


# ---------------------------------------------------------------------------
# Deserialization
# ---------------------------------------------------------------------------

def _coerce_date(value: Any) -> Optional[datetime.date]:
    if value is None:
        return None
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, str) and value:
        return datetime.date.fromisoformat(value)
    return None


def _coerce_datetime(value: Any) -> Optional[datetime.datetime]:
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value
    if isinstance(value, str) and value:
        return datetime.datetime.fromisoformat(value)
    return None


def _deserialize_vehicle_id(d: Dict) -> VehicleIdentification:
    valid = set(VehicleIdentification.__dataclass_fields__)
    return VehicleIdentification(**{k: v for k, v in d.items() if k in valid})


def _deserialize_fuel_economy(d: Dict) -> FuelEconomyData:
    valid = set(FuelEconomyData.__dataclass_fields__)
    return FuelEconomyData(**{k: v for k, v in d.items() if k in valid})


def _deserialize_vehicle(d: Dict) -> FleetVehicle:
    vehicle_id = _deserialize_vehicle_id(d.get("vehicle_id") or {})
    fuel_economy = _deserialize_fuel_economy(d.get("fuel_economy") or {})

    # Date/datetime fields require explicit coercion (asdict leaves them as
    # tagged dicts after _datetime_hook runs, or raw ISO strings on older files)
    acquisition_date = _coerce_date(d.get("acquisition_date"))
    retire_date = _coerce_date(d.get("retire_date"))
    last_quality_check = _coerce_datetime(d.get("last_quality_check"))

    skip = {
        "vehicle_id", "fuel_economy",
        "acquisition_date", "retire_date", "last_quality_check",
        "commercial_specs",
    }
    valid = set(FleetVehicle.__dataclass_fields__)
    scalar_fields = {k: v for k, v in d.items() if k in valid and k not in skip}

    return FleetVehicle(
        vehicle_id=vehicle_id,
        fuel_economy=fuel_economy,
        acquisition_date=acquisition_date,
        retire_date=retire_date,
        last_quality_check=last_quality_check,
        commercial_specs=None,
        **scalar_fields,
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def save_project(
    path: str,
    fleet: Fleet,
    scenario_results: Optional[Dict] = None,
) -> None:
    """Write the Fleet (and optional scenario_results) to a .fea JSON file.

    Args:
        path:             Destination file path (should end in .fea).
        fleet:            Fleet object to persist.
        scenario_results: Dict returned by compare_scenarios(), or None.
    """
    payload = {
        "version": PROJECT_VERSION,
        "fleet": {
            "name": fleet.name,
            "notes": fleet.notes,
            "fleet_type": fleet.fleet_type,
            "max_vehicles_per_year": fleet.max_vehicles_per_year,
            "creation_date": fleet.creation_date.isoformat() if fleet.creation_date else None,
            "last_modified": datetime.datetime.now().isoformat(),
            "vehicles": [_serialize_vehicle(v) for v in fleet.vehicles],
        },
        "scenario_results": scenario_results,
    }
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2, cls=_DatetimeEncoder)
    logger.info("Project saved to %s (%d vehicles)", path, len(fleet.vehicles))


def load_project(path: str) -> Tuple[Fleet, Optional[Dict]]:
    """Read a .fea file and reconstruct the Fleet.

    Returns:
        (fleet, scenario_results) — scenario_results is None when no analysis
        had been run before saving.

    Raises:
        ValueError: If the file version is unsupported.
        json.JSONDecodeError / IOError: On corrupt or missing file.
    """
    with open(path, "r", encoding="utf-8") as fh:
        data = json.load(fh, object_hook=_datetime_hook)

    version = data.get("version", 1)
    if version > PROJECT_VERSION:
        raise ValueError(
            f"Project file version {version} is newer than this app supports "
            f"(max {PROJECT_VERSION}). Please upgrade the application."
        )

    fd = data["fleet"]
    vehicles = [_deserialize_vehicle(vd) for vd in fd.get("vehicles", [])]

    creation_date = _coerce_datetime(fd.get("creation_date")) or datetime.datetime.now()

    fleet = Fleet(
        name=fd.get("name", "Loaded Fleet"),
        notes=fd.get("notes", ""),
        fleet_type=fd.get("fleet_type", "hpf"),
        max_vehicles_per_year=fd.get("max_vehicles_per_year", 0),
        creation_date=creation_date,
        last_modified=datetime.datetime.now(),
        vehicles=vehicles,
    )

    logger.info("Project loaded from %s (%d vehicles)", path, len(vehicles))
    return fleet, data.get("scenario_results")
