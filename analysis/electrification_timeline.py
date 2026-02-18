"""
analysis/electrification_timeline.py

Fleet-wide electrification year assignment.

Takes a list of vehicles (already classified by ACF compliance) and assigns
each a "Proposed EV Year" — a target year for electrification that:

  1. Prioritises older, higher-mileage, ACF-subject vehicles
  2. Spreads replacements roughly evenly across years (budget smoothing)
  3. Respects a user-configurable end year (default 2040)

This module MUST run after ACF classification (acf_compliance.py) because
it reads the internal ``_acf_code`` stored in each vehicle's custom_fields.
"""

import datetime
import logging
import math
from typing import List, Tuple

logger = logging.getLogger(__name__)

# Avoid circular import at module level
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from data.models import FleetVehicle


###############################################################################
# Scoring Weights
###############################################################################

# How much each factor contributes to the replacement priority score.
# Higher total score → replace sooner.
WEIGHT_AGE = 0.35
WEIGHT_MILEAGE = 0.25
WEIGHT_ANNUAL_USAGE = 0.15
# ACF boost is additive, not weighted (see below)

# Normalisation ceilings — vehicles at or above these values score 1.0
MAX_AGE_YEARS = 15
MAX_ODOMETER = 200_000
MAX_ANNUAL_MILEAGE = 20_000

# ACF category boosts (additive on top of the weighted score)
ACF_BOOST = {
    "B": 0.30,   # Subject to ACF — highest urgency
    "C": 0.10,   # Exempt body type — still useful to electrify
    "D": 0.10,   # Emergency — still useful to electrify
}


###############################################################################
# Public API
###############################################################################

def assign_electrification_years(
    vehicles: List['FleetVehicle'],
    end_year: int = 2040,
) -> None:
    """
    Assign a proposed electrification year to every vehicle in the list.

    Modifies ``vehicle.custom_fields["Proposed EV Year"]`` in-place.

    Args:
        vehicles: List of FleetVehicle objects (ACF classification must
                  already be populated in custom_fields["_acf_code"]).
        end_year: Final year of the electrification timeline (inclusive).
    """
    current_year = datetime.datetime.now().year
    if end_year <= current_year:
        logger.warning(
            f"End year {end_year} is not after current year {current_year}; "
            "skipping timeline assignment"
        )
        for v in vehicles:
            v.custom_fields["Proposed EV Year"] = "N/A"
        return

    # ── 1. Triage vehicles into groups ─────────────────────────────
    schedulable = []  # type: List[Tuple[float, FleetVehicle]]

    for vehicle in vehicles:
        acf_code = vehicle.custom_fields.get("_acf_code", "")

        # Already a ZEV — nothing to do
        if acf_code == "ZEV":
            vehicle.custom_fields["Proposed EV Year"] = "N/A"
            continue

        # Processing failed — can't schedule
        if not vehicle.processing_success:
            vehicle.custom_fields["Proposed EV Year"] = "N/A"
            continue

        # Missing ACF classification — warn and skip
        if not acf_code:
            logger.warning(
                f"Vehicle {vehicle.vin} missing ACF classification; "
                "cannot assign electrification year"
            )
            vehicle.custom_fields["Proposed EV Year"] = "N/A"
            continue

        # Light-duty exempt — not subject to timeline
        if acf_code == "A":
            vehicle.custom_fields["Proposed EV Year"] = "Exempt"
            continue

        # Everything else (B, C, D) gets scored and scheduled
        score = _score_vehicle(vehicle, acf_code)
        schedulable.append((score, vehicle))

    if not schedulable:
        logger.info("No vehicles eligible for electrification scheduling")
        return

    # ── 2. Sort by score descending (highest priority first) ───────
    schedulable.sort(key=lambda pair: pair[0], reverse=True)

    # ── 3. Budget-smooth across years ──────────────────────────────
    available_years = list(range(current_year + 1, end_year + 1))
    num_years = len(available_years)
    target_per_year = math.ceil(len(schedulable) / num_years)

    # Track how many vehicles have been assigned to each year
    year_counts = {yr: 0 for yr in available_years}

    for _score, vehicle in schedulable:
        acf_code = vehicle.custom_fields.get("_acf_code", "")

        # Find the earliest year that still has capacity
        assigned_year = None
        for yr in available_years:
            if year_counts[yr] < target_per_year:
                assigned_year = yr
                break

        # Fallback: if all years are "full", append to the last year
        if assigned_year is None:
            assigned_year = available_years[-1]

        year_counts[assigned_year] += 1
        vehicle.custom_fields["Proposed EV Year"] = str(assigned_year)

    # Log summary
    scheduled_count = len(schedulable)
    year_range = f"{available_years[0]}-{available_years[-1]}"
    logger.info(
        f"Electrification timeline: {scheduled_count} vehicles scheduled "
        f"across {year_range} (~{target_per_year}/year)"
    )


###############################################################################
# Scoring
###############################################################################

def _score_vehicle(vehicle: 'FleetVehicle', acf_code: str) -> float:
    """
    Calculate a replacement-priority score for a single vehicle.

    Higher score → should be replaced sooner.

    Components:
      - Age score (0-1): older vehicles score higher
      - Mileage score (0-1): higher-mileage vehicles score higher
      - Annual usage score (0-1): high-use vehicles benefit more from EV
      - ACF boost (additive): ACF-subject vehicles get a priority bump
      - Data-completeness penalty: vehicles missing all three metrics
        (age, odometer, annual mileage) are deprioritised so they don't
        leapfrog vehicles with real data just because of an ACF boost.

    Returns:
        Float score (roughly 0.0 to ~1.05).
    """
    # Age
    age = vehicle.age  # years since model year
    has_age = age > 0
    age_score = min(age / MAX_AGE_YEARS, 1.0) if has_age else 0.0

    # Odometer
    odo = vehicle.odometer
    has_odo = odo > 0
    mileage_score = min(odo / MAX_ODOMETER, 1.0) if has_odo else 0.0

    # Annual usage
    annual = vehicle.annual_mileage
    has_annual = annual > 0
    usage_score = min(annual / MAX_ANNUAL_MILEAGE, 1.0) if has_annual else 0.0

    # Weighted sum
    weighted = (
        WEIGHT_AGE * age_score
        + WEIGHT_MILEAGE * mileage_score
        + WEIGHT_ANNUAL_USAGE * usage_score
    )

    # ACF boost
    boost = ACF_BOOST.get(acf_code, 0.0)

    # Data-completeness penalty: if none of the three usage metrics are
    # available the weighted component is zero and only the ACF boost
    # remains.  Halve the boost in that case so these vehicles sort to
    # the back of their ACF tier rather than ahead of vehicles with real
    # data.  If at least one metric is present, apply the full boost.
    metrics_present = sum([has_age, has_odo, has_annual])
    if metrics_present == 0:
        boost *= 0.5

    return weighted + boost
