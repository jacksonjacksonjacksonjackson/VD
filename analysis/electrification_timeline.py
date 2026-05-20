"""
analysis/electrification_timeline.py

Fleet-wide electrification year assignment.

Takes a list of vehicles (already classified by ACF compliance) and assigns
each a "Proposed EV Year" — a target year for electrification that:

  1. For ALL schedulable vehicles (Cat B, C, D): score-based priority queue
     with GVWR-tiered urgency boost for Cat B.  CARB deadlines inform priority,
     not hard year assignment.  The result is an even procurement spread across
     the planning horizon that respects normal fleet-replacement budgets.
  2. Cat B CARB deadline years are stored in custom_fields["ACF Deadline Year"]
     for compliance reference and Milestone Option chart rendering.
  3. Spreads vehicles evenly across years (budget smoothing) with optional
     max_per_year capacity cap.
  4. Records a plain-English reason for every assignment so analysts can see
     exactly why each vehicle received its year or N/A.

Also populates an "ACF Relevance" field with a client-presentable description
of each vehicle's mandate status (used alongside the letter-code ACF Category).

Design notes
------------
CARB ACF High-Priority Fleet mandate structure (per GVWR class):

  ZEV Purchase Option milestones — each purchase of a covered vehicle in
  the applicable year must be a ZEV:
    Class 2b–4  (8,501–19,500 lbs): 10% by 2024, 25% by 2025, 50% by 2028,
                                     75% by 2031, 100% by 2033
    Class 5–8a (19,501–33,000 lbs): 10% by 2024, 20% by 2025, 30% by 2026,
                                     50% by 2028, 75% by 2031, 100% by 2035
    Class 8b     (33,001+ lbs):     10% by 2025, 20% by 2026, 30% by 2027,
                                     50% by 2030, 75% by 2033, 100% by 2040

  ZEV Milestone Option — 100% ZEV fleet target years by class:
    Class 2b–4:   2035
    Class 5–8a:   2039
    Class 8b:     2042

  For replacement planning purposes this module uses the 100% Milestone
  target year as the outer deadline for each class.  Vehicles with
  high priority scores are pulled forward in the queue; vehicles already
  at end-of-life (high age / high mileage) are placed in the earliest
  available bucket.

  Non-HPF fleets (state/local government etc.) have different schedules;
  the app does not currently distinguish HPF vs. non-HPF.  The analyst
  should review and adjust output accordingly.
"""

import datetime
import logging
import math
from typing import List, Optional, Tuple

logger = logging.getLogger(__name__)

# Avoid circular import at module level
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from data.models import FleetVehicle


###############################################################################
# ACF Compliance Deadline Table  (GVWR → 100% mandate year)
###############################################################################

# Keyed by (lower_lbs_exclusive, upper_lbs_inclusive).
# Each entry has:
#   "milestone_year"   – 100% ZEV Milestone Option year
#   "purchase_schedule"– list of (year, pct) ZEV Purchase Option checkpoints
#   "class_label"      – human-readable class description
ACF_DEADLINE_TABLE = {
    (8500, 19500): {
        "milestone_year": 2035,
        "purchase_schedule": [
            (2024, 10), (2025, 25), (2028, 50), (2031, 75), (2033, 100),
        ],
        "class_label": "Class 2b–4 (8,501–19,500 lbs)",
    },
    (19500, 33000): {
        "milestone_year": 2039,
        "purchase_schedule": [
            (2024, 10), (2025, 20), (2026, 30), (2028, 50), (2031, 75), (2035, 100),
        ],
        "class_label": "Class 5–8a (19,501–33,000 lbs)",
    },
    (33000, 999_999): {
        "milestone_year": 2042,
        "purchase_schedule": [
            (2025, 10), (2026, 20), (2027, 30), (2030, 50), (2033, 75), (2040, 100),
        ],
        "class_label": "Class 8b (33,001+ lbs)",
    },
}

# Fallback deadline for Category B vehicles whose GVWR is unknown.
# Using 2035 (Class 2b–4 / most common medium-duty milestone) is conservative.
ACF_DEADLINE_UNKNOWN_GVWR = 2035

# ---------------------------------------------------------------------------
# Non-HPF deadline table  (state/local government fleets not classified HPF)
# ---------------------------------------------------------------------------
# Milestone years are 1 year later than HPF per CARB ACF regulation.
# NOTE: Verify these values against current CARB regulatory text before
#       client delivery — thresholds can change as the regulation evolves.
ACF_DEADLINE_TABLE_NON_HPF = {
    (8500, 19500): {
        "milestone_year": 2036,
        "purchase_schedule": [
            (2025, 10), (2026, 25), (2029, 50), (2032, 75), (2034, 100),
        ],
        "class_label": "Class 2b–4 (8,501–19,500 lbs)",
    },
    (19500, 33000): {
        "milestone_year": 2040,
        "purchase_schedule": [
            (2025, 10), (2026, 20), (2027, 30), (2029, 50), (2032, 75), (2036, 100),
        ],
        "class_label": "Class 5–8a (19,501–33,000 lbs)",
    },
    (33000, 999_999): {
        "milestone_year": 2043,
        "purchase_schedule": [
            (2026, 10), (2027, 20), (2028, 30), (2031, 50), (2034, 75), (2041, 100),
        ],
        "class_label": "Class 8b (33,001+ lbs)",
    },
}
ACF_DEADLINE_UNKNOWN_GVWR_NON_HPF = 2036

# Convenience lookup: fleet_type string → (deadline_table, unknown_gvwr_fallback)
# "state_agency" uses the same table as "non_hpf" for now.
_DEADLINE_TABLES: dict = {
    "hpf":          (ACF_DEADLINE_TABLE,         ACF_DEADLINE_UNKNOWN_GVWR),
    "non_hpf":      (ACF_DEADLINE_TABLE_NON_HPF, ACF_DEADLINE_UNKNOWN_GVWR_NON_HPF),
    "state_agency": (ACF_DEADLINE_TABLE_NON_HPF, ACF_DEADLINE_UNKNOWN_GVWR_NON_HPF),
}


###############################################################################
# ACF Relevance Labels  (client-presentable plain-English)
###############################################################################

ACF_RELEVANCE = {
    "ZEV": "Already Zero-Emission — No Action Required",
    "A":   "Exempt — Light-Duty (Not Subject to ACF)",
    "B":   "ACF Mandate-Subject — Must Electrify",
    "C":   "Exempt — Body Type (ZEV Not Commercially Available)",
    "D":   "Emergency Vehicle — Exempt from Mandate",
    "":    "Classification Unavailable",
}


###############################################################################
# Scoring Weights  (used for non-deadline queue placement)
###############################################################################

# Age contributes most — a 2001 vehicle should almost always precede a 2018.
WEIGHT_AGE = 0.55
WEIGHT_MILEAGE = 0.25
WEIGHT_ANNUAL_USAGE = 0.10
# ACF boost is additive on top of the weighted score
ACF_BOOST = {
    "B": 0.30,
    "C": 0.10,
    "D": 0.10,
}

# Category B GVWR-tiered urgency boosts (additive on top of weighted score).
# Higher boost → placed earlier in the even-spread procurement queue.
# Tiers reflect CARB ACF deadline urgency: earlier mandate = higher boost.
ACF_BOOST_B_CLASS_2B4 = 0.35    # Class 2b–4 (8,501–19,500 lbs) — earliest deadline
ACF_BOOST_B_CLASS_5_8A = 0.25   # Class 5–8a (19,501–33,000 lbs) — middle deadline
ACF_BOOST_B_CLASS_8B = 0.15     # Class 8b  (33,001+ lbs) — latest deadline
ACF_BOOST_B_UNKNOWN = 0.30      # Unknown GVWR — conservative (Class 2b–4 level)

# Normalisation ceilings (vehicles at or above these values score 1.0)
MAX_AGE_YEARS = 15
MAX_ODOMETER = 200_000
MAX_ANNUAL_MILEAGE = 20_000


###############################################################################
# Public API
###############################################################################

def assign_electrification_years(
    vehicles: List['FleetVehicle'],
    end_year: int = 2040,
    fleet_type: str = "hpf",
    max_per_year: int = 0,
) -> None:
    """
    Assign a proposed electrification year to every vehicle in the list.

    All schedulable vehicles (Cat B, C, D) enter a score-based priority queue
    and are distributed evenly across the planning horizon.  Cat B vehicles
    receive a GVWR-tiered urgency boost so earlier-deadline vehicles naturally
    sort to earlier years.  CARB deadline years are stored for reference but
    do NOT hard-assign vehicles to a specific year.

    Modifies in-place:
      - ``vehicle.custom_fields["Proposed EV Year"]``
      - ``vehicle.custom_fields["EV Year Reason"]``
      - ``vehicle.custom_fields["ACF Relevance"]``
      - ``vehicle.custom_fields["ACF Deadline Year"]`` (Cat B only — CARB reference)

    Args:
        vehicles:     List of FleetVehicle objects (ACF classification must
                      already be populated in custom_fields["_acf_code"]).
        end_year:     Outer bound of the planning horizon (inclusive).
        fleet_type:   CARB fleet classification for deadline lookup.
                      One of "hpf" (default), "non_hpf", or "state_agency".
                      Unknown values fall back to "hpf".
        max_per_year: If > 0, cap the number of vehicles assigned per year
                      (capacity-constrained procurement budget).  Highest-
                      priority vehicles fill earlier years first.
    """
    current_year = datetime.datetime.now().year

    if end_year <= current_year:
        logger.warning(
            f"End year {end_year} is not after current year {current_year}; "
            "all non-exempt vehicles assigned N/A"
        )
        for v in vehicles:
            _set_result(v, "N/A", "N/A — Planning horizon has passed", "")
        return

    # ── 1. Triage ─────────────────────────────────────────────────────────────
    # All schedulable vehicles (B, C, D) enter score_queue.
    # Cat B receives GVWR-tiered urgency boost; CARB deadline year is stored
    # as "ACF Deadline Year" for reference only (does not drive assignment).
    score_queue: List[Tuple[float, 'FleetVehicle']] = []

    for vehicle in vehicles:
        acf_code = vehicle.custom_fields.get("_acf_code", "")

        # Populate ACF Relevance for every vehicle regardless of outcome
        vehicle.custom_fields["ACF Relevance"] = ACF_RELEVANCE.get(acf_code, ACF_RELEVANCE[""])

        # ── Already ZEV ───────────────────────────────────────────────────────
        if acf_code == "ZEV":
            _set_result(vehicle, "N/A", "N/A — Already a zero-emission vehicle", acf_code)
            continue

        # ── Processing failed ─────────────────────────────────────────────────
        if not vehicle.processing_success:
            _set_result(vehicle, "N/A", "N/A — Processing failed; vehicle data incomplete", acf_code)
            continue

        # ── Missing ACF classification ────────────────────────────────────────
        if not acf_code:
            logger.warning(f"Vehicle {vehicle.vin} missing ACF classification")
            _set_result(vehicle, "N/A", "N/A — ACF classification unavailable", "")
            continue

        # ── Category A: Light-duty exempt ─────────────────────────────────────
        if acf_code == "A":
            _set_result(
                vehicle, "Exempt",
                "Exempt — Light-duty vehicle (GVWR ≤ 8,500 lbs); not subject to ACF mandate",
                acf_code,
            )
            continue

        # ── Category B: ACF mandate-subject ───────────────────────────────────
        # Store CARB deadline year for compliance reference / Milestone chart.
        # Vehicle enters the score queue with GVWR-tiered urgency boost so
        # earlier-deadline vehicles naturally sort to earlier procurement years.
        if acf_code == "B":
            deadline_year, _ = _acf_deadline_for_vehicle(
                vehicle, current_year, fleet_type=fleet_type
            )
            vehicle.custom_fields["ACF Deadline Year"] = (
                str(deadline_year) if deadline_year is not None else "Unknown"
            )
            score = _score_vehicle(vehicle, acf_code)
            score_queue.append((score, vehicle))
            continue

        # ── Categories C and D: score queue ───────────────────────────────────
        score = _score_vehicle(vehicle, acf_code)
        score_queue.append((score, vehicle))

    # ── 2. Budget-smooth the score queue ──────────────────────────────────────
    if score_queue:
        score_queue.sort(key=lambda pair: pair[0], reverse=True)
        available_years = list(range(current_year + 1, end_year + 1))
        num_years = len(available_years)
        n = len(score_queue)

        if max_per_year > 0:
            # Capacity-constrained: fill years left-to-right, cap at max_per_year.
            # Highest-priority vehicles get the earliest slots.
            year_counts = {yr: 0 for yr in available_years}
            for score, vehicle in score_queue:
                acf_code = vehicle.custom_fields.get("_acf_code", "")
                assigned_year = available_years[-1]  # default: last year if all full
                for yr in available_years:
                    if year_counts[yr] < max_per_year:
                        assigned_year = yr
                        year_counts[yr] += 1
                        break
                _set_result(
                    vehicle,
                    str(assigned_year),
                    _score_reason(vehicle, acf_code, score, assigned_year),
                    acf_code,
                )
        elif n <= num_years:
            # Fewer (or equal) vehicles than years: space them evenly across the
            # full planning horizon.  Vehicle k (0 = highest priority) maps to:
            #   available_years[round(k * (num_years-1) / max(n-1, 1))]
            # This guarantees:
            #   • Highest-priority vehicle → first year of the horizon
            #   • Lowest-priority vehicle  → last year of the horizon
            #   • Others proportionally between (score order preserved)
            for k, (score, vehicle) in enumerate(score_queue):
                acf_code = vehicle.custom_fields.get("_acf_code", "")
                idx = 0 if n == 1 else round(k * (num_years - 1) / (n - 1))
                assigned_year = available_years[idx]
                _set_result(
                    vehicle,
                    str(assigned_year),
                    _score_reason(vehicle, acf_code, score, assigned_year),
                    acf_code,
                )
        else:
            # More vehicles than years: use floor-based per-year capacity so
            # every year in the horizon receives vehicles.
            base = n // num_years
            extra = n % num_years
            year_slots: List[int] = []
            for i, yr in enumerate(available_years):
                cap = base + (1 if i < extra else 0)
                year_slots.extend([yr] * cap)
            for (score, vehicle), assigned_year in zip(score_queue, year_slots):
                acf_code = vehicle.custom_fields.get("_acf_code", "")
                _set_result(
                    vehicle,
                    str(assigned_year),
                    _score_reason(vehicle, acf_code, score, assigned_year),
                    acf_code,
                )

    # ── 3. Log summary ────────────────────────────────────────────────────────
    logger.info(
        f"Electrification timeline: {len(score_queue)} vehicles queued by priority score "
        f"(Cat B with GVWR-tiered urgency boost; CARB deadlines stored for reference)"
    )


###############################################################################
# ACF Deadline Lookup
###############################################################################

def _acf_deadline_for_vehicle(
    vehicle: 'FleetVehicle', current_year: int, fleet_type: str = "hpf"
) -> Tuple[Optional[int], str]:
    """
    Return (deadline_year, reason_text) for a Category B vehicle.

    Uses GVWR to select the applicable mandate class from the deadline table
    appropriate for fleet_type ("hpf", "non_hpf", or "state_agency").
    Returns (None, "") if GVWR is unknown (caller falls through to score queue).
    """
    table, unknown_fallback = _DEADLINE_TABLES.get(fleet_type, _DEADLINE_TABLES["hpf"])
    gvwr = vehicle.vehicle_id.gvwr_pounds

    if gvwr <= 0:
        return None, ""

    for (low, high), entry in sorted(table.items()):
        if low < gvwr <= high:
            milestone = entry["milestone_year"]
            class_label = entry["class_label"]

            # Build purchase schedule note
            schedule_notes = ", ".join(
                f"{pct}% ZEV by {yr}" for yr, pct in entry["purchase_schedule"]
            )

            # Suggest the vehicle's individual replacement year.
            # Vehicles with high urgency scores are pulled toward current year;
            # all others get the full-100% milestone year.
            urgency = _score_vehicle(vehicle, "B")
            urgency_threshold = 0.55  # score above this = "high urgency"

            if urgency >= urgency_threshold:
                # Pull toward the earliest applicable purchase checkpoint
                replacement_year = _earliest_purchase_checkpoint(entry, current_year)
                reason = (
                    f"ACF mandate ({class_label}): high urgency score {urgency:.2f} → "
                    f"targeted for early replacement by {replacement_year}. "
                    f"Purchase milestones: {schedule_notes}. "
                    f"100% ZEV milestone: {milestone}."
                )
            else:
                replacement_year = milestone
                reason = (
                    f"ACF mandate ({class_label}): assigned 100% ZEV milestone year {milestone}. "
                    f"Purchase milestones: {schedule_notes}."
                )

            return replacement_year, reason

    # GVWR above all table ranges — treat as Class 8b
    entry = table[(33000, 999_999)]
    milestone = entry["milestone_year"]
    return milestone, (
        f"ACF mandate ({entry['class_label']}): assigned 100% ZEV milestone year {milestone}. "
        f"GVWR {gvwr:,.0f} lbs exceeds Class 8b threshold."
    )


def _earliest_purchase_checkpoint(entry: dict, current_year: int) -> int:
    """
    Return the earliest future purchase mandate checkpoint year for a fleet entry.
    Falls back to the milestone year if all checkpoints are in the past.
    """
    for yr, _pct in sorted(entry["purchase_schedule"]):
        if yr > current_year:
            return yr
    return entry["milestone_year"]


###############################################################################
# Scoring
###############################################################################

def _cat_b_boost(vehicle: 'FleetVehicle') -> float:
    """Return the urgency boost for a Category B vehicle based on GVWR class.

    Earlier CARB mandate deadline → higher boost → placed earlier in even spread.
    """
    gvwr = vehicle.vehicle_id.gvwr_pounds
    if gvwr <= 0:
        return ACF_BOOST_B_UNKNOWN
    elif gvwr <= 19500:
        return ACF_BOOST_B_CLASS_2B4    # Class 2b–4: earliest deadline
    elif gvwr <= 33000:
        return ACF_BOOST_B_CLASS_5_8A   # Class 5–8a: middle deadline
    else:
        return ACF_BOOST_B_CLASS_8B     # Class 8b: latest deadline


def _score_vehicle(vehicle: 'FleetVehicle', acf_code: str) -> float:
    """
    Calculate a replacement-priority score for a single vehicle.

    Higher score → should be replaced sooner.

    Components:
      - Age score (0–1): age weighted at WEIGHT_AGE (0.55).  Vehicles 15+ years
        score 1.0; newer vehicles score proportionally less.
      - Mileage score (0–1): WEIGHT_MILEAGE (0.25)
      - Annual usage score (0–1): WEIGHT_ANNUAL_USAGE (0.10)
      - ACF boost (additive): B=+0.30, C/D=+0.10
      - Data-completeness penalty: halves ACF boost when all three metrics absent

    Returns:
        Float score (roughly 0.0 to ~1.35).
    """
    age = vehicle.age
    has_age = age > 0
    age_score = min(age / MAX_AGE_YEARS, 1.0) if has_age else 0.0

    odo = vehicle.odometer
    has_odo = odo > 0
    mileage_score = min(odo / MAX_ODOMETER, 1.0) if has_odo else 0.0

    annual = vehicle.annual_mileage
    has_annual = annual > 0
    usage_score = min(annual / MAX_ANNUAL_MILEAGE, 1.0) if has_annual else 0.0

    weighted = (
        WEIGHT_AGE * age_score
        + WEIGHT_MILEAGE * mileage_score
        + WEIGHT_ANNUAL_USAGE * usage_score
    )

    if acf_code == "B":
        boost = _cat_b_boost(vehicle)
    else:
        boost = ACF_BOOST.get(acf_code, 0.0)

    metrics_present = sum([has_age, has_odo, has_annual])
    if metrics_present == 0:
        boost *= 0.5

    return weighted + boost


def _score_reason(
    vehicle: 'FleetVehicle',
    acf_code: str,
    score: float,
    assigned_year: int,
) -> str:
    """
    Build a plain-English reason string for a score-queue assignment.
    """
    parts = []

    age = vehicle.age
    if age > 0:
        parts.append(f"age {age:.0f} yrs")
    if vehicle.odometer > 0:
        parts.append(f"{vehicle.odometer:,.0f} mi odometer")
    if vehicle.annual_mileage > 0:
        parts.append(f"{vehicle.annual_mileage:,.0f} mi/yr")

    data_note = " · ".join(parts) if parts else "limited vehicle data"
    category_label = {
        "B": "ACF mandate-subject — queued by procurement urgency",
        "C": "Exempt body type — queued by replacement urgency",
        "D": "Emergency vehicle — queued by replacement urgency",
    }.get(acf_code, "Queued by replacement urgency")

    deadline_note = ""
    if acf_code == "B":
        deadline = vehicle.custom_fields.get("ACF Deadline Year", "")
        if deadline and deadline != "Unknown":
            deadline_note = f" CARB compliance deadline: {deadline}."

    return (
        f"{category_label}. Priority score {score:.2f} → assigned {assigned_year}. "
        f"Factors: {data_note}.{deadline_note}"
    )


###############################################################################
# Helpers
###############################################################################

def _set_result(
    vehicle: 'FleetVehicle',
    year_value: str,
    reason: str,
    acf_code: str,
) -> None:
    """Write Proposed EV Year and EV Year Reason to a vehicle's custom_fields."""
    vehicle.custom_fields["Proposed EV Year"] = year_value
    vehicle.custom_fields["EV Year Reason"] = reason
    # ACF Relevance is set once per vehicle in the triage loop;
    # only set here if not already populated (e.g. ZEV / failed vehicles).
    if "ACF Relevance" not in vehicle.custom_fields:
        vehicle.custom_fields["ACF Relevance"] = ACF_RELEVANCE.get(acf_code, ACF_RELEVANCE[""])
