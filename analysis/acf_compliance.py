"""
analysis/acf_compliance.py

CARB Advanced Clean Fleets (ACF) compliance classification.

Assigns each vehicle to one of four ACF categories based on GVWR,
body class, model/trim data, and department information:

    A  - Exempt (Light-Duty)      GVWR <= 8,500 lbs; not subject to ACF
    B  - Subject to ACF           Medium/heavy-duty; covered by regulation
    C  - Exempt (Body Type)       Body type on CARB ZEV Purchase Exemption List
    D  - Emergency Vehicle         Fire, police, ambulance, etc.
    ZEV - Zero-Emission Vehicle   Already electric; no action required
"""

import logging
import re
from typing import Tuple

logger = logging.getLogger(__name__)

# Avoid circular import at module level; FleetVehicle is only used for typing
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from data.models import FleetVehicle

###############################################################################
# Constants — CARB ACF Exemption Lists
###############################################################################

# Body type keywords from the CARB ZEV Purchase Exemption List.
# Vehicles with these body configurations can purchase ICE without
# pre-approval because ZEVs are not commercially available.
EXEMPT_BODY_TYPES = [
    "bucket",
    "boom",
    "dump",
    "flatbed",
    "stake bed",
    "stake",
    "refuse",
    "roll-off",
    "rolloff",
    "compactor",
    "service body",
    "street sweeper",
    "sweeper",
    "tank truck",
    "tanker",
    "tow truck",
    "tow",
    "water truck",
    "car carrier",
    "auto carrier",
    "concrete mixer",
    "cement mixer",
    "concrete pump",
    "crane",
    "drill rig",
    "drill",
    "vacuum truck",
    "vacuum",
]

# Emergency vehicle detection patterns.
# These are checked against various vehicle fields (case-insensitive).
# All keyword matching uses word-boundary regex (\b) to prevent partial
# matches (e.g. "fire" must not match "Firewall" or "Crossfire").

# Body class keywords — strong signals, matched with word boundaries.
EMERGENCY_BODY_KEYWORDS = [
    "ambulance",
    "fire apparatus",
    "fire truck",
    "fire engine",
    "rescue",           # NHTSA body class "Rescue Vehicle" is unambiguous
]

# Model/trim keywords that are STRONG signals on their own (no second
# signal required).  "ppv" and "ssv" are VIN-decoded designations that
# only appear on purpose-built emergency vehicles.
EMERGENCY_MODEL_KEYWORDS_STRONG = [
    "ppv",              # Police Pursuit Vehicle
    "ssv",              # Special Service Vehicle
    "special service",
    "responder",
]

# Model/trim keywords that are WEAK signals — common in civilian trims
# (e.g. "Police Interceptor Utility" sold to non-police fleets, Nissan
# "Patrol", Dodge "Pursuit").  These only trigger Category D when the
# vehicle ALSO has a matching emergency department name.
EMERGENCY_MODEL_KEYWORDS_WEAK = [
    "police",
    "interceptor",
    "pursuit",
    "patrol",
]

# Manufacturers that exclusively build emergency apparatus.
# Matched with word boundaries to avoid "hme" matching "scheme".
EMERGENCY_MAKES = [
    "pierce",           # Fire apparatus
    "e-one",
    "spartan",
    "ferrara",
    "rosenbauer",
    "seagrave",
    "american lafrance",
    "sutphen",
    "hme",              # HME Ahrens-Fox
    "kme",              # Kovatch Mobile Equipment
]

EMERGENCY_DEPARTMENT_KEYWORDS = [
    "police",
    "fire dept",
    "fire department",
    "fire district",
    "fire station",
    "fire rescue",
    r"\bems\b",         # word-boundary regex — avoids matching "systems"
    "sheriff",
    "public safety",
    "emergency",
    "marshal",
    "rescue",
    "paramedic",
]

# Light-duty passenger body classes (used as fallback when GVWR is missing)
LIGHT_DUTY_BODY_TYPES = [
    "sedan",
    "coupe",
    "hatchback",
    "wagon",
    "convertible",
    "crossover",
    "sport utility",
    "suv",
    "minivan",
    "passenger car",
]

# ZEV / electric fuel type indicators
ZEV_FUEL_TYPES = [
    "electric",
    "bev",
    "battery electric",
    "fuel cell",
    "hydrogen",
]

# Light-duty GVWR threshold (pounds)
LIGHT_DUTY_MAX_GVWR = 8500


###############################################################################
# Main Classification Function
###############################################################################

def classify_acf_vehicle(vehicle: 'FleetVehicle') -> Tuple[str, str, str]:
    """
    Classify a vehicle into a CARB ACF compliance category.

    Args:
        vehicle: A FleetVehicle with vehicle_id data populated.

    Returns:
        Tuple of (category_code, category_label, detail_text):
          - category_code: "ZEV", "A", "B", "C", or "D"
          - category_label: Human-readable label for the results table
          - detail_text: Explanation of the classification reasoning
    """
    vid = vehicle.vehicle_id
    fuel_lower = vid.fuel_type.lower().strip()
    body_lower = vid.body_class.lower().strip()
    model_lower = vid.model.lower().strip()
    make_lower = vid.make.lower().strip()
    series_lower = vid.series.lower().strip()
    trim_lower = vid.trim.lower().strip()
    dept_lower = vehicle.department.lower().strip()

    # ── 1. Already a zero-emission vehicle? ────────────────────────
    for zev_kw in ZEV_FUEL_TYPES:
        if zev_kw in fuel_lower:
            return (
                "ZEV",
                "Zero-Emission Vehicle",
                "Already ZEV — no action required"
            )

    # ── 2. Light-duty? ─────────────────────────────────────────────
    if _is_light_duty(vid, body_lower):
        gvwr_note = (
            f"GVWR {vid.gvwr_pounds:,.0f} lbs"
            if vid.gvwr_pounds > 0
            else "GVWR not reported; classified by body type"
        )
        return (
            "A",
            "Exempt — Light-Duty",
            f"{gvwr_note}; not subject to ACF"
        )

    # ── 3. Emergency vehicle? ──────────────────────────────────────
    emergency_match = _detect_emergency(
        body_lower, model_lower, make_lower,
        series_lower, trim_lower, dept_lower
    )
    if emergency_match:
        return (
            "D",
            "Emergency Vehicle",
            emergency_match
        )

    # ── 4. Exempt body type? ───────────────────────────────────────
    body_exemption = _detect_exempt_body(body_lower)
    if body_exemption:
        return (
            "C",
            "Exempt — Body Type",
            f"Body type: {vid.body_class}; {body_exemption}"
        )

    # ── 5. Default: subject to ACF ─────────────────────────────────
    gvwr_text = (
        f"GVWR {vid.gvwr_pounds:,.0f} lbs"
        if vid.gvwr_pounds > 0
        else "Medium/heavy-duty"
    )
    return (
        "B",
        "Subject to ACF",
        f"{gvwr_text}; covered by ACF regulation"
    )


###############################################################################
# Helper Functions
###############################################################################

def _is_light_duty(vid, body_lower: str) -> bool:
    """Determine if a vehicle is light-duty (exempt from ACF)."""
    # If GVWR is known, use the 8,500 lb threshold
    if vid.gvwr_pounds > 0:
        return vid.gvwr_pounds <= LIGHT_DUTY_MAX_GVWR

    # GVWR unknown — fall back to body class heuristic
    if vid.commercial_category == "Light Duty":
        return True

    # Check if body class looks like a passenger vehicle
    for passenger_type in LIGHT_DUTY_BODY_TYPES:
        if passenger_type in body_lower:
            return True

    # If we have no GVWR and no recognizable body class, we can't
    # confidently classify as light-duty.  Default to NOT light-duty
    # so the vehicle gets reviewed (safer for compliance purposes).
    return False


def _word_match(keyword: str, text: str) -> bool:
    """
    Check if *keyword* appears in *text* as a whole word / phrase.

    If the keyword already contains regex metacharacters (e.g. ``\\b``)
    it is used as-is; otherwise ``\\b`` word boundaries are added
    automatically.
    """
    if r"\b" in keyword:
        # Keyword is already a regex pattern
        return bool(re.search(keyword, text))
    pattern = r"\b" + re.escape(keyword) + r"\b"
    return bool(re.search(pattern, text))


def _is_emergency_department(dept_lower: str) -> bool:
    """Return True if the department name signals emergency use."""
    if not dept_lower:
        return False
    for kw in EMERGENCY_DEPARTMENT_KEYWORDS:
        if _word_match(kw, dept_lower):
            return True
    return False


def _detect_emergency(
    body_lower: str,
    model_lower: str,
    make_lower: str,
    series_lower: str,
    trim_lower: str,
    dept_lower: str,
) -> str:
    """
    Detect if a vehicle is an emergency vehicle.

    Uses word-boundary matching to avoid false positives from partial
    substring hits (e.g. "fire" inside "Crossfire").  Model keywords
    like "police" and "interceptor" are considered *weak* signals and
    only trigger when paired with an emergency department name.

    Returns:
        A description string if emergency vehicle detected, else empty string.
    """
    # 1. Body class — strong signal
    for kw in EMERGENCY_BODY_KEYWORDS:
        if _word_match(kw, body_lower):
            return f"Body class '{body_lower}' indicates emergency vehicle"

    # 2. Make — emergency apparatus manufacturers
    for kw in EMERGENCY_MAKES:
        if _word_match(kw, make_lower):
            return f"Make '{make_lower}' is an emergency vehicle manufacturer"

    combined_model = f"{model_lower} {series_lower} {trim_lower}"

    # 3. Strong model keywords (unambiguous on their own)
    for kw in EMERGENCY_MODEL_KEYWORDS_STRONG:
        if _word_match(kw, combined_model):
            return f"Model/trim '{combined_model.strip()}' indicates emergency vehicle"

    # 4. Weak model keywords — require a corroborating department signal
    has_emergency_dept = _is_emergency_department(dept_lower)
    if has_emergency_dept:
        for kw in EMERGENCY_MODEL_KEYWORDS_WEAK:
            if _word_match(kw, combined_model):
                return (
                    f"Model/trim '{combined_model.strip()}' + "
                    f"department '{dept_lower}' indicates emergency vehicle"
                )

    # 5. Department alone (no model signal needed for unambiguous depts)
    if has_emergency_dept:
        return f"Department '{dept_lower}' indicates emergency use"

    return ""


def _detect_exempt_body(body_lower: str) -> str:
    """
    Check if the vehicle's body type is on the CARB ZEV Purchase
    Exemption List (ZEV not commercially available in this configuration).

    Uses word-boundary matching to prevent partial hits (e.g. "stake"
    should not match inside "Mistake").

    Returns:
        Exemption reason string if exempt, else empty string.
    """
    for exempt_kw in EXEMPT_BODY_TYPES:
        if _word_match(exempt_kw, body_lower):
            return "ZEV not commercially available for this body configuration"

    return ""
