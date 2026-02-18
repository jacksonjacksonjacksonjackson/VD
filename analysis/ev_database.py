"""
ev_database.py

Static EV equivalent database and per-vehicle replacement recommendation engine.
Maps common fleet ICE vehicles to available EV replacements with pricing, range,
and capability data. Provides matching, recommendation generation, and fleet-wide
priority ranking.

Phase 9B of the Decision-Support & Client-Ready Output initiative.
"""

import logging
import re
from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional, Tuple

from data.models import FleetVehicle, Fleet
from settings import DEFAULT_ANNUAL_MILEAGE

logger = logging.getLogger(__name__)


###############################################################################
# EV Equivalent Data Model
###############################################################################

@dataclass
class EVEquivalent:
    """An available EV replacement option for a class of ICE vehicles."""
    ev_model: str                    # e.g. "Ford F-150 Lightning"
    ev_make: str                     # e.g. "Ford"
    msrp_low: float                  # Low-end MSRP ($)
    msrp_high: float                 # High-end MSRP ($)
    battery_kwh: float               # Battery capacity
    epa_range_miles: int             # EPA estimated range
    body_class: str                  # Matching body class category
    gvwr_min: int = 0               # Min GVWR this EV covers (lbs)
    gvwr_max: int = 99999           # Max GVWR this EV covers (lbs)
    cargo_cu_ft: float = 0.0        # Cargo volume
    towing_lbs: int = 0             # Max towing capacity
    payload_lbs: int = 0            # Max payload capacity
    availability_year: int = 2024    # Year available / expected
    ice_msrp_low: float = 0.0       # Comparable ICE MSRP low
    ice_msrp_high: float = 0.0      # Comparable ICE MSRP high
    notes: str = ""                  # Additional context
    match_keywords: List[str] = field(default_factory=list)  # Keywords for matching


###############################################################################
# EV Equivalents Database
# Organized by vehicle category. Prices are approximate 2024-2025 MSRPs.
###############################################################################

EV_DATABASE: List[EVEquivalent] = [
    # =========================================================================
    # SEDANS / COMPACT
    # =========================================================================
    EVEquivalent(
        ev_model="Tesla Model 3", ev_make="Tesla",
        msrp_low=38990, msrp_high=53990, battery_kwh=60, epa_range_miles=272,
        body_class="sedan", gvwr_max=8500, cargo_cu_ft=23,
        ice_msrp_low=26000, ice_msrp_high=35000,
        match_keywords=["camry", "accord", "altima", "fusion", "malibu", "civic",
                        "corolla", "sentra", "jetta", "elantra", "sonata", "sedan"],
    ),
    EVEquivalent(
        ev_model="Chevrolet Equinox EV", ev_make="Chevrolet",
        msrp_low=33900, msrp_high=43900, battery_kwh=85, epa_range_miles=319,
        body_class="sedan", gvwr_max=8500, cargo_cu_ft=26,
        ice_msrp_low=28000, ice_msrp_high=36000,
        match_keywords=["equinox", "escape", "rav4", "cr-v", "rogue", "tucson",
                        "compact suv", "crossover"],
    ),
    EVEquivalent(
        ev_model="Nissan Leaf / Ariya", ev_make="Nissan",
        msrp_low=28140, msrp_high=44000, battery_kwh=65, epa_range_miles=304,
        body_class="sedan", gvwr_max=8500,
        ice_msrp_low=22000, ice_msrp_high=32000,
        match_keywords=["leaf", "versa", "sentra", "hatchback", "compact"],
    ),

    # =========================================================================
    # SUVs / CROSSOVERS
    # =========================================================================
    EVEquivalent(
        ev_model="Ford Mustang Mach-E", ev_make="Ford",
        msrp_low=42995, msrp_high=63995, battery_kwh=91, epa_range_miles=312,
        body_class="suv", gvwr_max=8500, cargo_cu_ft=30, towing_lbs=3500,
        ice_msrp_low=35000, ice_msrp_high=52000,
        match_keywords=["explorer", "edge", "escape", "highlander", "4runner",
                        "pathfinder", "pilot", "blazer", "suv", "sport utility"],
    ),
    EVEquivalent(
        ev_model="Chevrolet Blazer EV", ev_make="Chevrolet",
        msrp_low=44995, msrp_high=61995, battery_kwh=102, epa_range_miles=324,
        body_class="suv", gvwr_max=8500, cargo_cu_ft=26, towing_lbs=4000,
        ice_msrp_low=36000, ice_msrp_high=50000,
        match_keywords=["blazer", "traverse", "equinox", "tahoe", "suburban"],
    ),
    EVEquivalent(
        ev_model="Tesla Model Y", ev_make="Tesla",
        msrp_low=44990, msrp_high=58990, battery_kwh=75, epa_range_miles=310,
        body_class="suv", gvwr_max=8500, cargo_cu_ft=68, towing_lbs=3500,
        ice_msrp_low=32000, ice_msrp_high=48000,
        match_keywords=["rav4", "cr-v", "forester", "outback", "cx-5",
                        "sportage", "crossover"],
    ),
    EVEquivalent(
        ev_model="Volkswagen ID.4", ev_make="Volkswagen",
        msrp_low=38995, msrp_high=53545, battery_kwh=82, epa_range_miles=275,
        body_class="suv", gvwr_max=8500, cargo_cu_ft=30,
        ice_msrp_low=30000, ice_msrp_high=42000,
        match_keywords=["tiguan", "atlas", "cx-50", "seltos"],
    ),

    # =========================================================================
    # PICKUP TRUCKS - Light Duty
    # =========================================================================
    EVEquivalent(
        ev_model="Ford F-150 Lightning", ev_make="Ford",
        msrp_low=49995, msrp_high=96995, battery_kwh=131, epa_range_miles=320,
        body_class="pickup", gvwr_min=0, gvwr_max=8500,
        towing_lbs=10000, payload_lbs=2000,
        ice_msrp_low=35000, ice_msrp_high=75000,
        match_keywords=["f-150", "f150", "f 150", "silverado 1500", "sierra 1500",
                        "ram 1500", "tundra", "titan", "light duty pickup",
                        "pickup", "crew cab"],
    ),
    EVEquivalent(
        ev_model="Chevrolet Silverado EV", ev_make="Chevrolet",
        msrp_low=57095, msrp_high=99995, battery_kwh=200, epa_range_miles=450,
        body_class="pickup", gvwr_min=0, gvwr_max=8500,
        towing_lbs=10000, payload_lbs=1800,
        ice_msrp_low=38000, ice_msrp_high=70000,
        match_keywords=["silverado", "sierra", "colorado", "canyon"],
    ),
    EVEquivalent(
        ev_model="Ram 1500 REV", ev_make="Ram",
        msrp_low=58995, msrp_high=75000, battery_kwh=168, epa_range_miles=350,
        body_class="pickup", gvwr_min=0, gvwr_max=8500,
        towing_lbs=14000, payload_lbs=2700,
        ice_msrp_low=38000, ice_msrp_high=65000,
        availability_year=2025,
        match_keywords=["ram 1500", "ram 2500"],
    ),

    # =========================================================================
    # PICKUP TRUCKS - Medium / Heavy Duty
    # =========================================================================
    EVEquivalent(
        ev_model="Ford F-600 / E-Transit Chassis", ev_make="Ford",
        msrp_low=75000, msrp_high=110000, battery_kwh=98, epa_range_miles=159,
        body_class="pickup", gvwr_min=8501, gvwr_max=22000,
        towing_lbs=7700, payload_lbs=3800,
        ice_msrp_low=45000, ice_msrp_high=75000,
        match_keywords=["f-250", "f250", "f-350", "f350", "f-450", "f450",
                        "f-550", "f550", "f-600", "silverado 2500",
                        "silverado 3500", "sierra 2500", "sierra 3500",
                        "ram 2500", "ram 3500", "medium duty pickup"],
    ),

    # =========================================================================
    # CARGO VANS
    # =========================================================================
    EVEquivalent(
        ev_model="Ford E-Transit", ev_make="Ford",
        msrp_low=51480, msrp_high=64580, battery_kwh=89, epa_range_miles=159,
        body_class="van", gvwr_max=14500, cargo_cu_ft=488,
        payload_lbs=3880,
        ice_msrp_low=42000, ice_msrp_high=55000,
        match_keywords=["transit", "e-transit", "cargo van", "transit 250",
                        "transit 350", "promaster", "sprinter", "nv200",
                        "nv2500", "nv3500"],
    ),
    EVEquivalent(
        ev_model="BrightDrop Zevo 600", ev_make="BrightDrop",
        msrp_low=60000, msrp_high=80000, battery_kwh=89, epa_range_miles=250,
        body_class="van", gvwr_max=14500, cargo_cu_ft=600,
        payload_lbs=2500,
        ice_msrp_low=38000, ice_msrp_high=55000,
        match_keywords=["express", "savana", "express 2500", "express 3500",
                        "savana 2500", "savana 3500", "cargo van"],
    ),
    EVEquivalent(
        ev_model="Mercedes eSprinter", ev_make="Mercedes-Benz",
        msrp_low=59900, msrp_high=74900, battery_kwh=113, epa_range_miles=260,
        body_class="van", gvwr_max=14500, cargo_cu_ft=488,
        payload_lbs=3200,
        ice_msrp_low=45000, ice_msrp_high=60000,
        match_keywords=["sprinter", "metris"],
    ),

    # =========================================================================
    # PASSENGER VANS / MINIVANS
    # =========================================================================
    EVEquivalent(
        ev_model="Volkswagen ID. Buzz", ev_make="Volkswagen",
        msrp_low=59995, msrp_high=69995, battery_kwh=91, epa_range_miles=234,
        body_class="van", gvwr_max=8500,
        ice_msrp_low=35000, ice_msrp_high=50000,
        match_keywords=["minivan", "sienna", "odyssey", "pacifica", "grand caravan",
                        "passenger van"],
    ),
    EVEquivalent(
        ev_model="Chrysler Pacifica PHEV", ev_make="Chrysler",
        msrp_low=52400, msrp_high=56400, battery_kwh=16, epa_range_miles=32,
        body_class="van", gvwr_max=8500,
        ice_msrp_low=38000, ice_msrp_high=50000,
        notes="PHEV, not full BEV. 32-mile electric range.",
        match_keywords=["pacifica", "town & country", "caravan"],
    ),

    # =========================================================================
    # MEDIUM-DUTY TRUCKS (Class 4-6, 14,001-26,000 lbs GVWR)
    # =========================================================================
    EVEquivalent(
        ev_model="Ford E-Transit Cutaway/Chassis", ev_make="Ford",
        msrp_low=55000, msrp_high=70000, battery_kwh=89, epa_range_miles=126,
        body_class="truck", gvwr_min=10000, gvwr_max=16000,
        payload_lbs=4400,
        ice_msrp_low=40000, ice_msrp_high=55000,
        match_keywords=["cutaway", "cab chassis", "box truck", "e-450",
                        "e450", "transit cutaway"],
    ),
    EVEquivalent(
        ev_model="Lightning eMotors Class 4-5", ev_make="Lightning eMotors",
        msrp_low=95000, msrp_high=150000, battery_kwh=120, epa_range_miles=120,
        body_class="truck", gvwr_min=14001, gvwr_max=19500,
        payload_lbs=6000,
        ice_msrp_low=55000, ice_msrp_high=85000,
        match_keywords=["class 4", "class 5", "hino", "isuzu", "npr", "nqr",
                        "fuso", "medium duty"],
    ),
    EVEquivalent(
        ev_model="Hino L6 EV / Class 6", ev_make="Hino",
        msrp_low=130000, msrp_high=180000, battery_kwh=148, epa_range_miles=124,
        body_class="truck", gvwr_min=19501, gvwr_max=26000,
        payload_lbs=8000,
        ice_msrp_low=65000, ice_msrp_high=100000,
        match_keywords=["class 6", "hino l6", "peterbilt 220ev", "freightliner em2"],
    ),

    # =========================================================================
    # HEAVY-DUTY TRUCKS (Class 7-8, >26,000 lbs GVWR)
    # =========================================================================
    EVEquivalent(
        ev_model="Freightliner eCascadia", ev_make="Freightliner",
        msrp_low=250000, msrp_high=400000, battery_kwh=438, epa_range_miles=230,
        body_class="truck", gvwr_min=26001, gvwr_max=82000,
        towing_lbs=82000,
        ice_msrp_low=130000, ice_msrp_high=180000,
        match_keywords=["class 7", "class 8", "cascadia", "semi", "tractor",
                        "freightliner", "kenworth", "peterbilt", "volvo vnr",
                        "international lt", "heavy duty"],
    ),
    EVEquivalent(
        ev_model="Peterbilt 579EV", ev_make="Peterbilt",
        msrp_low=280000, msrp_high=400000, battery_kwh=396, epa_range_miles=150,
        body_class="truck", gvwr_min=26001, gvwr_max=82000,
        ice_msrp_low=140000, ice_msrp_high=185000,
        match_keywords=["peterbilt", "579", "389"],
    ),
    EVEquivalent(
        ev_model="Volvo VNR Electric", ev_make="Volvo",
        msrp_low=260000, msrp_high=390000, battery_kwh=565, epa_range_miles=275,
        body_class="truck", gvwr_min=26001, gvwr_max=82000,
        ice_msrp_low=130000, ice_msrp_high=175000,
        match_keywords=["volvo", "vnr", "vnl", "mack"],
    ),

    # =========================================================================
    # BUSES
    # =========================================================================
    EVEquivalent(
        ev_model="Proterra ZX5 / BYD K9", ev_make="Proterra / BYD",
        msrp_low=550000, msrp_high=900000, battery_kwh=660, epa_range_miles=329,
        body_class="bus", gvwr_min=25000, gvwr_max=44000,
        ice_msrp_low=300000, ice_msrp_high=500000,
        match_keywords=["bus", "transit bus", "school bus", "shuttle",
                        "blue bird", "thomas", "ic bus", "gillig"],
    ),
    EVEquivalent(
        ev_model="GreenPower BEAST / Lion Electric", ev_make="GreenPower / Lion",
        msrp_low=350000, msrp_high=450000, battery_kwh=155, epa_range_miles=140,
        body_class="bus", gvwr_min=14000, gvwr_max=36000,
        ice_msrp_low=100000, ice_msrp_high=200000,
        notes="School bus segment",
        match_keywords=["school bus", "type c bus", "type d bus", "activity bus"],
    ),

    # =========================================================================
    # SPECIALTY / UTILITY
    # =========================================================================
    EVEquivalent(
        ev_model="Lightning eMotors Class 3 Utility", ev_make="Lightning eMotors",
        msrp_low=75000, msrp_high=120000, battery_kwh=80, epa_range_miles=110,
        body_class="truck", gvwr_min=8501, gvwr_max=14000,
        ice_msrp_low=45000, ice_msrp_high=75000,
        match_keywords=["utility", "service body", "bucket truck", "class 3",
                        "work truck", "service truck"],
    ),
]


###############################################################################
# Matching Engine
###############################################################################

def _normalize(text: str) -> str:
    """Lowercase, strip punctuation for matching."""
    return re.sub(r'[^a-z0-9 ]', '', (text or '').lower()).strip()


def _classify_body(vehicle: FleetVehicle) -> str:
    """Classify vehicle into broad body category for EV matching."""
    body = _normalize(vehicle.vehicle_id.body_class or '')
    model = _normalize(vehicle.vehicle_id.model or '')
    make = _normalize(vehicle.vehicle_id.make or '')

    # Bus detection
    if 'bus' in body or 'bus' in model:
        return 'bus'

    # Van detection
    van_keywords = ['van', 'transit', 'sprinter', 'promaster', 'express', 'savana',
                    'metris', 'nv200', 'nv2500', 'nv3500']
    if any(k in body for k in van_keywords) or any(k in model for k in van_keywords):
        return 'van'

    # Pickup detection
    pickup_keywords = ['pickup', 'crew cab', 'regular cab', 'extended cab',
                       'double cab', 'king cab', 'quad cab']
    pickup_models = ['f150', 'f 150', 'f250', 'f 250', 'f350', 'f 350',
                     'f450', 'f 450', 'f550', 'f 550',
                     'silverado', 'sierra', 'ram 1500', 'ram 2500', 'ram 3500',
                     'tundra', 'titan', 'colorado', 'canyon', 'tacoma',
                     'ranger', 'frontier', 'ridgeline', 'maverick', 'gladiator']
    if any(k in body for k in pickup_keywords) or any(k in model for k in pickup_models):
        return 'pickup'

    # Truck / cab chassis detection
    truck_keywords = ['truck', 'cab chassis', 'cutaway', 'chassis cab',
                      'incomplete', 'stripped chassis', 'stake', 'flatbed',
                      'dump', 'box truck']
    if any(k in body for k in truck_keywords):
        return 'truck'

    # SUV detection
    suv_keywords = ['suv', 'sport utility', 'utility']
    if any(k in body for k in suv_keywords):
        return 'suv'

    # Sedan / passenger car fallback
    sedan_keywords = ['sedan', 'coupe', 'hatchback', 'wagon', 'convertible',
                      'passenger car']
    if any(k in body for k in sedan_keywords):
        return 'sedan'

    # Default by GVWR
    gvwr = vehicle.vehicle_id.gvwr_pounds or 0
    if gvwr > 26000:
        return 'truck'
    elif gvwr > 14000:
        return 'truck'
    elif gvwr > 8500:
        return 'truck'
    else:
        return 'sedan'


def find_ev_equivalent(vehicle: FleetVehicle) -> Optional[Dict[str, Any]]:
    """
    Find the best EV equivalent for a given vehicle.

    Returns dict with:
        ev: EVEquivalent object
        fit_score: 0-100 match quality
        ev_price: midpoint MSRP
        ice_price: midpoint ICE MSRP
        rationale: human-readable match explanation
    Or None if no match found.
    """
    body_category = _classify_body(vehicle)
    gvwr = vehicle.vehicle_id.gvwr_pounds or 0
    model_norm = _normalize(vehicle.vehicle_id.model or '')
    make_norm = _normalize(vehicle.vehicle_id.make or '')
    combined = f"{make_norm} {model_norm}"

    best_match = None
    best_score = -1

    for ev in EV_DATABASE:
        score = 0

        # Body class match (required — skip if wrong category)
        if ev.body_class != body_category:
            continue

        score += 30  # Base body class match

        # GVWR range match
        if gvwr > 0 and ev.gvwr_min <= gvwr <= ev.gvwr_max:
            score += 25
        elif gvwr == 0:
            score += 10  # No GVWR data — partial credit
        else:
            continue  # GVWR outside range — skip

        # Keyword match (model/make name matching)
        keyword_matches = sum(1 for kw in ev.match_keywords if kw in combined)
        if keyword_matches >= 2:
            score += 30
        elif keyword_matches == 1:
            score += 20
        else:
            score += 5  # Category match but no keyword match

        # Make match bonus
        if ev.ev_make.lower() == make_norm:
            score += 10

        # Towing/payload capability match
        if ev.towing_lbs > 0 and body_category in ('pickup', 'truck'):
            score += 5

        if score > best_score:
            best_score = score
            best_match = ev

    if best_match is None:
        return None

    ev_price = (best_match.msrp_low + best_match.msrp_high) / 2
    ice_price = (best_match.ice_msrp_low + best_match.ice_msrp_high) / 2

    # Build rationale
    parts = []
    if body_category == 'pickup':
        parts.append(f"Pickup replacement")
    elif body_category == 'van':
        parts.append(f"Van replacement")
    elif body_category == 'truck':
        parts.append(f"Truck replacement")
    elif body_category == 'bus':
        parts.append(f"Bus replacement")
    elif body_category == 'suv':
        parts.append(f"SUV replacement")
    else:
        parts.append(f"Sedan replacement")

    parts.append(f"{best_match.epa_range_miles} mi range")

    if best_match.towing_lbs > 0:
        parts.append(f"{best_match.towing_lbs:,} lb towing")

    if best_match.notes:
        parts.append(best_match.notes)

    return {
        "ev": best_match,
        "ev_model": best_match.ev_model,
        "ev_make": best_match.ev_make,
        "fit_score": min(100, best_score),
        "ev_price": ev_price,
        "ice_price": ice_price,
        "ev_msrp_range": f"${best_match.msrp_low:,.0f}-${best_match.msrp_high:,.0f}",
        "epa_range": best_match.epa_range_miles,
        "battery_kwh": best_match.battery_kwh,
        "rationale": " | ".join(parts),
    }


###############################################################################
# Replacement Recommendations
###############################################################################

def generate_replacement_recommendation(
    vehicle: FleetVehicle,
    ev_match: Dict[str, Any],
    savings: Optional[Dict[str, float]] = None,
) -> Dict[str, Any]:
    """
    Generate a structured replacement recommendation combining vehicle data,
    EV match, and financial analysis.
    """
    annual_mileage = vehicle.annual_mileage or DEFAULT_ANNUAL_MILEAGE
    year = vehicle.vehicle_id.year or 0
    age = 2025 - year if year else 0

    # Rationale components
    rationale_parts = []
    if age >= 10:
        rationale_parts.append(f"aging ({age} years old)")
    if annual_mileage >= 15000:
        rationale_parts.append(f"high utilization ({annual_mileage:,.0f} mi/yr)")

    acf_code = vehicle.custom_fields.get("_acf_code", "")
    if acf_code == "B":
        rationale_parts.append("ACF compliance required")
    ev_year = vehicle.custom_fields.get("Proposed EV Year", "")
    if ev_year and ev_year not in ("N/A", "Exempt", ""):
        rationale_parts.append(f"scheduled for {ev_year}")

    if savings and savings.get("total_npv_savings", 0) > 0:
        rationale_parts.append(f"${savings['total_npv_savings']:,.0f} NPV savings")

    rec = {
        "vin": vehicle.vin,
        "current_vehicle": f"{vehicle.vehicle_id.year or ''} {vehicle.vehicle_id.make or ''} {vehicle.vehicle_id.model or ''}".strip(),
        "department": vehicle.department or "Unassigned",
        "annual_mileage": annual_mileage,
        "proposed_ev": ev_match["ev_model"],
        "ev_msrp_range": ev_match["ev_msrp_range"],
        "ev_price": ev_match["ev_price"],
        "ice_price": ev_match["ice_price"],
        "epa_range": ev_match["epa_range"],
        "fit_score": ev_match["fit_score"],
        "proposed_year": ev_year or "TBD",
        "rationale": "; ".join(rationale_parts) if rationale_parts else "Standard replacement candidate",
    }

    if savings:
        rec["annual_fuel_savings"] = savings.get("annual_fuel_savings", 0)
        rec["annual_maintenance_savings"] = savings.get("annual_maintenance_savings", 0)
        rec["total_npv_savings"] = savings.get("total_npv_savings", 0)
        rec["total_co2_reduction"] = savings.get("total_co2_reduction", 0)
        annual_total = savings.get("annual_fuel_savings", 0) + savings.get("annual_maintenance_savings", 0)
        premium = ev_match["ev_price"] - ev_match["ice_price"]
        rec["payback_years"] = round(premium / annual_total, 1) if annual_total > 0 else float('inf')
    else:
        rec["payback_years"] = None

    return rec


def get_priority_replacements(
    fleet: Fleet,
    electrification_analysis: Optional[Any] = None,
    n: int = 15,
    gas_price: float = 3.50,
    electricity_price: float = 0.13,
) -> List[Dict[str, Any]]:
    """
    Return top N priority replacement recommendations for a fleet.
    Sorted by total NPV savings (highest first).
    """
    from analysis.calculations import calculate_electrification_savings

    recommendations = []

    for vehicle in fleet.vehicles:
        if not vehicle.processing_success:
            continue
        if not vehicle.fuel_economy.combined_mpg:
            continue

        # Skip vehicles already electric
        fuel_type = _normalize(vehicle.vehicle_id.fuel_type or '')
        if any(kw in fuel_type for kw in ['electric', 'bev', 'battery', 'fuel cell']):
            continue

        # Find EV equivalent
        ev_match = find_ev_equivalent(vehicle)
        if ev_match is None:
            continue

        # Get savings from electrification analysis if available
        savings = None
        if electrification_analysis and vehicle.vin in electrification_analysis.vehicle_results:
            savings = electrification_analysis.vehicle_results[vehicle.vin]
        else:
            # Calculate on the fly
            savings = calculate_electrification_savings(
                vehicle=vehicle,
                gas_price=gas_price,
                electricity_price=electricity_price,
            )

        rec = generate_replacement_recommendation(vehicle, ev_match, savings)
        recommendations.append(rec)

    # Sort by NPV savings descending
    recommendations.sort(
        key=lambda r: r.get("total_npv_savings", 0),
        reverse=True
    )

    return recommendations[:n]


def match_fleet_ev_equivalents(fleet: Fleet) -> Dict[str, Dict[str, Any]]:
    """
    Match all vehicles in a fleet to EV equivalents.
    Returns dict of VIN -> match result (or None if no match).
    Used during processing to populate custom_fields.
    """
    results = {}
    for vehicle in fleet.vehicles:
        if not vehicle.processing_success:
            results[vehicle.vin] = None
            continue

        fuel_type = _normalize(vehicle.vehicle_id.fuel_type or '')
        if any(kw in fuel_type for kw in ['electric', 'bev', 'battery', 'fuel cell']):
            results[vehicle.vin] = None
            continue

        match = find_ev_equivalent(vehicle)
        results[vehicle.vin] = match

        # Store pricing in custom_fields for TCO computation
        if match:
            vehicle.custom_fields["_ev_purchase_price"] = match["ev_price"]
            vehicle.custom_fields["_ice_purchase_price"] = match["ice_price"]
            vehicle.custom_fields["EV Equivalent"] = match["ev_model"]
            vehicle.custom_fields["EV MSRP Range"] = match["ev_msrp_range"]
            vehicle.custom_fields["EV EPA Range"] = f"{match['epa_range']} mi"
            vehicle.custom_fields["EV Fit Score"] = f"{match['fit_score']}%"

    matched = sum(1 for v in results.values() if v is not None)
    logger.info(f"EV matching: {matched}/{len(results)} vehicles matched to EV equivalents")

    return results
