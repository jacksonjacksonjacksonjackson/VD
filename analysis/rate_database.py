"""
analysis/rate_database.py

State-level energy rates and available incentives for fleet electrification.

Provides default gas/electricity prices per state and federal/state incentive
programs so TCO calculations can reflect local conditions instead of relying
on national averages.

Phase 9H of the Decision-Support & Client-Ready Output initiative.
"""

import logging
from typing import Dict, List, Any, Optional

logger = logging.getLogger(__name__)


###############################################################################
# State Energy Rates (approximate averages, $/unit)
###############################################################################

# Source: EIA state-level averages (rounded for usability)
# gas_price: $/gallon regular unleaded
# electricity_price: $/kWh commercial rate
# demand_charge: $/kW/month (for DCFC installations)

STATE_ENERGY_RATES: Dict[str, Dict[str, float]] = {
    "AL": {"gas_price": 3.10, "electricity_price": 0.11, "demand_charge": 12.0},
    "AK": {"gas_price": 4.20, "electricity_price": 0.22, "demand_charge": 18.0},
    "AZ": {"gas_price": 3.40, "electricity_price": 0.11, "demand_charge": 14.0},
    "AR": {"gas_price": 3.05, "electricity_price": 0.10, "demand_charge": 11.0},
    "CA": {"gas_price": 4.80, "electricity_price": 0.20, "demand_charge": 20.0},
    "CO": {"gas_price": 3.30, "electricity_price": 0.12, "demand_charge": 13.0},
    "CT": {"gas_price": 3.60, "electricity_price": 0.21, "demand_charge": 17.0},
    "DE": {"gas_price": 3.30, "electricity_price": 0.13, "demand_charge": 13.0},
    "FL": {"gas_price": 3.40, "electricity_price": 0.12, "demand_charge": 12.0},
    "GA": {"gas_price": 3.15, "electricity_price": 0.11, "demand_charge": 12.0},
    "HI": {"gas_price": 4.90, "electricity_price": 0.33, "demand_charge": 25.0},
    "ID": {"gas_price": 3.50, "electricity_price": 0.09, "demand_charge": 10.0},
    "IL": {"gas_price": 3.60, "electricity_price": 0.12, "demand_charge": 14.0},
    "IN": {"gas_price": 3.25, "electricity_price": 0.11, "demand_charge": 12.0},
    "IA": {"gas_price": 3.15, "electricity_price": 0.11, "demand_charge": 11.0},
    "KS": {"gas_price": 3.10, "electricity_price": 0.11, "demand_charge": 12.0},
    "KY": {"gas_price": 3.10, "electricity_price": 0.10, "demand_charge": 11.0},
    "LA": {"gas_price": 3.05, "electricity_price": 0.10, "demand_charge": 11.0},
    "ME": {"gas_price": 3.50, "electricity_price": 0.17, "demand_charge": 15.0},
    "MD": {"gas_price": 3.40, "electricity_price": 0.14, "demand_charge": 14.0},
    "MA": {"gas_price": 3.55, "electricity_price": 0.22, "demand_charge": 18.0},
    "MI": {"gas_price": 3.35, "electricity_price": 0.13, "demand_charge": 13.0},
    "MN": {"gas_price": 3.25, "electricity_price": 0.12, "demand_charge": 12.0},
    "MS": {"gas_price": 3.00, "electricity_price": 0.10, "demand_charge": 11.0},
    "MO": {"gas_price": 3.05, "electricity_price": 0.10, "demand_charge": 11.0},
    "MT": {"gas_price": 3.40, "electricity_price": 0.10, "demand_charge": 11.0},
    "NE": {"gas_price": 3.15, "electricity_price": 0.10, "demand_charge": 11.0},
    "NV": {"gas_price": 3.80, "electricity_price": 0.11, "demand_charge": 13.0},
    "NH": {"gas_price": 3.40, "electricity_price": 0.19, "demand_charge": 16.0},
    "NJ": {"gas_price": 3.35, "electricity_price": 0.16, "demand_charge": 15.0},
    "NM": {"gas_price": 3.25, "electricity_price": 0.11, "demand_charge": 12.0},
    "NY": {"gas_price": 3.65, "electricity_price": 0.19, "demand_charge": 18.0},
    "NC": {"gas_price": 3.20, "electricity_price": 0.11, "demand_charge": 12.0},
    "ND": {"gas_price": 3.20, "electricity_price": 0.10, "demand_charge": 10.0},
    "OH": {"gas_price": 3.25, "electricity_price": 0.12, "demand_charge": 13.0},
    "OK": {"gas_price": 3.00, "electricity_price": 0.10, "demand_charge": 11.0},
    "OR": {"gas_price": 3.80, "electricity_price": 0.10, "demand_charge": 11.0},
    "PA": {"gas_price": 3.50, "electricity_price": 0.14, "demand_charge": 14.0},
    "RI": {"gas_price": 3.50, "electricity_price": 0.21, "demand_charge": 17.0},
    "SC": {"gas_price": 3.10, "electricity_price": 0.11, "demand_charge": 12.0},
    "SD": {"gas_price": 3.25, "electricity_price": 0.11, "demand_charge": 11.0},
    "TN": {"gas_price": 3.10, "electricity_price": 0.10, "demand_charge": 11.0},
    "TX": {"gas_price": 3.10, "electricity_price": 0.10, "demand_charge": 11.0},
    "UT": {"gas_price": 3.40, "electricity_price": 0.10, "demand_charge": 11.0},
    "VT": {"gas_price": 3.50, "electricity_price": 0.18, "demand_charge": 16.0},
    "VA": {"gas_price": 3.25, "electricity_price": 0.12, "demand_charge": 12.0},
    "WA": {"gas_price": 4.00, "electricity_price": 0.10, "demand_charge": 12.0},
    "WV": {"gas_price": 3.20, "electricity_price": 0.11, "demand_charge": 12.0},
    "WI": {"gas_price": 3.20, "electricity_price": 0.13, "demand_charge": 13.0},
    "WY": {"gas_price": 3.35, "electricity_price": 0.10, "demand_charge": 10.0},
    "DC": {"gas_price": 3.70, "electricity_price": 0.14, "demand_charge": 16.0},
}


###############################################################################
# Incentive Programs
###############################################################################

# Each incentive: name, amount ($ or % of vehicle price), vehicle_class filter,
# description, expiry note

FEDERAL_INCENTIVES = [
    {
        "name": "IRA Commercial Clean Vehicle Credit (45W)",
        "max_amount": 40000,
        "amount_rule": "lesser_of_pct_or_max",
        "pct_of_incremental": 30,  # 30% of incremental cost over ICE equivalent
        "vehicle_class": "all",  # light, medium, heavy, all
        "description": "Up to $40,000 for commercial clean vehicles (30% of incremental cost).",
        "note": "No income cap for commercial. Must be used in trade/business.",
    },
    {
        "name": "IRA Consumer Clean Vehicle Credit (30D)",
        "max_amount": 7500,
        "amount_rule": "fixed_max",
        "vehicle_class": "light",
        "description": "Up to $7,500 for new light-duty EVs meeting domestic content requirements.",
        "note": "Subject to MSRP caps ($55K cars, $80K trucks/SUVs) and income limits.",
    },
    {
        "name": "IRA Alternative Fuel Infrastructure Credit (30C)",
        "max_amount": 100000,
        "amount_rule": "pct_of_cost",
        "pct_of_cost": 30,
        "vehicle_class": "infrastructure",
        "description": "30% of charging infrastructure cost, up to $100,000 per location.",
        "note": "Must be in eligible census tract. Expires 2032.",
    },
]

# State incentives — top programs by dollar value
STATE_INCENTIVES: Dict[str, List[Dict[str, Any]]] = {
    "CA": [
        {
            "name": "HVIP (Hybrid & Zero-Emission Truck/Bus Voucher)",
            "max_amount": 120000,
            "vehicle_class": "medium_heavy",
            "description": "Vouchers for zero-emission trucks/buses. Amount varies by weight class.",
        },
        {
            "name": "CVRP (Clean Vehicle Rebate)",
            "max_amount": 7500,
            "vehicle_class": "light",
            "description": "Rebate for light-duty ZEVs for public fleets.",
        },
    ],
    "NY": [
        {
            "name": "NY Truck Voucher Incentive Program (NYTVIP)",
            "max_amount": 185000,
            "vehicle_class": "medium_heavy",
            "description": "Vouchers for Class 3-8 zero-emission vehicles.",
        },
        {
            "name": "Drive Clean Rebate",
            "max_amount": 2000,
            "vehicle_class": "light",
            "description": "$2,000 rebate for new EV purchases/leases.",
        },
    ],
    "CO": [
        {
            "name": "Colorado EV Tax Credit",
            "max_amount": 5000,
            "vehicle_class": "light",
            "description": "State tax credit for new light-duty EV purchases.",
        },
    ],
    "NJ": [
        {
            "name": "NJ Charge Up",
            "max_amount": 4000,
            "vehicle_class": "light",
            "description": "Point-of-sale incentive for new EV purchases.",
        },
    ],
    "MA": [
        {
            "name": "MOR-EV Rebate",
            "max_amount": 3500,
            "vehicle_class": "light",
            "description": "Rebate for battery-electric or fuel cell vehicles.",
        },
    ],
    "OR": [
        {
            "name": "Oregon Clean Vehicle Rebate",
            "max_amount": 7500,
            "vehicle_class": "light",
            "description": "Rebate for qualifying zero-emission vehicles.",
        },
    ],
    "TX": [
        {
            "name": "Texas Emissions Reduction Plan (TERP)",
            "max_amount": 60000,
            "vehicle_class": "medium_heavy",
            "description": "Grants for replacing older diesel vehicles with zero-emission.",
        },
    ],
    "WA": [
        {
            "name": "WA Clean Vehicles Sales Tax Exemption",
            "max_amount": 0,  # % based
            "vehicle_class": "all",
            "description": "Sales and use tax exemption for new/used EVs (up to $45K MSRP).",
        },
    ],
    "IL": [
        {
            "name": "Illinois EV Rebate",
            "max_amount": 4000,
            "vehicle_class": "light",
            "description": "Rebate for new EV purchases by Illinois residents.",
        },
    ],
}


###############################################################################
# Public API
###############################################################################

def get_rates_for_state(state_code: str) -> Dict[str, Any]:
    """
    Get energy rates for a specific state.

    Args:
        state_code: Two-letter state abbreviation (e.g., "CA", "NY")

    Returns:
        Dict with gas_price, electricity_price, demand_charge.
        Falls back to national averages if state not found.
    """
    state_code = state_code.upper().strip()
    rates = STATE_ENERGY_RATES.get(state_code)

    if rates:
        return {
            "state": state_code,
            "gas_price": rates["gas_price"],
            "electricity_price": rates["electricity_price"],
            "demand_charge": rates["demand_charge"],
            "source": "EIA state average (approximate)",
        }

    # National average fallback
    logger.warning(f"No rate data for state '{state_code}'; using national averages")
    return {
        "state": state_code,
        "gas_price": 3.50,
        "electricity_price": 0.13,
        "demand_charge": 13.0,
        "source": "National average (state data not available)",
    }


def get_federal_incentives(vehicle_class: str = "all") -> List[Dict[str, Any]]:
    """
    Get applicable federal incentive programs.

    Args:
        vehicle_class: "light", "medium_heavy", "infrastructure", or "all"

    Returns:
        List of applicable incentive dicts.
    """
    results = []
    for incentive in FEDERAL_INCENTIVES:
        ic = incentive["vehicle_class"]
        if ic == "all" or ic == vehicle_class or vehicle_class == "all":
            results.append(incentive)
    return results


def get_state_incentives(state_code: str, vehicle_class: str = "all") -> List[Dict[str, Any]]:
    """
    Get applicable state incentive programs.

    Args:
        state_code: Two-letter state abbreviation
        vehicle_class: "light", "medium_heavy", or "all"

    Returns:
        List of applicable incentive dicts. Empty list if no programs found.
    """
    state_code = state_code.upper().strip()
    programs = STATE_INCENTIVES.get(state_code, [])

    if vehicle_class == "all":
        return programs

    return [
        p for p in programs
        if p["vehicle_class"] == vehicle_class or p["vehicle_class"] == "all"
    ]


def get_all_incentives(state_code: str, vehicle_class: str = "all") -> Dict[str, Any]:
    """
    Get all applicable incentives (federal + state) with totals.

    Args:
        state_code: Two-letter state abbreviation
        vehicle_class: "light", "medium_heavy", "infrastructure", or "all"

    Returns:
        Dict with federal_incentives, state_incentives, max_federal, max_state, max_total.
    """
    federal = get_federal_incentives(vehicle_class)
    state = get_state_incentives(state_code, vehicle_class)

    max_federal = sum(i.get("max_amount", 0) for i in federal)
    max_state = sum(i.get("max_amount", 0) for i in state)

    return {
        "federal_incentives": federal,
        "state_incentives": state,
        "max_federal": max_federal,
        "max_state": max_state,
        "max_total": max_federal + max_state,
    }


def get_available_states() -> List[str]:
    """Return sorted list of state codes with rate data."""
    return sorted(STATE_ENERGY_RATES.keys())
