"""
analysis/scenarios.py

Scenario comparison engine for fleet electrification timelines.

Allows users to compare multiple "what-if" scenarios side-by-side:
- Different end years (aggressive 2030, moderate 2035, conservative 2040)
- ACF-only compliance scenarios
- Budget-constrained scenarios
- Custom vehicle filters

Each scenario produces yearly metrics (vehicles electrified, cost, CO₂ reduction)
that can be charted and compared.

Phase 9E of the Decision-Support & Client-Ready Output initiative.
"""

import logging
import math
import datetime
from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional, Callable, Tuple

from data.models import FleetVehicle, Fleet
from analysis.electrification_timeline import (
    _score_vehicle, ACF_BOOST,
    WEIGHT_AGE, WEIGHT_MILEAGE, WEIGHT_ANNUAL_USAGE,
    MAX_AGE_YEARS, MAX_ODOMETER, MAX_ANNUAL_MILEAGE,
)

logger = logging.getLogger(__name__)


###############################################################################
# Scenario Data Model
###############################################################################

@dataclass
class ElectrificationScenario:
    """A single electrification timeline scenario with configurable parameters."""
    name: str                          # Display name (e.g., "Aggressive")
    end_year: int = 2040               # Final year of timeline
    vehicle_filter: str = "all"        # "all", "acf_only", "medium_heavy_only",
                                       # "all_except_emergency"
    budget_per_year: float = 0.0       # Max $ per year (0 = unlimited)
    custom_weights: Optional[Dict[str, float]] = None  # Override scoring weights
    description: str = ""              # Human-readable description
    include_light_duty: bool = False   # Include Cat A (light-duty exempt) vehicles


# Preset scenario configurations
PRESET_SCENARIOS = {
    # ── Time-based presets (used by Present tab + PowerPoint export) ──────────
    "aggressive": ElectrificationScenario(
        name="Aggressive",
        end_year=2030,
        vehicle_filter="all",
        description="Full fleet electrification by 2030",
    ),
    "moderate": ElectrificationScenario(
        name="Moderate",
        end_year=2035,
        vehicle_filter="all",
        description="Full fleet electrification by 2035",
    ),
    "conservative": ElectrificationScenario(
        name="Conservative",
        end_year=2040,
        vehicle_filter="all",
        description="Full fleet electrification by 2040",
    ),
    "acf_compliance": ElectrificationScenario(
        name="ACF Compliance Only",
        end_year=2035,
        vehicle_filter="acf_only",
        description="Replace only CARB ACF-subject vehicles by 2035",
    ),

    # ── Scope-based presets (used by Analysis tab Scenario Comparison) ────────
    "minimum_compliance": ElectrificationScenario(
        name="Minimum Compliance",
        end_year=2040,
        vehicle_filter="acf_only",
        include_light_duty=False,
        description="Only ACF mandate-subject vehicles (Cat B: medium & heavy duty)",
    ),
    "all_except_emergency": ElectrificationScenario(
        name="All Excl. Emergency",
        end_year=2040,
        vehicle_filter="all_except_emergency",
        include_light_duty=True,
        description="All vehicles except emergency (Cat A+B+C)",
    ),
    "whole_fleet": ElectrificationScenario(
        name="Whole Fleet",
        end_year=2040,
        vehicle_filter="all",
        include_light_duty=True,
        description="Every vehicle in the fleet including emergency vehicles (Cat A+B+C+D)",
    ),
}

# Keys for the three scope-based scenarios shown in the Analysis tab
SCOPE_SCENARIO_KEYS = ("minimum_compliance", "all_except_emergency", "whole_fleet")


###############################################################################
# Scenario Engine
###############################################################################

def _filter_vehicles(
    vehicles: List[FleetVehicle],
    vehicle_filter: str,
) -> List[FleetVehicle]:
    """Filter vehicles based on scenario vehicle_filter setting."""
    if vehicle_filter == "all":
        return vehicles
    elif vehicle_filter == "acf_only":
        return [
            v for v in vehicles
            if v.custom_fields.get("_acf_code") == "B"
        ]
    elif vehicle_filter == "medium_heavy_only":
        return [
            v for v in vehicles
            if (v.vehicle_id.gvwr_pounds or 0) > 8500
        ]
    elif vehicle_filter == "all_except_emergency":
        # Include Cat A, B, C — exclude Cat D (emergency) and ZEV
        return [
            v for v in vehicles
            if v.custom_fields.get("_acf_code", "") not in ("D",)
        ]
    else:
        return vehicles


def _get_vehicle_ev_cost(vehicle: FleetVehicle) -> float:
    """Get EV purchase price from custom_fields, or estimate from GVWR."""
    ev_price = vehicle.custom_fields.get("_ev_purchase_price")
    if ev_price:
        return float(ev_price)

    # Rough estimate based on GVWR class
    gvwr = vehicle.vehicle_id.gvwr_pounds or 0
    if gvwr > 26000:
        return 250000  # Heavy-duty truck
    elif gvwr > 14000:
        return 120000  # Medium-duty
    elif gvwr > 8500:
        return 75000   # Light-medium
    else:
        return 45000   # Light-duty


def _get_vehicle_annual_co2(vehicle: FleetVehicle) -> float:
    """Get annual CO₂ emissions in metric tons."""
    annual_mileage = vehicle.annual_mileage or 12000
    co2_per_mile = vehicle.fuel_economy.co2_primary or 0
    if co2_per_mile <= 0:
        mpg = vehicle.fuel_economy.combined_mpg or 0
        co2_per_mile = 8900 / mpg if mpg > 0 else 0
    if co2_per_mile <= 0:
        return 0.0
    return (co2_per_mile * annual_mileage) / 1_000_000


def _get_vehicle_annual_savings(vehicle: FleetVehicle) -> float:
    """Estimate annual fuel + maintenance savings from electrification."""
    from analysis.calculations import (
        calculate_annual_fuel_cost, calculate_annual_ev_cost,
        DEFAULT_ICE_MAINTENANCE, DEFAULT_EV_MAINTENANCE,
        DEFAULT_ANNUAL_MILEAGE,
    )

    mileage = vehicle.annual_mileage or DEFAULT_ANNUAL_MILEAGE
    fuel_savings = calculate_annual_fuel_cost(vehicle) - calculate_annual_ev_cost(vehicle)
    maint_savings = mileage * (DEFAULT_ICE_MAINTENANCE - DEFAULT_EV_MAINTENANCE)
    return fuel_savings + maint_savings


def run_scenario(
    vehicles: List[FleetVehicle],
    scenario: ElectrificationScenario,
) -> Dict[str, Any]:
    """
    Run a single electrification scenario on the fleet.

    Returns a dict with:
        name: scenario name
        description: scenario description
        end_year: scenario end year
        vehicle_filter: filter type
        total_vehicles: number of vehicles in scope
        vehicles_per_year: {year: count}
        cumulative_vehicles: {year: cumulative count}
        cost_per_year: {year: $}
        cumulative_cost: {year: $}
        co2_reduction_per_year: {year: metric tons}
        cumulative_co2_reduction: {year: metric tons}
        savings_per_year: {year: $}
        cumulative_savings: {year: $}
        total_investment: total EV purchase cost
        total_annual_savings: annual savings at full deployment
        total_annual_co2_reduction: annual CO₂ reduction at full deployment
        summary_text: human-readable 1-2 sentence summary
    """
    current_year = datetime.datetime.now().year

    # Filter vehicles for this scenario
    eligible = _filter_vehicles(vehicles, scenario.vehicle_filter)

    # Further filter: only schedulable (not ZEV, not failed, has ACF code)
    schedulable = []
    for v in eligible:
        acf_code = v.custom_fields.get("_acf_code", "")
        if acf_code == "ZEV":
            continue  # Already electric — nothing to plan
        if acf_code == "":
            continue  # Classification unavailable
        if not v.processing_success:
            continue
        if acf_code == "A" and not scenario.include_light_duty:
            continue  # Skip light-duty unless scenario explicitly requests it
        score = _score_vehicle(v, acf_code)
        schedulable.append((score, v))

    if not schedulable:
        return _empty_result(scenario, current_year)

    # Sort by score descending (highest priority first)
    schedulable.sort(key=lambda pair: pair[0], reverse=True)

    # Budget-smooth across years
    end_year = scenario.end_year
    if end_year <= current_year:
        end_year = current_year + 1

    available_years = list(range(current_year + 1, end_year + 1))
    num_years = len(available_years)
    n = len(schedulable)

    # Assign vehicles to years
    assignments = []  # (year, vehicle)

    if scenario.budget_per_year > 0:
        # Budget-constrained path: greedy fill-from-start with floor-based
        # per-year capacities (avoids the ceil overcount that left tail years
        # empty when n is not a perfect multiple of num_years).
        base = max(1, n // num_years)
        extra = n % num_years if n >= num_years else n
        capacities = {
            yr: base + (1 if i < extra else 0)
            for i, yr in enumerate(available_years)
        }
        year_counts = {yr: 0 for yr in available_years}

        for _score, vehicle in schedulable:
            assigned_year = None
            ev_cost = _get_vehicle_ev_cost(vehicle)
            for yr in available_years:
                year_spend = sum(
                    _get_vehicle_ev_cost(v)
                    for y, v in assignments if y == yr
                )
                if (year_counts[yr] < capacities[yr] and
                        year_spend + ev_cost <= scenario.budget_per_year):
                    assigned_year = yr
                    break
            if assigned_year is None:
                assigned_year = available_years[-1]
            year_counts[assigned_year] += 1
            assignments.append((assigned_year, vehicle))
    else:
        # No budget constraint: spread vehicles evenly across the full horizon
        # using the same algorithm as assign_electrification_years().
        if n <= num_years:
            # Fewer (or equal) vehicles than years: linearly space across horizon.
            for k, (_score, vehicle) in enumerate(schedulable):
                idx = 0 if n == 1 else round(k * (num_years - 1) / (n - 1))
                assignments.append((available_years[idx], vehicle))
        else:
            # More vehicles than years: floor-based per-year capacity so all
            # years are used rather than just the first ceil(n/Y) years.
            base = n // num_years
            extra = n % num_years
            year_slots: List[int] = []
            for i, yr in enumerate(available_years):
                cap = base + (1 if i < extra else 0)
                year_slots.extend([yr] * cap)
            for (_score, vehicle), yr in zip(schedulable, year_slots):
                assignments.append((yr, vehicle))

    # Build yearly metrics
    vehicles_per_year = {}
    cost_per_year = {}
    co2_per_year = {}
    savings_per_year = {}

    for yr, v in assignments:
        vehicles_per_year[yr] = vehicles_per_year.get(yr, 0) + 1
        cost_per_year[yr] = cost_per_year.get(yr, 0) + _get_vehicle_ev_cost(v)
        co2_per_year[yr] = co2_per_year.get(yr, 0) + _get_vehicle_annual_co2(v)
        savings_per_year[yr] = savings_per_year.get(yr, 0) + _get_vehicle_annual_savings(v)

    # Build cumulative metrics
    cumul_vehicles = {}
    cumul_cost = {}
    cumul_co2 = {}
    cumul_savings = {}
    running_vehicles = 0
    running_cost = 0.0
    running_co2 = 0.0
    running_savings = 0.0

    for yr in available_years:
        running_vehicles += vehicles_per_year.get(yr, 0)
        running_cost += cost_per_year.get(yr, 0)
        running_co2 += co2_per_year.get(yr, 0)
        running_savings += savings_per_year.get(yr, 0)

        cumul_vehicles[yr] = running_vehicles
        cumul_cost[yr] = running_cost
        cumul_co2[yr] = running_co2
        cumul_savings[yr] = running_savings

    total_investment = running_cost
    total_annual_savings = running_savings
    total_co2 = running_co2

    # Summary text
    summary = (
        f"{scenario.name}: Electrify {len(schedulable)} vehicles by {end_year}. "
        f"Total investment: ${total_investment:,.0f}. "
        f"Annual savings at full deployment: ${total_annual_savings:,.0f}/yr. "
        f"CO₂ reduction: {total_co2:,.0f} MT/yr."
    )

    return {
        "name": scenario.name,
        "description": scenario.description,
        "end_year": end_year,
        "vehicle_filter": scenario.vehicle_filter,
        "total_vehicles": len(schedulable),
        "vehicles_per_year": vehicles_per_year,
        "cumulative_vehicles": cumul_vehicles,
        "cost_per_year": cost_per_year,
        "cumulative_cost": cumul_cost,
        "co2_reduction_per_year": co2_per_year,
        "cumulative_co2_reduction": cumul_co2,
        "savings_per_year": savings_per_year,
        "cumulative_savings": cumul_savings,
        "total_investment": total_investment,
        "total_annual_savings": total_annual_savings,
        "total_annual_co2_reduction": total_co2,
        "summary_text": summary,
    }


def _empty_result(scenario: ElectrificationScenario, current_year: int) -> Dict[str, Any]:
    """Return an empty result dict for a scenario with no eligible vehicles."""
    return {
        "name": scenario.name,
        "description": scenario.description,
        "end_year": scenario.end_year,
        "vehicle_filter": scenario.vehicle_filter,
        "total_vehicles": 0,
        "vehicles_per_year": {},
        "cumulative_vehicles": {},
        "cost_per_year": {},
        "cumulative_cost": {},
        "co2_reduction_per_year": {},
        "cumulative_co2_reduction": {},
        "savings_per_year": {},
        "cumulative_savings": {},
        "total_investment": 0,
        "total_annual_savings": 0,
        "total_annual_co2_reduction": 0,
        "summary_text": f"{scenario.name}: No eligible vehicles found.",
    }


def compare_scenarios(
    vehicles: List[FleetVehicle],
    scenario_names: Optional[List[str]] = None,
    custom_scenarios: Optional[List[ElectrificationScenario]] = None,
) -> Dict[str, Any]:
    """
    Run multiple scenarios and produce a side-by-side comparison.

    Args:
        vehicles: Fleet vehicles to analyze
        scenario_names: List of preset scenario keys (from PRESET_SCENARIOS)
        custom_scenarios: Additional custom scenario objects

    Returns:
        Dict with:
            scenarios: list of individual scenario results
            comparison_table: list of dicts for side-by-side display
            all_years: sorted list of all years across scenarios
            best_roi: name of scenario with best ROI
            lowest_cost: name of scenario with lowest total cost
            fastest: name of scenario with earliest completion
    """
    scenarios_to_run = []

    if scenario_names:
        for name in scenario_names:
            if name in PRESET_SCENARIOS:
                scenarios_to_run.append(PRESET_SCENARIOS[name])
            else:
                logger.warning(f"Unknown preset scenario: {name}")

    if custom_scenarios:
        scenarios_to_run.extend(custom_scenarios)

    if not scenarios_to_run:
        # Default: run all 4 presets
        scenarios_to_run = list(PRESET_SCENARIOS.values())

    results = []
    for scenario in scenarios_to_run:
        result = run_scenario(vehicles, scenario)
        results.append(result)

    if not results:
        return {"scenarios": [], "comparison_table": [], "all_years": []}

    # Build comparison table
    comparison = []
    for r in results:
        comparison.append({
            "name": r["name"],
            "vehicles": r["total_vehicles"],
            "end_year": r["end_year"],
            "total_investment": r["total_investment"],
            "total_annual_savings": r["total_annual_savings"],
            "total_annual_co2_reduction": r["total_annual_co2_reduction"],
            "payback_years": (
                r["total_investment"] / r["total_annual_savings"]
                if r["total_annual_savings"] > 0 else float('inf')
            ),
        })

    # Collect all years across scenarios
    all_years = set()
    for r in results:
        all_years.update(r["cumulative_vehicles"].keys())
    all_years = sorted(all_years)

    # Find best-in-class
    valid_results = [r for r in results if r["total_vehicles"] > 0]
    best_roi = ""
    lowest_cost = ""
    fastest = ""

    if valid_results:
        best_roi_r = max(
            valid_results,
            key=lambda r: r["total_annual_savings"] / max(r["total_investment"], 1)
        )
        best_roi = best_roi_r["name"]

        lowest_cost_r = min(valid_results, key=lambda r: r["total_investment"])
        lowest_cost = lowest_cost_r["name"]

        fastest_r = min(valid_results, key=lambda r: r["end_year"])
        fastest = fastest_r["name"]

    return {
        "scenarios": results,
        "comparison_table": comparison,
        "all_years": all_years,
        "best_roi": best_roi,
        "lowest_cost": lowest_cost,
        "fastest": fastest,
    }


def get_scenario_year_assignments(
    vehicles: List[FleetVehicle],
    scenario_name: str,
    fleet_type: str = "hpf",
) -> Dict[str, str]:
    """Return a {vin: year_str} mapping for every vehicle under a preset scenario.

    Works on deep copies so the original fleet is never mutated.
    Vehicles not eligible under the scenario's vehicle_filter are mapped to "—".

    Args:
        vehicles:     Fleet vehicles to evaluate.
        scenario_name: Key from PRESET_SCENARIOS.
        fleet_type:   CARB fleet classification for deadline lookup
                      ("hpf", "non_hpf", or "state_agency").
    """
    import copy
    from analysis.electrification_timeline import assign_electrification_years

    scenario = PRESET_SCENARIOS.get(scenario_name)
    if not scenario:
        return {}

    # Determine which VINs are in scope for this scenario's filter
    eligible_vins = {v.vin for v in _filter_vehicles(vehicles, scenario.vehicle_filter)}

    # Deep-copy to avoid mutating originals
    copies = copy.deepcopy(vehicles)
    assign_electrification_years(copies, end_year=scenario.end_year, fleet_type=fleet_type)

    result: Dict[str, str] = {}
    for vc in copies:
        if vc.vin in eligible_vins:
            result[vc.vin] = str(vc.custom_fields.get("Proposed EV Year", "N/A"))
        else:
            result[vc.vin] = "—"
    return result
