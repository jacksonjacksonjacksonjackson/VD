"""
calculations.py

Core analysis calculations for the Fleet Electrification Analyzer.
Implements functions for fuel cost analysis, emissions calculations,
and electrification ROI.
"""

import logging
import math
from datetime import datetime
from typing import Dict, List, Any, Optional, Union, Tuple, Set

from settings import (
    DEFAULT_GAS_PRICE,
    DEFAULT_ELECTRICITY_PRICE,
    DEFAULT_EV_EFFICIENCY,
    DEFAULT_ANNUAL_MILEAGE,
    DEFAULT_VEHICLE_LIFESPAN,
    DEFAULT_BATTERY_DEGRADATION,
    DEFAULT_ICE_MAINTENANCE,
    DEFAULT_EV_MAINTENANCE
)
from data.models import FleetVehicle, Fleet, ElectrificationAnalysis, ChargingAnalysis, EmissionsInventory

# Set up module logger
logger = logging.getLogger(__name__)

###############################################################################
# Fuel and Emissions Calculations
###############################################################################

def calculate_annual_fuel_cost(vehicle: FleetVehicle, gas_price: float = DEFAULT_GAS_PRICE) -> float:
    """
    Calculate annual fuel cost for a vehicle.
    
    Args:
        vehicle: Vehicle to analyze
        gas_price: Price of gasoline in $/gallon
        
    Returns:
        Annual fuel cost in dollars
    """
    # Use vehicle's annual mileage if available, otherwise default
    annual_mileage = vehicle.annual_mileage or DEFAULT_ANNUAL_MILEAGE
    
    # MPG from fuel economy data
    mpg = vehicle.fuel_economy.combined_mpg
    
    # If no MPG data, return 0
    if not mpg or mpg <= 0:
        return 0.0
    
    # Calculate gallons used
    gallons = annual_mileage / mpg
    
    # Calculate cost
    return gallons * gas_price

def calculate_annual_ev_cost(vehicle: FleetVehicle, electricity_price: float = DEFAULT_ELECTRICITY_PRICE,
                           ev_efficiency: float = DEFAULT_EV_EFFICIENCY) -> float:
    """
    Calculate annual cost if the vehicle were electric.
    
    Args:
        vehicle: Vehicle to analyze
        electricity_price: Price of electricity in $/kWh
        ev_efficiency: EV energy usage in kWh/mile
        
    Returns:
        Annual electricity cost in dollars
    """
    # Use vehicle's annual mileage if available, otherwise default
    annual_mileage = vehicle.annual_mileage or DEFAULT_ANNUAL_MILEAGE
    
    # Calculate kWh used
    kwh = annual_mileage * ev_efficiency
    
    # Calculate cost
    return kwh * electricity_price

def calculate_annual_co2_emissions(vehicle: FleetVehicle) -> float:
    """
    Calculate annual CO2 emissions for a vehicle in metric tons.
    
    Args:
        vehicle: Vehicle to analyze
        
    Returns:
        Annual CO2 emissions in metric tons
    """
    # Use vehicle's annual mileage if available, otherwise default
    annual_mileage = vehicle.annual_mileage or DEFAULT_ANNUAL_MILEAGE
    
    # CO2 emissions in g/mile
    co2_per_mile = vehicle.fuel_economy.co2_primary
    
    # If no emissions data, estimate from MPG
    if not co2_per_mile or co2_per_mile <= 0:
        mpg = vehicle.fuel_economy.combined_mpg
        if mpg and mpg > 0:
            # Approximate CO2 emissions based on MPG (gasoline: ~8.9 kg CO2 per gallon)
            co2_per_mile = 8900 / mpg
        else:
            return 0.0
    
    # Calculate total emissions in metric tons (g -> kg -> metric ton)
    return (co2_per_mile * annual_mileage) / 1000000

def calculate_ev_emissions(annual_mileage: float, electricity_intensity: float = 0.4) -> float:
    """
    Calculate annual CO2 emissions for an electric vehicle in metric tons.
    
    Args:
        annual_mileage: Annual mileage
        electricity_intensity: Grid emissions intensity in kg CO2/kWh
        
    Returns:
        Annual CO2 emissions in metric tons
    """
    # Calculate kWh used
    kwh = annual_mileage * DEFAULT_EV_EFFICIENCY
    
    # Calculate emissions in metric tons
    return (kwh * electricity_intensity) / 1000

def calculate_emissions_reduction(vehicle: FleetVehicle, 
                               electricity_intensity: float = 0.4) -> float:
    """
    Calculate CO2 emissions reduction from converting a vehicle to electric.
    
    Args:
        vehicle: Vehicle to analyze
        electricity_intensity: Grid emissions intensity in kg CO2/kWh
        
    Returns:
        Annual CO2 emissions reduction in metric tons
    """
    # Current emissions
    current_emissions = calculate_annual_co2_emissions(vehicle)
    
    # Emissions if electric
    annual_mileage = vehicle.annual_mileage or DEFAULT_ANNUAL_MILEAGE
    ev_emissions = calculate_ev_emissions(annual_mileage, electricity_intensity)
    
    # Calculate reduction
    return max(0, current_emissions - ev_emissions)

###############################################################################
# TCO and ROI Calculations
###############################################################################

def calculate_electrification_savings(
    vehicle: FleetVehicle,
    gas_price: float = DEFAULT_GAS_PRICE,
    electricity_price: float = DEFAULT_ELECTRICITY_PRICE,
    ev_efficiency: float = DEFAULT_EV_EFFICIENCY,
    analysis_years: int = DEFAULT_VEHICLE_LIFESPAN,
    discount_rate: float = 5.0,
    battery_degradation: float = DEFAULT_BATTERY_DEGRADATION,
    ice_maintenance: float = DEFAULT_ICE_MAINTENANCE,
    ev_maintenance: float = DEFAULT_EV_MAINTENANCE
) -> Dict[str, float]:
    """
    Calculate comprehensive savings from electrifying a vehicle.
    
    Args:
        vehicle: Vehicle to analyze
        gas_price: Price of gasoline in $/gallon
        electricity_price: Price of electricity in $/kWh
        ev_efficiency: EV energy usage in kWh/mile
        analysis_years: Number of years to analyze
        discount_rate: Annual discount rate for NPV calculation (%)
        battery_degradation: Annual battery degradation rate (%)
        ice_maintenance: ICE maintenance cost per mile ($)
        ev_maintenance: EV maintenance cost per mile ($)
        
    Returns:
        Dictionary of savings by category and totals
    """
    # Annual mileage
    annual_mileage = vehicle.annual_mileage or DEFAULT_ANNUAL_MILEAGE
    
    # Initialize results
    results = {
        "annual_fuel_savings": 0.0,
        "total_fuel_savings": 0.0,
        "annual_maintenance_savings": 0.0,
        "total_maintenance_savings": 0.0,
        "total_npv_savings": 0.0,
        "annual_co2_reduction": 0.0,
        "total_co2_reduction": 0.0
    }
    
    # Calculate annual fuel costs
    ice_fuel_cost = calculate_annual_fuel_cost(vehicle, gas_price)
    ev_fuel_cost = calculate_annual_ev_cost(vehicle, electricity_price, ev_efficiency)
    
    # Annual fuel savings
    annual_fuel_savings = ice_fuel_cost - ev_fuel_cost
    results["annual_fuel_savings"] = annual_fuel_savings
    
    # Annual maintenance savings
    annual_maintenance_savings = annual_mileage * (ice_maintenance - ev_maintenance)
    results["annual_maintenance_savings"] = annual_maintenance_savings
    
    # Annual emissions reduction
    annual_co2_reduction = calculate_emissions_reduction(vehicle)
    results["annual_co2_reduction"] = annual_co2_reduction
    
    # Calculate NPV of savings over analysis period
    total_npv = 0.0
    total_fuel_savings = 0.0
    total_maintenance_savings = 0.0
    total_co2_reduction = 0.0
    
    for year in range(1, analysis_years + 1):
        # Adjust EV efficiency for battery degradation
        degraded_efficiency = ev_efficiency * (1 + (battery_degradation/100) * (year-1))
        
        # Calculate fuel costs with degradation
        year_ev_fuel_cost = (annual_mileage * degraded_efficiency * electricity_price)
        year_fuel_savings = ice_fuel_cost - year_ev_fuel_cost
        
        # Add to totals
        total_fuel_savings += year_fuel_savings
        total_maintenance_savings += annual_maintenance_savings
        total_co2_reduction += annual_co2_reduction
        
        # Calculate NPV of this year's savings
        year_savings = year_fuel_savings + annual_maintenance_savings
        year_npv = year_savings / ((1 + discount_rate/100) ** year)
        total_npv += year_npv
    
    # Store totals
    results["total_fuel_savings"] = total_fuel_savings
    results["total_maintenance_savings"] = total_maintenance_savings
    results["total_npv_savings"] = total_npv
    results["total_co2_reduction"] = total_co2_reduction
    
    return results

def calculate_ev_roi(
    vehicle: FleetVehicle,
    ev_purchase_price: float,
    ice_purchase_price: float,
    gas_price: float = DEFAULT_GAS_PRICE,
    electricity_price: float = DEFAULT_ELECTRICITY_PRICE,
    ev_efficiency: float = DEFAULT_EV_EFFICIENCY,
    analysis_years: int = DEFAULT_VEHICLE_LIFESPAN,
    ice_maintenance: float = DEFAULT_ICE_MAINTENANCE,
    ev_maintenance: float = DEFAULT_EV_MAINTENANCE
) -> Dict[str, Any]:
    """
    Calculate ROI metrics for replacing vehicle with an electric equivalent.
    
    Args:
        vehicle: Vehicle to analyze
        ev_purchase_price: Purchase price of EV equivalent
        ice_purchase_price: Purchase price of new ICE equivalent
        gas_price: Price of gasoline in $/gallon
        electricity_price: Price of electricity in $/kWh
        ev_efficiency: EV energy usage in kWh/mile
        analysis_years: Number of years to analyze
        ice_maintenance: ICE maintenance cost per mile ($)
        ev_maintenance: EV maintenance cost per mile ($)
        
    Returns:
        Dictionary with ROI metrics
    """
    # Get savings from electrification
    savings = calculate_electrification_savings(
        vehicle=vehicle,
        gas_price=gas_price,
        electricity_price=electricity_price,
        ev_efficiency=ev_efficiency,
        analysis_years=analysis_years,
        ice_maintenance=ice_maintenance,
        ev_maintenance=ev_maintenance
    )
    
    # Calculate price premium for EV
    price_premium = ev_purchase_price - ice_purchase_price
    
    # Calculate simple payback period (years)
    annual_savings = savings["annual_fuel_savings"] + savings["annual_maintenance_savings"]
    
    if annual_savings <= 0:
        payback_years = float('inf')
    else:
        payback_years = price_premium / annual_savings
    
    # Calculate ROI over the analysis period
    if price_premium <= 0:
        roi_percent = float('inf')  # If EV costs less, infinite ROI
    else:
        roi_percent = (savings["total_npv_savings"] / price_premium) * 100
    
    # Calculate TCO for both options
    ice_tco = ice_purchase_price + (annual_savings * analysis_years) 
    ev_tco = ev_purchase_price
    
    # Return results
    return {
        "price_premium": price_premium,
        "annual_savings": annual_savings,
        "payback_years": payback_years,
        "roi_percent": roi_percent,
        "ice_tco": ice_tco,
        "ev_tco": ev_tco,
        "tco_savings": ice_tco - ev_tco,
        "co2_reduction": savings["total_co2_reduction"]
    }

###############################################################################
# Fleet Analysis
###############################################################################

def analyze_fleet_electrification(
    fleet: Fleet,
    gas_price: float = DEFAULT_GAS_PRICE,
    electricity_price: float = DEFAULT_ELECTRICITY_PRICE,
    ev_efficiency: float = DEFAULT_EV_EFFICIENCY,
    analysis_years: int = DEFAULT_VEHICLE_LIFESPAN,
    discount_rate: float = 5.0
) -> ElectrificationAnalysis:
    """
    Analyze electrification potential for an entire fleet.
    
    Args:
        fleet: Fleet to analyze
        gas_price: Price of gasoline in $/gallon
        electricity_price: Price of electricity in $/kWh
        ev_efficiency: EV energy usage in kWh/mile
        analysis_years: Number of years for analysis
        discount_rate: Annual discount rate for NPV calculations (%)
        
    Returns:
        ElectrificationAnalysis object with results
    """
    # Initialize analysis object
    analysis = ElectrificationAnalysis(
        fleet_name=fleet.name,
        gas_price=gas_price,
        electricity_price=electricity_price,
        ev_efficiency=ev_efficiency,
        analysis_period=analysis_years,
        discount_rate=discount_rate
    )
    
    # Analyze each vehicle
    total_co2_savings = 0.0
    total_fuel_cost_savings = 0.0
    total_maintenance_savings = 0.0
    total_savings = 0.0
    prioritized_vehicles = []
    
    for vehicle in fleet.vehicles:
        # Skip vehicles without sufficient data
        if not vehicle.fuel_economy.combined_mpg:
            continue
        
        # Calculate savings
        savings = calculate_electrification_savings(
            vehicle=vehicle,
            gas_price=gas_price,
            electricity_price=electricity_price,
            ev_efficiency=ev_efficiency,
            analysis_years=analysis_years,
            discount_rate=discount_rate
        )
        
        # Store vehicle results
        analysis.vehicle_results[vehicle.vin] = {
            "annual_fuel_savings": savings["annual_fuel_savings"],
            "total_fuel_savings": savings["total_fuel_savings"],
            "annual_maintenance_savings": savings["annual_maintenance_savings"],
            "total_maintenance_savings": savings["total_maintenance_savings"],
            "total_npv_savings": savings["total_npv_savings"],
            "annual_co2_reduction": savings["annual_co2_reduction"],
            "total_co2_reduction": savings["total_co2_reduction"],
            "make": vehicle.vehicle_id.make,
            "model": vehicle.vehicle_id.model,
            "year": vehicle.vehicle_id.year,
            "annual_mileage": vehicle.annual_mileage or DEFAULT_ANNUAL_MILEAGE,
            "mpg": vehicle.fuel_economy.combined_mpg,
            "display_name": vehicle.display_name
        }
        
        # Add to totals
        total_co2_savings += savings["total_co2_reduction"]
        total_fuel_cost_savings += savings["total_fuel_savings"]
        total_maintenance_savings += savings["total_maintenance_savings"]
        total_savings += savings["total_npv_savings"]
        
        # Add to prioritization list
        prioritized_vehicles.append((vehicle.vin, savings["total_npv_savings"]))
    
    # Sort vehicles by savings for prioritization
    prioritized_vehicles.sort(key=lambda x: x[1], reverse=True)
    analysis.prioritized_vehicles = [vin for vin, _ in prioritized_vehicles]
    
    # Set total results
    analysis.co2_savings = total_co2_savings
    analysis.fuel_cost_savings = total_fuel_cost_savings
    analysis.maintenance_savings = total_maintenance_savings
    analysis.total_savings = total_savings
    
    # Calculate fleet-wide payback period (simple approximation)
    if total_savings > 0:
        # Calculate annual savings
        annual_savings = total_fuel_cost_savings / analysis_years
        annual_savings += total_maintenance_savings / analysis_years
        
        # Assuming $15,000 per vehicle EV premium
        ev_premium = len(fleet.vehicles) * 15000
        
        # Calculate payback
        analysis.payback_period = ev_premium / annual_savings if annual_savings > 0 else float('inf')
    else:
        analysis.payback_period = float('inf')
    
    return analysis

def create_emissions_inventory(fleet: Fleet) -> EmissionsInventory:
    """
    Create a greenhouse gas emissions inventory for the fleet.
    
    Args:
        fleet: Fleet to analyze
        
    Returns:
        EmissionsInventory object with results
    """
    # Initialize inventory object
    inventory = EmissionsInventory(
        fleet_name=fleet.name,
        inventory_year=datetime.now().year
    )
    
    # Calculate emissions by category
    by_department = {}
    by_vehicle_type = {}
    by_fuel_type = {}
    total_emissions = 0.0
    
    for vehicle in fleet.vehicles:
        # Calculate annual emissions
        emissions = calculate_annual_co2_emissions(vehicle)
        
        # Skip vehicles with no emissions data
        if emissions <= 0:
            continue
        
        # Add to total
        total_emissions += emissions
        
        # Add to department breakdown
        dept = vehicle.department or "Unassigned"
        by_department[dept] = by_department.get(dept, 0.0) + emissions
        
        # Add to vehicle type breakdown
        vtype = vehicle.vehicle_id.body_class or "Unknown"
        by_vehicle_type[vtype] = by_vehicle_type.get(vtype, 0.0) + emissions
        
        # Add to fuel type breakdown
        ftype = vehicle.vehicle_id.fuel_type or "Unknown"
        by_fuel_type[ftype] = by_fuel_type.get(ftype, 0.0) + emissions
    
    # Update inventory object
    inventory.total_emissions = total_emissions
    inventory.by_department = by_department
    inventory.by_vehicle_type = by_vehicle_type
    inventory.by_fuel_type = by_fuel_type
    
    # Set baseline and target years (mock data)
    current_year = datetime.now().year
    inventory.baseline_year = current_year - 1
    inventory.target_year = current_year + 10
    inventory.reduction_target = 50.0  # 50% reduction target
    
    # Add mock historical data (for demo purposes)
    inventory.historical_data = {
        current_year - 3: total_emissions * 1.1,
        current_year - 2: total_emissions * 1.05,
        current_year - 1: total_emissions * 1.02,
        current_year: total_emissions
    }
    
    # Add projected emissions (for demo purposes)
    target_emissions = total_emissions * 0.5  # 50% reduction
    
    inventory.projected_emissions = {
        current_year: total_emissions,
        current_year + 5: total_emissions * 0.75,
        current_year + 10: target_emissions
    }
    
    return inventory

###############################################################################
# Charging Infrastructure Analysis
###############################################################################

def analyze_charging_needs(
    fleet: Fleet,
    daily_usage_pattern: str = "standard",
    charging_window: Tuple[int, int] = (18, 6),  # 6 PM to 6 AM
    ev_battery_capacity: float = 60.0,  # kWh
    level2_charging_rate: float = 7.2,  # kW
    dcfc_charging_rate: float = 50.0,  # kW
    level2_cost: float = 4000.0,  # $ per port
    dcfc_cost: float = 50000.0  # $ per port
) -> ChargingAnalysis:
    """
    Analyze charging infrastructure needs for fleet electrification.
    
    Args:
        fleet: Fleet to analyze
        daily_usage_pattern: Usage pattern ("standard", "extended", "24-hour")
        charging_window: Hours available for charging (start, end)
        ev_battery_capacity: Average EV battery capacity in kWh
        level2_charging_rate: Level 2 charging power in kW
        dcfc_charging_rate: DC fast charging power in kW
        level2_cost: Cost per Level 2 charger
        dcfc_cost: Cost per DC fast charger
        
    Returns:
        ChargingAnalysis object with results
    """
    # Initialize analysis object
    analysis = ChargingAnalysis(
        fleet_name=fleet.name,
        daily_usage_pattern=daily_usage_pattern,
        charging_window=charging_window
    )
    
    # Count total vehicles
    total_vehicles = len(fleet.vehicles)
    if total_vehicles == 0:
        return analysis
    
    # Calculate daily energy needs
    total_daily_miles = 0.0
    for vehicle in fleet.vehicles:
        annual_mileage = vehicle.annual_mileage or DEFAULT_ANNUAL_MILEAGE
        daily_miles = annual_mileage / 365.0
        total_daily_miles += daily_miles
    
    # Calculate daily energy requirement (kWh)
    daily_energy = total_daily_miles * DEFAULT_EV_EFFICIENCY
    
    # Adjust charging window based on usage pattern
    charging_hours = 0
    start, end = charging_window
    if start < end:
        charging_hours = end - start
    else:
        charging_hours = (24 - start) + end
    
    if daily_usage_pattern == "extended":
        # Reduce available charging hours
        charging_hours *= 0.7
    elif daily_usage_pattern == "24-hour":
        # Multi-shift operation
        charging_hours = 8  # Assume 8 hours available
    
    # Calculate chargers needed based on energy needs
    hourly_energy = daily_energy / max(1, charging_hours)
    
    # Level 2 chargers needed
    level2_chargers = math.ceil(hourly_energy / level2_charging_rate)
    
    # DC fast chargers for emergency/opportunity charging
    dcfc_chargers = max(1, math.ceil(total_vehicles * 0.1))  # At least 10% of fleet size
    
    # Adjust for multi-shift operations
    if daily_usage_pattern == "24-hour":
        level2_chargers = math.ceil(level2_chargers * 1.5)
        dcfc_chargers = max(1, math.ceil(total_vehicles * 0.2))  # At least 20% of fleet size
    
    # Update analysis object
    analysis.level2_chargers_needed = level2_chargers
    analysis.dcfc_chargers_needed = dcfc_chargers
    
    # Calculate maximum power required
    analysis.max_power_required = (level2_chargers * level2_charging_rate) + (dcfc_chargers * dcfc_charging_rate)
    
    # Estimate installation cost
    analysis.estimated_installation_cost = (level2_chargers * level2_cost) + (dcfc_chargers * dcfc_cost)
    
    # Create recommended layout (simple example)
    analysis.recommended_layout = {
        "zones": [
            {
                "name": "Main Depot",
                "level2_chargers": level2_chargers,
                "dcfc_chargers": dcfc_chargers,
                "power_required": analysis.max_power_required
            }
        ],
        "phasing": [
            {
                "phase": 1,
                "level2_chargers": max(1, math.ceil(level2_chargers * 0.5)),
                "dcfc_chargers": max(1, math.ceil(dcfc_chargers * 0.5)),
                "estimated_cost": max(1, math.ceil((level2_chargers * 0.5) * level2_cost + (dcfc_chargers * 0.5) * dcfc_cost))
            },
            {
                "phase": 2,
                "level2_chargers": level2_chargers - max(1, math.ceil(level2_chargers * 0.5)),
                "dcfc_chargers": dcfc_chargers - max(1, math.ceil(dcfc_chargers * 0.5)),
                "estimated_cost": analysis.estimated_installation_cost - max(1, math.ceil((level2_chargers * 0.5) * level2_cost + (dcfc_chargers * 0.5) * dcfc_cost))
            }
        ]
    }
    
    return analysis