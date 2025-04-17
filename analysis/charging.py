import math

def analyze_charging_needs(fleet, daily_usage_pattern, charging_window, charging_power, power_level_type):
    """
    Analyze charging infrastructure needs for a fleet.
    
    Args:
        fleet (dict): Fleet information including vehicle count and types
        daily_usage_pattern (str): Pattern of daily usage ('overnight' or 'opportunity')
        charging_window (tuple): Start and end hours for charging window
        charging_power (float): Power rating of chargers in kW
        power_level_type (str): Type of power level (LP, MP, HP, VHP)
    
    Returns:
        dict: Analysis results including energy needs, chargers required, and schedule
    """
    results = {
        'total_energy_needed': 0,
        'peak_power_demand': 0,
        'chargers_needed': 0,
        'charging_schedule': [],
        'power_level': power_level_type,
        'charging_power': charging_power
    }
    
    # Calculate total daily energy needs
    total_energy = sum(vehicle.daily_kwh for vehicle in fleet.vehicles)
    results['total_energy_needed'] = total_energy
    
    # Calculate charging window duration
    start_hour, end_hour = charging_window
    if end_hour < start_hour:  # Window crosses midnight
        window_hours = (24 - start_hour) + end_hour
    else:
        window_hours = end_hour - start_hour
    
    # Calculate required chargers based on energy needs and charging window
    energy_per_charger = charging_power * window_hours * 0.85  # Assuming 85% efficiency
    min_chargers = math.ceil(total_energy / energy_per_charger)
    
    # Add buffer based on usage pattern
    if daily_usage_pattern == 'overnight':
        buffer_factor = 1.1  # 10% buffer for overnight charging
    else:  # opportunity charging needs more buffer due to time constraints
        buffer_factor = 1.3  # 30% buffer for opportunity charging
    
    results['chargers_needed'] = math.ceil(min_chargers * buffer_factor)
    results['peak_power_demand'] = results['chargers_needed'] * charging_power
    
    # Generate simplified charging schedule
    schedule = []
    vehicles_per_charger = math.ceil(len(fleet.vehicles) / results['chargers_needed'])
    current_hour = start_hour
    
    for vehicle in fleet.vehicles:
        charge_duration = math.ceil(vehicle.daily_kwh / (charging_power * 0.85))
        schedule.append({
            'vehicle_id': vehicle.id,
            'start_hour': current_hour,
            'duration': charge_duration,
            'power': charging_power
        })
        
        if len(schedule) % vehicles_per_charger == 0:
            current_hour = (current_hour + charge_duration) % 24
    
    results['charging_schedule'] = schedule
    
    return results 