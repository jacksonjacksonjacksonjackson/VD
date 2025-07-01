#!/usr/bin/env python3
"""
Diagnostic Test Script for VIN Processing Issues

This script tests the VIN processing pipeline with diagnostic logging
to identify the root causes of data-related issues.
"""

import logging
import sys
import os

# Add the project root to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from data.providers import VehicleDataProvider
from utils import validate_vin_detailed

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(levelname)s: %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

# Set up diagnostic logger
diagnostic_logger = logging.getLogger("diagnostic")
diagnostic_logger.setLevel(logging.DEBUG)

def test_vin_processing():
    """Test VIN processing with diagnostic logging."""
    
    # Test VINs - mix of different vehicle types and potential problem cases
    test_vins = [
        "1FTFW1ET5DFC10312",  # Ford F-150 (should work well)
        "1GCEK14T74Z123456",  # Chevy Silverado (commercial vehicle)
        "1G1BE5SM2G7123456",  # Chevy Malibu (passenger car)
        "5NPE24AF4GH123456",  # Hyundai Sonata (foreign make)
        "1FAHP2D83EG123456",  # Ford Fusion (discontinued model)
        "INVALID123456789",   # Invalid VIN (should fail gracefully)
        "1234567890ABCDEFG"   # Invalid VIN format
    ]
    
    print("=" * 80)
    print("DIAGNOSTIC TEST: VIN Processing Pipeline")
    print("=" * 80)
    print()
    
    # Initialize the provider
    provider = VehicleDataProvider(cache_enabled=False)  # Disable cache for testing
    
    successful_count = 0
    failed_count = 0
    partial_count = 0
    
    for i, vin in enumerate(test_vins, 1):
        print(f"\n{'='*60}")
        print(f"TEST {i}/{len(test_vins)}: {vin}")
        print(f"{'='*60}")
        
        # First validate VIN format
        is_valid, validation_error = validate_vin_detailed(vin)
        print(f"VIN VALIDATION: {'PASS' if is_valid else 'FAIL'}")
        if not is_valid:
            print(f"VALIDATION ERROR: {validation_error}")
            failed_count += 1
            continue
        
        try:
            # Process the VIN through the full pipeline
            success, data, error = provider.get_vehicle_by_vin(vin)
            
            if success:
                vehicle_id = data.get('vehicle_id', {})
                fuel_economy = data.get('fuel_economy', {})
                match_confidence = data.get('match_confidence', 0)
                
                # Check data completeness
                has_basic_data = all([
                    vehicle_id.get('year'),
                    vehicle_id.get('make'), 
                    vehicle_id.get('model')
                ])
                
                has_fuel_data = fuel_economy.get('combined_mpg', 0) > 0
                
                print(f"\nRESULT SUMMARY:")
                print(f"  SUCCESS: {success}")
                print(f"  BASIC DATA: {'✓' if has_basic_data else '✗'}")
                print(f"  FUEL DATA: {'✓' if has_fuel_data else '✗'}")
                print(f"  MATCH CONFIDENCE: {match_confidence}%")
                print(f"  VEHICLE: {vehicle_id.get('year')} {vehicle_id.get('make')} {vehicle_id.get('model')}")
                print(f"  MPG: {fuel_economy.get('combined_mpg', 'N/A')}")
                
                if has_basic_data and has_fuel_data:
                    successful_count += 1
                    print(f"  STATUS: ✓ COMPLETE SUCCESS")
                elif has_basic_data:
                    partial_count += 1
                    print(f"  STATUS: ⚠ PARTIAL SUCCESS (No fuel data)")
                else:
                    failed_count += 1
                    print(f"  STATUS: ✗ MISSING BASIC DATA")
                    
            else:
                failed_count += 1
                print(f"\nRESULT SUMMARY:")
                print(f"  SUCCESS: {success}")
                print(f"  ERROR: {error}")
                print(f"  STATUS: ✗ FAILED")
                
        except Exception as e:
            failed_count += 1
            print(f"\nEXCEPTION OCCURRED:")
            print(f"  ERROR: {str(e)}")
            print(f"  STATUS: ✗ EXCEPTION")
            
        print(f"\n{'-'*60}")
    
    # Final summary
    print(f"\n{'='*80}")
    print("DIAGNOSTIC TEST SUMMARY")
    print(f"{'='*80}")
    print(f"Total VINs Tested: {len(test_vins)}")
    print(f"Complete Success: {successful_count}")
    print(f"Partial Success: {partial_count}")
    print(f"Failed: {failed_count}")
    print(f"Success Rate: {(successful_count / len(test_vins)) * 100:.1f}%")
    print(f"{'='*80}")
    
    return successful_count, partial_count, failed_count

def test_specific_issues():
    """Test specific scenarios that might cause issues."""
    
    print("\n" + "="*80)
    print("SPECIFIC ISSUE TESTING")
    print("="*80)
    
    provider = VehicleDataProvider(cache_enabled=False)
    
    # Test 1: Popular vehicle that should work
    print("\nTEST 1: Popular Vehicle (2020 Ford F-150)")
    print("-" * 50)
    vin1 = "1FTFW1ET5LFC10312"  # 2020 Ford F-150
    success, data, error = provider.get_vehicle_by_vin(vin1)
    
    # Test 2: Less common vehicle
    print("\nTEST 2: Less Common Vehicle")
    print("-" * 50)
    vin2 = "JM1BL1SF4A1234567"  # Mazda3
    success, data, error = provider.get_vehicle_by_vin(vin2)
    
    # Test 3: Commercial vehicle
    print("\nTEST 3: Commercial Vehicle")
    print("-" * 50)
    vin3 = "1GCEK14T74Z123456"  # Chevy Silverado
    success, data, error = provider.get_vehicle_by_vin(vin3)

if __name__ == "__main__":
    print("Starting VIN Processing Diagnostic Test...")
    print("This will help identify the root causes of data-related issues.")
    print()
    
    # Run main diagnostic test
    successful, partial, failed = test_vin_processing()
    
    # Run specific issue tests
    test_specific_issues()
    
    print("\nDiagnostic test complete!")
    print("Review the DIAGNOSTIC log messages above to identify issues.") 