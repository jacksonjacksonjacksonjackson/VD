"""
processor.py

Data processing pipeline for the Fleet Electrification Analyzer.
Handles CSV input/output, threading, and parallel processing.
"""

import os
import csv
import time
import datetime
import threading
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Any, Optional, Callable, Tuple, Set, Iterator
from pathlib import Path

from settings import MAX_THREADS, COLUMN_NAME_MAP, ALL_FUEL_ECONOMY_FIELDS
from utils import safe_cast, validate_vin, timestamp
from data.models import FleetVehicle, VehicleIdentification, FuelEconomyData, Fleet
from data.providers import VehicleDataProvider

# Set up module logger
logger = logging.getLogger(__name__)

###############################################################################
# CSV Input/Output
###############################################################################

class CsvReader:
    """Handles reading VINs and optional data from CSV files."""
    
    def __init__(self, file_path: str):
        """
        Initialize the CSV reader.
        
        Args:
            file_path: Path to the CSV file
        """
        self.file_path = file_path
    
    def read_vins(self) -> List[str]:
        """
        Read VINs from the CSV file.
        
        Returns:
            List of valid VINs
        """
        vins = []
        
        try:
            with open(self.file_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                
                # Check for VIN column
                vin_column = self._find_vin_column(reader.fieldnames)
                if not vin_column:
                    logger.error("No VIN column found in CSV")
                    return []
                
                # Read VINs
                for row in reader:
                    vin = row.get(vin_column, "").strip()
                    if vin and validate_vin(vin):
                        vins.append(vin)
                    elif vin:
                        logger.warning(f"Skipping invalid VIN: {vin}")
        
        except Exception as e:
            logger.error(f"Error reading CSV file: {e}")
            return []
        
        return vins
    
    def read_data(self) -> Dict[str, Dict[str, Any]]:
        """
        Read VINs and all additional data from the CSV file.
        
        Returns:
            Dictionary mapping VINs to their associated data
        """
        data = {}
        
        try:
            with open(self.file_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                
                # Check for VIN column
                vin_column = self._find_vin_column(reader.fieldnames)
                if not vin_column:
                    logger.error("No VIN column found in CSV")
                    return {}
                
                # Read all rows
                for row in reader:
                    vin = row.get(vin_column, "").strip()
                    if not vin:
                        continue
                    
                    # Store all data for this VIN
                    data[vin] = {k: v.strip() for k, v in row.items() if k != vin_column}
        
        except Exception as e:
            logger.error(f"Error reading CSV file: {e}")
            return {}
        
        return data
    
    def _find_vin_column(self, fieldnames: Optional[List[str]]) -> Optional[str]:
        """
        Find the column containing VINs in the CSV.
        
        Args:
            fieldnames: List of column names
            
        Returns:
            Name of the VIN column or None if not found
        """
        if not fieldnames:
            return None
        
        # Common VIN column names
        vin_columns = ['VIN', 'VINs', 'VIN Number', 'Vehicle Identification Number']
        
        # Check for exact matches
        for col in vin_columns:
            if col in fieldnames:
                return col
        
        # Check for case-insensitive matches
        lower_fieldnames = [f.lower() for f in fieldnames]
        for col in vin_columns:
            if col.lower() in lower_fieldnames:
                idx = lower_fieldnames.index(col.lower())
                return fieldnames[idx]
        
        # Check for partial matches
        for field in fieldnames:
            if 'vin' in field.lower():
                return field
        
        return None


class CsvWriter:
    """Handles writing fleet vehicle data to CSV files."""
    
    def __init__(self, file_path: str):
        """
        Initialize the CSV writer.
        
        Args:
            file_path: Path to the output CSV file
        """
        self.file_path = file_path
    
    def write_vehicles(self, vehicles: List[FleetVehicle], 
                      fields: Optional[List[str]] = None) -> bool:
        """
        Write vehicles to CSV file.
        
        Args:
            vehicles: List of FleetVehicle objects
            fields: List of fields to include (None for all)
            
        Returns:
            True if successful
        """
        if not vehicles:
            logger.warning("No vehicles to write")
            return False
        
        try:
            # Determine fields to write
            if not fields:
                # Use all fields from the first vehicle as a reference
                sample = vehicles[0].to_row_dict()
                fields = list(sample.keys())
            
            # Create friendly field names for CSV header
            header = [COLUMN_NAME_MAP.get(field, field) for field in fields]
            
            with open(self.file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # Write header
                writer.writerow(header)
                
                # Write vehicle data
                for vehicle in vehicles:
                    row_dict = vehicle.to_row_dict()
                    row = [row_dict.get(field, "") for field in fields]
                    writer.writerow(row)
            
            return True
            
        except Exception as e:
            logger.error(f"Error writing CSV file: {e}")
            return False
    
    def write_results(self, results: Dict[str, Dict[str, Any]],
                     fields: Optional[List[str]] = None) -> bool:
        """
        Write raw processing results to CSV file.
        
        Args:
            results: Dictionary mapping VINs to result data
            fields: List of fields to include (None for all)
            
        Returns:
            True if successful
        """
        if not results:
            logger.warning("No results to write")
            return False
        
        try:
            # Determine fields to write
            if not fields:
                # Collect all possible fields from all results
                all_fields = set()
                for vin, result in results.items():
                    if result.get("success") and "data" in result:
                        all_fields.update(self._flatten_dict(result["data"]).keys())
                
                fields = ["VIN", "success", "error"] + sorted(all_fields)
            
            with open(self.file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # Write header
                writer.writerow(fields)
                
                # Write result data
                for vin, result in results.items():
                    row = [vin, result.get("success", False), result.get("error", "")]
                    
                    # Add data fields if successful
                    flat_data = self._flatten_dict(result.get("data", {}))
                    for field in fields[3:]:  # Skip VIN, success, error
                        row.append(flat_data.get(field, ""))
                    
                    writer.writerow(row)
            
            return True
            
        except Exception as e:
            logger.error(f"Error writing results CSV: {e}")
            return False
    
    def _flatten_dict(self, data: Dict[str, Any], prefix: str = "") -> Dict[str, Any]:
        """
        Flatten a nested dictionary.
        
        Args:
            data: Dictionary to flatten
            prefix: Prefix for nested keys
            
        Returns:
            Flattened dictionary
        """
        result = {}
        
        for key, value in data.items():
            new_key = f"{prefix}_{key}" if prefix else key
            
            if isinstance(value, dict):
                # Recursively flatten nested dictionaries
                result.update(self._flatten_dict(value, new_key))
            else:
                # Add leaf values
                result[new_key] = value
        
        return result


###############################################################################
# Processing Pipeline
###############################################################################

class ProcessingPipeline:
    """
    Handles the end-to-end process of reading VINs, fetching data,
    and creating FleetVehicle objects.
    """
    
    def __init__(self, input_path: str, output_path: str = "", 
                max_threads: int = MAX_THREADS):
        """
        Initialize the processing pipeline.
        
        Args:
            input_path: Path to input CSV
            output_path: Path to output CSV (empty for no output)
            max_threads: Maximum number of worker threads
        """
        self.input_path = input_path
        self.output_path = output_path
        self.max_threads = max_threads
        
        self.provider = VehicleDataProvider(cache_enabled=True)
        self.stop_event = threading.Event()
        self.vehicles = []
        self.results = {}
    
    def process(self, 
               log_callback: Optional[Callable[[str], None]] = None,
               progress_callback: Optional[Callable[[int, int], None]] = None,
               done_callback: Optional[Callable[[List[FleetVehicle]], None]] = None) -> None:
        """
        Run the processing pipeline.
        
        Args:
            log_callback: Callback for logging messages
            progress_callback: Callback for progress updates
            done_callback: Callback for completion notification
        """
        # Reset state
        self.stop_event.clear()
        self.vehicles = []
        self.results = {}
        
        # Logging function (use callback if provided, else log to console)
        def log(message: str) -> None:
            logger.info(message)
            if log_callback:
                log_callback(message)
        
        log(f"Starting processing at {timestamp()}")
        
        # Step 1: Read VINs and additional data from CSV
        csv_reader = CsvReader(self.input_path)
        vins = csv_reader.read_vins()
        
        if not vins:
            log("No valid VINs found in input file")
            if done_callback:
                done_callback([])
            return
        
        log(f"Found {len(vins)} valid VINs")
        
        # Read any additional data
        additional_data = csv_reader.read_data()
        
        # Step 2: Process VINs in parallel
        log(f"Processing with {self.max_threads} threads")
        total = len(vins)
        processed = 0
        
        # Process in batches for better progress reporting
        batch_size = min(100, max(10, total // 10))
        
        with ThreadPoolExecutor(max_workers=self.max_threads) as executor:
            # Submit all VINs for processing
            future_to_vin = {
                executor.submit(self._process_single_vin, vin): vin 
                for vin in vins
            }
            
            # Process results as they complete
            for future in as_completed(future_to_vin):
                if self.stop_event.is_set():
                    log("Processing stopped by user")
                    break
                
                vin = future_to_vin[future]
                
                try:
                    success, result, vehicle = future.result()
                    
                    # Store result and vehicle if successful
                    self.results[vin] = {
                        "success": success,
                        "data": result,
                        "error": "" if success else result.get("error", "Unknown error")
                    }
                    
                    if success and vehicle:
                        # Add any additional data from the CSV
                        if vin in additional_data:
                            self._add_additional_data(vehicle, additional_data[vin])
                        
                        self.vehicles.append(vehicle)
                
                except Exception as e:
                    logger.error(f"Error processing VIN {vin}: {e}")
                    self.results[vin] = {
                        "success": False,
                        "data": {},
                        "error": str(e)
                    }
                
                # Update progress
                processed += 1
                if progress_callback:
                    progress_callback(processed, total)
                
                # Log batch completion
                if processed % batch_size == 0 or processed == total:
                    log(f"Processed {processed}/{total} VINs ({processed/total:.1%})")
        
        # Step 3: Write output if path specified
        if self.output_path and not self.stop_event.is_set():
            csv_writer = CsvWriter(self.output_path)
            success = csv_writer.write_vehicles(self.vehicles)
            
            if success:
                log(f"Wrote {len(self.vehicles)} vehicles to {self.output_path}")
            else:
                log(f"Failed to write output file {self.output_path}")
        
        # Step 4: Create Fleet object
        fleet = self._create_fleet()
        
        # Step 5: Call done callback
        if done_callback and not self.stop_event.is_set():
            done_callback(self.vehicles)
        
        log(f"Finished processing at {timestamp()}")
    
    def stop(self) -> None:
        """Stop the processing pipeline."""
        self.stop_event.set()
    
    def _process_single_vin(self, vin: str) -> Tuple[bool, Dict[str, Any], Optional[FleetVehicle]]:
        """
        Process a single VIN.
        
        Args:
            vin: VIN to process
            
        Returns:
            Tuple of (success, result_dict, FleetVehicle or None)
        """
        try:
            # Skip if stop requested
            if self.stop_event.is_set():
                return False, {"error": "Processing stopped"}, None
            
            # Get vehicle data
            success, data, error = self.provider.get_vehicle_by_vin(vin)
            
            if not success:
                return False, {"error": error}, None
            
            # Create vehicle object
            vehicle = self._create_vehicle_from_data(vin, data)
            
            return True, data, vehicle
            
        except Exception as e:
            logger.error(f"Error processing VIN {vin}: {e}")
            return False, {"error": str(e)}, None
    
    def _create_vehicle_from_data(self, vin: str, data: Dict[str, Any]) -> FleetVehicle:
        """
        Create a FleetVehicle object from API response data.
        
        Args:
            vin: Vehicle VIN
            data: Data from vehicle data provider
            
        Returns:
            FleetVehicle object
        """
        # Create vehicle identification
        vehicle_id = VehicleIdentification.from_dict(data.get("vehicle_id", {}))
        vehicle_id.vin = vin  # Ensure VIN is set
        
        # Create fuel economy data
        fuel_economy = FuelEconomyData.from_dict(data.get("fuel_economy", {}))
        
        # Create fleet vehicle
        vehicle = FleetVehicle(
            vin=vin,
            vehicle_id=vehicle_id,
            fuel_economy=fuel_economy,
            match_confidence=data.get("match_confidence", 0.0),
            assumed_vehicle_id=data.get("assumed_vehicle_id", ""),
            assumed_vehicle_text=data.get("assumed_vehicle_text", "")
        )
        
        return vehicle
    
    def _add_additional_data(self, vehicle: FleetVehicle, data: Dict[str, str]) -> None:
        """
        Add additional data from CSV to a vehicle.
        
        Args:
            vehicle: Vehicle to update
            data: Additional data from CSV
        """
        # Process known fields
        for field, value in data.items():
            if not value.strip():
                continue
                
            if field.lower() in ["odometer", "odo", "mileage"]:
                vehicle.odometer = safe_cast(value, float, 0.0)
            
            elif field.lower() in ["annual_mileage", "annual mileage", "yearly_mileage"]:
                vehicle.annual_mileage = safe_cast(value, float, 0.0)
            
            elif field.lower() in ["asset_id", "asset id", "asset_number"]:
                vehicle.asset_id = value
            
            elif field.lower() in ["department", "dept", "division"]:
                vehicle.department = value
            
            elif field.lower() in ["location", "site", "facility"]:
                vehicle.location = value
            
            else:
                # Store in custom fields
                vehicle.custom_fields[field] = value
        
        # Calculate annual mileage if not provided but odometer is
        if vehicle.odometer > 0 and vehicle.annual_mileage == 0:
            try:
                year = int(vehicle.vehicle_id.year)
                current_year = datetime.datetime.now().year
                years = max(1, current_year - year)
                vehicle.annual_mileage = vehicle.odometer / years
            except (ValueError, ZeroDivisionError):
                pass
    
    def _create_fleet(self) -> Fleet:
        """
        Create a Fleet object from processed vehicles.
        
        Returns:
            Fleet object
        """
        # Create a default name based on date/time and number of vehicles
        fleet_name = f"Fleet Analysis - {datetime.datetime.now().strftime('%Y-%m-%d')} ({len(self.vehicles)} vehicles)"
        
        # Create fleet
        fleet = Fleet(
            name=fleet_name,
            vehicles=self.vehicles,
            creation_date=datetime.datetime.now(),
            last_modified=datetime.datetime.now()
        )
        
        return fleet


###############################################################################
# Batch Processor
###############################################################################

class BatchProcessor:
    """
    Handles batch processing of multiple VINs with progress tracking,
    logging, and error handling.
    """
    
    def __init__(self, max_threads: int = MAX_THREADS):
        """
        Initialize the batch processor.
        
        Args:
            max_threads: Maximum number of worker threads
        """
        self.max_threads = max_threads
        self.provider = VehicleDataProvider(cache_enabled=True)
        self.stop_event = threading.Event()
        self.current_pipeline = None
    
    def process_file(self, input_path: str, output_path: str,
                   log_callback: Optional[Callable[[str], None]] = None,
                   progress_callback: Optional[Callable[[int, int], None]] = None,
                   done_callback: Optional[Callable[[List[FleetVehicle]], None]] = None) -> None:
        """
        Process a CSV file in a background thread.
        
        Args:
            input_path: Path to input CSV
            output_path: Path to output CSV
            log_callback: Callback for logging messages
            progress_callback: Callback for progress updates
            done_callback: Callback for completion notification
        """
        # Reset stop event
        self.stop_event.clear()
        
        # Create processing pipeline
        self.current_pipeline = ProcessingPipeline(
            input_path=input_path,
            output_path=output_path,
            max_threads=self.max_threads
        )
        
        # Start processing thread
        thread = threading.Thread(
            target=self._run_pipeline,
            args=(self.current_pipeline, log_callback, progress_callback, done_callback),
            daemon=True
        )
        
        thread.start()
    
    def stop(self) -> None:
        """Stop the current processing job."""
        self.stop_event.set()
        if self.current_pipeline:
            self.current_pipeline.stop()
    
    def _run_pipeline(self, pipeline: ProcessingPipeline,
                    log_callback: Optional[Callable[[str], None]],
                    progress_callback: Optional[Callable[[int, int], None]],
                    done_callback: Optional[Callable[[List[FleetVehicle]], None]]) -> None:
        """
        Run the processing pipeline in a background thread.
        
        Args:
            pipeline: Processing pipeline to run
            log_callback: Callback for logging messages
            progress_callback: Callback for progress updates
            done_callback: Callback for completion notification
        """
        try:
            pipeline.process(
                log_callback=log_callback,
                progress_callback=progress_callback,
                done_callback=done_callback
            )
        except Exception as e:
            logger.error(f"Error in processing pipeline: {e}")
            if log_callback:
                log_callback(f"Error: {e}")
            
            # Call done callback with empty list to signal completion
            if done_callback:
                done_callback([])