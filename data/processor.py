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
from dataclasses import dataclass, field
import pandas as pd
import re

from settings import MAX_THREADS, COLUMN_NAME_MAP, ALL_FUEL_ECONOMY_FIELDS, ADDITIONAL_DATA_MAPPINGS
from utils import safe_cast, validate_vin, validate_vin_detailed, timestamp, ErrorCommunicator, ContextHelp
from data.models import FleetVehicle, VehicleIdentification, FuelEconomyData, Fleet
from data.providers import VehicleDataProvider

# Set up module logger
logger = logging.getLogger(__name__)

###############################################################################
# CSV File Validation and Preview
###############################################################################

@dataclass
class FileValidationResult:
    """Result of CSV file validation with enhanced error details."""
    valid: bool
    error_message: str = ""
    warning_messages: List[str] = None
    total_rows: int = 0
    column_count: int = 0
    vin_column: str = ""
    valid_vins: int = 0
    invalid_vins: int = 0
    detected_encoding: str = "utf-8"
    sample_data: List[Dict] = None
    additional_columns: Dict[str, str] = None
    error_category: str = ""  # For enhanced error handling
    suggested_fixes: List[str] = None  # User-friendly suggestions
    
    # Additional attributes expected by UI
    mapped_columns: Dict[str, str] = None  # Original column -> standard field mapping
    unmapped_columns: List[str] = None  # Columns that weren't mapped to standard fields
    sample_rows: List[Dict] = None  # Sample data rows for preview
    columns: List[str] = None  # List of column names
    fleet_management_fields: List[str] = None  # Detected fleet management fields
    sample_valid_vins: List[str] = None  # Sample of valid VINs for display
    sample_invalid_vins: List[str] = None  # Sample of invalid VINs with errors
    
    def __post_init__(self):
        if self.warning_messages is None:
            self.warning_messages = []
        if self.sample_data is None:
            self.sample_data = []
        if self.additional_columns is None:
            self.additional_columns = {}
        if self.suggested_fixes is None:
            self.suggested_fixes = []
        if self.mapped_columns is None:
            self.mapped_columns = {}
        if self.unmapped_columns is None:
            self.unmapped_columns = []
        if self.sample_rows is None:
            self.sample_rows = []
        if self.columns is None:
            self.columns = []
        if self.fleet_management_fields is None:
            self.fleet_management_fields = []
        if self.sample_valid_vins is None:
            self.sample_valid_vins = []
        if self.sample_invalid_vins is None:
            self.sample_invalid_vins = []

class CsvFileValidator:
    """Enhanced CSV file validator with user-friendly error messages."""
    
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.result = FileValidationResult(valid=False)
    
    def validate_and_preview(self) -> FileValidationResult:
        """
        Validate CSV file and return detailed results with user-friendly error messages.
        
        Returns:
            FileValidationResult with enhanced error details
        """
        try:
            # Check file existence
            if not os.path.exists(self.filepath):
                self.result.error_message, self.result.suggested_fixes = ErrorCommunicator.get_file_error_message(
                    self.filepath, "not_found"
                )
                self.result.error_category = "file_access"
                return self.result
            
            # Detect encoding
            self.result.detected_encoding = self._detect_encoding()
            
            # Read and validate file
            try:
                df = pd.read_csv(self.filepath, encoding=self.result.detected_encoding)
            except PermissionError:
                self.result.error_message, self.result.suggested_fixes = ErrorCommunicator.get_file_error_message(
                    self.filepath, "permission"
                )
                self.result.error_category = "file_access"
                return self.result
            except UnicodeDecodeError:
                self.result.error_message, self.result.suggested_fixes = ErrorCommunicator.get_file_error_message(
                    self.filepath, "encoding"
                )
                self.result.error_category = "file_access"
                return self.result
            except Exception as e:
                self.result.error_message, self.result.suggested_fixes = ErrorCommunicator.get_file_error_message(
                    self.filepath, "format"
                )
                self.result.error_category = "validation"
                return self.result
            
            # Basic file info
            self.result.total_rows = len(df)
            self.result.column_count = len(df.columns)
            
            # Check for empty file
            if self.result.total_rows == 0:
                self.result.error_message = f"The file '{os.path.basename(self.filepath)}' is empty or has no data rows."
                self.result.suggested_fixes = [
                    "Add VIN data to the file",
                    "Check if the file was saved correctly",
                    "Use the 'Download Sample CSV' button to see the expected format"
                ]
                self.result.error_category = "validation"
                return self.result
            
            # Find VIN column
            vin_column = self._find_vin_column(df)
            if not vin_column:
                self.result.error_message = f"No VIN column found in '{os.path.basename(self.filepath)}'."
                self.result.suggested_fixes = [
                    "Add a column header named 'VIN', 'VINs', or 'Vehicle_ID'",
                    "Check that the first row contains column headers",
                    "Use the 'Download Sample CSV' button to see the expected format"
                ]
                self.result.error_category = "validation"
                ContextHelp.show_help_dialog(None, "csv_format", "CSV Format Help")
                return self.result
            
            self.result.vin_column = vin_column
            
            # Validate VINs
            self._validate_vins(df, vin_column)
            
            # Map additional columns
            self._map_additional_columns(df)
            
            # Create sample data
            self._create_sample_data(df)
            
            # Final validation
            if self.result.valid_vins == 0:
                self.result.error_message = f"No valid VINs found in '{os.path.basename(self.filepath)}'."
                self.result.suggested_fixes = [
                    "Check that VINs are exactly 17 characters long",
                    "Verify VINs contain only letters and numbers (no I, O, Q)",
                    "Remove any placeholder or test VINs",
                    "Use real vehicle VINs from registration documents"
                ]
                self.result.error_category = "vin_format"
                return self.result
            
            # Success
            self.result.valid = True
            
            # Add warnings for invalid VINs
            if self.result.invalid_vins > 0:
                invalid_percentage = (self.result.invalid_vins / self.result.total_rows) * 100
                self.result.warning_messages.append(
                    f"{self.result.invalid_vins} invalid VINs ({invalid_percentage:.1f}%) will be included with error messages"
                )
            
            return self.result
            
        except Exception as e:
            logger.error(f"Unexpected error validating file {self.filepath}: {e}")
            self.result.error_message = f"An unexpected error occurred while validating the file: {str(e)}"
            self.result.suggested_fixes = [
                "Check if the file is corrupted",
                "Try opening the file in Excel to verify it's readable",
                "Contact support if the problem persists"
            ]
            self.result.error_category = "processing"
            return self.result
    
    def _detect_encoding(self) -> str:
        """
        Detect the encoding of the CSV file.
        
        Returns:
            Detected encoding string
        """
        encodings_to_try = ['utf-8', 'utf-8-sig', 'latin1', 'cp1252']
        
        for encoding in encodings_to_try:
            try:
                with open(self.filepath, 'r', encoding=encoding) as f:
                    f.read(1024)  # Read first 1KB
                return encoding
            except UnicodeDecodeError:
                continue
        
        # Default to utf-8 if all fail
        return 'utf-8'
    
    def _find_vin_column(self, df: pd.DataFrame) -> Optional[str]:
        """
        Find the column containing VINs in the CSV.
        
        Args:
            df: Pandas DataFrame containing the CSV data
            
        Returns:
            Name of the VIN column or None if not found
        """
        # Priority order: exact matches, case-insensitive, partial matches
        vin_patterns = [
            # Exact matches (highest priority)
            ['VIN', 'VINs', 'Vehicle_Identification_Number', 'VIN_Number'],
            # Common variations
            ['Vehicle Identification Number', 'VIN Number', 'Vehicle ID', 'VehicleID'],
            # Partial matches (lowest priority)
            ['vin', 'vehicle', 'identification']
        ]
        
        # Check exact matches first
        for pattern_group in vin_patterns[:2]:
            for pattern in pattern_group:
                # Exact match
                if pattern in df.columns:
                    return pattern
                
                # Case-insensitive match
                for column in df.columns:
                    if column.lower() == pattern.lower():
                        return column
        
        # Check partial matches
        for column in df.columns:
            column_lower = column.lower()
            if 'vin' in column_lower or ('vehicle' in column_lower and 'id' in column_lower):
                return column
        
        return None
    
    def _validate_vins(self, df: pd.DataFrame, vin_column: str) -> None:
        """
        Validate VINs in the CSV file.
        
        Args:
            df: Pandas DataFrame containing the CSV data
            vin_column: Name of the VIN column
        """
        self.result.valid_vins = 0
        self.result.invalid_vins = 0
        self.result.sample_valid_vins = []  # Initialize the lists
        self.result.sample_invalid_vins = []
        
        for index, row in df.iterrows():
            vin_value = row[vin_column]
            
            # Handle NaN values (empty cells) properly
            if pd.isna(vin_value) or vin_value is None:
                vin = ""
            else:
                # Convert to string and strip whitespace
                vin = str(vin_value).strip()
            
            if vin:
                is_valid, error_msg = validate_vin_detailed(vin)
                if is_valid:
                    self.result.valid_vins += 1
                    if len(self.result.sample_valid_vins) < 3:
                        self.result.sample_valid_vins.append(vin)
                else:
                    self.result.invalid_vins += 1
                    if len(self.result.sample_invalid_vins) < 3:
                        self.result.sample_invalid_vins.append(f"{vin} ({error_msg})")
        
        # Add warnings for invalid VINs
        if self.result.invalid_vins > 0:
            invalid_percentage = (self.result.invalid_vins / self.result.total_rows) * 100
            self.result.warning_messages.append(
                f"{self.result.invalid_vins} invalid VINs ({invalid_percentage:.1f}%) will be included with error messages"
            )
    
    def _map_additional_columns(self, df: pd.DataFrame) -> None:
        """
        Map CSV column names to standardized field names using comprehensive mappings.
        
        Args:
            df: Pandas DataFrame containing the CSV data
        """
        self.result.additional_columns = {}
        self.result.mapped_columns = {}
        self.result.unmapped_columns = []
        
        # Create reverse mapping from input column names to standard names
        column_mapping = {}
        for standard_field, variants in ADDITIONAL_DATA_MAPPINGS.items():
            for variant in variants:
                column_mapping[variant.lower()] = standard_field
        
        # Map each input column to a standard field
        for column in df.columns:
            if column == self.result.vin_column:
                continue  # Skip VIN column
                
            column_lower = column.lower().strip()
            
            # Check if this column maps to a known standard field
            if column_lower in column_mapping:
                standard_field = column_mapping[column_lower]
                self.result.additional_columns[column] = standard_field
                self.result.mapped_columns[column] = standard_field
            else:
                # Keep original column name for unmapped fields
                self.result.additional_columns[column] = column
                self.result.unmapped_columns.append(column)
        
        # Populate fleet management fields (mapped standard fields)
        self.result.fleet_management_fields = list(set(self.result.mapped_columns.values()))
        
        # Add informational messages about additional data
        if self.result.mapped_columns:
            mapped_count = len(self.result.mapped_columns)
            self.result.warning_messages.append(
                f"Detected {mapped_count} fleet management columns that will be auto-mapped"
            )
        
        if self.result.unmapped_columns:
            unmapped_count = len(self.result.unmapped_columns)
            self.result.warning_messages.append(
                f"Found {unmapped_count} custom columns that will be preserved as-is"
            )
    
    def _create_sample_data(self, df: pd.DataFrame) -> None:
        """
        Create sample data for preview.
        
        Args:
            df: Pandas DataFrame containing the CSV data
        """
        # Store columns list for UI
        self.result.columns = df.columns.tolist()
        
        # Create sample data (original format)
        self.result.sample_data = df.head(10).to_dict(orient='records')
        
        # Create sample rows (same data, different name for UI compatibility)
        self.result.sample_rows = self.result.sample_data.copy()


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
        self.validator = CsvFileValidator(file_path)
    
    def validate_file(self) -> FileValidationResult:
        """
        Validate the CSV file before processing.
        
        Returns:
            FileValidationResult with validation details
        """
        return self.validator.validate_and_preview()
    
    def read_vins(self) -> List[str]:
        """
        Read VINs from the CSV file.
        
        Returns:
            List of all VINs and placeholders (preserves ALL rows)
        """
        vins = []
        
        try:
            # First validate the file
            validation = self.validate_file()
            if not validation.valid:
                logger.error(f"CSV validation failed: {validation.error_message}")
                return []
            
            # Use detected encoding
            with open(self.file_path, 'r', newline='', encoding=validation.detected_encoding) as f:
                reader = csv.DictReader(f)
                
                # Add detailed logging
                logger.info(f"🔧 DEBUG: Starting to read VINs from {self.file_path}")
                logger.info(f"🔧 DEBUG: VIN column detected: {validation.vin_column}")
                
                # Read ALL rows including invalid ones - preserve original order
                for row_idx, row in enumerate(reader):
                    vin_value = row.get(validation.vin_column, "")
                    
                    # Handle different types of empty values
                    if vin_value is None or vin_value == "" or (isinstance(vin_value, float) and pd.isna(vin_value)):
                        vin = ""
                    else:
                        vin = str(vin_value).strip()
                    
                    if vin:
                        # Add all VINs (valid and invalid) to preserve them in results
                        vins.append(vin)
                        logger.debug(f"🔧 DEBUG: Row {row_idx + 1}: Added VIN {vin}")
                    else:
                        # Handle rows with missing VINs - create placeholder to preserve row
                        if any(v for v in row.values() if v is not None and str(v).strip()):  # Row has other data but no VIN
                            placeholder_vin = f"MISSING_VIN_ROW_{row_idx + 1}"
                            vins.append(placeholder_vin)
                            logger.warning(f"🔧 DEBUG: Row {row_idx + 1}: Missing VIN but has other data, created placeholder: {placeholder_vin}")
                        else:
                            # Completely empty row - create placeholder to maintain row count
                            placeholder_vin = f"EMPTY_ROW_{row_idx + 1}"
                            vins.append(placeholder_vin)
                            logger.warning(f"🔧 DEBUG: Row {row_idx + 1}: Empty row, created placeholder: {placeholder_vin}")
                
                logger.info(f"🔧 DEBUG: Total VINs read from CSV: {len(vins)}")
                valid_count = sum(1 for vin in vins if not (vin.startswith("MISSING_VIN_ROW_") or vin.startswith("EMPTY_ROW_")))
                placeholder_count = len(vins) - valid_count
                logger.info(f"🔧 DEBUG: Valid VINs: {valid_count}, Placeholder VINs: {placeholder_count}")
        
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
            # First validate the file
            validation = self.validate_file()
            if not validation.valid:
                logger.error(f"CSV validation failed: {validation.error_message}")
                return {}
            
            # Use detected encoding
            with open(self.file_path, 'r', newline='', encoding=validation.detected_encoding) as f:
                reader = csv.DictReader(f)
                
                # Read all rows
                for row in reader:
                    vin = row.get(validation.vin_column, "").strip()
                    if not vin:
                        continue
                    
                    # Store all data for this VIN
                    data[vin] = {k: v.strip() for k, v in row.items() if k != validation.vin_column}
        
        except Exception as e:
            logger.error(f"Error reading CSV file: {e}")
            return {}
        
        return data
    
    def _find_vin_column(self, fieldnames: Optional[List[str]]) -> Optional[str]:
        """
        Find the column containing VINs in the CSV.
        DEPRECATED: Use CsvFileValidator._find_vin_column_enhanced instead
        
        Args:
            fieldnames: List of column names
            
        Returns:
            Name of the VIN column or None if not found
        """
        # Use the enhanced validator method
        if fieldnames:
            validator = CsvFileValidator("")
            return validator._find_vin_column(fieldnames)
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
        
        # Track VIN order for preserving input order in results
        self.vin_order = []
        self.vin_to_index = {}
    
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
        
        # Track VIN order for preserving input order in results
        self.vin_order = vins.copy()
        self.vin_to_index = {vin: idx for idx, vin in enumerate(vins)}
        
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
                    
                    # Add detailed logging for vehicle processing
                    logger.info(f"🔧 DEBUG: Processed VIN {vin} - Success: {success}")
                    if vehicle:
                        logger.info(f"🔧 DEBUG: Vehicle created - Processing success: {vehicle.processing_success}, Error: '{vehicle.processing_error}'")
                    else:
                        logger.warning(f"🔧 DEBUG: No vehicle object returned for VIN {vin}")
                    
                    # Store result
                    self.results[vin] = {
                        "success": success,
                        "data": result,
                        "error": "" if success else result.get("error", "Unknown error")
                    }
                    
                    # Always add vehicle to results (successful or failed)
                    if vehicle:
                        # Add any additional data from the CSV
                        if vin in additional_data:
                            self._add_additional_data(vehicle, additional_data[vin])
                            logger.debug(f"🔧 DEBUG: Added additional data to vehicle {vin}")
                        
                        self.vehicles.append(vehicle)
                        logger.debug(f"🔧 DEBUG: Added vehicle to collection. Total vehicles: {len(self.vehicles)}")
                    else:
                        logger.error(f"🔧 DEBUG: Vehicle object is None for VIN {vin} - this should not happen!")
                
                except Exception as e:
                    logger.error(f"Error processing VIN {vin}: {e}")
                    logger.error(f"🔧 DEBUG: Exception details: {type(e).__name__}: {str(e)}")
                    self.results[vin] = {
                        "success": False,
                        "data": {},
                        "error": str(e)
                    }
                    
                    # Create a failed vehicle for this exception
                    failed_vehicle = self._create_failed_vehicle(vin, f"Processing Exception: {str(e)}", self.vin_to_index.get(vin, processed))
                    self.vehicles.append(failed_vehicle)
                    logger.info(f"🔧 DEBUG: Created failed vehicle for exception. Total vehicles: {len(self.vehicles)}")
                
                # Update progress
                processed += 1
                if progress_callback:
                    progress_callback(processed, total)
                
                # Log batch completion
                if processed % batch_size == 0 or processed == total:
                    log(f"Processed {processed}/{total} VINs ({processed/total:.1%})")
        
        # Step 3: Write output if path specified
        if self.output_path and not self.stop_event.is_set():
            logger.info(f"🔧 DEBUG: Writing {len(self.vehicles)} vehicles to {self.output_path}")
            csv_writer = CsvWriter(self.output_path)
            success = csv_writer.write_vehicles(self.vehicles)
            
            if success:
                log(f"Wrote {len(self.vehicles)} vehicles to {self.output_path}")
                logger.info(f"🔧 DEBUG: Successfully wrote {len(self.vehicles)} vehicles to CSV")
            else:
                log(f"Failed to write output file {self.output_path}")
                logger.error(f"🔧 DEBUG: Failed to write vehicles to CSV")
        
        # Step 4: Create Fleet object
        fleet = self._create_fleet()
        
        # Step 5: Call done callback
        if done_callback and not self.stop_event.is_set():
            logger.info(f"🔧 DEBUG: Calling done_callback with {len(self.vehicles)} vehicles")
            logger.info(f"🔧 DEBUG: done_callback invoked from thread: {threading.current_thread().name}")
            logger.info(f"🔧 DEBUG: Is main thread: {threading.current_thread() == threading.main_thread()}")
            
            for i, vehicle in enumerate(self.vehicles[:5]):  # Log first 5 vehicles for debugging
                logger.info(f"🔧 DEBUG: Vehicle {i+1}: VIN={vehicle.vin}, Success={vehicle.processing_success}, Error='{vehicle.processing_error}'")
            if len(self.vehicles) > 5:
                logger.info(f"🔧 DEBUG: ... and {len(self.vehicles) - 5} more vehicles")
            
            try:
                logger.info(f"🔧 DEBUG: About to call done_callback function")
                done_callback(self.vehicles)
                logger.info(f"🔧 DEBUG: done_callback completed successfully")
            except Exception as callback_e:
                logger.error(f"🔧 DEBUG: Exception in done_callback: {callback_e}")
                logger.error(f"🔧 DEBUG: done_callback exception type: {type(callback_e).__name__}")
                import traceback
                logger.error(f"🔧 DEBUG: done_callback traceback: {traceback.format_exc()}")
                raise  # Re-raise to see the full error chain
        
        log(f"Finished processing at {timestamp()}")
        logger.info(f"🔧 DEBUG: Processing complete. Final vehicle count: {len(self.vehicles)}")
    
    def stop(self) -> None:
        """Stop the processing pipeline."""
        self.stop_event.set()
    
    def _process_single_vin(self, vin: str) -> Tuple[bool, Dict[str, Any], Optional[FleetVehicle]]:
        """
        Process a single VIN and ALWAYS return a vehicle object (even for failures).
        This ensures ALL input rows are preserved in the results.
        
        Args:
            vin: VIN to process (could be valid VIN, invalid VIN, or placeholder)
            
        Returns:
            Tuple of (success, result_dict, FleetVehicle) - FleetVehicle is NEVER None
        """
        # Set input order index
        input_order_index = self.vin_to_index.get(vin, 0)
        
        try:
            # Skip if stop requested
            if self.stop_event.is_set():
                return False, {"error": "Processing stopped"}, self._create_failed_vehicle(
                    vin, "Processing was stopped by user", input_order_index)
            
            # Handle placeholder VINs for missing VIN rows
            if vin.startswith("MISSING_VIN_ROW_"):
                return False, {"error": "Missing VIN"}, self._create_failed_vehicle(
                    vin, "VIN is missing from this row", input_order_index)
            
            # Handle placeholder VINs for empty rows
            if vin.startswith("EMPTY_ROW_"):
                return False, {"error": "Empty row"}, self._create_failed_vehicle(
                    vin, "This row was empty in the input CSV", input_order_index)
            
            # Validate VIN format first
            logger.info(f"🔧 DIAGNOSTIC: VIN VALIDATION START: {vin}")
            is_valid, validation_error = validate_vin_detailed(vin)
            if not is_valid:
                logger.error(f"🔧 DIAGNOSTIC: VIN VALIDATION FAILED: {vin} - {validation_error}")
                return False, {"error": validation_error}, self._create_failed_vehicle(
                    vin, f"Invalid VIN format: {validation_error}", input_order_index)
            logger.info(f"🔧 DIAGNOSTIC: VIN VALIDATION PASSED: {vin}")
            
            # Get vehicle data from APIs
            logger.info(f"🔧 DIAGNOSTIC: API CALL START: {vin}")
            success, data, error = self.provider.get_vehicle_by_vin(vin)
            
            if not success:
                logger.error(f"🔧 DIAGNOSTIC: API CALL FAILED: {vin} - {error}")
                return False, {"error": error}, self._create_failed_vehicle(
                    vin, f"API Error: {error}", input_order_index)
            logger.info(f"🔧 DIAGNOSTIC: API CALL SUCCESS: {vin}")
            
            # Create successful vehicle object
            logger.info(f"🔧 DIAGNOSTIC: VEHICLE CREATION START: {vin}")
            vehicle = self._create_vehicle_from_data(vin, data)
            vehicle.input_order_index = input_order_index
            vehicle.processing_success = True
            vehicle.processing_error = ""
            logger.info(f"🔧 DIAGNOSTIC: VEHICLE CREATION SUCCESS: {vin}")
            
            # Calculate data quality score
            logger.info(f"🔧 DIAGNOSTIC: QUALITY SCORE CALCULATION START: {vin}")
            logger.info(f"🔧 DIAGNOSTIC: ENGINE DISPLACEMENT VALUE: {vin} - '{vehicle.vehicle_id.engine_displacement}'")
            logger.info(f"🔧 DIAGNOSTIC: ENGINE CYLINDERS VALUE: {vin} - '{vehicle.vehicle_id.engine_cylinders}'")
            
            try:
                vehicle.data_quality_score = self._calculate_quality_score(vehicle)
                logger.info(f"🔧 DIAGNOSTIC: QUALITY SCORE CALCULATION SUCCESS: {vin} - Score: {vehicle.data_quality_score}")
            except Exception as quality_error:
                # Quality score calculation failed, but VIN is still valid - use default score
                logger.warning(f"🔧 DIAGNOSTIC: QUALITY SCORE CALCULATION FAILED: {vin} - {quality_error}")
                logger.warning(f"🔧 DIAGNOSTIC: Using default quality score for valid VIN: {vin}")
                vehicle.data_quality_score = 50.0  # Default moderate score
            
            return True, data, vehicle
            
        except Exception as e:
            import traceback
            logger.error(f"🔧 DIAGNOSTIC: EXCEPTION CAUGHT IN _process_single_vin: {vin}")
            logger.error(f"🔧 DIAGNOSTIC: EXCEPTION TYPE: {type(e).__name__}")
            logger.error(f"🔧 DIAGNOSTIC: EXCEPTION MESSAGE: {str(e)}")
            logger.error(f"🔧 DIAGNOSTIC: EXCEPTION TRACEBACK: {traceback.format_exc()}")
            
            # For data processing errors on valid VINs, create a partial success vehicle
            # This prevents valid VINs from being marked as "Invalid VIN"
            try:
                # Try to create a basic vehicle with minimal data
                basic_vehicle = FleetVehicle(vin=vin)
                basic_vehicle.input_order_index = input_order_index
                basic_vehicle.processing_success = False
                basic_vehicle.processing_error = f"Data processing error: {str(e)}"
                basic_vehicle.data_quality_score = 0.0
                
                # Set make/model to indicate processing issue rather than invalid VIN
                basic_vehicle.vehicle_id.make = "Processing Error"
                basic_vehicle.vehicle_id.model = "Data unavailable"
                basic_vehicle.vehicle_id.year = "Check logs"
                
                logger.info(f"🔧 DIAGNOSTIC: CREATED PARTIAL VEHICLE FOR PROCESSING ERROR: {vin}")
                return False, {"error": str(e)}, basic_vehicle
                
            except Exception as creation_error:
                # Fallback to failed vehicle creation
                logger.error(f"🔧 DIAGNOSTIC: FAILED TO CREATE PARTIAL VEHICLE: {vin} - {creation_error}")
                return False, {"error": str(e)}, self._create_failed_vehicle(
                    vin, f"Processing Exception: {str(e)}", input_order_index)
    
    def _create_failed_vehicle(self, vin: str, error_msg: str, input_order_index: int) -> FleetVehicle:
        """
        Create a FleetVehicle object for failed processing with comprehensive error details.
        
        Args:
            vin: The VIN (or placeholder) that failed
            error_msg: Detailed error message
            input_order_index: Original position in input file
            
        Returns:
            FleetVehicle object with error information populated
        """
        vehicle = FleetVehicle(vin=vin)
        vehicle.input_order_index = input_order_index
        vehicle.processing_success = False
        vehicle.processing_error = error_msg
        vehicle.data_quality_score = 0.0
        
        # Set some basic identifying information for failed vehicles
        if vin.startswith("MISSING_VIN_ROW_"):
            # Extract row number for display
            row_num = vin.replace("MISSING_VIN_ROW_", "")
            vehicle.vin = f"Row {row_num} (No VIN)"
            vehicle.vehicle_id.make = "Missing VIN"
            vehicle.vehicle_id.model = "Check input data"
            vehicle.vehicle_id.year = "N/A"
        elif vin.startswith("EMPTY_ROW_"):
            # Extract row number for display
            row_num = vin.replace("EMPTY_ROW_", "")
            vehicle.vin = f"Row {row_num} (Empty)"
            vehicle.vehicle_id.make = "Empty Row"
            vehicle.vehicle_id.model = "No data provided"
            vehicle.vehicle_id.year = "N/A"
        elif len(vin) != 17:
            vehicle.vehicle_id.make = "Invalid VIN"
            vehicle.vehicle_id.model = f"Length: {len(vin)} (need 17)"
            vehicle.vehicle_id.year = "N/A"
        else:
            # VIN is 17 characters but failed validation
            vehicle.vehicle_id.make = "Invalid VIN"
            vehicle.vehicle_id.model = "Failed validation"
            vehicle.vehicle_id.year = "N/A"
        
        # Add timestamp for when the error occurred
        vehicle.processing_date = datetime.datetime.now()
        
        return vehicle
    
    def _calculate_quality_score(self, vehicle: FleetVehicle) -> float:
        """
        Calculate a comprehensive data quality score (0-100) for a vehicle.
        Enhanced for Step 13 with commercial vehicle fields and consistency checks.
        """
        score = 0.0
        
        # === CORE VEHICLE DATA (35 points) ===
        # Essential identification fields
        if vehicle.vehicle_id.year: score += 8
        if vehicle.vehicle_id.make: score += 8  
        if vehicle.vehicle_id.model: score += 8
        if vehicle.vehicle_id.fuel_type: score += 6
        if vehicle.vehicle_id.body_class: score += 5
        
        # === FUEL ECONOMY DATA (25 points) ===
        # MPG completeness
        if vehicle.fuel_economy.combined_mpg > 0: score += 12
        if vehicle.fuel_economy.city_mpg > 0: score += 6
        if vehicle.fuel_economy.highway_mpg > 0: score += 6
        # CO2 data
        if vehicle.fuel_economy.co2_primary > 0: score += 1
        
        # === COMMERCIAL VEHICLE DATA (15 points) ===
        # GVWR and classification data
        if vehicle.vehicle_id.gvwr_pounds > 0: score += 4
        if vehicle.vehicle_id.commercial_category: score += 3
        if vehicle.vehicle_id.engine_power_hp: score += 2
        if vehicle.vehicle_id.engine_type: score += 2
        if vehicle.vehicle_id.vehicle_class: score += 2
        if vehicle.vehicle_id.series or vehicle.vehicle_id.trim: score += 2
        
        # === TECHNICAL DETAILS (10 points) ===
        if vehicle.vehicle_id.engine_displacement: score += 3
        if vehicle.vehicle_id.engine_cylinders: score += 2
        if vehicle.vehicle_id.transmission: score += 2
        if vehicle.vehicle_id.drive_type: score += 2
        if vehicle.vehicle_id.fuel_type_secondary: score += 1
        
        # === MATCH CONFIDENCE (10 points) ===
        # API matching confidence
        confidence_score = vehicle.match_confidence / 10.0  # Convert 0-100 to 0-10
        score += confidence_score
        
        # === DATA CONSISTENCY BONUS (5 points) ===
        consistency_bonus = self._calculate_consistency_bonus(vehicle)
        score += consistency_bonus
        
        return min(score, 100.0)
    
    def _calculate_consistency_bonus(self, vehicle: FleetVehicle) -> float:
        """
        Calculate bonus points for data consistency across sources.
        
        Args:
            vehicle: Vehicle to assess
            
        Returns:
            Bonus points (0-5) for data consistency
        """
        bonus = 0.0
        
        # Check VIN vs year consistency (basic checksum validation)
        if vehicle.vehicle_id.year and len(vehicle.vin) == 17:
            try:
                vin_year_digit = vehicle.vin[9]  # 10th position indicates model year
                year_int = int(vehicle.vehicle_id.year)
                
                # VIN year encoding is complex, but we can do basic checks
                if year_int >= 2010:  # Modern VINs use A-Z for 2010-2039
                    bonus += 1.0
                elif year_int >= 1980:  # Earlier system
                    bonus += 0.5
            except (ValueError, IndexError):
                pass
        
        # Check fuel type consistency with commercial classification
        if vehicle.vehicle_id.is_commercial and vehicle.vehicle_id.is_diesel:
            # Commercial + diesel is common and expected
            bonus += 1.0
        elif not vehicle.vehicle_id.is_commercial and not vehicle.vehicle_id.is_diesel:
            # Passenger car without diesel is common
            bonus += 0.5
        
        # Check GVWR vs body class consistency
        if vehicle.vehicle_id.gvwr_pounds > 0 and vehicle.vehicle_id.body_class:
            body_lower = vehicle.vehicle_id.body_class.lower()
            
            # Light duty vehicles should have reasonable GVWR
            if vehicle.vehicle_id.gvwr_pounds <= 8500:
                if any(term in body_lower for term in ['sedan', 'coupe', 'hatchback', 'suv', 'wagon']):
                    bonus += 1.0
            # Heavy duty vehicles
            elif vehicle.vehicle_id.gvwr_pounds > 19500:
                if any(term in body_lower for term in ['truck', 'bus', 'commercial', 'chassis']):
                    bonus += 1.0
        
        # Check MPG vs vehicle size consistency
        if vehicle.fuel_economy.combined_mpg > 0 and vehicle.vehicle_id.gvwr_pounds > 0:
            # Heavier vehicles typically have lower MPG
            if vehicle.vehicle_id.gvwr_pounds > 8500 and vehicle.fuel_economy.combined_mpg < 25:
                bonus += 0.5  # Expected for heavy vehicles
            elif vehicle.vehicle_id.gvwr_pounds <= 6000 and vehicle.fuel_economy.combined_mpg > 20:
                bonus += 0.5  # Expected for lighter vehicles
        
        # Check engine specs consistency
        if vehicle.vehicle_id.engine_displacement and vehicle.vehicle_id.engine_cylinders:
            try:
                logger.info(f"🔧 DIAGNOSTIC: ENGINE SPEC CONSISTENCY CHECK: {vehicle.vin}")
                # Extract numeric value from engine displacement, handling formats like "361/479"
                displacement_str = str(vehicle.vehicle_id.engine_displacement).strip()
                cylinders_str = str(vehicle.vehicle_id.engine_cylinders).strip()
                
                logger.info(f"🔧 DIAGNOSTIC: DISPLACEMENT STRING: {vehicle.vin} - '{displacement_str}'")
                logger.info(f"🔧 DIAGNOSTIC: CYLINDERS STRING: {vehicle.vin} - '{cylinders_str}'")
                
                # Extract first numeric value for displacement
                disp_match = re.search(r'(\d+\.?\d*)', displacement_str)
                if disp_match:
                    logger.info(f"🔧 DIAGNOSTIC: DISPLACEMENT REGEX MATCH: {vehicle.vin} - '{disp_match.group(1)}'")
                    displacement = float(disp_match.group(1))
                    cylinders = int(cylinders_str)
                    
                    logger.info(f"🔧 DIAGNOSTIC: DISPLACEMENT PARSED: {vehicle.vin} - {displacement}")
                    logger.info(f"🔧 DIAGNOSTIC: CYLINDERS PARSED: {vehicle.vin} - {cylinders}")
                    
                    # Reasonable displacement per cylinder (0.3-1.0L typically)
                    displacement_per_cylinder = displacement / cylinders
                    logger.info(f"🔧 DIAGNOSTIC: DISPLACEMENT PER CYLINDER: {vehicle.vin} - {displacement_per_cylinder}")
                    if 0.3 <= displacement_per_cylinder <= 1.0:
                        bonus += 0.5
                        logger.info(f"🔧 DIAGNOSTIC: ENGINE SPEC CONSISTENCY BONUS AWARDED: {vehicle.vin}")
                else:
                    logger.warning(f"🔧 DIAGNOSTIC: NO DISPLACEMENT REGEX MATCH: {vehicle.vin} - '{displacement_str}'")
            except (ValueError, ZeroDivisionError, AttributeError) as e:
                logger.error(f"🔧 DIAGNOSTIC: ENGINE SPEC CONSISTENCY ERROR: {vehicle.vin} - {type(e).__name__}: {str(e)}")
                pass
        
        return min(bonus, 5.0)
    
    def _create_vehicle_from_data(self, vin: str, data: Dict[str, Any]) -> FleetVehicle:
        """
        Create a FleetVehicle object from API response data.
        
        Args:
            vin: Vehicle VIN
            data: Data from vehicle data provider
            
        Returns:
            FleetVehicle object
        """
        # Get vehicle identification data and ensure VIN is included
        vehicle_id_data = data.get("vehicle_id", {}).copy()
        vehicle_id_data["vin"] = vin  # Always set the VIN to avoid constructor errors
        
        # Create vehicle identification
        vehicle_id = VehicleIdentification.from_dict(vehicle_id_data)
        
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
        Add additional data from CSV to a vehicle using comprehensive column mapping.
        
        Args:
            vehicle: Vehicle to update
            data: Additional data from CSV
        """
        # Create a mapping of standardized field names to actual values
        mapped_data = self._map_additional_columns(data)
        
        # Process mapped fields to vehicle attributes
        for standard_field, value in mapped_data.items():
            if not value.strip():
                continue
                
            # Handle known fields with proper type conversion
            if standard_field == "odometer":
                vehicle.odometer = safe_cast(value, float, 0.0)
            
            elif standard_field == "annual_mileage":
                vehicle.annual_mileage = safe_cast(value, float, 0.0)
            
            elif standard_field == "asset_id":
                vehicle.asset_id = value
            
            elif standard_field == "department":
                vehicle.department = value
            
            elif standard_field == "location":
                vehicle.location = value
            
            elif standard_field == "acquisition_date":
                # Try to parse date
                try:
                    # Support common date formats
                    for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%Y/%m/%d"]:
                        try:
                            vehicle.acquisition_date = datetime.datetime.strptime(value, fmt).date()
                            break
                        except ValueError:
                            continue
                except Exception:
                    # Store as custom field if date parsing fails
                    vehicle.custom_fields[standard_field] = value
            
            elif standard_field == "retire_date":
                # Try to parse date
                try:
                    for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%Y/%m/%d"]:
                        try:
                            vehicle.retire_date = datetime.datetime.strptime(value, fmt).date()
                            break
                        except ValueError:
                            continue
                except Exception:
                    # Store as custom field if date parsing fails
                    vehicle.custom_fields[standard_field] = value
            
            else:
                # Store all other fields in custom_fields
                vehicle.custom_fields[standard_field] = value
        
        # Calculate annual mileage if not provided but odometer is
        if vehicle.odometer > 0 and vehicle.annual_mileage == 0:
            try:
                year = int(vehicle.vehicle_id.year)
                current_year = datetime.datetime.now().year
                years = max(1, current_year - year)
                vehicle.annual_mileage = vehicle.odometer / years
            except (ValueError, ZeroDivisionError):
                pass
    
    def _map_additional_columns(self, data: Dict[str, str]) -> Dict[str, str]:
        """
        Map CSV column names to standardized field names using comprehensive mappings.
        
        Args:
            data: Raw additional data from CSV with original column names
            
        Returns:
            Dictionary with standardized field names and values
        """
        mapped_data = {}
        
        # Create reverse mapping from input column names to standard names
        column_mapping = {}
        for standard_field, variants in ADDITIONAL_DATA_MAPPINGS.items():
            for variant in variants:
                column_mapping[variant.lower()] = standard_field
        
        # Map each input column to a standard field
        for original_column, value in data.items():
            original_lower = original_column.lower().strip()
            
            # Check if this column maps to a known standard field
            if original_lower in column_mapping:
                standard_field = column_mapping[original_lower]
                mapped_data[standard_field] = value
            else:
                # Keep original column name for unmapped fields
                mapped_data[original_column] = value
        
        return mapped_data
    
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