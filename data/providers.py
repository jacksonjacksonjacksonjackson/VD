"""
providers.py

Data retrieval services and API clients for the Fleet Electrification Analyzer.
Implements clients for VIN decoding and Fuel Economy data retrieval.
"""

import os
import time
import requests
import xml.etree.ElementTree as ET
import logging
from typing import Dict, List, Any, Optional, Union, Tuple
from urllib.parse import quote

from settings import (
    NHTSA_BASE_URL, 
    NHTSA_BATCH_URL,
    FUELECONOMY_BASE_URL, 
    FUELECONOMY_MENU_URL,
    API_TIMEOUT, 
    MAX_RETRIES, 
    RETRY_DELAY,
    RATE_LIMIT_DELAY
)
from utils import Cache, case_insensitive_equal, normalize_vehicle_model
from data.models import (
    VehicleIdentification, 
    FuelEconomyData,
    VinDecoderResponse,
    FuelEconomyResponse
)

# Set up module logger
logger = logging.getLogger(__name__)

###############################################################################
# API Clients
###############################################################################

class BaseApiClient:
    """Base class for API clients with common functionality."""
    
    def __init__(self, cache_enabled: bool = True):
        """
        Initialize the API client.
        
        Args:
            cache_enabled: Whether to use caching
        """
        self.cache_enabled = cache_enabled
        self.cache = Cache() if cache_enabled else None
        self.session = requests.Session()
    
    def _get_from_cache(self, cache_key: str) -> Optional[Any]:
        """Get data from cache if enabled and exists."""
        if not self.cache_enabled or not self.cache:
            return None
        return self.cache.get(cache_key)
    
    def _save_to_cache(self, cache_key: str, data: Any) -> None:
        """Save data to cache if enabled."""
        if not self.cache_enabled or not self.cache:
            return
        self.cache.set(cache_key, data)
    
    def _make_request(self, url: str, params: Optional[Dict[str, Any]] = None, 
                     method: str = "GET", data: Optional[Any] = None,
                     headers: Optional[Dict[str, str]] = None) -> Tuple[bool, Any, str]:
        """
        Make an HTTP request with retries and error handling.
        
        Args:
            url: URL to request
            params: Query parameters
            method: HTTP method
            data: Request body data
            headers: Request headers
            
        Returns:
            Tuple of (success, response_data, error_message)
        """
        # Apply rate limiting
        time.sleep(RATE_LIMIT_DELAY)
        
        # Initialize retry counter
        retries = 0
        last_error = ""
        
        while retries <= MAX_RETRIES:
            try:
                response = self.session.request(
                    method=method,
                    url=url,
                    params=params,
                    data=data,
                    headers=headers,
                    timeout=API_TIMEOUT
                )
                
                # Raise an exception for 4XX or 5XX responses
                response.raise_for_status()
                
                # Return the successful response
                return True, response, ""
                
            except requests.exceptions.HTTPError as e:
                last_error = f"HTTP Error: {e}"
                logger.warning(f"Request failed: {last_error}")
                
            except requests.exceptions.ConnectionError as e:
                last_error = f"Connection Error: {e}"
                logger.warning(f"Request failed: {last_error}")
                
            except requests.exceptions.Timeout as e:
                last_error = f"Timeout Error: {e}"
                logger.warning(f"Request timed out: {last_error}")
                
            except requests.exceptions.RequestException as e:
                last_error = f"Request Error: {e}"
                logger.warning(f"Request failed: {last_error}")
            
            # Increment retry counter and wait before retrying
            retries += 1
            if retries <= MAX_RETRIES:
                wait_time = RETRY_DELAY * retries
                logger.info(f"Retrying in {wait_time} seconds (attempt {retries}/{MAX_RETRIES})")
                time.sleep(wait_time)
        
        # All retries failed
        return False, None, last_error


class VinDecoderClient(BaseApiClient):
    """Client for decoding VINs using the NHTSA API."""
    
    def decode_vin(self, vin: str) -> VinDecoderResponse:
        """
        Decode a single VIN using the NHTSA API.
        
        Args:
            vin: Vehicle Identification Number to decode
            
        Returns:
            VinDecoderResponse object with results or error
        """
        if not vin:
            return VinDecoderResponse(
                success=False,
                error_message="VIN cannot be empty",
                vin=vin
            )
        
        # Check cache first
        cache_key = f"vin_{vin}"
        cached_data = self._get_from_cache(cache_key)
        if cached_data:
            logger.debug(f"Using cached data for VIN {vin}")
            return cached_data
        
        # Build URL
        url = f"{NHTSA_BASE_URL}{quote(vin)}?format=json"
        
        # Make request
        success, response, error = self._make_request(url)
        
        if not success:
            return VinDecoderResponse(
                success=False,
                error_message=error,
                vin=vin
            )
        
        # Parse response
        try:
            data = response.json()
            results = data.get("Results", [])
            
            if not results:
                return VinDecoderResponse(
                    success=False,
                    error_message="No results returned from API",
                    vin=vin
                )
            
            # NHTSA API returns a list, but we only need the first item
            result = results[0]
            
            # Check if we got essential data
            year = result.get("ModelYear", "").strip()
            make = result.get("Make", "").strip()
            model = result.get("Model", "").strip()
            
            if not (year and make and model):
                return VinDecoderResponse(
                    success=False,
                    error_message="Incomplete vehicle data returned",
                    vin=vin,
                    data=result
                )
            
            # Success
            response_obj = VinDecoderResponse(
                success=True,
                vin=vin,
                data=result
            )
            
            # Cache the result
            self._save_to_cache(cache_key, response_obj)
            
            return response_obj
            
        except Exception as e:
            logger.error(f"Error parsing VIN decoder response: {e}")
            return VinDecoderResponse(
                success=False,
                error_message=f"Error parsing response: {e}",
                vin=vin
            )
    
    def decode_batch(self, vins: List[str]) -> Dict[str, VinDecoderResponse]:
        """
        Decode multiple VINs in a batch.
        
        Args:
            vins: List of VINs to decode
            
        Returns:
            Dictionary mapping VINs to their respective VinDecoderResponse objects
        """
        # Initialize results dictionary
        results = {}
        
        # First, check cache for each VIN and collect those not in cache
        uncached_vins = []
        for vin in vins:
            cache_key = f"vin_{vin}"
            cached_data = self._get_from_cache(cache_key)
            
            if cached_data:
                results[vin] = cached_data
            else:
                uncached_vins.append(vin)
        
        # If all VINs were in cache, return early
        if not uncached_vins:
            return results
        
        # Batch size for API requests (NHTSA supports up to 50)
        BATCH_SIZE = 50
        
        # Process uncached VINs in batches
        for i in range(0, len(uncached_vins), BATCH_SIZE):
            batch = uncached_vins[i:i+BATCH_SIZE]
            
            # For small batches (< 5), use individual requests for better reliability
            if len(batch) < 5:
                for vin in batch:
                    results[vin] = self.decode_vin(vin)
                continue
            
            # For larger batches, use the batch API
            try:
                # Format batch data
                batch_data = ";".join(batch)
                headers = {"Content-Type": "application/x-www-form-urlencoded"}
                data = f"DATA={quote(batch_data)}&format=json"
                
                # Make batch request
                success, response, error = self._make_request(
                    url=NHTSA_BATCH_URL,
                    method="POST",
                    data=data,
                    headers=headers
                )
                
                if not success:
                    # Fallback to individual requests
                    logger.warning(f"Batch request failed: {error}. Falling back to individual requests.")
                    for vin in batch:
                        results[vin] = self.decode_vin(vin)
                    continue
                
                # Parse batch response
                data = response.json()
                batch_results = data.get("Results", [])
                
                # Process each result
                for result in batch_results:
                    result_vin = result.get("VIN", "").strip()
                    if not result_vin or result_vin not in batch:
                        continue
                    
                    # Check for essential data
                    year = result.get("ModelYear", "").strip()
                    make = result.get("Make", "").strip()
                    model = result.get("Model", "").strip()
                    
                    if year and make and model:
                        response_obj = VinDecoderResponse(
                            success=True,
                            vin=result_vin,
                            data=result
                        )
                    else:
                        response_obj = VinDecoderResponse(
                            success=False,
                            error_message="Incomplete vehicle data returned",
                            vin=result_vin,
                            data=result
                        )
                    
                    # Cache and store the result
                    cache_key = f"vin_{result_vin}"
                    self._save_to_cache(cache_key, response_obj)
                    results[result_vin] = response_obj
                
                # Handle any missing VINs in the response
                for vin in batch:
                    if vin not in results:
                        results[vin] = VinDecoderResponse(
                            success=False,
                            error_message="VIN not found in batch response",
                            vin=vin
                        )
                
            except Exception as e:
                # Fallback to individual requests on batch failure
                logger.error(f"Error in batch VIN decoding: {e}")
                for vin in batch:
                    results[vin] = self.decode_vin(vin)
        
        return results


class FuelEconomyClient(BaseApiClient):
    """Client for retrieving fuel economy data from FuelEconomy.gov API."""
    
    def fetch_menu(self, endpoint: str, params: Optional[Dict[str, str]] = None) -> List[Dict[str, str]]:
        """
        Fetch menu items from the FuelEconomy.gov menu API.
        
        Args:
            endpoint: API endpoint (e.g., 'make', 'model', 'options')
            params: Query parameters
            
        Returns:
            List of menu items (dictionaries with 'value' and 'text' keys)
        """
        if not params:
            params = {}
        
        # Build cache key
        param_str = ":".join(f"{k}={v}" for k, v in sorted(params.items()))
        cache_key = f"menu_{endpoint}_{param_str}"
        
        # Check cache
        cached_data = self._get_from_cache(cache_key)
        if cached_data:
            return cached_data
        
        # Build URL
        url = f"{FUELECONOMY_MENU_URL}/{endpoint}"
        
        # Make request
        success, response, error = self._make_request(url, params=params)
        
        if not success:
            logger.warning(f"Error fetching menu {endpoint}: {error}")
            return []
        
        # Parse XML response
        try:
            root = ET.fromstring(response.content)
            items = []
            
            for item in root.findall("menuItem"):
                val = item.findtext("value") or ""
                txt = item.findtext("text") or ""
                
                if val and txt:
                    items.append({"value": val, "text": txt})
            
            # Cache the result
            self._save_to_cache(cache_key, items)
            
            return items
            
        except Exception as e:
            logger.error(f"Error parsing menu response: {e}")
            return []
    
    def fetch_vehicle_details(self, vehicle_id: str) -> FuelEconomyResponse:
        """
        Fetch detailed fuel economy data for a specific vehicle.
        
        Args:
            vehicle_id: FuelEconomy.gov vehicle ID
            
        Returns:
            FuelEconomyResponse object with results or error
        """
        if not vehicle_id:
            return FuelEconomyResponse(
                success=False,
                error_message="Vehicle ID cannot be empty",
                vehicle_id=vehicle_id
            )
        
        # Check cache
        cache_key = f"vehicle_{vehicle_id}"
        cached_data = self._get_from_cache(cache_key)
        if cached_data:
            return cached_data
        
        # Build URL
        url = f"{FUELECONOMY_BASE_URL}/{vehicle_id}"
        
        # Make request
        success, response, error = self._make_request(url)
        
        if not success:
            return FuelEconomyResponse(
                success=False,
                error_message=error,
                vehicle_id=vehicle_id
            )
        
        # Parse XML response
        try:
            root = ET.fromstring(response.content)
            vehicle_data = {}
            
            for child in root:
                vehicle_data[child.tag] = child.text
            
            # Create response object
            response_obj = FuelEconomyResponse(
                success=True,
                vehicle_id=vehicle_id,
                data=vehicle_data
            )
            
            # Cache the result
            self._save_to_cache(cache_key, response_obj)
            
            return response_obj
            
        except Exception as e:
            logger.error(f"Error parsing vehicle details: {e}")
            return FuelEconomyResponse(
                success=False,
                error_message=f"Error parsing response: {e}",
                vehicle_id=vehicle_id
            )
    
    def find_vehicle_matches(self, year: str, make: str, model: str, 
                          engine_disp: str = None) -> List[Dict[str, str]]:
        """
        Find matching vehicles in the FuelEconomy.gov database.
        
        Args:
            year: Vehicle model year
            make: Vehicle manufacturer
            model: Vehicle model
            engine_disp: Engine displacement (optional)
            
        Returns:
            List of matching vehicle options
        """
        if not year or not make or not model:
            logger.warning("Missing required vehicle information")
            return []
        
        # Find matching make
        makes = self.fetch_menu("make", params={"year": year})
        matched_make = None
        
        for m in makes:
            if case_insensitive_equal(m["text"], make):
                matched_make = m["text"]
                break
        
        if not matched_make:
            logger.warning(f"Make '{make}' not found for year {year}")
            return []
        
        # Find matching models
        models_list = self.fetch_menu("model", params={"year": year, "make": matched_make})
        if not models_list:
            logger.warning(f"No models found for {matched_make} {year}")
            return []
        
        # Normalize the search model name
        normalized_model = normalize_vehicle_model(model)
        
        # Find model matches (exact or partial)
        exact_matches = [m for m in models_list 
                        if normalize_vehicle_model(m["text"]) == normalized_model]
        
        if exact_matches:
            matched_models = exact_matches
        else:
            matched_models = [m for m in models_list 
                            if normalized_model in normalize_vehicle_model(m["text"])]
        
        if not matched_models:
            logger.warning(f"Model '{model}' not found for {matched_make} {year}")
            return []
        
        # Collect all options for matching models
        all_options = []
        
        for mm in matched_models:
            options = self.fetch_menu("options", params={
                "year": year, 
                "make": matched_make, 
                "model": mm["text"]
            })
            
            if not options:
                continue
            
            # Filter options by model and engine displacement
            filtered_options = self._filter_options(options, model, engine_disp)
            all_options.extend(filtered_options)
        
        # Remove duplicates
        seen_values = set()
        unique_options = []
        
        for opt in all_options:
            if opt["value"] not in seen_values:
                unique_options.append(opt)
                seen_values.add(opt["value"])
        
        return unique_options
    
    def _filter_options(self, options: List[Dict[str, str]], model: str, 
                       engine_disp: Optional[str] = None) -> List[Dict[str, str]]:
        """
        Filter vehicle options by model and engine displacement.
        
        Args:
            options: List of options to filter
            model: Model to match
            engine_disp: Engine displacement to match (optional)
            
        Returns:
            Filtered list of options
        """
        # Normalize the model name for comparison
        normalized_model = normalize_vehicle_model(model)
        
        # Filter by model first
        model_filtered = [o for o in options 
                         if normalized_model in normalize_vehicle_model(o["text"])]
        
        # If no matches, return original options
        if not model_filtered:
            model_filtered = options
        
        # Filter by engine displacement if provided
        if engine_disp:
            disp_str = engine_disp.lower()
            disp_filtered = [o for o in model_filtered if disp_str in o["text"].lower()]
            
            # Return displacement-filtered results if any, otherwise keep model filtered
            if disp_filtered:
                return disp_filtered
        
        return model_filtered
    
    def pick_best_match(self, options: List[Dict[str, str]], year: str, make: str, 
                      model: str, engine_disp: Optional[str] = None) -> Optional[Dict[str, str]]:
        """
        Pick the best matching option from a list of vehicle options.
        
        Args:
            options: List of options to choose from
            year: Vehicle year
            make: Vehicle make
            model: Vehicle model
            engine_disp: Engine displacement (optional)
            
        Returns:
            Best matching option or None if no options
        """
        if not options:
            return None
        
        # Convert to lowercase for comparison
        lower_year = str(year).lower()
        lower_model = model.lower()
        lower_disp = engine_disp.lower() if engine_disp else ""
        
        # Score each option
        best_score = -1
        best_option = None
        
        for opt in options:
            text = opt["text"].lower()
            score = 0
            
            # Basic matching criteria
            if lower_model in text:
                score += 10
            if lower_disp and lower_disp in text:
                score += 5
            if lower_year in text:
                score += 3
            
            # Prefer simpler descriptions
            if len(text) < 50:
                score += 1
            
            # Prefer options that mention trim levels
            if "trim" in text or "level" in text:
                score += 2
            
            # Common trim indicators
            for trim in ["base", "standard", "le", "se", "xle", "limited", "sport", "touring"]:
                if f" {trim}" in text or f"{trim} " in text:
                    score += 2
            
            # Update best match if score is higher
            if score > best_score:
                best_score = score
                best_option = opt
        
        return best_option

###############################################################################
# Vehicle Data Provider
###############################################################################

class VehicleDataProvider:
    """
    Combines VIN decoding and fuel economy data retrieval.
    Provides a unified interface for getting comprehensive vehicle data.
    """
    
    def __init__(self, cache_enabled: bool = True):
        """
        Initialize the vehicle data provider.
        
        Args:
            cache_enabled: Whether to use caching
        """
        self.vin_client = VinDecoderClient(cache_enabled=cache_enabled)
        self.fe_client = FuelEconomyClient(cache_enabled=cache_enabled)
    
    def get_vehicle_by_vin(self, vin: str) -> Tuple[bool, Dict[str, Any], str]:
        """
        Get comprehensive vehicle data by VIN.
        
        Args:
            vin: Vehicle Identification Number
            
        Returns:
            Tuple of (success, data, error_message)
        """
        # Step 1: Decode VIN
        vin_response = self.vin_client.decode_vin(vin)
        
        if not vin_response.success:
            return False, {}, vin_response.error_message
        
        vehicle_id = vin_response.to_vehicle_id()
        
        # Step 2: Find matching vehicles in FuelEconomy.gov
        options = self.fe_client.find_vehicle_matches(
            year=vehicle_id.year,
            make=vehicle_id.make,
            model=vehicle_id.model,
            engine_disp=vehicle_id.engine_displacement
        )
        
        # No matches found
        if not options:
            # Return basic data from VIN decoder without fuel economy
            return True, {
                "vin": vin,
                "vehicle_id": vehicle_id.to_dict(),
                "fuel_economy": FuelEconomyData().to_dict(),
                "match_confidence": 0.0,
                "assumed_vehicle_id": "",
                "assumed_vehicle_text": ""
            }, "No matching vehicles found in fuel economy database"
        
        # Step 3: Pick best match
        best_match = self.fe_client.pick_best_match(
            options=options,
            year=vehicle_id.year,
            make=vehicle_id.make,
            model=vehicle_id.model,
            engine_disp=vehicle_id.engine_displacement
        )
        
        # Use first option if no best match found
        if not best_match:
            best_match = options[0]
        
        # Step 4: Get detailed fuel economy data
        fe_response = self.fe_client.fetch_vehicle_details(best_match["value"])
        
        if not fe_response.success:
            # Return basic data without fuel economy
            return True, {
                "vin": vin,
                "vehicle_id": vehicle_id.to_dict(),
                "fuel_economy": FuelEconomyData().to_dict(),
                "match_confidence": 30.0,  # Low confidence
                "assumed_vehicle_id": best_match["value"],
                "assumed_vehicle_text": best_match["text"]
            }, f"Error retrieving fuel economy data: {fe_response.error_message}"
        
        # Step 5: Combine all data
        fuel_economy = fe_response.to_fuel_economy()
        
        # Calculate match confidence
        match_confidence = self._calculate_match_confidence(
            vehicle_id, fuel_economy.raw_data, best_match["text"]
        )
        
        # Return complete data
        return True, {
            "vin": vin,
            "vehicle_id": vehicle_id.to_dict(),
            "fuel_economy": fuel_economy.to_dict(),
            "match_confidence": match_confidence,
            "assumed_vehicle_id": best_match["value"],
            "assumed_vehicle_text": best_match["text"]
        }, ""
    
    def get_vehicles_by_vins(self, vins: List[str]) -> Dict[str, Dict[str, Any]]:
        """
        Get vehicle data for multiple VINs.
        
        Args:
            vins: List of VINs to process
            
        Returns:
            Dictionary mapping VINs to their respective data
        """
        results = {}
        
        # Step 1: Decode all VINs in batch
        vin_responses = self.vin_client.decode_batch(vins)
        
        # Process each VIN individually
        for vin in vins:
            vin_response = vin_responses.get(vin)
            
            if not vin_response or not vin_response.success:
                error_msg = vin_response.error_message if vin_response else "VIN decoding failed"
                results[vin] = {
                    "success": False,
                    "error": error_msg,
                    "data": {}
                }
                continue
            
            # Get complete vehicle data
            success, data, error = self.get_vehicle_by_vin(vin)
            
            results[vin] = {
                "success": success,
                "error": error,
                "data": data
            }
        
        return results
    
    def _calculate_match_confidence(self, vehicle_id: VehicleIdentification, 
                                   fe_data: Dict[str, Any], match_text: str) -> float:
        """
        Calculate confidence score for the match between VIN data and fuel economy data.
        
        Args:
            vehicle_id: Vehicle identification from VIN
            fe_data: Fuel economy data
            match_text: Text description of the matched vehicle
            
        Returns:
            Confidence score (0-100)
        """
        score = 50.0  # Start with moderate confidence
        
        # Basic year, make, model match
        if vehicle_id.year == fe_data.get("year", ""):
            score += 15.0
        
        if case_insensitive_equal(vehicle_id.make, fe_data.get("make", "")):
            score += 15.0
        
        # Model might be partial match
        if vehicle_id.model.lower() in fe_data.get("model", "").lower():
            score += 10.0
        
        # Engine displacement match
        if vehicle_id.engine_displacement:
            fe_displ = fe_data.get("displ", "")
            try:
                vin_displ = float(vehicle_id.engine_displacement)
                fe_displ_val = float(fe_displ)
                
                if abs(vin_displ - fe_displ_val) < 0.1:
                    score += 5.0
            except (ValueError, TypeError):
                pass
        
        # Cylinders match
        if vehicle_id.engine_cylinders:
            fe_cyl = fe_data.get("cylinders", "")
            if vehicle_id.engine_cylinders == fe_cyl:
                score += 5.0
        
        # Cap the score at 100
        return min(100.0, score)