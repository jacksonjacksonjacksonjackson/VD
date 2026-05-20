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
import re

from settings import (
    NHTSA_BASE_URL,
    NHTSA_BATCH_URL,
    FUELECONOMY_BASE_URL,
    FUELECONOMY_MENU_URL,
    API_TIMEOUT,
    MAX_RETRIES,
    RETRY_DELAY,
    RATE_LIMIT_DELAY,
    DEFAULT_CACHE_FILE,
    MATCHING_WEIGHTS,
    MIN_MATCH_CONFIDENCE,
)
from utils import Cache, case_insensitive_equal, normalize_vehicle_model
from data.models import (
    VehicleIdentification,
    FuelEconomyData,
    VinDecoderResponse,
    FuelEconomyResponse
)

# Register dataclass types for cache serialization/deserialization
Cache.register_type(VinDecoderResponse)
Cache.register_type(FuelEconomyResponse)

# Set up module logger
logger = logging.getLogger(__name__)

###############################################################################
# API Clients
###############################################################################

class BaseApiClient:
    """Base class for API clients with common functionality."""

    def __init__(self, cache_enabled: bool = True, shared_cache: Optional[Cache] = None):
        """
        Initialize the API client.

        Args:
            cache_enabled: Whether to use caching
            shared_cache: Optional shared Cache instance (enables disk persistence
                          when multiple clients share the same cache)
        """
        self.cache_enabled = cache_enabled
        self.cache = shared_cache if shared_cache else (Cache() if cache_enabled else None)
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
            return cached_data

        # Build URL and make request
        url = f"{NHTSA_BASE_URL}{quote(vin)}?format=json"
        success, response, error = self._make_request(url)

        if not success:
            logger.warning(f"VIN decode API failed for {vin}: {error}")
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

            result = results[0]

            # Check if we got essential data
            year = result.get("ModelYear", "").strip()
            make = result.get("Make", "").strip()
            model = result.get("Model", "").strip()

            if not (year and make and model):
                logger.warning(f"Incomplete VIN data for {vin}: year='{year}', make='{make}', model='{model}'")
                return VinDecoderResponse(
                    success=False,
                    error_message="Incomplete vehicle data returned",
                    vin=vin,
                    data=result
                )

            # Success - cache and return
            vin_response = VinDecoderResponse(
                success=True,
                vin=vin,
                data=result
            )
            self._save_to_cache(cache_key, vin_response)
            return vin_response

        except Exception as e:
            logger.error(f"Error parsing VIN decoder response for {vin}: {e}")
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
                          engine_disp: str = None,
                          is_diesel: bool = False) -> List[Dict[str, str]]:
        """
        Find vehicle matches in the FuelEconomy.gov database.

        Args:
            year: Model year
            make: Manufacturer
            model: Model name
            engine_disp: Engine displacement (optional)
            is_diesel: Whether the VIN indicates a diesel engine

        Returns:
            List of matching vehicle options
        """
        # Step 1: Get available models for this year and make
        available_models = self.fetch_menu("model", {"year": year, "make": make})

        if not available_models:
            logger.debug(f"No models found for {year} {make}")
            return []

        # Step 2: Find the best matching model
        best_model_match = self._find_best_model_match(available_models, model)

        if not best_model_match:
            logger.debug(f"No model match for '{model}' in {year} {make}")
            return []

        # Step 3: Get vehicle options for the matched model
        options = self.fetch_menu("options", {
            "year": year,
            "make": make,
            "model": best_model_match["value"]
        })

        if not options:
            return []

        # Step 4: Filter by engine displacement if specified
        if engine_disp:
            filtered_options = self._filter_options_by_engine(options, engine_disp)
            if filtered_options:
                options = filtered_options
            # If no engine match, keep all options as fallback

        # Step 5: Filter by fuel type for diesel vehicles
        if is_diesel:
            diesel_options = self._filter_options_by_fuel_type(options, diesel=True)
            if diesel_options:
                return diesel_options
            # No diesel options found — fall back to all options.
            # Caller (get_vehicle_by_vin) will detect the mismatch.
            logger.info(f"No diesel options for {year} {make} {model}; falling back to gasoline")

        return options

    @staticmethod
    def _filter_options_by_fuel_type(options: List[Dict[str, str]],
                                     diesel: bool = True) -> List[Dict[str, str]]:
        """Filter vehicle options by fuel type keywords in the option text."""
        diesel_keywords = {"diesel", "biodiesel", "b20"}
        filtered = []
        for opt in options:
            text_lower = opt.get("text", "").lower()
            has_diesel_kw = any(kw in text_lower for kw in diesel_keywords)
            if diesel and has_diesel_kw:
                filtered.append(opt)
            elif not diesel and not has_diesel_kw:
                filtered.append(opt)
        return filtered
    

    
    def _find_best_model_match(self, available_models: List[Dict[str, str]], target_model: str) -> Optional[Dict[str, str]]:
        """
        Find the best matching model from available models.
        
        Args:
            available_models: List of available models from API
            target_model: Target model name to match
            
        Returns:
            Best matching model or None if no match found
        """
        if not available_models or not target_model:
            return None
        
        target_normalized = normalize_vehicle_model(target_model).lower()

        # Try exact match first
        for model in available_models:
            model_value = model.get("value", "").lower()
            model_text = model.get("text", "").lower()

            if target_normalized == model_value or target_normalized == model_text:
                return model
        
        # Try partial match
        best_match = None
        best_score = 0
        
        for model in available_models:
            model_value = model.get("value", "").lower()
            model_text = model.get("text", "").lower()
            
            score = 0
            
            # Check if target is in model name
            if target_normalized in model_value:
                score += 10
            if target_normalized in model_text:
                score += 8
            
            # Check if model name is in target (for cases like "F-150" vs "F150")
            if model_value in target_normalized:
                score += 6
            if model_text in target_normalized:
                score += 4
            
            # Prefer shorter matches (less specific variants)
            if score > 0:
                score -= len(model_text) * 0.1
            
            if score > best_score:
                best_score = score
                best_match = model
        
        return best_match
    
    def _filter_options_by_engine(self, options: List[Dict[str, str]], engine_disp: str) -> List[Dict[str, str]]:
        """
        Filter vehicle options by engine displacement.
        
        Args:
            options: List of vehicle options
            engine_disp: Engine displacement to match
            
        Returns:
            Filtered list of options
        """
        if not engine_disp:
            return options
        
        filtered = []
        engine_normalized = engine_disp.lower().replace(" ", "")

        for option in options:
            option_text = option.get("text", "").lower()

            # Check for direct match
            if engine_disp.lower() in option_text:
                filtered.append(option)
                continue

            # Check for normalized match (e.g., "2.0L" vs "2.0 L")
            if engine_normalized in option_text.replace(" ", ""):
                filtered.append(option)
                continue

            # Check for displacement patterns (e.g., "2.0" in "2.0 L, Turbo")
            import re
            engine_pattern = re.escape(engine_disp.replace("L", "").replace("l", "").strip())
            if re.search(rf"{engine_pattern}\s*l", option_text, re.IGNORECASE):
                filtered.append(option)

        return filtered
    
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
        # Create a single persistent cache shared by both API clients.
        # On init it loads previously cached API responses from disk so
        # re-processing the same fleet is nearly instant.
        if cache_enabled:
            self._shared_cache = Cache(file_path=DEFAULT_CACHE_FILE)
        else:
            self._shared_cache = None

        self.vin_client = VinDecoderClient(
            cache_enabled=cache_enabled, shared_cache=self._shared_cache)
        self.fe_client = FuelEconomyClient(
            cache_enabled=cache_enabled, shared_cache=self._shared_cache)

    def save_cache(self) -> bool:
        """Persist the shared API cache to disk. Call after processing completes."""
        if self._shared_cache:
            return self._shared_cache.save_to_disk()
        return False
    
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
            engine_disp=vehicle_id.engine_displacement,
            is_diesel=vehicle_id.is_diesel
        )

        # No matches found — return basic data without fuel economy
        if not options:
            return True, {
                "vin": vin,
                "vehicle_id": vehicle_id.to_dict(),
                "fuel_economy": FuelEconomyData().to_dict(),
                "match_confidence": 0.0,
                "assumed_vehicle_id": "",
                "assumed_vehicle_text": "",
                "fuel_type_mismatch": False
            }, "No matching vehicles found in fuel economy database"

        # Step 3: Pick best match
        best_match = self.fe_client.pick_best_match(
            options=options,
            year=vehicle_id.year,
            make=vehicle_id.make,
            model=vehicle_id.model,
            engine_disp=vehicle_id.engine_displacement
        )

        if not best_match:
            best_match = options[0]

        # Step 4: Detect fuel type mismatch (diesel VIN matched to gas data)
        fuel_type_mismatch = False
        if vehicle_id.is_diesel:
            match_text_lower = best_match.get("text", "").lower()
            diesel_kws = {"diesel", "biodiesel", "b20"}
            if not any(kw in match_text_lower for kw in diesel_kws):
                fuel_type_mismatch = True
                logger.warning(
                    f"Diesel vehicle {vin} matched to gasoline data: {best_match.get('text', '')}"
                )

        # Step 5: Get detailed fuel economy data
        fe_response = self.fe_client.fetch_vehicle_details(best_match["value"])

        if not fe_response.success:
            return True, {
                "vin": vin,
                "vehicle_id": vehicle_id.to_dict(),
                "fuel_economy": FuelEconomyData().to_dict(),
                "match_confidence": 30.0,
                "assumed_vehicle_id": best_match["value"],
                "assumed_vehicle_text": best_match["text"],
                "fuel_type_mismatch": fuel_type_mismatch
            }, f"Error retrieving fuel economy data: {fe_response.error_message}"

        # Step 6: Combine all data
        fuel_economy = fe_response.to_fuel_economy()
        match_confidence = self._calculate_match_confidence(
            vehicle_id, fuel_economy.raw_data, best_match["text"]
        )

        # Penalize confidence for fuel type mismatch
        if fuel_type_mismatch:
            match_confidence = max(0.0, match_confidence - 15.0)

        # Warn when confidence falls below the configured threshold
        if match_confidence < MIN_MATCH_CONFIDENCE:
            logger.warning(
                f"Low-confidence match for {vin} ({match_confidence:.0f}% < "
                f"{MIN_MATCH_CONFIDENCE}% threshold): {best_match.get('text', '')}"
            )

        return True, {
            "vin": vin,
            "vehicle_id": vehicle_id.to_dict(),
            "fuel_economy": fuel_economy.to_dict(),
            "match_confidence": match_confidence,
            "assumed_vehicle_id": best_match["value"],
            "assumed_vehicle_text": best_match["text"],
            "fuel_type_mismatch": fuel_type_mismatch
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

        Scoring uses MATCHING_WEIGHTS from settings.py so thresholds are tunable
        without code changes.

        Returns:
            Confidence score (0-100)
        """
        w = MATCHING_WEIGHTS
        score = 0.0

        # --- Year / make / model ---
        year_ok = vehicle_id.year == fe_data.get("year", "")
        make_ok = case_insensitive_equal(vehicle_id.make, fe_data.get("make", ""))
        model_ok = bool(vehicle_id.model and
                        vehicle_id.model.lower() in fe_data.get("model", "").lower())

        if year_ok and make_ok and model_ok:
            score += w.get("year_make_model", 80)
        elif year_ok and make_ok:
            score += w.get("year_make", 60)
        elif make_ok and model_ok:
            score += w.get("make_model", 50)

        # --- Engine displacement ---
        if vehicle_id.engine_displacement:
            fe_displ = fe_data.get("displ", "")
            try:
                vin_m = re.search(r'(\d+\.?\d*)', str(vehicle_id.engine_displacement).strip())
                fe_m = re.search(r'(\d+\.?\d*)', str(fe_displ).strip())
                if vin_m and fe_m and abs(float(vin_m.group(1)) - float(fe_m.group(1))) < 0.1:
                    score += w.get("displacement_match", 15)
            except (ValueError, TypeError, AttributeError) as exc:
                logger.debug(
                    f"Could not compare engine displacements: "
                    f"VIN '{vehicle_id.engine_displacement}' vs FE '{fe_displ}': {exc}"
                )

        # --- Engine match bonus (displacement + cylinders both hit) ---
        cyl_ok = bool(vehicle_id.engine_cylinders and
                      vehicle_id.engine_cylinders == str(fe_data.get("cylinders", "")))
        if cyl_ok:
            score += w.get("cylinders_match", 10)

        # Award engine_match bonus only when both displacement and cylinders agree
        if vehicle_id.engine_displacement and cyl_ok:
            score += w.get("engine_match", 20)

        # --- Fuel type ---
        if vehicle_id.fuel_type:
            fe_fuel = fe_data.get("fuelType1", "") or fe_data.get("fuelType", "")
            if fe_fuel and case_insensitive_equal(vehicle_id.fuel_type, fe_fuel):
                score += w.get("fuel_type_match", 10)

        # --- Drive type ---
        if vehicle_id.drive_type:
            fe_drive = fe_data.get("drive", "")
            if fe_drive and case_insensitive_equal(vehicle_id.drive_type, fe_drive):
                score += w.get("drive_match", 5)

        # --- Transmission ---
        if vehicle_id.transmission:
            fe_tranny = fe_data.get("trany", "")
            if fe_tranny and vehicle_id.transmission.lower() in fe_tranny.lower():
                score += w.get("transmission_match", 5)

        return min(100.0, score)