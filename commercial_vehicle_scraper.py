"""
commercial_vehicle_scraper.py

Intelligent Commercial Vehicle Web Scraping Engine for Fleet Electrification Analyzer.
Enhances data coverage for Class 3-8 commercial vehicles through intelligent web scraping.
"""

import os
import re
import time
import json
import hashlib
import logging
import requests
import pandas as pd
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Tuple, Union
from urllib.parse import urlparse, urljoin, quote
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
import warnings

# Web scraping libraries
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, WebDriverException

# PDF processing
try:
    import pdfplumber
except ImportError:
    pdfplumber = None
    warnings.warn("pdfplumber not installed - PDF extraction will be disabled")

# Import existing application modules
from data.providers import VehicleDataProvider, BaseApiClient
from data.models import VehicleIdentification, FuelEconomyData, FleetVehicle
from utils import Cache, safe_cast, normalize_vehicle_model
from settings import CACHE_DIR, MAX_THREADS, SCRAPING_CONFIG

# Configure logging
logger = logging.getLogger("commercial_scraper")
logger.setLevel(logging.INFO)

# Use scraping configuration from settings.py
# Local fallback config if import fails
LOCAL_SCRAPING_CONFIG = {
    "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "request_timeout": 10,
    "selenium_timeout": 15,
    "rate_limit_delay": 1.0,  # seconds between requests to same domain
    "max_retries": 3,
    "cache_expiry_hours": 168,  # 1 week for scraped data
    "confidence_threshold": 0.7,  # Minimum confidence for estimated values
}

###############################################################################
# Data Models for Commercial Vehicles
###############################################################################

@dataclass
class CommercialVehicleSpecs:
    """Extended specifications for commercial vehicles."""
    
    # Payload and capacity
    payload_capacity_lbs: Optional[float] = None
    towing_capacity_lbs: Optional[float] = None
    gcwr_lbs: Optional[float] = None  # Gross Combined Weight Rating
    front_gawr_lbs: Optional[float] = None  # Front Gross Axle Weight Rating
    rear_gawr_lbs: Optional[float] = None  # Rear Gross Axle Weight Rating
    
    # Engine specifications
    engine_torque_lb_ft: Optional[float] = None
    engine_torque_rpm: Optional[int] = None
    max_hp_rpm: Optional[int] = None
    engine_manufacturer: str = ""
    engine_model: str = ""
    
    # Fuel system
    fuel_tank_capacity_gal: Optional[float] = None
    def_tank_capacity_gal: Optional[float] = None  # Diesel Exhaust Fluid
    
    # Dimensions
    wheelbase_inches: Optional[float] = None
    overall_length_inches: Optional[float] = None
    overall_width_inches: Optional[float] = None
    overall_height_inches: Optional[float] = None
    cargo_length_inches: Optional[float] = None
    cargo_width_inches: Optional[float] = None
    cargo_height_inches: Optional[float] = None
    
    # Configuration
    cab_configuration: str = ""  # Regular, Extended, Crew, etc.
    bed_length: str = ""  # Short, Standard, Long
    axle_configuration: str = ""  # 4x2, 4x4, 6x4, etc.
    axle_ratio: str = ""
    
    # Operational classification
    duty_cycle: str = ""  # Urban Delivery, Long Haul, Construction, etc.
    vocation: str = ""  # Specific use case
    electrification_suitability: str = ""  # High, Medium, Low
    recommended_ev_alternatives: List[str] = field(default_factory=list)
    
    # Data quality metrics
    data_source: str = ""  # Where data was scraped from
    data_confidence: float = 0.0  # Confidence score (0-1)
    is_estimated: bool = False  # Whether values are estimated
    last_updated: datetime = field(default_factory=datetime.now)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization."""
        return {
            k: v.isoformat() if isinstance(v, datetime) else v
            for k, v in self.__dict__.items()
        }

@dataclass
class ScrapingResult:
    """Result from a scraping operation."""
    success: bool
    source_tier: int  # 1, 2, or 3
    source_name: str
    url: str
    data: Dict[str, Any] = field(default_factory=dict)
    specs: Optional[CommercialVehicleSpecs] = None
    error_message: str = ""
    confidence_score: float = 0.0
    extraction_time: float = 0.0  # seconds

###############################################################################
# Pattern Recognition and Data Extraction
###############################################################################

class SpecificationExtractor:
    """Intelligent pattern-based specification extraction."""
    
    # Common specification patterns
    SPEC_PATTERNS = {
        'gvwr': [
            r'GVWR[:\s]+([0-9,]+)\s*(?:lbs?|pounds?)',
            r'Gross Vehicle Weight Rating[:\s]+([0-9,]+)',
            r'Max\.\s+GVW[:\s]+([0-9,]+)',
        ],
        'payload': [
            r'(?:Max\.?\s+)?Payload[:\s]+([0-9,]+)\s*(?:lbs?|pounds?)',
            r'Payload Capacity[:\s]+([0-9,]+)',
            r'Maximum Payload[:\s]+([0-9,]+)',
        ],
        'towing': [
            r'(?:Max\.?\s+)?Towing[:\s]+([0-9,]+)\s*(?:lbs?|pounds?)',
            r'Towing Capacity[:\s]+([0-9,]+)',
            r'Maximum Trailer Weight[:\s]+([0-9,]+)',
        ],
        'torque': [
            r'(?:Peak\s+)?Torque[:\s]+([0-9,]+)\s*(?:lb[\.-]ft|ft[\.-]lbs?)',
            r'([0-9,]+)\s*(?:lb[\.-]ft|ft[\.-]lbs?)\s+(?:of\s+)?torque',
            r'Maximum Torque[:\s]+([0-9,]+)',
        ],
        'fuel_capacity': [
            r'Fuel (?:Tank )?Capacity[:\s]+([0-9\.]+)\s*(?:gal|gallons?)',
            r'([0-9\.]+)[\s\-](?:gal|gallon)\s+fuel tank',
            r'Fuel Tank[:\s]+([0-9\.]+)',
        ],
        'wheelbase': [
            r'Wheelbase[:\s]+([0-9\.]+)\s*(?:in|inches|")',
            r'([0-9\.]+)[\s\-](?:in|inch)\s+wheelbase',
            r'WB[:\s]+([0-9\.]+)',
        ],
        # MPG and fuel economy patterns
        'mpg_combined': [
            r'Combined MPG[:\s]*([0-9]+)\s*combined',
            r'([0-9]+)\s*combined city/highway',
            r'EPA.*Combined.*?([0-9]+)',
            r'Combined[:\s]+([0-9]+)',
            r'([0-9]+)\s*mpg combined',
        ],
        'mpg_city': [
            r'City MPG[:\s]*([0-9]+)\s*city',
            r'([0-9]+)\s*city',
            r'City[:\s]+([0-9]+)',
            r'([0-9]+)\s*mpg city',
        ],
        'mpg_highway': [
            r'Highway MPG[:\s]*([0-9]+)\s*highway',
            r'([0-9]+)\s*highway',
            r'Highway[:\s]+([0-9]+)',
            r'([0-9]+)\s*mpg highway',
        ],
    }
    
    def extract_from_html(self, html: str, soup: Optional[BeautifulSoup] = None) -> Dict[str, Any]:
        """Extract specifications from HTML content."""
        if not soup:
            soup = BeautifulSoup(html, 'html.parser')
        
        specs = {}
        
        # Method 1: Look for specification tables
        specs.update(self._extract_from_tables(soup))
        
        # Method 2: Look for specification lists
        specs.update(self._extract_from_lists(soup))
        
        # Method 3: Pattern matching in text
        specs.update(self._extract_from_text(soup.get_text()))
        
        # Method 4: Look for structured data (JSON-LD, microdata)
        specs.update(self._extract_structured_data(soup))
        
        return specs
    
    def _extract_from_tables(self, soup: BeautifulSoup) -> Dict[str, Any]:
        """Extract data from HTML tables."""
        specs = {}
        
        # Find all tables that might contain specifications
        tables = soup.find_all('table')
        
        for table in tables:
            # Check if table likely contains specs (heuristic)
            text = table.get_text().lower()
            if any(term in text for term in ['specification', 'specs', 'gvwr', 'payload', 'capacity']):
                rows = table.find_all('tr')
                
                for row in rows:
                    cells = row.find_all(['td', 'th'])
                    if len(cells) >= 2:
                        label = cells[0].get_text().strip()
                        value = cells[1].get_text().strip()
                        
                        # Try to match known specification types
                        normalized_specs = self._normalize_specification(label, value)
                        specs.update(normalized_specs)
        
        return specs
    
    def _extract_from_lists(self, soup: BeautifulSoup) -> Dict[str, Any]:
        """Extract data from HTML lists (ul, dl)."""
        specs = {}
        
        # Definition lists often contain specifications
        dl_elements = soup.find_all('dl')
        for dl in dl_elements:
            dt_elements = dl.find_all('dt')
            dd_elements = dl.find_all('dd')
            
            for dt, dd in zip(dt_elements, dd_elements):
                label = dt.get_text().strip()
                value = dd.get_text().strip()
                normalized_specs = self._normalize_specification(label, value)
                specs.update(normalized_specs)
        
        # Also check unordered lists with specific patterns
        ul_elements = soup.find_all('ul', class_=re.compile(r'spec|feature|detail', re.I))
        for ul in ul_elements:
            for li in ul.find_all('li'):
                text = li.get_text().strip()
                # Look for "Label: Value" patterns
                if ':' in text:
                    parts = text.split(':', 1)
                    if len(parts) == 2:
                        label, value = parts
                        normalized_specs = self._normalize_specification(label.strip(), value.strip())
                        specs.update(normalized_specs)
        
        return specs
    
    def _extract_from_text(self, text: str) -> Dict[str, Any]:
        """Extract specifications using regex patterns."""
        specs = {}
        
        for spec_type, patterns in self.SPEC_PATTERNS.items():
            for pattern in patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    value = match.group(1).replace(',', '')
                    try:
                        numeric_value = float(value)
                        specs[spec_type] = numeric_value
                        break
                    except ValueError:
                        continue
        
        return specs
    
    def _extract_structured_data(self, soup: BeautifulSoup) -> Dict[str, Any]:
        """Extract JSON-LD or microdata structured data."""
        specs = {}
        
        # Look for JSON-LD scripts
        scripts = soup.find_all('script', type='application/ld+json')
        for script in scripts:
            try:
                data = json.loads(script.string)
                if isinstance(data, dict):
                    # Look for vehicle-related structured data
                    if data.get('@type') in ['Car', 'Vehicle', 'Product']:
                        specs.update(self._parse_structured_vehicle_data(data))
            except json.JSONDecodeError:
                continue
        
        return specs
    
    def _normalize_specification(self, label: str, value: str) -> Dict[str, Any]:
        """Normalize a specification label/value pair."""
        normalized = {}
        label_lower = label.lower()
        
        # Map common labels to standardized field names
        if 'gvwr' in label_lower or 'gross vehicle weight' in label_lower:
            try:
                numeric_value = float(re.sub(r'[^0-9.]', '', value))
                normalized['gvwr'] = numeric_value
            except ValueError:
                pass
        
        elif 'payload' in label_lower:
            try:
                numeric_value = float(re.sub(r'[^0-9.]', '', value))
                normalized['payload'] = numeric_value
            except ValueError:
                pass
        
        elif 'towing' in label_lower or 'trailer' in label_lower:
            try:
                numeric_value = float(re.sub(r'[^0-9.]', '', value))
                normalized['towing'] = numeric_value
            except ValueError:
                pass
        
        elif 'torque' in label_lower:
            try:
                numeric_value = float(re.sub(r'[^0-9.]', '', value))
                normalized['torque'] = numeric_value
            except ValueError:
                pass
        
        elif 'fuel' in label_lower and 'capacity' in label_lower:
            try:
                numeric_value = float(re.sub(r'[^0-9.]', '', value))
                normalized['fuel_capacity'] = numeric_value
            except ValueError:
                pass
        
        elif 'wheelbase' in label_lower:
            try:
                numeric_value = float(re.sub(r'[^0-9.]', '', value))
                normalized['wheelbase'] = numeric_value
            except ValueError:
                pass
        
        return normalized
    
    def _parse_structured_vehicle_data(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Parse vehicle data from structured data formats."""
        specs = {}
        
        # Common structured data fields
        if 'vehicleSpecification' in data:
            vehicle_spec = data['vehicleSpecification']
            if isinstance(vehicle_spec, dict):
                for key, value in vehicle_spec.items():
                    normalized = self._normalize_specification(key, str(value))
                    specs.update(normalized)
        
        # Additional properties that might contain specs
        for prop in ['additionalProperty', 'specification', 'feature']:
            if prop in data:
                prop_data = data[prop]
                if isinstance(prop_data, list):
                    for item in prop_data:
                        if isinstance(item, dict) and 'name' in item and 'value' in item:
                            normalized = self._normalize_specification(
                                item['name'], str(item['value'])
                            )
                            specs.update(normalized)
        
        return specs

###############################################################################
# Web Scraping Engine
###############################################################################

class CommercialVehicleScraper:
    """Main scraping engine for commercial vehicle data."""
    
    def __init__(self, cache_enabled: bool = True, use_selenium: bool = False):
        """
        Initialize the scraping engine.
        
        Args:
            cache_enabled: Whether to use caching
            use_selenium: Whether to use Selenium for dynamic content
        """
        self.cache_enabled = cache_enabled
        # Use global configuration with fallback
        self.config = SCRAPING_CONFIG if 'SCRAPING_CONFIG' in globals() else LOCAL_SCRAPING_CONFIG
        self.cache = Cache(expiry_seconds=self.config['cache_expiry_hours'] * 3600)
        self.use_selenium = use_selenium
        self.extractor = SpecificationExtractor()
        self.session = requests.Session()
        self.session.headers.update({'User-Agent': self.config['user_agent']})
        
        # Graceful degradation settings
        self.graceful_degradation = self.config.get('graceful_degradation', True)
        self.enable_estimates = self.config.get('enable_estimates', True)
        
        # Rate limiting tracking
        self.last_request_time = {}
        
        # Initialize Selenium driver if needed
        self.driver = None
        if use_selenium:
            self._init_selenium_driver()
        
        # Tier 1 Sources - Manufacturer websites
        self.tier1_sources = {
            'ford_commercial': {
                'base_url': 'https://www.ford.com/commercial-trucks/',
                'models': ['f-150', 'f-250', 'f-350', 'f-450', 'f-550', 'f-600', 'f-650', 'f-750',
                          'transit', 'transit-connect', 'e-transit'],
                'scraper': self._scrape_ford_commercial
            },
            'freightliner': {
                'base_url': 'https://www.freightliner.com/',
                'models': ['cascadia', 'cascadia-126', 'ecascadia', 'm2-106', 'm2-112'],
                'scraper': self._scrape_freightliner
            },
            'peterbilt': {
                'base_url': 'https://www.peterbilt.com/',
                'models': ['579', '567', '389', '367', '365', '348', '337', '220'],
                'scraper': self._scrape_peterbilt
            },
            'kenworth': {
                'base_url': 'https://www.kenworth.com/',
                'models': ['t680', 't880', 'w990', 't370', 't270', 'k270', 'k370'],
                'scraper': self._scrape_kenworth
            },
            'mack': {
                'base_url': 'https://www.macktrucks.com/',
                'models': ['anthem', 'pinnacle', 'granite', 'lr', 'md', 'terra-pro'],
                'scraper': self._scrape_mack
            },
            'volvo': {
                'base_url': 'https://www.volvotrucks.us/',
                'models': ['vnl', 'vnr', 'vnx', 'vhd'],
                'scraper': self._scrape_volvo
            },
            'international': {
                'base_url': 'https://www.internationaltrucks.com/',
                'models': ['lt', 'lonestar', 'rh', 'hx', 'hv', 'mv', 'cv', 'durastar'],
                'scraper': self._scrape_international
            },
            'isuzu': {
                'base_url': 'https://www.isuzucv.com/',
                'models': ['npr', 'npr-hd', 'nqr', 'nrr', 'ftr'],
                'scraper': self._scrape_isuzu
            }
        }
        
        # Tier 2 Sources - Government/Industry databases
        self.tier2_sources = {
            'epa_fueleconomy': {
                'base_url': 'https://fueleconomy.gov/',
                'scraper': self._scrape_epa_smartway  # Renamed but same method
            },
            'carb_clean_truck': {
                'base_url': 'https://ww2.arb.ca.gov/our-work/programs/clean-truck-check',
                'scraper': self._scrape_carb
            }
        }
        
        # Tier 3 Sources - Community data and dealer sites
        self.tier3_sources = {
            'fuelly_community': {
                'base_url': 'https://www.fuelly.com/',
                'scraper': self._scrape_fuelly
            },
            'commercial_truck_trader': {
                'base_url': 'https://www.commercialtrucktrader.com/',
                'scraper': self._scrape_truck_trader
            },
            'truck_paper': {
                'base_url': 'https://www.truckpaper.com/',
                'scraper': self._scrape_truck_paper
            }
        }
    
    def _init_selenium_driver(self):
        """Initialize Selenium WebDriver for dynamic content."""
        try:
            options = Options()
            options.add_argument('--headless')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument(f'user-agent={SCRAPING_CONFIG["user_agent"]}')
            
            self.driver = webdriver.Chrome(options=options)
            logger.info("Selenium WebDriver initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize Selenium: {e}")
            self.use_selenium = False
    
    def scrape_vehicle(self, vehicle_id: VehicleIdentification) -> ScrapingResult:
        """
        Main entry point for scraping vehicle data.
        
        Args:
            vehicle_id: Vehicle identification object
            
        Returns:
            ScrapingResult with extracted data
        """
        # Generate cache key
        cache_key = f"scraped_{vehicle_id.vin}_{vehicle_id.year}_{vehicle_id.make}_{vehicle_id.model}"
        
        # Check cache first
        if self.cache_enabled:
            cached_result = self.cache.get(cache_key)
            if cached_result:
                logger.info(f"Using cached scraping result for {vehicle_id.display_name}")
                return cached_result
        
        # Try tiered scraping
        result = None
        
        # Tier 1: Manufacturer websites
        result = self._try_tier1_sources(vehicle_id)
        if result and result.success:
            logger.info(f"Successfully scraped from Tier 1 source: {result.source_name}")
            if self.cache_enabled:
                self.cache.set(cache_key, result)
            return result
        
        # Tier 2: Government/Industry databases
        result = self._try_tier2_sources(vehicle_id)
        if result and result.success:
            # Check if the Tier 2 data is actually useful (not garbage)
            if self._has_useful_mpg_data(result):
                logger.info(f"Successfully scraped from Tier 2 source: {result.source_name}")
                if self.cache_enabled:
                        self.cache.set(cache_key, result)
                return result
            else:
                logger.warning(f"Tier 2 source {result.source_name} returned invalid MPG data, continuing to Tier 3")
        
        # Tier 3: Review sites and dealers
        result = self._try_tier3_sources(vehicle_id)
        if result and result.success:
            logger.info(f"Successfully scraped from Tier 3 source: {result.source_name}")
            if self.cache_enabled:
                self.cache.set(cache_key, result)
            return result
        
        # If all scraping failed, try pattern-based estimation
        logger.warning(f"All scraping tiers failed for {vehicle_id.display_name}, attempting estimation")
        estimated_result = self._estimate_specifications(vehicle_id)
        
        if self.cache_enabled and estimated_result:
            self.cache.set(cache_key, estimated_result)
        
        return estimated_result
    
    def _try_tier1_sources(self, vehicle_id: VehicleIdentification) -> Optional[ScrapingResult]:
        """Try scraping from Tier 1 manufacturer sources."""
        make_lower = vehicle_id.make.lower()
        
        # Map vehicle makes to scraper sources
        make_mapping = {
            'ford': 'ford_commercial',
            'freightliner': 'freightliner',
            'peterbilt': 'peterbilt',
            'kenworth': 'kenworth',
            'mack': 'mack',
            'volvo': 'volvo',
            'international': 'international',
            'navistar': 'international',
            'isuzu': 'isuzu'
        }
        
        source_key = make_mapping.get(make_lower)
        if source_key and source_key in self.tier1_sources:
            source = self.tier1_sources[source_key]
            try:
                start_time = time.time()
                result = source['scraper'](vehicle_id, source)
                result.extraction_time = time.time() - start_time
                result.source_tier = 1
                return result
            except Exception as e:
                logger.error(f"Error scraping {source_key}: {e}")
        
        return None
    
    def _try_tier2_sources(self, vehicle_id: VehicleIdentification) -> Optional[ScrapingResult]:
        """Try scraping from Tier 2 government/industry sources."""
        for source_name, source in self.tier2_sources.items():
            try:
                start_time = time.time()
                result = source['scraper'](vehicle_id, source)
                if result and result.success:
                    result.extraction_time = time.time() - start_time
                    result.source_tier = 2
                    return result
            except Exception as e:
                logger.error(f"Error scraping {source_name}: {e}")
        
        return None
    
    def _try_tier3_sources(self, vehicle_id: VehicleIdentification) -> Optional[ScrapingResult]:
        """Try scraping from Tier 3 review/dealer sources."""
        for source_name, source in self.tier3_sources.items():
            try:
                start_time = time.time()
                result = source['scraper'](vehicle_id, source)
                if result and result.success:
                    result.extraction_time = time.time() - start_time
                    result.source_tier = 3
                    return result
            except Exception as e:
                logger.error(f"Error scraping {source_name}: {e}")
        
        return None
    
    def _rate_limit(self, domain: str):
        """Apply rate limiting for a domain."""
        if domain in self.last_request_time:
            elapsed = time.time() - self.last_request_time[domain]
            if elapsed < SCRAPING_CONFIG['rate_limit_delay']:
                time.sleep(SCRAPING_CONFIG['rate_limit_delay'] - elapsed)
        
        self.last_request_time[domain] = time.time()
    
    def _fetch_page(self, url: str, use_selenium: bool = False) -> Optional[str]:
        """Fetch page content with error handling."""
        domain = urlparse(url).netloc
        self._rate_limit(domain)
        
        try:
            if use_selenium and self.driver:
                self.driver.get(url)
                WebDriverWait(self.driver, SCRAPING_CONFIG['selenium_timeout']).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
                return self.driver.page_source
            else:
                response = self.session.get(url, timeout=SCRAPING_CONFIG['request_timeout'])
                response.raise_for_status()
                return response.text
        except Exception as e:
            logger.error(f"Error fetching {url}: {e}")
            return None
    
    # Manufacturer-specific scrapers
    def _scrape_ford_commercial(self, vehicle_id: VehicleIdentification, source: Dict) -> ScrapingResult:
        """Scrape Ford commercial vehicle specifications."""
        model_normalized = normalize_vehicle_model(vehicle_id.model)
        
        # Build URL for specific model
        url = None
        for model in source['models']:
            if model.replace('-', '') in model_normalized.lower():
                url = urljoin(source['base_url'], f"{model}/specs/")
                break
        
        if not url:
            return ScrapingResult(
                success=False,
                source_tier=1,
                source_name="Ford Commercial",
                url=source['base_url'],
                error_message="Model not found in Ford commercial lineup"
            )
        
        # Fetch and parse page
        html = self._fetch_page(url, use_selenium=True)
        if not html:
            return ScrapingResult(
                success=False,
                source_tier=1,
                source_name="Ford Commercial",
                url=url,
                error_message="Failed to fetch page"
            )
        
        # Extract specifications
        soup = BeautifulSoup(html, 'html.parser')
        raw_specs = self.extractor.extract_from_html(html, soup)
        
        # Create CommercialVehicleSpecs object
        specs = CommercialVehicleSpecs(
            payload_capacity_lbs=raw_specs.get('payload'),
            towing_capacity_lbs=raw_specs.get('towing'),
            engine_torque_lb_ft=raw_specs.get('torque'),
            fuel_tank_capacity_gal=raw_specs.get('fuel_capacity'),
            wheelbase_inches=raw_specs.get('wheelbase'),
            data_source="Ford Commercial Website",
            data_confidence=0.9 if raw_specs else 0.3
        )
        
        # Determine duty cycle based on model
        if 'transit' in model_normalized:
            specs.duty_cycle = "Urban Delivery"
            specs.vocation = "Delivery Van"
            specs.electrification_suitability = "High"
            specs.recommended_ev_alternatives = ["Ford E-Transit", "Rivian EDV", "BrightDrop Zevo"]
        elif any(f in model_normalized for f in ['f150', 'f250', 'f350']):
            specs.duty_cycle = "Mixed Use"
            specs.vocation = "Pickup Truck"
            specs.electrification_suitability = "Medium"
            specs.recommended_ev_alternatives = ["Ford F-150 Lightning", "Rivian R1T", "Chevrolet Silverado EV"]
        elif any(f in model_normalized for f in ['f450', 'f550', 'f650', 'f750']):
            specs.duty_cycle = "Heavy Duty"
            specs.vocation = "Chassis Cab"
            specs.electrification_suitability = "Low"
            specs.recommended_ev_alternatives = ["Consider hybrid options", "Wait for heavy-duty EV development"]
        
        return ScrapingResult(
            success=bool(raw_specs),
            source_tier=1,
            source_name="Ford Commercial",
            url=url,
            data=raw_specs,
            specs=specs,
            confidence_score=specs.data_confidence
        )
    
    def _scrape_freightliner(self, vehicle_id: VehicleIdentification, source: Dict) -> ScrapingResult:
        """Scrape Freightliner specifications."""
        # Implementation would follow similar pattern to Ford
        # This is a placeholder showing the structure
        return ScrapingResult(
            success=False,
            source_tier=1,
            source_name="Freightliner",
            url=source['base_url'],
            error_message="Freightliner scraper not yet implemented"
        )
    
    def _scrape_peterbilt(self, vehicle_id: VehicleIdentification, source: Dict) -> ScrapingResult:
        """Scrape Peterbilt specifications."""
        return ScrapingResult(
            success=False,
            source_tier=1,
            source_name="Peterbilt",
            url=source['base_url'],
            error_message="Peterbilt scraper not yet implemented"
        )
    
    def _scrape_kenworth(self, vehicle_id: VehicleIdentification, source: Dict) -> ScrapingResult:
        """Scrape Kenworth specifications."""
        return ScrapingResult(
            success=False,
            source_tier=1,
            source_name="Kenworth",
            url=source['base_url'],
            error_message="Kenworth scraper not yet implemented"
        )
    
    def _scrape_mack(self, vehicle_id: VehicleIdentification, source: Dict) -> ScrapingResult:
        """Scrape Mack specifications."""
        return ScrapingResult(
            success=False,
            source_tier=1,
            source_name="Mack",
            url=source['base_url'],
            error_message="Mack scraper not yet implemented"
        )
    
    def _scrape_volvo(self, vehicle_id: VehicleIdentification, source: Dict) -> ScrapingResult:
        """Scrape Volvo Trucks specifications."""
        return ScrapingResult(
            success=False,
            source_tier=1,
            source_name="Volvo Trucks",
            url=source['base_url'],
            error_message="Volvo scraper not yet implemented"
        )
    
    def _scrape_international(self, vehicle_id: VehicleIdentification, source: Dict) -> ScrapingResult:
        """Scrape International Trucks specifications."""
        return ScrapingResult(
            success=False,
            source_tier=1,
            source_name="International",
            url=source['base_url'],
            error_message="International scraper not yet implemented"
        )
    
    def _scrape_isuzu(self, vehicle_id: VehicleIdentification, source: Dict) -> ScrapingResult:
        """Scrape Isuzu Commercial specifications."""
        return ScrapingResult(
            success=False,
            source_tier=1,
            source_name="Isuzu Commercial",
            url=source['base_url'],
            error_message="Isuzu scraper not yet implemented"
        )
    
    def _scrape_epa_smartway(self, vehicle_id: VehicleIdentification, source: Dict) -> ScrapingResult:
        """Scrape EPA fueleconomy.gov database for MPG data."""
        # Build URL for fueleconomy.gov search
        year = int(vehicle_id.year)  # Convert to int for year arithmetic
        make = vehicle_id.make.replace(' ', '%20')
        
        # Normalize model names for EPA database compatibility
        model = self._normalize_epa_model_name(vehicle_id.model, vehicle_id.make)
        model = model.replace(' ', '%20')
        
        # Try the direct model search URL first
        search_url = f"https://fueleconomy.gov/feg/bymodel/{year}_{make}_{model}.shtml"
        
        logger.info(f"Searching EPA fueleconomy.gov: {search_url}")
        
        # Fetch the page
        html = self._fetch_page(search_url, use_selenium=False)
        if not html:
            # Try nearby years (commercial vehicles often have data in adjacent years)
            nearby_years = [year - 1, year + 1, year - 2]
            for try_year in nearby_years:
                try_url = f"https://fueleconomy.gov/feg/bymodel/{try_year}_{make}_{model}.shtml"
                logger.info(f"Trying nearby year: {try_url}")
                html = self._fetch_page(try_url, use_selenium=False)
                if html:
                    search_url = try_url  # Update for logging
                    break
            
            # Try alternate search if all direct URLs fail
            if not html:
                search_url = f"https://fueleconomy.gov/feg/PowerSearch.do?action=noform&year1={year-2}&year2={year+2}&make={make}&model={model}"
                logger.info(f"Trying power search: {search_url}")
                html = self._fetch_page(search_url, use_selenium=False)
        
        if not html:
            return ScrapingResult(
            success=False,
            source_tier=2,
                source_name="EPA FuelEconomy.gov",
                url=search_url,
                error_message="Failed to fetch EPA page"
            )
        
        # Extract MPG data from the page
        soup = BeautifulSoup(html, 'html.parser')
        raw_specs = self.extractor.extract_from_html(html, soup)
        
        # Look for EPA-specific patterns in tables
        mpg_data = self._extract_epa_mpg_data(soup, vehicle_id)
        raw_specs.update(mpg_data)
        
        if not any(key in raw_specs for key in ['mpg_combined', 'mpg_city', 'mpg_highway']):
            return ScrapingResult(
                success=False,
                source_tier=2,
                source_name="EPA FuelEconomy.gov",
                url=search_url,
                error_message="No MPG data found on EPA page"
            )
        
        # Create specifications object with MPG data
        specs = CommercialVehicleSpecs(
            data_source="EPA FuelEconomy.gov",
            data_confidence=0.95,  # EPA data is highly reliable
            is_estimated=False
        )
        
        # Store MPG data in the result for integration
        mpg_result = {
            'combined_mpg': raw_specs.get('mpg_combined'),
            'city_mpg': raw_specs.get('mpg_city'),
            'highway_mpg': raw_specs.get('mpg_highway'),
            'fuel_capacity': raw_specs.get('fuel_capacity'),
        }
        
        logger.info(f"Found EPA MPG data: Combined={mpg_result.get('combined_mpg')}, City={mpg_result.get('city_mpg')}, Highway={mpg_result.get('highway_mpg')}")
        
        return ScrapingResult(
            success=True,
            source_tier=2,
            source_name="EPA FuelEconomy.gov",
            url=search_url,
            data=mpg_result,
            specs=specs,
            confidence_score=0.95
        )
    
    def _extract_epa_mpg_data(self, soup: BeautifulSoup, vehicle_id: VehicleIdentification) -> Dict[str, Any]:
        """Extract MPG data specifically from EPA fueleconomy.gov pages."""
        mpg_data = {}
        
        # Look for EPA's specific table format
        tables = soup.find_all('table')
        for table in tables:
            rows = table.find_all('tr')
            for row in rows:
                cells = row.find_all(['td', 'th'])
                if len(cells) >= 2:
                    text = ' '.join(cell.get_text().strip() for cell in cells)
                    
                    # EPA uses specific patterns like "Combined MPG:13 combined city/highway MPG"
                    combined_match = re.search(r'Combined MPG[:\s]*(\d+)\s*combined', text, re.IGNORECASE)
                    if combined_match:
                        mpg_data['mpg_combined'] = int(combined_match.group(1))
                    
                    city_match = re.search(r'City MPG[:\s]*(\d+)\s*city', text, re.IGNORECASE)
                    if city_match:
                        mpg_data['mpg_city'] = int(city_match.group(1))
                    
                    highway_match = re.search(r'Highway MPG[:\s]*(\d+)\s*highway', text, re.IGNORECASE)
                    if highway_match:
                        mpg_data['mpg_highway'] = int(highway_match.group(1))
        
        # Also check for patterns in the general text
        page_text = soup.get_text()
        
        # Additional EPA patterns from the research
        patterns = [
            (r'Combined MPG[:\s]*(\d+)', 'mpg_combined'),
            (r'(\d+)\s*combined city/highway', 'mpg_combined'),
            (r'City MPG[:\s]*(\d+)', 'mpg_city'),
            (r'(\d+)\s*city', 'mpg_city'), 
            (r'Highway MPG[:\s]*(\d+)', 'mpg_highway'),
            (r'(\d+)\s*highway', 'mpg_highway'),
        ]
        
        for pattern, key in patterns:
            if key not in mpg_data:  # Don't overwrite already found values
                            match = re.search(pattern, page_text, re.IGNORECASE)
            if match:
                try:
                    mpg_value = int(match.group(1))
                    # Validate MPG values - reject obviously wrong ones
                    if self._is_valid_mpg(mpg_value, key):
                        mpg_data[key] = mpg_value
                except (ValueError, IndexError):
                    continue
        
        return mpg_data
    
    def _extract_fuelly_mpg_data(self, soup: BeautifulSoup, vehicle_id: VehicleIdentification) -> Dict[str, Any]:
        """Extract MPG data from Fuelly.com pages."""
        mpg_data = {}
        
        # Look for the main MPG display on Fuelly pages
        # Pattern: "8.89 with a 0.29 MPG margin of error"
        mpg_text = soup.get_text()
        
        # Extract average MPG
        import re
        avg_mpg_match = re.search(r'gets a combined Avg MPG of ([\d\.]+)', mpg_text)
        if avg_mpg_match:
            combined_mpg = float(avg_mpg_match.group(1))
            mpg_data['mpg_combined'] = combined_mpg
            logger.info(f"Found Fuelly combined MPG: {combined_mpg}")
        
        # Alternative pattern: "X.X Avg MPG" 
        if not mpg_data.get('mpg_combined'):
            avg_mpg_match = re.search(r'([\d\.]+) Avg MPG', mpg_text)
            if avg_mpg_match:
                combined_mpg = float(avg_mpg_match.group(1))
                mpg_data['mpg_combined'] = combined_mpg
                logger.info(f"Found Fuelly average MPG (alt pattern): {combined_mpg}")
        
        # Extract number of vehicles and fuel-ups for confidence scoring
        vehicles_match = re.search(r'Based on data from (\d+) vehicles', mpg_text)
        fuel_ups_match = re.search(r'(\d+) fuel-ups', mpg_text)
        
        if vehicles_match and fuel_ups_match:
            vehicles_count = int(vehicles_match.group(1))
            fuel_ups_count = int(fuel_ups_match.group(1))
            mpg_data['vehicles_count'] = vehicles_count
            mpg_data['fuel_ups_count'] = fuel_ups_count
            mpg_data['mpg_range'] = f"{vehicles_count} vehicles, {fuel_ups_count} fuel-ups"
            logger.info(f"Fuelly data confidence: {vehicles_count} vehicles, {fuel_ups_count} fuel-ups")
        
        # Extract individual vehicle MPG ranges from the page
        individual_mpgs = re.findall(r'(\d+\.\d+)\s*Avg MPG', mpg_text)
        if individual_mpgs:
            mpg_values = [float(mpg) for mpg in individual_mpgs]
            if mpg_values:
                mpg_data['mpg_min'] = min(mpg_values)
                mpg_data['mpg_max'] = max(mpg_values)
                logger.info(f"Fuelly MPG range: {mpg_data['mpg_min']} - {mpg_data['mpg_max']}")
        
        # Look for year-specific data
        year_pattern = f"{vehicle_id.year}.*?([\\d\\.]+) Avg MPG"
        year_match = re.search(year_pattern, mpg_text)
        if year_match:
            year_mpg = float(year_match.group(1))
            mpg_data['mpg_year_specific'] = year_mpg
            # Use year-specific data if available
            mpg_data['mpg_combined'] = year_mpg
            logger.info(f"Found {vehicle_id.year}-specific MPG: {year_mpg}")
        
        return mpg_data
    
    def _is_valid_mpg(self, mpg_value: int, mpg_type: str) -> bool:
        """Validate if an MPG value is reasonable for commercial vehicles."""
        # Reasonable MPG ranges for commercial vehicles
        if mpg_type == 'mpg_combined':
            return 3 <= mpg_value <= 25  # Commercial vehicles: 3-25 MPG combined
        elif mpg_type == 'mpg_city':
            return 2 <= mpg_value <= 20  # City MPG typically lower
        elif mpg_type == 'mpg_highway':
            return 4 <= mpg_value <= 30  # Highway MPG typically higher
        
        return False
    
    def _has_useful_mpg_data(self, result: ScrapingResult) -> bool:
        """Check if scraping result contains useful MPG data."""
        if not result.data:
            return False
            
        # Check if any MPG values are present and valid
        mpg_fields = ['mpg_combined', 'mpg_city', 'mpg_highway', 'combined_mpg', 'city_mpg', 'highway_mpg']
        for field in mpg_fields:
            if field in result.data:
                mpg_value = result.data[field]
                if mpg_value and isinstance(mpg_value, (int, float)):
                    # Use the validation logic - normalize field name
                    if 'combined' in field:
                        mpg_type = 'mpg_combined'
                    elif 'city' in field:
                        mpg_type = 'mpg_city'
                    elif 'highway' in field:
                        mpg_type = 'mpg_highway'
                    else:
                        mpg_type = field
                    
                    if self._is_valid_mpg(int(mpg_value), mpg_type):
                        return True
        
        return False
    
    def _normalize_epa_model_name(self, model: str, make: str) -> str:
        """Normalize model names to match EPA fueleconomy.gov database naming conventions."""
        model_lower = model.lower()
        make_lower = make.lower()
        
        # Ford commercial vehicle mappings
        if make_lower == 'ford':
            if 'e-150' in model_lower or 'e150' in model_lower:
                return 'Econoline'
            elif 'e-250' in model_lower or 'e250' in model_lower:
                return 'Econoline'
            elif 'e-350' in model_lower or 'e350' in model_lower:
                return 'Econoline'
            elif 'e-450' in model_lower or 'e450' in model_lower:
                return 'Econoline'
            elif 'transit' in model_lower:
                return 'Transit'
        
        # Chevrolet/GMC commercial vehicle mappings
        elif make_lower in ['chevrolet', 'gmc']:
            if 'express' in model_lower:
                return 'Express'
            elif 'savana' in model_lower:
                return 'Savana'
        
        # Ram commercial vehicle mappings  
        elif make_lower in ['ram', 'dodge']:
            if 'promaster' in model_lower:
                return 'ProMaster'
            elif 'ram' in model_lower and 'van' in model_lower:
                return 'Ram Van'
        
        # Mercedes commercial vehicle mappings
        elif make_lower == 'mercedes-benz':
            if 'sprinter' in model_lower:
                return 'Sprinter'
            elif 'metris' in model_lower:
                return 'Metris'
        
        # Default: return original model
        return model
    
    def _normalize_fuelly_model_name(self, model: str, make: str) -> str:
        """Normalize model names to match Fuelly.com URL format."""
        model_lower = model.lower()
        make_lower = make.lower()
        
        # Convert to Fuelly URL format (underscores, no hyphens)
        model_clean = model_lower.replace('-', '_').replace(' ', '_')
        
        # Ford model mappings for Fuelly (pickups use /car/ not /truck/)
        if make_lower == 'ford':
            if 'f-350' in model_lower or 'f350' in model_lower:
                return 'f-350_super_duty'  # Fuelly uses this exact format
            elif 'f-250' in model_lower or 'f250' in model_lower:
                return 'f-250_super_duty'  # Consistent with F-350 format
            elif 'f-150' in model_lower or 'f150' in model_lower:
                return 'f-150'  # F-150 doesn't use super duty suffix
            elif 'f-550' in model_lower or 'f550' in model_lower:
                return 'f-550_super_duty'  # Consistent with F-350 format
            elif 'e-250' in model_lower or 'e250' in model_lower:
                return 'e-250'  # Van models likely different
            elif 'e-350' in model_lower or 'e350' in model_lower:
                return 'e-350'  # Van models likely different
            elif 'transit' in model_lower:
                if 'connect' in model_lower:
                    return 'transit_connect'  # Fix: was missing _connect
                else:
                    return 'transit'
        
        # RAM model mappings for Fuelly
        elif make_lower == 'ram':
            if '3500' in model_lower:
                return '3500'
            elif '2500' in model_lower:
                return '2500'
            elif '1500' in model_lower:
                return '1500'
        
        # Freightliner model mappings for Fuelly
        if make_lower == 'freightliner':
            if 'm2' in model_lower:
                return 'm2_106' if '106' not in model_lower else 'm2_106'
            elif 'cascadia' in model_lower:
                return 'cascadia'
            elif 'century' in model_lower:
                return 'century_class'
        
        # International model mappings for Fuelly
        elif make_lower == 'international':
            if 'ma025' in model_lower or 'ma25' in model_lower:
                return 'durastar'  # MA025 is often called DuraStar
            elif 'prostar' in model_lower:
                return 'prostar'
            elif 'lonestar' in model_lower:
                return 'lonestar'
        
        # RAM/Dodge commercial mappings
        elif make_lower in ['ram', 'dodge']:
            if 'promaster' in model_lower:
                if '1500' in model_lower:
                    return 'promaster_1500'
                elif '2500' in model_lower:
                    return 'promaster_2500'  
                elif '3500' in model_lower:
                    return 'promaster_3500'
                else:
                    return 'promaster'
        
        # Chevrolet/GMC mappings (use exact Fuelly model names)
        elif make_lower in ['chevrolet', 'gmc']:
            if 'silverado' in model_lower:
                if '3500' in model_lower:
                    return 'silverado_3500_hd'  # Confirmed: fuelly.com/car/chevrolet/silverado_3500_hd
                elif '2500' in model_lower:
                    return 'silverado_2500_hd'  # Confirmed: fuelly.com/car/chevrolet/silverado_2500_hd
                elif 'hd' in model_lower:
                    return 'silverado_hd'  # Generic HD fallback
                else:
                    return 'silverado'
            elif 'sierra' in model_lower:
                if '3500' in model_lower:
                    return 'sierra_3500_hd'
                elif '2500' in model_lower:
                    return 'sierra_2500_hd' 
                elif 'hd' in model_lower:
                    return 'sierra_hd'
                else:
                    return 'sierra'
        
        return model_clean
    
    def _scrape_carb(self, vehicle_id: VehicleIdentification, source: Dict) -> ScrapingResult:
        """Scrape CARB Clean Truck Check."""
        return ScrapingResult(
            success=False,
            source_tier=2,
            source_name="CARB Clean Truck Check",
            url=source['base_url'],
            error_message="CARB scraper not yet implemented"
        )
    
    def _scrape_truck_trader(self, vehicle_id: VehicleIdentification, source: Dict) -> ScrapingResult:
        """Scrape Commercial Truck Trader."""
        return ScrapingResult(
            success=False,
            source_tier=3,
            source_name="Commercial Truck Trader",
            url=source['base_url'],
            error_message="Truck Trader scraper not yet implemented"
        )
    
    def _scrape_truck_paper(self, vehicle_id: VehicleIdentification, source: Dict) -> ScrapingResult:
        """Scrape TruckPaper."""
        return ScrapingResult(
            success=False,
            source_tier=3,
            source_name="TruckPaper",
            url=source['base_url'],
            error_message="TruckPaper scraper not yet implemented"
        )
    
    def _is_valid_fuelly_page(self, html: str) -> bool:
        """Check if a Fuelly page contains valid vehicle data."""
        if not html:
            return False
            
        # Check for indicators of a valid vehicle page
        valid_indicators = [
            'avg mpg',
            'fuel-ups',
            'miles tracked',
            'vehicles',
            'combined avg mpg'
        ]
        
        # Check for redirect/error indicators
        invalid_indicators = [
            'invalid model',
            'page not found',
            'no vehicles found',
            'browse similar vehicles'  # Redirect page
        ]
        
        html_lower = html.lower()
        
        # Must have at least one valid indicator
        has_valid = any(indicator in html_lower for indicator in valid_indicators)
        
        # Must not have any invalid indicators  
        has_invalid = any(indicator in html_lower for indicator in invalid_indicators)
        
        return has_valid and not has_invalid
    
    def _scrape_fuelly(self, vehicle_id: VehicleIdentification, source: Dict) -> ScrapingResult:
        """Scrape Fuelly.com for real-world community MPG data."""
        year = int(vehicle_id.year)
        make = vehicle_id.make.lower()
        model = self._normalize_fuelly_model_name(vehicle_id.model, vehicle_id.make)
        
        # Universal approach: Try BOTH sections AND multiple model variations
        # No guessing, no assumptions - try everything and let Fuelly tell us what works
        
        # Generate multiple model name variations to try
        model_variations = [
            model,  # Original normalized model
        ]
        
        # Add common variations for ALL models
        if '_' in model:
            # Try without underscores: "transit_connect" → "transit-connect", "transitconnect"
            model_variations.extend([
                model.replace('_', '-'),
                model.replace('_', '')
            ])
        
        if '-' in model:
            # Try with underscores: "f-350" → "f_350"
            model_variations.append(model.replace('-', '_'))
            
        # Add simplified version (remove suffixes)
        base_model = model.split('_')[0].split('-')[0]  # Get first part only
        if base_model != model and len(base_model) >= 2:
            model_variations.append(base_model)
        
        # Remove duplicates while preserving order
        seen = set()
        model_variations = [x for x in model_variations if not (x in seen or seen.add(x))]
        
        # Generate all URL combinations
        urls_to_try = []
        for model_var in model_variations:
            # Try both sections for each model variation
            urls_to_try.extend([
                f"https://www.fuelly.com/car/{make}/{model_var}/{year}",
                f"https://www.fuelly.com/car/{make}/{model_var}",
                f"https://www.fuelly.com/truck/{make}/{model_var}/{year}",
                f"https://www.fuelly.com/truck/{make}/{model_var}",
            ])
        
        html = None
        tried_url = None
        
        for fuelly_url in urls_to_try:
            logger.info(f"Searching Fuelly.com: {fuelly_url}")
            tried_url = fuelly_url
            html = self._fetch_page(fuelly_url, use_selenium=False)
            
            # Check for success indicators
            if html and self._is_valid_fuelly_page(html):
                logger.info(f"Found valid Fuelly page: {fuelly_url}")
                break  # Found working URL
            else:
                logger.debug(f"URL failed or invalid page: {fuelly_url}")
        
        # Set the final URL for result reporting
        fuelly_url = tried_url
        
        if not html:
            return ScrapingResult(
                success=False,
                source_tier=3,
                source_name="Fuelly Community",
                url=fuelly_url,
                error_message="Failed to fetch Fuelly page"
            )
        
        # Extract MPG data from Fuelly page
        soup = BeautifulSoup(html, 'html.parser')
        mpg_data = self._extract_fuelly_mpg_data(soup, vehicle_id)
        
        if not any(key in mpg_data for key in ['mpg_combined', 'mpg_city', 'mpg_highway']):
            return ScrapingResult(
                success=False,
                source_tier=3,
                source_name="Fuelly Community",
                url=fuelly_url,
                error_message="No MPG data found on Fuelly page"
            )
        
        # Create specifications object with community MPG data
        specs = CommercialVehicleSpecs(
            data_source="Fuelly Community",
            data_confidence=0.85,  # Community data is generally reliable
            is_estimated=False
        )
        
        logger.info(f"Found Fuelly MPG data: Combined={mpg_data.get('mpg_combined')}, Range={mpg_data.get('mpg_range', 'N/A')}")
        
        return ScrapingResult(
            success=True,
            source_tier=3,
            source_name="Fuelly Community",
            url=fuelly_url,
            data=mpg_data,
            specs=specs,
            confidence_score=0.85
        )
    
    def _estimate_specifications(self, vehicle_id: VehicleIdentification) -> ScrapingResult:
        """Estimate specifications based on similar vehicles."""
        specs = CommercialVehicleSpecs(
            is_estimated=True,
            data_source="Pattern-based estimation",
            data_confidence=0.5
        )
        
        # Use GVWR to estimate other values
        if vehicle_id.gvwr_pounds > 0:
            gvwr = vehicle_id.gvwr_pounds
            
            # Rough estimation formulas based on industry patterns
            specs.payload_capacity_lbs = gvwr * 0.35  # Payload typically 30-40% of GVWR
            specs.gcwr_lbs = gvwr * 2.2  # GCWR often 2-2.5x GVWR for trucks
            
            # Estimate towing based on GVWR class
            if gvwr <= 10000:
                specs.towing_capacity_lbs = gvwr * 0.8
            elif gvwr <= 19500:
                specs.towing_capacity_lbs = gvwr * 1.2
            else:
                specs.towing_capacity_lbs = gvwr * 1.5
            
            # Estimate fuel capacity based on GVWR
            if gvwr <= 8500:
                specs.fuel_tank_capacity_gal = 26
            elif gvwr <= 14000:
                specs.fuel_tank_capacity_gal = 40
            elif gvwr <= 26000:
                specs.fuel_tank_capacity_gal = 50
            else:
                specs.fuel_tank_capacity_gal = 100
            
            # Classify duty cycle
            if gvwr <= 8500:
                specs.duty_cycle = "Light Duty"
                specs.electrification_suitability = "High"
            elif gvwr <= 19500:
                specs.duty_cycle = "Medium Duty"
                specs.electrification_suitability = "Medium"
            else:
                specs.duty_cycle = "Heavy Duty"
                specs.electrification_suitability = "Low"
        
        # Use body class for vocation
        body_lower = vehicle_id.body_class.lower()
        if 'van' in body_lower:
            specs.vocation = "Delivery/Service Van"
            specs.electrification_suitability = "High"
        elif 'pickup' in body_lower or 'truck' in body_lower:
            specs.vocation = "Pickup/Work Truck"
            specs.electrification_suitability = "Medium"
        elif 'bus' in body_lower:
            specs.vocation = "Passenger Transport"
            specs.electrification_suitability = "High"
        
        return ScrapingResult(
            success=True,
            source_tier=0,  # Estimation tier
            source_name="Pattern Estimation",
            url="",
            data={
                'payload': specs.payload_capacity_lbs,
                'towing': specs.towing_capacity_lbs,
                'fuel_capacity': specs.fuel_tank_capacity_gal,
                'gcwr': specs.gcwr_lbs
            },
            specs=specs,
            confidence_score=specs.data_confidence
        )
    
    def cleanup(self):
        """Clean up resources."""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass

###############################################################################
# Enhanced Vehicle Data Provider
###############################################################################

class EnhancedCommercialVehicleProvider(VehicleDataProvider):
    """
    Enhanced vehicle data provider with intelligent web scraping for commercial vehicles.
    Extends the base VehicleDataProvider with scraping capabilities.
    """
    
    def __init__(self, cache_enabled: bool = True, enable_scraping: bool = True, 
                 use_selenium: bool = False):
        """
        Initialize the enhanced provider.
        
        Args:
            cache_enabled: Whether to use caching
            enable_scraping: Whether to enable web scraping
            use_selenium: Whether to use Selenium for dynamic content
        """
        super().__init__(cache_enabled=cache_enabled)
        
        self.enable_scraping = enable_scraping
        self.scraper = None
        
        if enable_scraping:
            self.scraper = CommercialVehicleScraper(
                cache_enabled=cache_enabled,
                use_selenium=use_selenium
            )
        
        # Enhanced logging for commercial vehicles
        self.commercial_logger = logging.getLogger("commercial_vehicles")
        self.commercial_logger.setLevel(logging.INFO)
    
    def get_vehicle_by_vin(self, vin: str) -> Tuple[bool, Dict[str, Any], str]:
        """
        Get comprehensive vehicle data by VIN with enhanced commercial vehicle support.
        
        Args:
            vin: Vehicle Identification Number
            
        Returns:
            Tuple of (success, data, error_message)
        """
        # First, try the standard API approach
        success, data, error = super().get_vehicle_by_vin(vin)
        
        # If basic data was retrieved but appears incomplete for a commercial vehicle
        if success and self._is_commercial_vehicle(data):
            self.commercial_logger.info(f"Commercial vehicle detected: {vin}")
            
            # Check data completeness
            completeness = self._assess_data_completeness(data)
            self.commercial_logger.info(f"Data completeness: {completeness:.1%}")
            
            # If data is incomplete and scraping is enabled
            if completeness < 0.8 and self.enable_scraping and self.scraper:
                self.commercial_logger.info(f"Attempting to enhance data through web scraping")
                
                # Create VehicleIdentification from existing data
                vehicle_id = VehicleIdentification.from_dict(data.get('vehicle_id', {}))
                
                # Attempt scraping
                scraping_result = self.scraper.scrape_vehicle(vehicle_id)
                
                if scraping_result and scraping_result.success:
                    # Merge scraped data with existing data
                    data = self._merge_scraped_data(data, scraping_result)
                    self.commercial_logger.info(
                        f"Successfully enhanced data from {scraping_result.source_name} "
                        f"(Tier {scraping_result.source_tier})"
                    )
                else:
                    self.commercial_logger.warning(
                        f"Scraping failed or returned no additional data for {vin}"
                    )
        
        return success, data, error
    
    def _is_commercial_vehicle(self, data: Dict[str, Any]) -> bool:
        """Determine if vehicle is commercial based on available data."""
        vehicle_id = data.get('vehicle_id', {})
        
        # Check GVWR
        gvwr = vehicle_id.get('gvwr_pounds', 0)
        if gvwr > 8500:
            return True
        
        # Check body class
        body_class = vehicle_id.get('body_class', '').lower()
        commercial_indicators = ['truck', 'van', 'bus', 'chassis', 'cab', 'commercial']
        if any(indicator in body_class for indicator in commercial_indicators):
            return True
        
        # Check make
        make = vehicle_id.get('make', '').lower()
        commercial_makes = ['freightliner', 'peterbilt', 'kenworth', 'mack', 'volvo trucks',
                          'international', 'isuzu', 'hino', 'western star']
        if make in commercial_makes:
            return True
        
        # Check model
        model = vehicle_id.get('model', '').lower()
        commercial_models = ['f250', 'f350', 'f450', 'f550', 'f650', 'f750',
                           'silverado 2500', 'silverado 3500', 'sierra 2500', 'sierra 3500',
                           'ram 2500', 'ram 3500', 'transit', 'sprinter', 'promaster']
        if any(cm in model for cm in commercial_models):
            return True
        
        return False
    
    def _assess_data_completeness(self, data: Dict[str, Any]) -> float:
        """Assess how complete the vehicle data is."""
        vehicle_id = data.get('vehicle_id', {})
        fuel_economy = data.get('fuel_economy', {})
        
        # Critical fields for commercial vehicles
        critical_fields = [
            ('gvwr_pounds', vehicle_id),
            ('commercial_category', vehicle_id),
            ('engine_power_hp', vehicle_id),
            ('engine_type', vehicle_id),
            ('combined_mpg', fuel_economy),
            ('fuel_type', vehicle_id),
            ('transmission', vehicle_id),
            ('drive_type', vehicle_id)
        ]
        
        # Count populated fields
        populated = sum(1 for field, source in critical_fields 
                       if source.get(field) and str(source.get(field)).strip())
        
        return populated / len(critical_fields)
    
    def _merge_scraped_data(self, existing_data: Dict[str, Any], 
                          scraping_result: ScrapingResult) -> Dict[str, Any]:
        """Merge scraped data with existing data."""
        merged = existing_data.copy()
        
        if scraping_result.specs:
            specs = scraping_result.specs
            
            # Add scraped specifications to vehicle_id
            vehicle_id = merged.get('vehicle_id', {})
            
            # Add payload and towing capacity
            if specs.payload_capacity_lbs:
                vehicle_id['payload_capacity_lbs'] = specs.payload_capacity_lbs
            if specs.towing_capacity_lbs:
                vehicle_id['towing_capacity_lbs'] = specs.towing_capacity_lbs
            
            # Add engine torque
            if specs.engine_torque_lb_ft:
                vehicle_id['engine_torque_lb_ft'] = specs.engine_torque_lb_ft
            
            # Add fuel capacity
            if specs.fuel_tank_capacity_gal:
                vehicle_id['fuel_tank_capacity_gal'] = specs.fuel_tank_capacity_gal
            
            # Add dimensions
            if specs.wheelbase_inches:
                vehicle_id['wheelbase_inches'] = specs.wheelbase_inches
            
            # Add operational classification
            vehicle_id['duty_cycle'] = specs.duty_cycle
            vehicle_id['vocation'] = specs.vocation
            vehicle_id['electrification_suitability'] = specs.electrification_suitability
            
            # Add data quality metrics
            vehicle_id['scraping_source'] = specs.data_source
            vehicle_id['scraping_confidence'] = specs.data_confidence
            vehicle_id['data_is_estimated'] = specs.is_estimated
            
            merged['vehicle_id'] = vehicle_id
            
            # Add recommended EV alternatives
            if specs.recommended_ev_alternatives:
                merged['ev_alternatives'] = specs.recommended_ev_alternatives
            
            # Update match confidence based on scraping success
            if 'match_confidence' in merged:
                # Boost confidence if we got good scraped data
                if specs.data_confidence > 0.7:
                    merged['match_confidence'] = min(100, merged['match_confidence'] + 20)
        
        # Integrate MPG data if found in scraping result
        if scraping_result.data:
            result_data = scraping_result.data
            fuel_economy = merged.get('fuel_economy', {})
            
            # Update fuel economy with scraped MPG data
            # Handle both formats: mpg_combined/combined_mpg, mpg_city/city_mpg, etc.
            combined_mpg = result_data.get('combined_mpg') or result_data.get('mpg_combined')
            if combined_mpg:
                fuel_economy['combined_mpg'] = combined_mpg
                self.commercial_logger.info(f"Updated combined MPG to {combined_mpg} from {scraping_result.source_name}")
            
            city_mpg = result_data.get('city_mpg') or result_data.get('mpg_city')
            if city_mpg:
                fuel_economy['city_mpg'] = city_mpg
                self.commercial_logger.info(f"Updated city MPG to {city_mpg} from {scraping_result.source_name}")
            
            highway_mpg = result_data.get('highway_mpg') or result_data.get('mpg_highway')
            if highway_mpg:
                fuel_economy['highway_mpg'] = highway_mpg
                self.commercial_logger.info(f"Updated highway MPG to {highway_mpg} from {scraping_result.source_name}")
            
            # Update fuel economy in merged data
            merged['fuel_economy'] = fuel_economy
            
            # Also update fuel capacity if found
            if result_data.get('fuel_capacity'):
                vehicle_id = merged.get('vehicle_id', {})
                vehicle_id['fuel_tank_capacity_gal'] = result_data['fuel_capacity']
                merged['vehicle_id'] = vehicle_id
        
        return merged
    
    def get_commercial_vehicle_analysis(self, vin: str) -> Dict[str, Any]:
        """
        Get comprehensive commercial vehicle analysis including electrification potential.
        
        Args:
            vin: Vehicle Identification Number
            
        Returns:
            Dictionary with analysis results
        """
        # Get enhanced vehicle data
        success, data, error = self.get_vehicle_by_vin(vin)
        
        if not success:
            return {
                'success': False,
                'error': error,
                'analysis': {}
            }
        
        # Perform commercial vehicle analysis
        vehicle_id = data.get('vehicle_id', {})
        fuel_economy = data.get('fuel_economy', {})
        
        analysis = {
            'vehicle_classification': self._classify_commercial_vehicle(vehicle_id),
            'operational_profile': self._determine_operational_profile(vehicle_id),
            'electrification_assessment': self._assess_electrification_potential(vehicle_id, fuel_economy),
            'tco_comparison': self._calculate_tco_comparison(vehicle_id, fuel_economy),
            'recommended_alternatives': data.get('ev_alternatives', []),
            'data_quality': self._calculate_enhanced_quality_score(data)
        }
        
        return {
            'success': True,
            'data': data,
            'analysis': analysis
        }
    
    def _classify_commercial_vehicle(self, vehicle_id: Dict[str, Any]) -> Dict[str, Any]:
        """Classify commercial vehicle by various criteria."""
        gvwr = vehicle_id.get('gvwr_pounds', 0)
        
        # DOT classification
        if gvwr <= 6000:
            dot_class = "Class 1"
        elif gvwr <= 10000:
            dot_class = "Class 2"
        elif gvwr <= 14000:
            dot_class = "Class 3"
        elif gvwr <= 16000:
            dot_class = "Class 4"
        elif gvwr <= 19500:
            dot_class = "Class 5"
        elif gvwr <= 26000:
            dot_class = "Class 6"
        elif gvwr <= 33000:
            dot_class = "Class 7"
        else:
            dot_class = "Class 8"
        
        return {
            'dot_class': dot_class,
            'gvwr_category': vehicle_id.get('commercial_category', ''),
            'duty_cycle': vehicle_id.get('duty_cycle', ''),
            'vocation': vehicle_id.get('vocation', ''),
            'is_diesel': vehicle_id.get('is_diesel', False)
        }
    
    def _determine_operational_profile(self, vehicle_id: Dict[str, Any]) -> Dict[str, Any]:
        """Determine operational profile for the vehicle."""
        body_class = vehicle_id.get('body_class', '').lower()
        model = vehicle_id.get('model', '').lower()
        gvwr = vehicle_id.get('gvwr_pounds', 0)
        
        # Default profile
        profile = {
            'typical_daily_miles': 100,
            'typical_route_type': 'Mixed',
            'idle_time_percentage': 20,
            'stop_frequency': 'Medium',
            'load_factor': 0.5
        }
        
        # Adjust based on vehicle type
        if 'van' in body_class or 'transit' in model:
            profile.update({
                'typical_daily_miles': 80,
                'typical_route_type': 'Urban',
                'idle_time_percentage': 30,
                'stop_frequency': 'High',
                'load_factor': 0.3
            })
        elif gvwr > 26000:  # Heavy duty
            profile.update({
                'typical_daily_miles': 250,
                'typical_route_type': 'Highway',
                'idle_time_percentage': 10,
                'stop_frequency': 'Low',
                'load_factor': 0.7
            })
        elif 'pickup' in body_class:
            profile.update({
                'typical_daily_miles': 60,
                'typical_route_type': 'Mixed',
                'idle_time_percentage': 25,
                'stop_frequency': 'Medium',
                'load_factor': 0.4
            })
        
        return profile
    
    def _assess_electrification_potential(self, vehicle_id: Dict[str, Any], 
                                         fuel_economy: Dict[str, Any]) -> Dict[str, Any]:
        """Assess vehicle's potential for electrification."""
        gvwr = vehicle_id.get('gvwr_pounds', 0)
        mpg = fuel_economy.get('combined_mpg', 0)
        duty_cycle = vehicle_id.get('duty_cycle', '')
        
        # Start with base score
        score = 50
        
        # Weight class factor (lighter vehicles easier to electrify)
        if gvwr <= 10000:
            score += 20
        elif gvwr <= 19500:
            score += 10
        elif gvwr <= 26000:
            score += 0
        else:
            score -= 20
        
        # Fuel efficiency factor (less efficient = more savings potential)
        if mpg > 0:
            if mpg < 10:
                score += 15
            elif mpg < 15:
                score += 10
            elif mpg < 20:
                score += 5
        
        # Duty cycle factor
        if duty_cycle in ['Urban Delivery', 'Light Duty']:
            score += 15
        elif duty_cycle == 'Medium Duty':
            score += 5
        elif duty_cycle == 'Heavy Duty':
            score -= 10
        
        # Determine suitability level
        if score >= 80:
            suitability = "Excellent"
            recommendation = "Strong candidate for immediate electrification"
        elif score >= 60:
            suitability = "Good"
            recommendation = "Good candidate for electrification with current technology"
        elif score >= 40:
            suitability = "Moderate"
            recommendation = "Consider electrification as technology improves"
        else:
            suitability = "Limited"
            recommendation = "Wait for advances in heavy-duty EV technology"
        
        return {
            'score': score,
            'suitability': suitability,
            'recommendation': recommendation,
            'barriers': self._identify_electrification_barriers(vehicle_id),
            'benefits': self._identify_electrification_benefits(vehicle_id, fuel_economy)
        }
    
    def _identify_electrification_barriers(self, vehicle_id: Dict[str, Any]) -> List[str]:
        """Identify barriers to electrification."""
        barriers = []
        
        gvwr = vehicle_id.get('gvwr_pounds', 0)
        towing = vehicle_id.get('towing_capacity_lbs', 0)
        
        if gvwr > 26000:
            barriers.append("Limited heavy-duty EV options available")
        
        if towing > 10000:
            barriers.append("High towing requirements may reduce EV range significantly")
        
        if vehicle_id.get('duty_cycle') == 'Long Haul':
            barriers.append("Long-distance operations exceed current EV range capabilities")
        
        if not barriers:
            barriers.append("No significant barriers identified")
        
        return barriers
    
    def _identify_electrification_benefits(self, vehicle_id: Dict[str, Any], 
                                          fuel_economy: Dict[str, Any]) -> List[str]:
        """Identify benefits of electrification."""
        benefits = []
        
        mpg = fuel_economy.get('combined_mpg', 0)
        
        if mpg < 15:
            benefits.append("Significant fuel cost savings potential")
        
        if vehicle_id.get('duty_cycle') in ['Urban Delivery', 'Light Duty']:
            benefits.append("Ideal duty cycle for electric operation")
        
        if vehicle_id.get('is_diesel'):
            benefits.append("Eliminate diesel emissions in urban areas")
        
        benefits.append("Reduced maintenance costs")
        benefits.append("Potential incentives and grants available")
        
        return benefits
    
    def _calculate_tco_comparison(self, vehicle_id: Dict[str, Any], 
                                 fuel_economy: Dict[str, Any]) -> Dict[str, Any]:
        """Calculate simplified TCO comparison."""
        # Simplified TCO calculation
        annual_miles = 30000  # Default commercial vehicle annual mileage
        years = 7  # Typical commercial vehicle ownership period
        
        mpg = fuel_economy.get('combined_mpg', 15)  # Default if not available
        fuel_price = 3.50  # $/gallon
        electricity_price = 0.13  # $/kWh
        ev_efficiency = 2.0  # miles/kWh for commercial EV
        
        # Current vehicle costs
        ice_fuel_cost_annual = (annual_miles / mpg) * fuel_price
        ice_maintenance_annual = annual_miles * 0.15  # $0.15/mile for diesel
        ice_total_7yr = (ice_fuel_cost_annual + ice_maintenance_annual) * years
        
        # EV equivalent costs
        ev_fuel_cost_annual = (annual_miles / ev_efficiency) * electricity_price
        ev_maintenance_annual = annual_miles * 0.08  # $0.08/mile for EV
        ev_total_7yr = (ev_fuel_cost_annual + ev_maintenance_annual) * years
        
        savings_7yr = ice_total_7yr - ev_total_7yr
        
        return {
            'ice_annual_operating_cost': ice_fuel_cost_annual + ice_maintenance_annual,
            'ev_annual_operating_cost': ev_fuel_cost_annual + ev_maintenance_annual,
            'total_7yr_savings': savings_7yr,
            'annual_savings': savings_7yr / years,
            'break_even_years': 3.5  # Simplified estimate
        }
    
    def _calculate_enhanced_quality_score(self, data: Dict[str, Any]) -> float:
        """Calculate enhanced quality score including scraped data."""
        base_score = 0.0
        
        vehicle_id = data.get('vehicle_id', {})
        fuel_economy = data.get('fuel_economy', {})
        
        # Basic fields (40 points)
        basic_fields = ['year', 'make', 'model', 'fuel_type', 'body_class']
        for field in basic_fields:
            if vehicle_id.get(field):
                base_score += 8
        
        # Commercial fields (30 points)
        commercial_fields = ['gvwr_pounds', 'commercial_category', 'payload_capacity_lbs',
                           'towing_capacity_lbs', 'duty_cycle', 'vocation']
        for field in commercial_fields:
            if vehicle_id.get(field):
                base_score += 5
        
        # Fuel economy (20 points)
        if fuel_economy.get('combined_mpg', 0) > 0:
            base_score += 20
        
        # Scraped data bonus (10 points)
        if vehicle_id.get('scraping_confidence', 0) > 0.7:
            base_score += 10
        elif vehicle_id.get('scraping_confidence', 0) > 0.5:
            base_score += 5
        
        return min(100, base_score)
    
    def cleanup(self):
        """Clean up resources."""
        if self.scraper:
            self.scraper.cleanup()

###############################################################################
# Integration with Existing Application
###############################################################################

def integrate_enhanced_provider():
    """
    Function to integrate the enhanced provider with the existing application.
    This replaces the standard VehicleDataProvider in processor.py
    """
    
    # Modify processor.py to use EnhancedCommercialVehicleProvider
    integration_code = '''
    # In processor.py, replace the VehicleDataProvider import and initialization:
    
    # Old code:
    # from data.providers import VehicleDataProvider
    # self.provider = VehicleDataProvider(cache_enabled=True)
    
    # New code:
    from commercial_vehicle_scraper import EnhancedCommercialVehicleProvider
    
    # In ProcessingPipeline.__init__:
    self.provider = EnhancedCommercialVehicleProvider(
        cache_enabled=True,
        enable_scraping=True,
        use_selenium=False  # Set to True for sites requiring JavaScript
    )
    
    # Add cleanup in ProcessingPipeline:
    def cleanup(self):
        """Clean up resources including scraper."""
        if hasattr(self.provider, 'cleanup'):
            self.provider.cleanup()
    '''
    
    return integration_code

# Example usage and testing
if __name__ == "__main__":
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # Test the enhanced provider
    provider = EnhancedCommercialVehicleProvider(
        cache_enabled=True,
        enable_scraping=True,
        use_selenium=False
    )
    
    # Test with a commercial vehicle VIN
    test_vin = "1FTFW1ET5DFC10312"  # Example Ford F-150
    
    print("Testing Enhanced Commercial Vehicle Provider")
    print("=" * 50)
    
    # Get enhanced data
    success, data, error = provider.get_vehicle_by_vin(test_vin)
    
    if success:
        print(f"✓ Successfully retrieved data for VIN: {test_vin}")
        print(f"  Make: {data.get('vehicle_id', {}).get('make')}")
        print(f"  Model: {data.get('vehicle_id', {}).get('model')}")
        print(f"  Year: {data.get('vehicle_id', {}).get('year')}")
        print(f"  GVWR: {data.get('vehicle_id', {}).get('gvwr_pounds')} lbs")
        print(f"  Commercial Category: {data.get('vehicle_id', {}).get('commercial_category')}")
        
        # Get commercial analysis
        analysis = provider.get_commercial_vehicle_analysis(test_vin)
        if analysis['success']:
            print("\nCommercial Vehicle Analysis:")
            print(f"  Classification: {analysis['analysis']['vehicle_classification']}")
            print(f"  Electrification Score: {analysis['analysis']['electrification_assessment']['score']}/100")
            print(f"  Suitability: {analysis['analysis']['electrification_assessment']['suitability']}")
            print(f"  Data Quality: {analysis['analysis']['data_quality']:.1f}%")
    else:
        print(f"✗ Failed to retrieve data: {error}")
    
    # Cleanup
    provider.cleanup()