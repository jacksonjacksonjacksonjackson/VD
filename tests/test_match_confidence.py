"""
test_match_confidence.py

Tests for VehicleDataProvider._calculate_match_confidence().
Verifies that scoring uses MATCHING_WEIGHTS from settings and that
individual signal components contribute the right points.
"""

import pytest
from unittest.mock import MagicMock

from data.providers import VehicleDataProvider
from data.models import VehicleIdentification
from settings import MATCHING_WEIGHTS, MIN_MATCH_CONFIDENCE


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _provider():
    """Return a VehicleDataProvider with caching disabled (no disk I/O)."""
    p = VehicleDataProvider.__new__(VehicleDataProvider)
    return p


def _vid(**kw) -> VehicleIdentification:
    defaults = dict(
        vin="1HGBH41JXMN109186",
        year="2020",
        make="Ford",
        model="F-150",
        fuel_type="Gasoline",
        body_class="Pickup",
        gvwr="Class 2: 6,001 - 10,000 lb (2,722 - 4,536 kg)",
    )
    defaults.update(kw)
    return VehicleIdentification(**defaults)


def _fe(**kw) -> dict:
    """Minimal fuel economy data dict."""
    defaults = dict(
        year="2020",
        make="Ford",
        model="F-150",
        displ="3.5",
        cylinders="6",
        fuelType1="Gasoline",
        drive="4-Wheel Drive",
        trany="Automatic (S10)",
    )
    defaults.update(kw)
    return defaults


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class TestMatchConfidenceWeights:

    def test_full_year_make_model_match_awards_year_make_model_weight(self):
        provider = _provider()
        vid = _vid(year="2020", make="Ford", model="F-150")
        fe = _fe(year="2020", make="Ford", model="F-150")
        score = provider._calculate_match_confidence(vid, fe, "2020 Ford F-150")
        assert score >= MATCHING_WEIGHTS["year_make_model"]

    def test_year_make_only_awards_year_make_weight(self):
        provider = _provider()
        vid = _vid(year="2020", make="Ford", model="Maverick")
        # model does NOT appear in fe model string
        fe = _fe(year="2020", make="Ford", model="F-150")
        score = provider._calculate_match_confidence(vid, fe, "2020 Ford F-150")
        # Should get year_make but not year_make_model
        expected_min = MATCHING_WEIGHTS["year_make"]
        assert score >= expected_min
        assert score < MATCHING_WEIGHTS["year_make_model"]

    def test_no_year_make_model_overlap_scores_below_threshold(self):
        # Year, make, and model all miss — score stays well below the
        # year_make_model bucket (80), whatever component signals remain.
        provider = _provider()
        vid = _vid(year="2018", make="Chevy", model="Silverado",
                   engine_cylinders="", fuel_type="", drive_type="",
                   transmission="", engine_displacement="")
        fe = _fe(year="2020", make="Ford", model="F-150",
                 cylinders="", fuelType1="", drive="", trany="", displ="")
        score = provider._calculate_match_confidence(vid, fe, "2020 Ford F-150")
        assert score == 0.0

    def test_displacement_match_adds_weight(self):
        provider = _provider()
        vid = _vid(engine_displacement="3.5")
        fe = _fe(displ="3.5")
        score_with = provider._calculate_match_confidence(vid, fe, "")
        vid_no_disp = _vid(engine_displacement="")
        score_without = provider._calculate_match_confidence(vid_no_disp, fe, "")
        assert score_with > score_without

    def test_displacement_mismatch_does_not_add_weight(self):
        provider = _provider()
        vid = _vid(engine_displacement="5.0")
        fe = _fe(displ="3.5")
        # Only year+make+model bonus, no displacement bonus
        vid_no_disp = _vid(engine_displacement="")
        score_mismatch = provider._calculate_match_confidence(vid, fe, "")
        score_no_disp = provider._calculate_match_confidence(vid_no_disp, fe, "")
        assert score_mismatch == score_no_disp

    def test_cylinders_match_adds_weight(self):
        provider = _provider()
        vid = _vid(engine_cylinders="6")
        fe = _fe(cylinders="6")
        score_with = provider._calculate_match_confidence(vid, fe, "")
        vid_no_cyl = _vid(engine_cylinders="")
        score_without = provider._calculate_match_confidence(vid_no_cyl, fe, "")
        assert score_with >= score_without + MATCHING_WEIGHTS["cylinders_match"]

    def test_fuel_type_match_adds_weight(self):
        provider = _provider()
        vid = _vid(fuel_type="Gasoline")
        fe = _fe(fuelType1="Gasoline")
        score_with = provider._calculate_match_confidence(vid, fe, "")
        vid_no_fuel = _vid(fuel_type="")
        score_without = provider._calculate_match_confidence(vid_no_fuel, fe, "")
        assert score_with >= score_without + MATCHING_WEIGHTS["fuel_type_match"]

    def test_drive_type_match_adds_weight(self):
        provider = _provider()
        vid = _vid(drive_type="4-Wheel Drive")
        fe = _fe(drive="4-Wheel Drive")
        score_with = provider._calculate_match_confidence(vid, fe, "")
        vid_no_drive = _vid(drive_type="")
        score_without = provider._calculate_match_confidence(vid_no_drive, fe, "")
        assert score_with >= score_without + MATCHING_WEIGHTS["drive_match"]

    def test_transmission_match_adds_weight(self):
        provider = _provider()
        vid = _vid(transmission="Automatic")
        fe = _fe(trany="Automatic (S10)")
        score_with = provider._calculate_match_confidence(vid, fe, "")
        vid_no_trany = _vid(transmission="")
        score_without = provider._calculate_match_confidence(vid_no_trany, fe, "")
        assert score_with >= score_without + MATCHING_WEIGHTS["transmission_match"]

    def test_score_capped_at_100(self):
        provider = _provider()
        vid = _vid(
            year="2020", make="Ford", model="F-150",
            engine_displacement="3.5", engine_cylinders="6",
            fuel_type="Gasoline", drive_type="4-Wheel Drive",
            transmission="Automatic",
        )
        fe = _fe()
        score = provider._calculate_match_confidence(vid, fe, "2020 Ford F-150 3.5L 6-cyl")
        assert score <= 100.0

    def test_score_non_negative(self):
        provider = _provider()
        vid = _vid(year="1999", make="Unknown", model="Unknown")
        fe = _fe(year="2020", make="Ford", model="F-150")
        score = provider._calculate_match_confidence(vid, fe, "")
        assert score >= 0.0

    def test_min_match_confidence_is_positive_integer(self):
        """Sanity check that the settings value is usable."""
        assert isinstance(MIN_MATCH_CONFIDENCE, (int, float))
        assert MIN_MATCH_CONFIDENCE > 0
