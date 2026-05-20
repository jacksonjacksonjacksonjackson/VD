"""Tests for vehicle model normalization in utils.py."""

import pytest
from utils import normalize_vehicle_model


class TestNormalizeVehicleModel:
    """Tests for normalize_vehicle_model()."""

    # ── Ford F-series ─────────────────────────────────────────────

    def test_f150_with_dash(self):
        assert normalize_vehicle_model("F-150") == "f150"

    def test_f150_with_space(self):
        assert normalize_vehicle_model("F 150") == "f150"

    def test_f150_no_separator(self):
        assert normalize_vehicle_model("F150") == "f150"

    def test_f250_normalised(self):
        result = normalize_vehicle_model("F-250 Super Duty")
        assert "f250" in result
        assert "super" not in result  # "super duty" removed

    def test_f350_normalised(self):
        assert normalize_vehicle_model("F-350") == "f350"

    # ── Ford E-series ─────────────────────────────────────────────

    def test_e350_with_dash(self):
        assert normalize_vehicle_model("E-350") == "e350"

    def test_e250_with_space(self):
        assert normalize_vehicle_model("E 250") == "e250"

    # ── Chevrolet / GMC ───────────────────────────────────────────

    def test_silverado_1500(self):
        result = normalize_vehicle_model("Silverado 1500")
        assert "silverado1500" in result

    def test_sierra_2500(self):
        result = normalize_vehicle_model("Sierra-2500")
        assert "sierra2500" in result

    # ── Ram ────────────────────────────────────────────────────────

    def test_ram_1500(self):
        result = normalize_vehicle_model("Ram 1500")
        assert "ram1500" in result

    def test_ram_3500(self):
        result = normalize_vehicle_model("Ram-3500")
        assert "ram3500" in result

    # ── Express / Savana ──────────────────────────────────────────

    def test_express_2500(self):
        result = normalize_vehicle_model("Express 2500")
        assert "express2500" in result

    def test_savana_3500(self):
        result = normalize_vehicle_model("Savana-3500")
        assert "savana3500" in result

    # ── General normalisation ─────────────────────────────────────

    def test_lowercase(self):
        result = normalize_vehicle_model("TRANSIT")
        assert result == "transit"

    def test_strips_whitespace(self):
        result = normalize_vehicle_model("  F-150  ")
        assert result == "f150"

    def test_special_chars_removed(self):
        result = normalize_vehicle_model("F-150 (4WD)")
        assert "(" not in result
        assert ")" not in result
