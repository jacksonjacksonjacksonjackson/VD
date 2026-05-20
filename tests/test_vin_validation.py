"""Tests for VIN validation logic in utils.py."""

import pytest
from utils import validate_vin_detailed


class TestValidateVinDetailed:
    """Tests for validate_vin_detailed()."""

    # ── Valid VINs ─────────────────────────────────────────────────

    def test_valid_17_char_vin(self):
        valid, msg = validate_vin_detailed("1HGBH41JXMN109186")
        assert valid is True
        assert msg == ""

    def test_valid_vin_with_spaces_stripped(self):
        valid, msg = validate_vin_detailed("1HGB H41J XMN1 0918 6")
        assert valid is True

    def test_valid_vin_with_dashes_stripped(self):
        valid, msg = validate_vin_detailed("1HGBH-41JX-MN109-186")
        assert valid is True

    def test_valid_vin_lowercased(self):
        valid, msg = validate_vin_detailed("1hgbh41jxmn109186")
        assert valid is True

    # ── Empty / missing ───────────────────────────────────────────

    def test_empty_string(self):
        valid, msg = validate_vin_detailed("")
        assert valid is False
        assert "required" in msg.lower()

    def test_none_value(self):
        valid, msg = validate_vin_detailed(None)
        assert valid is False

    def test_whitespace_only(self):
        valid, msg = validate_vin_detailed("   ")
        assert valid is False

    # ── Wrong length ──────────────────────────────────────────────

    def test_too_short(self):
        valid, msg = validate_vin_detailed("1HGBH41JXM")
        assert valid is False
        assert "too short" in msg.lower()

    def test_too_long(self):
        valid, msg = validate_vin_detailed("1HGBH41JXMN109186X")
        assert valid is False
        assert "too long" in msg.lower()

    # ── Invalid characters (I, O, Q) ─────────────────────────────

    def test_contains_I(self):
        # Replace a valid char with 'I'
        valid, msg = validate_vin_detailed("1IGBH41JXMN109186")
        assert valid is False
        assert "I" in msg

    def test_contains_O(self):
        valid, msg = validate_vin_detailed("1HGBHO1JXMN109186")
        assert valid is False
        assert "O" in msg

    def test_contains_Q(self):
        valid, msg = validate_vin_detailed("1HGBH41QXMN109186")
        assert valid is False
        assert "Q" in msg

    # ── Placeholder / test data detection ─────────────────────────

    def test_all_zeros(self):
        valid, msg = validate_vin_detailed("00000000000000000")
        assert valid is False
        assert "test data" in msg.lower() or "placeholder" in msg.lower()

    def test_all_ones(self):
        valid, msg = validate_vin_detailed("11111111111111111")
        assert valid is False
        assert "test data" in msg.lower()

    def test_excessive_zeros(self):
        # 9+ zeros should flag as placeholder
        valid, msg = validate_vin_detailed("10000000000000002")
        assert valid is False
        assert "placeholder" in msg.lower()
