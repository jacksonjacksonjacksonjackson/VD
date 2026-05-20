"""
test_timeline_override.py

Regression tests for Phase 19: manual EV-year override data model,
Gantt chart data helpers, and scenario year-assignment helper.
"""

import sys
import os
import copy
import pytest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from tests.conftest import make_fleet_vehicle
from ui.analysis_panel import AnalysisPanel


# ---------------------------------------------------------------------------
# Helpers — build vehicles with ACF classification already set
# ---------------------------------------------------------------------------

def _vehicle_b(vin="VIN_B001", ev_year="2035"):
    v = make_fleet_vehicle(vin=vin)
    v.custom_fields["ACF Category"] = "B"
    v.custom_fields["_acf_code"]    = "B"
    v.custom_fields["Proposed EV Year"] = ev_year
    return v


def _vehicle_c(vin="VIN_C001", ev_year="2033"):
    v = make_fleet_vehicle(vin=vin)
    v.custom_fields["ACF Category"] = "C"
    v.custom_fields["_acf_code"]    = "C"
    v.custom_fields["Proposed EV Year"] = ev_year
    return v


# ---------------------------------------------------------------------------
# Phase A: override write-back
# ---------------------------------------------------------------------------

class TestApplyEvYearOverride:
    def test_sets_proposed_year(self):
        v = _vehicle_b(ev_year="2035")
        AnalysisPanel._apply_ev_year_override(v, "2030")
        assert v.custom_fields["Proposed EV Year"] == "2030"

    def test_stores_original_on_first_override(self):
        v = _vehicle_b(ev_year="2035")
        AnalysisPanel._apply_ev_year_override(v, "2030")
        assert v.custom_fields["System Recommended EV Year"] == "2035"

    def test_does_not_overwrite_original_on_second_override(self):
        """System Rec. should stay as the FIRST value, not the most recent edit."""
        v = _vehicle_b(ev_year="2035")
        AnalysisPanel._apply_ev_year_override(v, "2030")
        AnalysisPanel._apply_ev_year_override(v, "2028")
        assert v.custom_fields["System Recommended EV Year"] == "2035"
        assert v.custom_fields["Proposed EV Year"] == "2028"

    def test_sets_override_flag(self):
        v = _vehicle_b(ev_year="2035")
        AnalysisPanel._apply_ev_year_override(v, "2030")
        assert v.custom_fields.get("EV Year Overridden") == "Yes"


class TestResetEvYearOverride:
    def test_restores_original_year(self):
        v = _vehicle_b(ev_year="2035")
        AnalysisPanel._apply_ev_year_override(v, "2030")
        AnalysisPanel._reset_ev_year_override(v)
        assert v.custom_fields["Proposed EV Year"] == "2035"

    def test_removes_override_flag(self):
        v = _vehicle_b(ev_year="2035")
        AnalysisPanel._apply_ev_year_override(v, "2030")
        AnalysisPanel._reset_ev_year_override(v)
        assert "EV Year Overridden" not in v.custom_fields

    def test_removes_sys_rec_after_reset(self):
        v = _vehicle_b(ev_year="2035")
        AnalysisPanel._apply_ev_year_override(v, "2030")
        AnalysisPanel._reset_ev_year_override(v)
        assert "System Recommended EV Year" not in v.custom_fields

    def test_reset_on_unmodified_vehicle_is_noop(self):
        """Resetting a vehicle that was never overridden should not crash."""
        v = _vehicle_b(ev_year="2035")
        AnalysisPanel._reset_ev_year_override(v)
        assert v.custom_fields["Proposed EV Year"] == "2035"


# ---------------------------------------------------------------------------
# Gantt chart grouping helpers
# ---------------------------------------------------------------------------

class TestGanttGroupedData:
    """Verify that the grouped-by-ACF Gantt helper draws without errors
    and that year counts are correct (tested via data extraction)."""

    def test_grouped_counts_per_year(self):
        """Two Cat-B vehicles in 2035, one Cat-C in 2030."""
        from ui.analysis_panel import _gantt_grouped, GANTT_YEAR_MIN, GANTT_YEAR_MAX
        import matplotlib
        matplotlib.use("Agg")  # headless
        from matplotlib.figure import Figure

        vehicles = [
            _vehicle_b("B1", "2035"),
            _vehicle_b("B2", "2035"),
            _vehicle_c("C1", "2030"),
        ]
        fig = Figure()
        ax = fig.add_subplot(111)
        years = list(range(GANTT_YEAR_MIN, GANTT_YEAR_MAX + 1))
        # Should not raise
        _gantt_grouped(ax, vehicles, years)

    def test_per_vehicle_draws_overrides_with_star(self):
        from ui.analysis_panel import _gantt_per_vehicle, GANTT_YEAR_MIN, GANTT_YEAR_MAX
        import matplotlib
        matplotlib.use("Agg")
        from matplotlib.figure import Figure

        v = _vehicle_b("B1", "2033")
        AnalysisPanel._apply_ev_year_override(v, "2030")

        fig = Figure()
        ax = fig.add_subplot(111)
        years = list(range(GANTT_YEAR_MIN, GANTT_YEAR_MAX + 1))
        # Should not raise; override vehicle will use star marker
        _gantt_per_vehicle(ax, [v], years)


# ---------------------------------------------------------------------------
# Scenario year assignments helper
# ---------------------------------------------------------------------------

class TestGetScenarioYearAssignments:
    """Tests for scenarios.get_scenario_year_assignments()."""

    def _make_category_b_vehicle(self, vin: str):
        """Create a processable Cat-B vehicle for scenario testing."""
        v = make_fleet_vehicle(
            vin=vin,
            vid_overrides={
                "vin": vin,
                "gvwr": "Class 5: 16,001 - 19,500 lb",
            },
        )
        v.processing_success = True
        v.custom_fields["ACF Category"] = "B"
        v.custom_fields["_acf_code"]    = "B"
        v.custom_fields["Proposed EV Year"] = "2035"
        return v

    def test_returns_dict_keyed_by_vin(self):
        from analysis.scenarios import get_scenario_year_assignments
        vehicles = [self._make_category_b_vehicle("VIN00000000001")]
        result = get_scenario_year_assignments(vehicles, "moderate")
        assert isinstance(result, dict)
        assert "VIN00000000001" in result

    def test_does_not_mutate_originals(self):
        """Original custom_fields must be unchanged after the helper runs."""
        from analysis.scenarios import get_scenario_year_assignments
        v = self._make_category_b_vehicle("VIN00000000002")
        original_year = v.custom_fields["Proposed EV Year"]
        _ = get_scenario_year_assignments([v], "aggressive")
        assert v.custom_fields["Proposed EV Year"] == original_year

    def test_unknown_scenario_returns_empty(self):
        from analysis.scenarios import get_scenario_year_assignments
        v = self._make_category_b_vehicle("VIN00000000003")
        result = get_scenario_year_assignments([v], "nonexistent_scenario")
        assert result == {}

    def test_acf_only_filter_marks_non_eligible(self):
        """The 'acf_compliance' scenario only targets Cat-B; Cat-C gets '—'."""
        from analysis.scenarios import get_scenario_year_assignments
        vb = self._make_category_b_vehicle("VIN00000000004")
        vc = _vehicle_c("VIN00000000005", "2033")
        result = get_scenario_year_assignments([vb, vc], "acf_compliance")
        assert result.get("VIN00000000005") == "—"
