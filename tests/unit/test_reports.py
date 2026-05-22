"""Unit tests for `bridge/reports.py::ExecutiveReportBuilder`.

Verifies the layout produces the expected cell values + sheet
replacement behaviour. COM-level formatting (colours, font sizes,
chart rendering) can't be unit-tested without real Excel; we rely on
the real-Excel test for that. What we cover here:

- Sheet creation + replacement (idempotent re-runs)
- Title / subtitle land in the right cells
- Headline numbers land in the right cells with the right values
- Stats table rows are correct
- Callouts are generated from the data and land below the stats
- Adaptive layout (charts succeed even when samples list is empty)
- The MCP-tool passthrough wiring
"""

from __future__ import annotations

from typing import Any
from unittest.mock import MagicMock

import pytest

from modelrisk_mcp.bridge.reports import (
    DriversReportBuilder,
    DriversReportResult,
    ExecutiveReportBuilder,
    ExecutiveReportResult,
    default_subtitle,
)
from modelrisk_mcp.schemas.results import (
    SensitivityEntry,
    SensitivityRanking,
    SimulationResult,
)

# ----------------------------------------------------------------------
# Fake xlwings shapes (reuse the pattern from test_charts.py)
# ----------------------------------------------------------------------


class _FakeRange:
    def __init__(self, sheet: _FakeSheet, ref: str) -> None:
        self._sheet = sheet
        self._ref = ref
        self._value: Any = None

    @property
    def value(self) -> Any:
        return self._sheet.cells.get(self._ref)

    @value.setter
    def value(self, v: Any) -> None:
        self._sheet.cells[self._ref] = v

    @property
    def api(self) -> Any:
        return MagicMock()

    @property
    def column_width(self) -> int:
        return self._sheet.col_widths.get(self._ref[0], 0)

    @column_width.setter
    def column_width(self, w: int) -> None:
        self._sheet.col_widths[self._ref[0]] = w

    @property
    def row_height(self) -> int:
        return 0

    @row_height.setter
    def row_height(self, h: int) -> None:
        pass

    def merge(self) -> None:
        pass


class _FakeChart:
    def __init__(self) -> None:
        self.source_range: str | None = None
        self.name = "Chart"
        self.api: Any = (MagicMock(), MagicMock())

    def set_source_data(self, rng: Any) -> None:
        self.source_range = rng._ref


class _FakeChartsCollection:
    def __init__(self) -> None:
        self.added: list[_FakeChart] = []

    def add(self, *, left: int, top: int, width: int, height: int) -> _FakeChart:
        c = _FakeChart()
        self.added.append(c)
        return c


class _FakeSheet:
    def __init__(self, name: str, parent: _FakeSheets) -> None:
        self.name = name
        self._parent = parent
        self.cells: dict[str, Any] = {}
        self.col_widths: dict[str, int] = {}
        self.charts = _FakeChartsCollection()
        self.deleted = False
        self.activated = False

    def range(self, ref: str) -> _FakeRange:
        return _FakeRange(self, ref)

    def delete(self) -> None:
        self.deleted = True
        if self in self._parent._sheets:
            self._parent._sheets.remove(self)

    def activate(self) -> None:
        self.activated = True


class _FakeSheets:
    def __init__(self) -> None:
        self._sheets: list[_FakeSheet] = []
        self._sheets.append(_FakeSheet("Sheet1", self))

    def __iter__(self) -> Any:
        return iter(self._sheets)

    def __getitem__(self, key: int | str) -> _FakeSheet:
        if isinstance(key, int):
            return self._sheets[key]
        for s in self._sheets:
            if s.name == key:
                return s
        raise KeyError(key)

    def add(self, name: str, after: _FakeSheet | None = None) -> _FakeSheet:
        s = _FakeSheet(name, self)
        self._sheets.append(s)
        return s


class _FakeBook:
    def __init__(self) -> None:
        self.sheets = _FakeSheets()


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------


def _make_result(
    name: str = "Profit",
    mean: float = 1000.0,
    stdev: float = 200.0,
    iterations: int = 5000,
) -> SimulationResult:
    return SimulationResult(
        output_name=name,
        iterations=iterations,
        mean=mean,
        stdev=stdev,
        variance=stdev**2,
        skewness=0.0,
        kurtosis=0.0,
        min=mean - 3 * stdev,
        max=mean + 3 * stdev,
        percentiles={
            0.05: mean - 1.645 * stdev,
            0.10: mean - 1.282 * stdev,
            0.25: mean - 0.674 * stdev,
            0.50: mean,
            0.75: mean + 0.674 * stdev,
            0.90: mean + 1.282 * stdev,
            0.95: mean + 1.645 * stdev,
        },
    )


def _make_sensitivity(
    output: str = "Profit",
    entries: list[tuple[str, float]] | None = None,
) -> SensitivityRanking:
    entries = entries or [("DemandUnits", 0.75), ("UnitCost", -0.42), ("Discount", 0.15)]
    return SensitivityRanking(
        output_name=output,
        entries=[
            SensitivityEntry(
                input_name=n, correlation=c, regression_coefficient=c * 0.9,
            )
            for n, c in entries
        ],
        iterations=1000,
    )


# ----------------------------------------------------------------------
# ExecutiveReportBuilder tests
# ----------------------------------------------------------------------


class TestExecutiveReportBuilder:
    def test_creates_sheet_with_title_and_subtitle(self) -> None:
        book = _FakeBook()
        result = ExecutiveReportBuilder.build(
            book,
            sheet_name="Report",
            title="NPV under uncertainty",
            subtitle="5,000 iterations · seed 1 · 2026-05-22",
            primary_output="Profit",
            primary_result=_make_result(),
            primary_samples=[1000.0] * 100,
            sensitivity=_make_sensitivity(),
        )
        assert isinstance(result, ExecutiveReportResult)
        sheet = book.sheets["Report"]
        assert sheet.cells.get("A1") == "NPV under uncertainty"
        assert sheet.cells.get("A2") == "5,000 iterations · seed 1 · 2026-05-22"
        assert sheet.activated is True

    def test_headline_numbers_in_expected_cells(self) -> None:
        book = _FakeBook()
        result_data = _make_result(mean=1000.0, stdev=200.0)
        ExecutiveReportBuilder.build(
            book,
            sheet_name="Report",
            title="t", subtitle="s",
            primary_output="Profit",
            primary_result=result_data,
            primary_samples=[],
            sensitivity=_make_sensitivity(),
        )
        sheet = book.sheets["Report"]
        assert sheet.cells.get("A6") == 1000.0  # Mean
        assert sheet.cells.get("C6") == 1000.0  # P50 = mean for symmetric
        assert sheet.cells.get("E6") == 200.0   # Stdev
        # Headline label row
        assert sheet.cells.get("A5") == "MEAN"
        assert sheet.cells.get("E5") == "STDEV"

    def test_stats_table_includes_secondary_outputs(self) -> None:
        book = _FakeBook()
        primary = _make_result("Profit", 1000.0, 200.0)
        secondary = [
            _make_result("Cost", 500.0, 50.0),
            _make_result("Revenue", 1500.0, 250.0),
        ]
        ExecutiveReportBuilder.build(
            book,
            sheet_name="Report",
            title="t", subtitle="s",
            primary_output="Profit",
            primary_result=primary,
            primary_samples=[],
            sensitivity=_make_sensitivity(),
            secondary_results=secondary,
        )
        sheet = book.sheets["Report"]
        # Stats table starts at row 26 (header), data at 27+
        assert sheet.cells.get("A26") == "Output"
        assert sheet.cells.get("A27") == "Profit"      # primary first
        assert sheet.cells.get("A28") == "Cost"
        assert sheet.cells.get("A29") == "Revenue"

    def test_callouts_generated_from_data(self) -> None:
        book = _FakeBook()
        # CV = 200/1000 = 0.2 → moderate volatility callout
        primary = _make_result("Profit", mean=1000.0, stdev=200.0)
        ExecutiveReportBuilder.build(
            book,
            sheet_name="Report",
            title="t", subtitle="s",
            primary_output="Profit",
            primary_result=primary,
            primary_samples=[],
            sensitivity=_make_sensitivity(),
        )
        sheet = book.sheets["Report"]
        # Callouts should appear below the stats table.
        # Stats table ends at row 27 (one data row), plus padding (2) = row 29
        # CALLOUTS header on first callout row.
        callout_cells = [
            v for k, v in sheet.cells.items()
            if k.startswith("A") and isinstance(v, str)
            and v.startswith("•")
        ]
        assert len(callout_cells) >= 2
        # Should mention "90% confidence" in one
        assert any("90% confidence" in c for c in callout_cells)
        # Should mention the top driver
        assert any("DemandUnits" in c for c in callout_cells)

    def test_replaces_existing_sheet(self) -> None:
        book = _FakeBook()
        book.sheets.add("Report")
        existing = book.sheets["Report"]
        existing.cells["A1"] = "stale"

        ExecutiveReportBuilder.build(
            book,
            sheet_name="Report",
            title="fresh", subtitle="s",
            primary_output="Profit",
            primary_result=_make_result(),
            primary_samples=[],
            sensitivity=_make_sensitivity(),
        )
        # Old sheet deleted.
        assert existing.deleted is True
        # New sheet has the fresh content.
        new_sheet = book.sheets["Report"]
        assert new_sheet is not existing
        assert new_sheet.cells.get("A1") == "fresh"

    def test_empty_samples_doesnt_crash(self) -> None:
        """If get_samples returned [] (no .vmrs / empty result), the
        histogram chart silently fails but the rest of the report
        still renders."""
        book = _FakeBook()
        result = ExecutiveReportBuilder.build(
            book,
            sheet_name="Report",
            title="t", subtitle="s",
            primary_output="Profit",
            primary_result=_make_result(),
            primary_samples=[],
            sensitivity=_make_sensitivity(),
        )
        # Histogram chart skipped (0); tornado succeeds (1).
        assert result.chart_count == 1

    def test_empty_sensitivity_skips_tornado(self) -> None:
        book = _FakeBook()
        empty_sens = SensitivityRanking(output_name="Profit", entries=[])
        result = ExecutiveReportBuilder.build(
            book,
            sheet_name="Report",
            title="t", subtitle="s",
            primary_output="Profit",
            primary_result=_make_result(),
            primary_samples=[100.0] * 50,
            sensitivity=empty_sens,
        )
        # Histogram succeeds, tornado skipped.
        assert result.chart_count == 1

    def test_high_cv_triggers_volatility_callout(self) -> None:
        book = _FakeBook()
        # CV = 800/1000 = 0.8 → high-volatility callout
        primary = _make_result("Profit", mean=1000.0, stdev=800.0)
        ExecutiveReportBuilder.build(
            book,
            sheet_name="Report",
            title="t", subtitle="s",
            primary_output="Profit",
            primary_result=primary,
            primary_samples=[],
            sensitivity=_make_sensitivity(),
        )
        sheet = book.sheets["Report"]
        callout_strings = [
            v for v in sheet.cells.values()
            if isinstance(v, str) and v.startswith("•")
        ]
        assert any("High volatility" in c for c in callout_strings)

    def test_headline_summary_one_liner(self) -> None:
        book = _FakeBook()
        result = ExecutiveReportBuilder.build(
            book,
            sheet_name="Report",
            title="t", subtitle="s",
            primary_output="Profit",
            primary_result=_make_result(mean=1000.0, stdev=200.0),
            primary_samples=[],
            sensitivity=_make_sensitivity(),
            contingency_percentile=0.90,
        )
        assert "Profit" in result.headline_summary
        assert "1,000.00" in result.headline_summary
        assert "P90" in result.headline_summary


# ----------------------------------------------------------------------
# default_subtitle
# ----------------------------------------------------------------------


class TestDefaultSubtitle:
    def test_includes_iteration_count(self) -> None:
        assert "5,000 iterations" in default_subtitle(5000)

    def test_includes_seed_when_provided(self) -> None:
        s = default_subtitle(1000, seed=42)
        assert "seed 42" in s

    def test_omits_seed_when_none(self) -> None:
        s = default_subtitle(1000)
        assert "seed" not in s

    def test_includes_date(self) -> None:
        import re
        s = default_subtitle(1000)
        # YYYY-MM-DD pattern
        assert re.search(r"\d{4}-\d{2}-\d{2}", s) is not None


# ----------------------------------------------------------------------
# MCP-tool passthrough
# ----------------------------------------------------------------------


class TestBuildExecutiveReportTool:
    def test_passes_args_through(self) -> None:
        from modelrisk_mcp.tools import reading, workflows

        bridge = MagicMock()
        bridge.build_executive_report.return_value = ExecutiveReportResult(
            sheet_name="Report",
            primary_output="Profit",
            secondary_outputs=("Cost",),
            chart_count=2,
            callout_count=4,
            headline_summary="Profit: mean 1,000.00",
        )
        reading.set_bridge_for_testing(bridge)
        try:
            out = workflows.build_executive_report(
                primary_output="Profit",
                title="My Report",
                secondary_outputs=["Cost"],
                contingency_percentile=0.95,
                top_drivers=3,
                workbook_name="m.xlsx",
            )
            bridge.build_executive_report.assert_called_once_with(
                "Profit",
                workbook="m.xlsx",
                title="My Report",
                subtitle=None,
                secondary_outputs=["Cost"],
                contingency_percentile=0.95,
                top_drivers=3,
                sheet_name="Executive_Report",
            )
            assert out["sheet_name"] == "Report"
            assert out["secondary_outputs"] == ["Cost"]
            assert out["chart_count"] == 2
            assert out["callout_count"] == 4
            assert "1,000.00" in out["headline_summary"]
        finally:
            reading.set_bridge_for_testing(None)


# ----------------------------------------------------------------------
# DriversReportBuilder tests
# ----------------------------------------------------------------------


class TestDriversReportBuilder:
    def test_builds_sheet_with_title_and_findings(self) -> None:
        book = _FakeBook()
        sens = _make_sensitivity(
            "Profit",
            entries=[
                ("WidgetCost", -0.72),
                ("UnitsSold", 0.55),
                ("Discount", 0.18),
            ],
        )
        result = DriversReportBuilder.build(
            book,
            sheet_name="Drivers",
            output_name="Profit",
            sensitivity=sens,
            iterations=5000,
        )
        assert isinstance(result, DriversReportResult)
        sheet = book.sheets["Drivers"]
        assert "Profit" in sheet.cells["A1"]
        assert "5,000 iterations" in sheet.cells["A2"]
        assert sheet.cells["A4"] == "KEY FINDINGS"
        assert result.drivers_analyzed == 3
        assert result.top_driver == "WidgetCost"
        assert result.top_correlation == -0.72
        assert sheet.activated is True

    def test_findings_name_top_driver_with_direction(self) -> None:
        book = _FakeBook()
        sens = _make_sensitivity(
            "Profit", entries=[("WidgetCost", -0.72)],
        )
        DriversReportBuilder.build(
            book,
            sheet_name="Drivers",
            output_name="Profit",
            sensitivity=sens,
            iterations=1000,
        )
        sheet = book.sheets["Drivers"]
        # First finding should mention the driver name + that it
        # LOWERS the output (negative correlation).
        first_finding_text = sheet.cells.get("A5", "") or ""
        assert "WidgetCost" in first_finding_text
        assert "lowers" in first_finding_text
        assert "Profit" in first_finding_text

    def test_findings_mention_top_n_variance_share(self) -> None:
        book = _FakeBook()
        # Two strong drivers — top-2 r² = 0.5184 + 0.4225 ≈ 0.94 → ~94%
        sens = _make_sensitivity(
            "Profit",
            entries=[("A", -0.72), ("B", 0.65), ("C", 0.05)],
        )
        DriversReportBuilder.build(
            book,
            sheet_name="Drivers",
            output_name="Profit",
            sensitivity=sens,
            iterations=1000,
        )
        sheet = book.sheets["Drivers"]
        coverage_finding = sheet.cells.get("A6", "") or ""
        assert "top 3" in coverage_finding.lower() or "top-3" in coverage_finding.lower() or "top " in coverage_finding.lower()

    def test_driver_table_populated(self) -> None:
        book = _FakeBook()
        sens = _make_sensitivity(
            "Profit",
            entries=[("A", -0.7), ("B", 0.4), ("C", 0.1)],
        )
        DriversReportBuilder.build(
            book,
            sheet_name="Drivers",
            output_name="Profit",
            sensitivity=sens,
            iterations=1000,
        )
        sheet = book.sheets["Drivers"]
        # Header
        assert sheet.cells["G11"] == "Input"
        assert sheet.cells["H11"] == "Correlation (r)"
        assert sheet.cells["I11"] == "|r|"
        assert sheet.cells["J11"] == "Variance share"
        # Data row 1 (strongest)
        assert sheet.cells["G12"] == "A"
        assert sheet.cells["H12"] == -0.7
        assert abs(sheet.cells["I12"] - 0.7) < 1e-9
        assert abs(sheet.cells["J12"] - 0.49) < 1e-9  # r² = 0.49

    def test_recommendations_tier_by_correlation_strength(self) -> None:
        book = _FakeBook()
        sens = _make_sensitivity(
            "Profit",
            entries=[
                ("StrongDriver", -0.7),     # focus
                ("StrongDriver2", 0.5),     # focus
                ("ModerateDriver", 0.25),   # monitor
                ("WeakDriver", 0.05),       # deprioritise
            ],
        )
        DriversReportBuilder.build(
            book,
            sheet_name="Drivers",
            output_name="Profit",
            sensitivity=sens,
            iterations=1000,
        )
        sheet = book.sheets["Drivers"]
        # The "Focus mitigation on" row lives at RECOMMEND_HEADER_ROW + 1.
        focus_row = DriversReportBuilder.RECOMMEND_HEADER_ROW + 1
        monitor_row = focus_row + 1
        deprioritise_row = focus_row + 2
        assert "StrongDriver" in sheet.cells[f"B{focus_row}"]
        assert "StrongDriver2" in sheet.cells[f"B{focus_row}"]
        assert "ModerateDriver" in sheet.cells[f"B{monitor_row}"]
        assert "WeakDriver" in sheet.cells[f"B{deprioritise_row}"]

    def test_empty_sensitivity_doesnt_crash(self) -> None:
        book = _FakeBook()
        from modelrisk_mcp.schemas.results import SensitivityRanking
        empty = SensitivityRanking(output_name="Profit", entries=[])
        result = DriversReportBuilder.build(
            book,
            sheet_name="Drivers",
            output_name="Profit",
            sensitivity=empty,
            iterations=0,
        )
        assert result.drivers_analyzed == 0
        assert result.top_driver is None
        # A finding still gets written, explaining the empty result.
        sheet = book.sheets["Drivers"]
        first_finding = sheet.cells.get("A5", "")
        assert "No drivers" in first_finding

    def test_concentration_label(self) -> None:
        book = _FakeBook()
        # Concentrated: one input dominates
        sens = _make_sensitivity(
            "Profit",
            entries=[("A", -0.9), ("B", 0.1), ("C", 0.05)],
        )
        result = DriversReportBuilder.build(
            book,
            sheet_name="Drivers",
            output_name="Profit",
            sensitivity=sens,
            iterations=1000,
        )
        assert result.concentration == "concentrated"

    def test_replaces_existing_sheet(self) -> None:
        book = _FakeBook()
        book.sheets.add("Drivers")
        existing = book.sheets["Drivers"]
        existing.cells["A1"] = "stale"

        DriversReportBuilder.build(
            book,
            sheet_name="Drivers",
            output_name="Profit",
            sensitivity=_make_sensitivity(),
            iterations=1000,
        )
        assert existing.deleted is True


class TestBuildDriversReportTool:
    def test_passes_args_through(self) -> None:
        from modelrisk_mcp.tools import reading, workflows

        bridge = MagicMock()
        bridge.build_drivers_report.return_value = DriversReportResult(
            sheet_name="Drivers_Report",
            output_name="Profit",
            drivers_analyzed=4,
            top_driver="WidgetCost",
            top_correlation=-0.72,
            concentration="moderate",
            headline_finding=(
                "WidgetCost is the strong driver of Profit (r = -0.72); "
                "4 inputs analyzed."
            ),
        )
        reading.set_bridge_for_testing(bridge)
        try:
            out = workflows.build_drivers_report(
                output_name="Profit",
                title="What drives Profit",
                workbook_name="m.xlsx",
            )
            bridge.build_drivers_report.assert_called_once_with(
                "Profit",
                workbook="m.xlsx",
                title="What drives Profit",
                subtitle=None,
                sheet_name="Drivers_Report",
            )
            assert out["top_driver"] == "WidgetCost"
            assert out["concentration"] == "moderate"
            assert "Profit" in out["headline_finding"]
        finally:
            reading.set_bridge_for_testing(None)


_ = pytest  # keep import used
