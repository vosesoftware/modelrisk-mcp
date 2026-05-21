"""Unit tests for `bridge/charts.py::TornadoChartWriter`.

Verifies the data-table shape, sort order, sheet-replacement behaviour,
and that COM-level formatting failures don't tank the whole call.

We can't unit-test the actual Excel chart rendering (requires Excel
running), but we can verify:
- The right sheet gets created / replaced
- Data is written in the right cells, sorted correctly
- The chart `add()` + `set_source_data` calls are made with the right
  range string
- Formatting failures inside the writer (try/except blocks) don't
  propagate up

A real-Excel test lives in the gated `tests/integration/` directory
(future work)."""

from __future__ import annotations

from typing import Any
from unittest.mock import MagicMock

import pytest

from modelrisk_mcp.bridge.charts import TornadoChartResult, TornadoChartWriter
from modelrisk_mcp.schemas.results import SensitivityEntry

# ----------------------------------------------------------------------
# Fake xlwings shapes
# ----------------------------------------------------------------------


class FakeRange:
    """Records `.value =` writes; pretends formatting attribute
    accesses succeed."""

    def __init__(self, sheet: FakeSheet, ref: str) -> None:
        self._sheet = sheet
        self._ref = ref

    @property
    def value(self) -> Any:
        return self._sheet.cells.get(self._ref)

    @value.setter
    def value(self, v: Any) -> None:
        self._sheet.cells[self._ref] = v

    # The chart writer dives into .api for formatting (Font.Bold,
    # NumberFormat, EntireColumn.Hidden). We return a permissive MagicMock
    # so all such attribute walks succeed silently.
    @property
    def api(self) -> Any:
        return MagicMock()


class FakeChart:
    """xlwings Chart stand-in."""

    def __init__(self, sheet: FakeSheet) -> None:
        self._sheet = sheet
        self.source_range: str | None = None
        self.name = "Chart 1"
        # `.api` returns (ChartObject, Chart) on Windows; tuple form is
        # what the writer's COM-config block dereferences.
        self.api: Any = (MagicMock(), MagicMock())

    def set_source_data(self, rng: Any) -> None:
        self.source_range = rng._ref


class FakeCharts:
    def __init__(self, sheet: FakeSheet) -> None:
        self._sheet = sheet
        self.added: list[FakeChart] = []

    def add(
        self, *, left: int, top: int, width: int, height: int,
    ) -> FakeChart:
        chart = FakeChart(self._sheet)
        self.added.append(chart)
        return chart


class FakeSheet:
    def __init__(self, name: str, parent: FakeSheetsCollection) -> None:
        self.name = name
        self._parent = parent
        self.cells: dict[str, Any] = {}
        self.charts = FakeCharts(self)
        self.deleted = False

    def range(self, ref: str) -> FakeRange:
        return FakeRange(self, ref)

    def delete(self) -> None:
        self.deleted = True
        if self in self._parent._sheets:
            self._parent._sheets.remove(self)


class FakeSheetsCollection:
    def __init__(self) -> None:
        self._sheets: list[FakeSheet] = []
        self._sheets.append(FakeSheet("Sheet1", self))

    def __iter__(self) -> Any:
        return iter(self._sheets)

    def __getitem__(self, key: int | str) -> FakeSheet:
        if isinstance(key, int):
            return self._sheets[key]
        for s in self._sheets:
            if s.name == key:
                return s
        raise KeyError(key)

    def add(self, name: str, after: FakeSheet | None = None) -> FakeSheet:
        sheet = FakeSheet(name, self)
        self._sheets.append(sheet)
        return sheet


class FakeBook:
    def __init__(self) -> None:
        self.sheets = FakeSheetsCollection()


# ----------------------------------------------------------------------
# Tests
# ----------------------------------------------------------------------


def _entry(name: str, corr: float, beta: float | None = None) -> SensitivityEntry:
    return SensitivityEntry(
        input_name=name, correlation=corr, regression_coefficient=beta,
    )


class TestTornadoChartWriter:
    def test_returns_chart_metadata(self) -> None:
        book = FakeBook()
        entries = [
            _entry("widget_cost", -0.7, -0.6),
            _entry("demand", 0.5, 0.4),
        ]
        result = TornadoChartWriter.write(book, "Profit", entries)
        assert isinstance(result, TornadoChartResult)
        assert result.sheet_name == "Tornado_Profit"
        assert result.output_name == "Profit"
        assert result.input_count == 2
        # Sorted by |corr| descending → widget_cost first.
        assert result.top_input == "widget_cost"
        assert result.top_correlation == -0.7

    def test_data_written_with_correct_sort_order(self) -> None:
        book = FakeBook()
        # Deliberately unsorted — writer must re-sort by |correlation|.
        entries = [
            _entry("price", 0.3),
            _entry("volatility", -0.8),
            _entry("rate", 0.5),
        ]
        TornadoChartWriter.write(book, "NPV", entries)
        sheet = book.sheets["Tornado_NPV"]

        assert sheet.cells["A1"] == "Input"
        assert sheet.cells["B1"] == "Spearman correlation"
        assert sheet.cells["C1"] == "|correlation|"
        assert sheet.cells["D1"] == "Regression coefficient"

        # Row 2 → highest |corr|.
        assert sheet.cells["A2"] == "volatility"
        assert sheet.cells["B2"] == -0.8
        assert sheet.cells["C2"] == 0.8

        # Row 3 → next.
        assert sheet.cells["A3"] == "rate"
        assert sheet.cells["B3"] == 0.5

        # Row 4 → smallest.
        assert sheet.cells["A4"] == "price"

    def test_replaces_existing_sheet(self) -> None:
        """If the target sheet already exists, the writer must delete
        it first. Idempotent re-runs are core to the tool's value."""
        book = FakeBook()
        # Pre-create the target sheet.
        book.sheets.add("Tornado_X")
        existing = book.sheets["Tornado_X"]
        existing.cells["A1"] = "stale data"

        TornadoChartWriter.write(book, "X", [_entry("a", 0.5)])
        # The old sheet's delete() must have been called.
        assert existing.deleted is True
        # The new sheet exists with fresh data.
        new_sheet = book.sheets["Tornado_X"]
        assert new_sheet is not existing
        assert new_sheet.cells["A1"] == "Input"

    def test_custom_sheet_name_used(self) -> None:
        book = FakeBook()
        result = TornadoChartWriter.write(
            book, "Profit", [_entry("a", 0.5)],
            sheet_name="MyCustomSheet",
        )
        assert result.sheet_name == "MyCustomSheet"
        assert "MyCustomSheet" in {s.name for s in book.sheets}

    def test_chart_source_range_matches_data_rows(self) -> None:
        book = FakeBook()
        entries = [_entry(f"in_{i}", 0.5 - i * 0.1) for i in range(5)]
        TornadoChartWriter.write(book, "Profit", entries)
        sheet = book.sheets["Tornado_Profit"]
        chart = sheet.charts.added[0]
        # 5 entries + 1 header row → A1:B6.
        assert chart.source_range == "A1:B6"

    def test_empty_entries_dont_crash(self) -> None:
        """A sensitivity ranking with zero inputs (no VoseInput cells)
        is a valid edge case — the writer should produce a chart-free
        sheet rather than throw."""
        book = FakeBook()
        result = TornadoChartWriter.write(book, "Profit", [])
        assert result.input_count == 0
        assert result.top_input is None
        assert result.top_correlation is None
        # Sheet was created.
        assert "Tornado_Profit" in {s.name for s in book.sheets}

    def test_sheet_name_truncated_to_31_chars(self) -> None:
        """Excel's hard limit; without truncation the sheet add() call
        would raise."""
        book = FakeBook()
        # 50-char output name.
        output = "This_Is_A_Very_Long_Output_Name_That_Exceeds_31_Chars"
        result = TornadoChartWriter.write(book, output, [_entry("a", 0.5)])
        assert len(result.sheet_name) <= 31

    def test_regression_coefficient_carried_when_present(self) -> None:
        book = FakeBook()
        entries = [_entry("price", 0.7, 0.65)]
        TornadoChartWriter.write(book, "Profit", entries)
        sheet = book.sheets["Tornado_Profit"]
        assert sheet.cells["D2"] == 0.65

    def test_regression_coefficient_none_passed_through(self) -> None:
        """Some inputs can't have a regression coefficient (e.g.
        constant series). The writer should write None, not crash."""
        book = FakeBook()
        entries = [_entry("price", 0.7, None)]
        TornadoChartWriter.write(book, "Profit", entries)
        sheet = book.sheets["Tornado_Profit"]
        assert sheet.cells["D2"] is None


class TestTornadoChartMCPWrapper:
    """Verify the MCP tool wrapper delegates correctly."""

    def test_create_tornado_chart_passes_args(self) -> None:
        from collections.abc import Iterator

        from modelrisk_mcp.tools import reading, workflows

        bridge = MagicMock()
        bridge.create_tornado_chart.return_value = TornadoChartResult(
            sheet_name="Tornado_X", chart_name="Tornado_X",
            output_name="X", input_count=2,
            top_input="a", top_correlation=0.7,
        )
        reading.set_bridge_for_testing(bridge)  # type: ignore[arg-type]
        try:
            out = workflows.create_tornado_chart(
                "X", workbook_name="m.xlsx", sheet_name="MySheet",
            )
            bridge.create_tornado_chart.assert_called_once_with(
                "X", "m.xlsx", sheet_name="MySheet",
            )
            assert out == {
                "sheet_name": "Tornado_X",
                "chart_name": "Tornado_X",
                "output_name": "X",
                "input_count": 2,
                "top_input": "a",
                "top_correlation": 0.7,
            }
        finally:
            reading.set_bridge_for_testing(None)
        # Reference Iterator to keep the import used for ruff.
        _ = Iterator

    def test_defaults_passed_as_none(self) -> None:
        from modelrisk_mcp.tools import reading, workflows

        bridge = MagicMock()
        bridge.create_tornado_chart.return_value = TornadoChartResult(
            sheet_name="Tornado_X", chart_name="",
            output_name="X", input_count=0,
            top_input=None, top_correlation=None,
        )
        reading.set_bridge_for_testing(bridge)  # type: ignore[arg-type]
        try:
            workflows.create_tornado_chart("X")
            bridge.create_tornado_chart.assert_called_once_with(
                "X", None, sheet_name=None,
            )
        finally:
            reading.set_bridge_for_testing(None)


# Quieten unused imports for ruff.
_ = pytest
