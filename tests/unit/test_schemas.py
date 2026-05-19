"""Tests for the Pydantic schemas in `modelrisk_mcp.schemas`."""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from modelrisk_mcp.schemas import (
    CellRef,
    DistributionParameter,
    InsertDistributionRequest,
    RangeInfo,
    SimulationResult,
    WorkbookInfo,
)


class TestCellRef:
    def test_basic(self) -> None:
        ref = CellRef(workbook="book.xlsx", sheet="Sheet1", cell="B12")
        assert ref.cell == "B12"
        assert ref.a1 == "Sheet1!B12"

    def test_lowercase_cell_normalised(self) -> None:
        ref = CellRef(workbook="book", sheet="s", cell="b12")
        assert ref.cell == "B12"

    def test_invalid_cell_rejected(self) -> None:
        with pytest.raises(ValidationError):
            CellRef(workbook="book", sheet="s", cell="not-a-cell")

    def test_parse_with_sheet_prefix(self) -> None:
        ref = CellRef.parse(workbook="b", default_sheet="Sheet1", ref="MySheet!C7")
        assert ref.sheet == "MySheet"
        assert ref.cell == "C7"

    def test_parse_without_sheet_uses_default(self) -> None:
        ref = CellRef.parse(workbook="b", default_sheet="Sheet1", ref="C7")
        assert ref.sheet == "Sheet1"
        assert ref.cell == "C7"


class TestRangeInfo:
    def test_basic(self) -> None:
        ri = RangeInfo(workbook="b", sheet="s", range_ref="A1:B2")
        assert ri.range_ref == "A1:B2"

    def test_single_cell_range_accepted(self) -> None:
        ri = RangeInfo(workbook="b", sheet="s", range_ref="A1")
        assert ri.range_ref == "A1"

    def test_invalid_range_rejected(self) -> None:
        with pytest.raises(ValidationError):
            RangeInfo(workbook="b", sheet="s", range_ref="1A:2B")


class TestWorkbookInfo:
    def test_default_sheets_empty(self) -> None:
        wb = WorkbookInfo(name="book.xlsx", path="C:/book.xlsx")
        assert wb.sheets == []
        assert wb.active_sheet is None


class TestDistributionParameter:
    def test_number_value(self) -> None:
        p = DistributionParameter(name="mu", value=5)
        assert p.value == 5

    def test_string_value(self) -> None:
        p = DistributionParameter(name="mu", value="B12")
        assert p.value == "B12"

    def test_array_value(self) -> None:
        p = DistributionParameter(name="values", value=[1, 2, 3])
        assert p.value == [1, 2, 3]


class TestInsertDistributionRequest:
    def test_dry_run_defaults_true(self) -> None:
        """Spec §11.1 / §7.2 — building tools default to dry_run=True."""
        req = InsertDistributionRequest(
            cell=CellRef(workbook="b", sheet="s", cell="B1"),
            function_name="VoseNormal",
            parameters=[
                DistributionParameter(name="mu", value=0),
                DistributionParameter(name="sigma", value=1),
            ],
        )
        assert req.dry_run is True

    def test_duplicate_param_names_rejected(self) -> None:
        with pytest.raises(ValidationError):
            InsertDistributionRequest(
                cell=CellRef(workbook="b", sheet="s", cell="B1"),
                function_name="VoseNormal",
                parameters=[
                    DistributionParameter(name="mu", value=0),
                    DistributionParameter(name="mu", value=1),
                ],
            )


class TestSimulationResult:
    def test_basic(self) -> None:
        r = SimulationResult(
            output_name="Profit",
            iterations=10_000,
            mean=100.0,
            stdev=15.0,
            min=20.0,
            max=200.0,
            percentiles={0.05: 70.0, 0.95: 130.0},
        )
        assert r.iterations == 10_000
        assert r.percentiles[0.05] == 70.0

    def test_negative_iterations_rejected(self) -> None:
        with pytest.raises(ValidationError):
            SimulationResult(
                output_name="x",
                iterations=-1,
                mean=0.0,
                stdev=0.0,
                min=0.0,
                max=0.0,
            )
