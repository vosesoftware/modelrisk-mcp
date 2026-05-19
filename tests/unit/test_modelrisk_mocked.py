"""Tests for `ModelRiskBridge` read-only methods.

`ExcelBridge` is mocked entirely; we feed `iterate_cells` a synthetic
sequence of CellInfo objects and confirm the bridge correctly extracts
inputs, outputs, distribution cells, and hard-coded numeric inputs.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import Any

import pytest

from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.schemas.workbook import CellInfo, CellRef


def make_cell(
    cell: str,
    *,
    sheet: str = "Sheet1",
    workbook: str = "book.xlsx",
    formula: str = "",
    value: Any = None,
) -> CellInfo:
    return CellInfo(
        ref=CellRef(workbook=workbook, sheet=sheet, cell=cell),
        formula=formula,
        value=value,
        cell_type=(
            "formula" if formula else (
                "number" if isinstance(value, (int, float))
                and not isinstance(value, bool)
                else ("empty" if value is None else "text")
            )
        ),
    )


class FakeExcelBridge:
    """Stand-in for `ExcelBridge` that yields a fixed list of cells."""

    def __init__(self, cells: list[CellInfo]) -> None:
        self._cells = cells

    def iterate_cells(
        self,
        workbook: str,
        predicate: Any = None,
        *,
        sheet: str | None = None,
    ) -> Iterator[CellInfo]:
        for c in self._cells:
            if sheet is not None and c.ref.sheet != sheet:
                continue
            if predicate is None or predicate(c):
                yield c


@pytest.fixture
def cells() -> list[CellInfo]:
    return [
        make_cell("A1", value=42),  # plain numeric (hard-coded input candidate)
        make_cell("A2", value=10),  # another hard-coded
        make_cell("A3", value="not used", formula=""),  # text, ignored
        make_cell(
            "B1",
            formula='=VoseInput("Demand")+VoseModPERT(A1,A2,100)',
        ),
        make_cell(
            "B2",
            formula='=VoseOutput("Profit")+B1-50',
            value=75.0,
        ),
        make_cell("C1", formula="=VoseNormal(0,1)"),
        make_cell("C2", formula="=SUM(A1:A2)"),  # plain Excel, not Vose
        make_cell(
            "D1",
            formula="=VoseTimeGBM(10,0.05,0.2)",
        ),
        make_cell("Z9", value=999),  # numeric but not referenced
    ]


def test_list_inputs(cells: list[CellInfo]) -> None:
    bridge = ModelRiskBridge(FakeExcelBridge(cells))  # type: ignore[arg-type]
    inputs = bridge.list_inputs("book.xlsx")
    assert len(inputs) == 1
    assert inputs[0].name == "Demand"
    assert inputs[0].ref.cell == "B1"


def test_list_outputs(cells: list[CellInfo]) -> None:
    bridge = ModelRiskBridge(FakeExcelBridge(cells))  # type: ignore[arg-type]
    outputs = bridge.list_outputs("book.xlsx")
    assert len(outputs) == 1
    assert outputs[0].name == "Profit"
    assert outputs[0].ref.cell == "B2"


def test_list_distributions_finds_all_dist_categories(
    cells: list[CellInfo],
) -> None:
    bridge = ModelRiskBridge(FakeExcelBridge(cells))  # type: ignore[arg-type]
    dists = bridge.list_distributions("book.xlsx")
    cells_found = sorted(d.ref.cell for d in dists)
    # B1: wrapped VoseModPERT, B2: just VoseOutput (no distribution!),
    # C1: VoseNormal, D1: VoseTimeGBM
    assert "B1" in cells_found  # has VoseModPERT
    assert "C1" in cells_found  # has VoseNormal
    assert "D1" in cells_found  # has VoseTimeGBM


def test_list_distributions_detects_wrappers(cells: list[CellInfo]) -> None:
    bridge = ModelRiskBridge(FakeExcelBridge(cells))  # type: ignore[arg-type]
    dists = bridge.list_distributions("book.xlsx")
    b1 = next(d for d in dists if d.ref.cell == "B1")
    assert b1.has_input_wrapper is True
    assert b1.has_output_wrapper is False
    assert b1.function_name == "VoseModPERT"


def test_list_distributions_skips_non_vose(cells: list[CellInfo]) -> None:
    bridge = ModelRiskBridge(FakeExcelBridge(cells))  # type: ignore[arg-type]
    dists = bridge.list_distributions("book.xlsx")
    found_cells = {d.ref.cell for d in dists}
    assert "C2" not in found_cells  # =SUM(...) is not a distribution cell


def test_extract_top_level_args_simple() -> None:
    args = ModelRiskBridge._extract_top_level_args(
        "=VoseNormal(0,1)", "VoseNormal"
    )
    assert args == ["0", "1"]


def test_extract_top_level_args_respects_nested_parens() -> None:
    args = ModelRiskBridge._extract_top_level_args(
        "=VoseNormal(IF(A1>0,A1,0),1)",
        "VoseNormal",
    )
    assert args == ["IF(A1>0,A1,0)", "1"]


def test_extract_top_level_args_respects_strings() -> None:
    args = ModelRiskBridge._extract_top_level_args(
        '=VoseInput("a,b,c")+VoseNormal(0,1)',
        "VoseInput",
    )
    # The comma inside the string must not split the argument list.
    assert args == ['"a,b,c"']


def test_find_hard_coded_inputs(cells: list[CellInfo]) -> None:
    bridge = ModelRiskBridge(FakeExcelBridge(cells))  # type: ignore[arg-type]
    candidates = bridge.find_hard_coded_inputs("book.xlsx")
    cell_refs = sorted(c.cell for c in candidates)
    # A1 and A2 are referenced by =VoseModPERT(A1,A2,100) and =SUM(A1:A2).
    # Z9 is never referenced; not a candidate.
    assert "A1" in cell_refs
    assert "A2" in cell_refs
    assert "Z9" not in cell_refs


def test_is_modelrisk_loaded_returns_bool() -> None:
    bridge = ModelRiskBridge(FakeExcelBridge([]))  # type: ignore[arg-type]
    result = bridge.is_modelrisk_loaded()
    assert isinstance(result, bool)
