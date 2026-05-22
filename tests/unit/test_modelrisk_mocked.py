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


# ----------------------------------------------------------------------
# Bug #20 post-condition verification + #21 restore_deterministic_state
# ----------------------------------------------------------------------


class _PostCondFakeExcel(FakeExcelBridge):
    """Extends `FakeExcelBridge` with the workbook-level methods that
    `run_simulation` calls — `get_active_workbook` and
    `recalculate_workbook`. The latter records its calls so tests can
    assert the auto-restore fires on post-condition failure (#21)."""

    def __init__(
        self, cells: list[CellInfo], *, active_name: str = "book.xlsx",
    ) -> None:
        super().__init__(cells)
        self._active_name = active_name
        self.recalculate_calls: list[str] = []

    def get_active_workbook(self) -> Any:
        from modelrisk_mcp.schemas.workbook import WorkbookInfo
        return WorkbookInfo(
            name=self._active_name, path="", sheets=["Sheet1"],
        )

    def recalculate_workbook(self, workbook: str) -> None:
        self.recalculate_calls.append(workbook)


class _FakeVmrsHandle:
    """Stand-in for `VmrsHandle`. `lookup_var_id` returns the configured
    map; opening / closing are no-ops."""

    def __init__(self, var_ids: dict[str, int | None]) -> None:
        self._var_ids = var_ids

    def __enter__(self) -> _FakeVmrsHandle:
        return self

    def __exit__(self, *exc: Any) -> None:
        pass

    def lookup_var_id(self, name: str) -> int | None:
        return self._var_ids.get(name)


class _FakeMrService:
    def __init__(self, var_ids: dict[str, int | None]) -> None:
        self._var_ids = var_ids
        self.open_calls: list[str] = []

    def open_vmrs(self, path: str) -> _FakeVmrsHandle:
        self.open_calls.append(path)
        return _FakeVmrsHandle(self._var_ids)


class _FakeSimulationController:
    """Minimal SimulationController stand-in. `run_simulation` returns
    a canned `SimulationRunResult` so the bridge wiring is exercised
    without touching Excel / the XLL surface."""

    def __init__(self, vmrs_path: str = r"C:\tmp\book.vmrs") -> None:
        self._vmrs_path = vmrs_path

    def run_simulation(
        self,
        *,
        workbook_name: str | None,
        samples: int,
        seed: int,
        save_to: str | None,
        output_names: tuple[str, ...] = (),
    ) -> Any:
        from modelrisk_mcp.bridge.simulation import SimulationRunResult
        # Stash for tests that want to assert the bridge populated this.
        self.last_output_names = output_names
        return SimulationRunResult(
            workbook_name=workbook_name or "book.xlsx",
            vmrs_path=self._vmrs_path,
            iterations=samples,
        )


def _voseoutput_cells() -> list[CellInfo]:
    """Cells containing one VoseOutput so list_outputs returns it."""
    return [
        make_cell("B2", formula='=VoseOutput("Profit")+B1-50', value=75.0),
    ]


def test_run_simulation_post_condition_satisfied_when_output_resolves() -> None:
    """Happy path: at least one expected output name resolves to a
    var_id in the produced .vmrs → no error, no auto-restore."""
    excel = _PostCondFakeExcel(_voseoutput_cells())
    bridge = ModelRiskBridge(excel)  # type: ignore[arg-type]
    bridge._simulation = _FakeSimulationController()  # type: ignore[assignment]
    bridge._mrservice = _FakeMrService({"Profit": 7})  # type: ignore[assignment]
    result = bridge.run_simulation(workbook="book.xlsx", samples=1000)
    assert result.vmrs_path == r"C:\tmp\book.vmrs"
    assert excel.recalculate_calls == []  # no recovery needed


def test_run_simulation_post_condition_fails_when_no_output_registered() -> None:
    """Bug #20 regression: .vmrs lacks the expected output names →
    `run_simulation` raises SimulationFailedError instead of
    pretending success."""
    from modelrisk_mcp.errors import SimulationFailedError

    excel = _PostCondFakeExcel(_voseoutput_cells())
    bridge = ModelRiskBridge(excel)  # type: ignore[arg-type]
    bridge._simulation = _FakeSimulationController()  # type: ignore[assignment]
    # Profit not in the .vmrs → post-condition fails.
    bridge._mrservice = _FakeMrService({"Profit": None})  # type: ignore[assignment]
    with pytest.raises(SimulationFailedError) as exc:
        bridge.run_simulation(workbook="book.xlsx", samples=1000)
    msg = str(exc.value)
    assert "does not register any of the expected" in msg
    assert "Profit" in msg
    # Auto-restore fires on failure (bug #21 recovery path).
    assert excel.recalculate_calls == ["book.xlsx"]


def test_run_simulation_skips_verification_when_no_outputs_declared() -> None:
    """A workbook with zero VoseOutput cells can't be verified — but
    that's not a failure, it's a no-op (the sim just won't have
    anything to report)."""
    excel = _PostCondFakeExcel([])  # no VoseOutputs
    bridge = ModelRiskBridge(excel)  # type: ignore[arg-type]
    bridge._simulation = _FakeSimulationController()  # type: ignore[assignment]
    bridge._mrservice = _FakeMrService({})  # type: ignore[assignment]
    result = bridge.run_simulation(workbook="book.xlsx", samples=1000)
    assert result.vmrs_path == r"C:\tmp\book.vmrs"


def test_restore_deterministic_state_calls_recalc() -> None:
    """Bug #21 recovery tool: triggers a full recalculation of the
    given workbook (or the active one if omitted)."""
    excel = _PostCondFakeExcel([])
    bridge = ModelRiskBridge(excel)  # type: ignore[arg-type]
    out = bridge.restore_deterministic_state("book.xlsx")
    assert out == {"workbook_name": "book.xlsx", "recalculated": True}
    assert excel.recalculate_calls == ["book.xlsx"]


def test_run_simulation_does_not_filter_xll_outputs() -> None:
    """alpha.32 reversal of alpha.18. Round-7 testing on Vose's
    `Inputs Outputs.xlsx` sample (which has an expression-based
    output name) showed that explicitly populating `output_names`
    acts as a FILTER on the XLL side — only outputs matching the
    list get registered in the .vmrs. The alpha.18 fix was a
    coincidental work-around; the real culprit was bug #29 (XLL
    not registered), fixed properly in alpha.27. With #29 fixed,
    passing empty `output_names` lets the XLL auto-scan and
    register every VoseOutput, including expression-named ones."""
    excel = _PostCondFakeExcel(_voseoutput_cells())
    bridge = ModelRiskBridge(excel)  # type: ignore[arg-type]
    sim = _FakeSimulationController()
    bridge._simulation = sim  # type: ignore[assignment]
    bridge._mrservice = _FakeMrService({"Profit": 7})  # type: ignore[assignment]
    bridge.run_simulation(workbook="book.xlsx", samples=1000)
    # alpha.32: we no longer pre-populate output_names. The XLL
    # scans the workbook and registers all VoseOutputs.
    assert sim.last_output_names == ()


def test_restore_deterministic_state_defaults_to_active_workbook() -> None:
    excel = _PostCondFakeExcel([], active_name="active.xlsx")
    bridge = ModelRiskBridge(excel)  # type: ignore[arg-type]
    out = bridge.restore_deterministic_state()
    assert out["workbook_name"] == "active.xlsx"
    assert excel.recalculate_calls == ["active.xlsx"]
