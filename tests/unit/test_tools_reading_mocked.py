"""Phase 2 reading-tool tests with a mocked bridge.

We bypass `set_bridge_for_testing` to inject a `ModelRiskBridge` over
fake Excel + fake ResultsCom objects, so the tools execute their real
logic but never touch Excel/COM.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import Any

import numpy as np
import pytest

from modelrisk_mcp.bridge.catalogue import load_catalogue
from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.bridge.results import ResultsReader
from modelrisk_mcp.errors import SimulationFailedError, WorkbookNotFoundError
from modelrisk_mcp.schemas.workbook import CellInfo, CellRef, WorkbookInfo
from modelrisk_mcp.tools import reading

# ----------------------------------------------------------------------
# Fakes
# ----------------------------------------------------------------------


def _cell(
    ref_cell: str,
    *,
    sheet: str = "Sheet1",
    workbook: str = "book.xlsx",
    formula: str = "",
    value: Any = None,
) -> CellInfo:
    return CellInfo(
        ref=CellRef(workbook=workbook, sheet=sheet, cell=ref_cell),
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


class FakeExcel:
    """Minimal in-memory Excel substitute for the reading tools."""

    def __init__(
        self,
        workbooks: list[WorkbookInfo],
        cells: dict[str, list[CellInfo]],
    ) -> None:
        self._workbooks = workbooks
        self._cells = cells  # keyed by workbook name

    # --- ExcelBridge surface ---------------------------------------

    def list_workbooks(self) -> list[WorkbookInfo]:
        return list(self._workbooks)

    def get_active_workbook(self) -> WorkbookInfo:
        if not self._workbooks:
            raise WorkbookNotFoundError("No active workbook.")
        return self._workbooks[0]

    def get_cell(self, workbook: str, sheet: str, cell: str) -> CellInfo:
        for c in self._cells.get(workbook, []):
            if c.ref.sheet == sheet and c.ref.cell == cell.upper():
                return c
        return _cell(cell.upper(), sheet=sheet, workbook=workbook)

    def read_range(self, workbook: str, sheet: str, range_ref: str) -> Any:
        from modelrisk_mcp.schemas.workbook import RangeInfo
        return RangeInfo(
            workbook=workbook,
            sheet=sheet,
            range_ref=range_ref,
            values=[[None]],
            formulas=[[""]],
        )

    def iterate_cells(
        self,
        workbook: str,
        predicate: Any = None,
        *,
        sheet: str | None = None,
    ) -> Iterator[CellInfo]:
        for c in self._cells.get(workbook, []):
            if sheet is not None and c.ref.sheet != sheet:
                continue
            if predicate is None or predicate(c):
                yield c


class FakeSimVariable:
    def __init__(self, name: str, samples: np.ndarray) -> None:
        self._name = name
        self._samples = samples

    def GetName(self) -> str:  # noqa: N802
        return self._name

    def GetMean(self) -> float:  # noqa: N802
        return float(self._samples.mean())

    def GetVariance(self) -> float:  # noqa: N802
        return float(self._samples.var(ddof=1))

    def GetStDev(self) -> float:  # noqa: N802
        return float(self._samples.std(ddof=1))

    def GetSkewness(self) -> float:  # noqa: N802
        return 0.0

    def GetKurtosis(self) -> float:  # noqa: N802
        return 3.0

    def GetPercentile(self, p: float) -> float:  # noqa: N802
        return float(np.quantile(self._samples, p))

    def GetSamples(self) -> Any:  # noqa: N802
        return self._samples.tolist()


class FakeResultsCom:
    def __init__(
        self,
        inputs: list[FakeSimVariable],
        outputs: list[FakeSimVariable],
    ) -> None:
        self._inputs = inputs
        self._outputs = outputs

    def sim_outputs(self) -> list[FakeSimVariable]:
        return list(self._outputs)

    def sim_inputs(self) -> list[FakeSimVariable]:
        return list(self._inputs)


# ----------------------------------------------------------------------
# Fixtures
# ----------------------------------------------------------------------


@pytest.fixture
def workbook_data() -> tuple[
    list[WorkbookInfo], dict[str, list[CellInfo]]
]:
    workbooks = [
        WorkbookInfo(
            name="book.xlsx",
            path="C:/book.xlsx",
            sheets=["Sheet1", "Inputs"],
            active_sheet="Sheet1",
        ),
        WorkbookInfo(
            name="other.xlsx",
            path="C:/other.xlsx",
            sheets=["Main"],
            active_sheet="Main",
        ),
    ]
    cells: dict[str, list[CellInfo]] = {
        "book.xlsx": [
            _cell("A1", value=42),
            _cell("A2", value=10),
            _cell("B1", formula='=VoseInput("Demand")+VoseModPERT(A1,A2,100)'),
            _cell("B2", formula='=VoseOutput("Profit")+B1-50', value=75.0),
            _cell("C1", formula="=VoseNormal(0,1)"),
            _cell("C2", formula="=SUM(A1:A2)"),
            _cell("D1", formula="=VoseTimeGBM(10,0.05,0.2)"),
        ],
        "other.xlsx": [],
    }
    return workbooks, cells


@pytest.fixture
def bridge_with_data(
    workbook_data: tuple[list[WorkbookInfo], dict[str, list[CellInfo]]],
) -> Iterator[ModelRiskBridge]:
    workbooks, cells = workbook_data
    rng = np.random.default_rng(42)
    n = 200
    demand = rng.normal(100, 10, n)
    revenue = demand * 5 + rng.normal(0, 5, n)
    profit = revenue - 200
    results_com = FakeResultsCom(
        inputs=[FakeSimVariable("Demand", demand)],
        outputs=[
            FakeSimVariable("Revenue", revenue),
            FakeSimVariable("Profit", profit),
        ],
    )
    bridge = ModelRiskBridge(
        excel=FakeExcel(workbooks, cells),  # type: ignore[arg-type]
        catalogue=load_catalogue(),
        results=ResultsReader(com=results_com),
    )
    reading.set_bridge_for_testing(bridge)
    yield bridge
    reading.set_bridge_for_testing(None)


# ----------------------------------------------------------------------
# Workbook-level tools
# ----------------------------------------------------------------------


def test_list_open_workbooks(bridge_with_data: ModelRiskBridge) -> None:
    result = reading.list_open_workbooks()
    names = [w.name for w in result]
    assert names == ["book.xlsx", "other.xlsx"]


def test_get_active_workbook(bridge_with_data: ModelRiskBridge) -> None:
    result = reading.get_active_workbook()
    assert result.name == "book.xlsx"
    assert result.active_sheet == "Sheet1"


def test_get_workbook_summary(bridge_with_data: ModelRiskBridge) -> None:
    summary = reading.get_workbook_summary("book.xlsx")
    assert summary.workbook == "book.xlsx"
    assert summary.sheets == ["Sheet1", "Inputs"]
    assert summary.input_count == 1   # B1 has VoseInput
    assert summary.output_count == 1  # B2 has VoseOutput
    assert summary.distribution_count == 3  # B1, C1, D1
    assert summary.formula_cell_count == 5
    assert summary.numeric_cell_count == 2  # A1, A2


def test_get_workbook_summary_empty_workbook(
    bridge_with_data: ModelRiskBridge,
) -> None:
    summary = reading.get_workbook_summary("other.xlsx")
    assert summary.input_count == 0
    assert summary.output_count == 0
    assert summary.distribution_count == 0


# ----------------------------------------------------------------------
# Per-cell read tools
# ----------------------------------------------------------------------


def test_list_modelrisk_inputs(bridge_with_data: ModelRiskBridge) -> None:
    inputs = reading.list_modelrisk_inputs("book.xlsx")
    assert len(inputs) == 1
    assert inputs[0].name == "Demand"
    assert inputs[0].ref.cell == "B1"


def test_list_modelrisk_outputs(bridge_with_data: ModelRiskBridge) -> None:
    outputs = reading.list_modelrisk_outputs("book.xlsx")
    assert len(outputs) == 1
    assert outputs[0].name == "Profit"
    assert outputs[0].ref.cell == "B2"


def test_list_distributions(bridge_with_data: ModelRiskBridge) -> None:
    dists = reading.list_distributions("book.xlsx")
    cells = sorted(d.ref.cell for d in dists)
    assert cells == ["B1", "C1", "D1"]
    b1 = next(d for d in dists if d.ref.cell == "B1")
    assert b1.function_name == "VoseModPERT"
    assert b1.has_input_wrapper is True


def test_list_distributions_with_sheet_filter(
    bridge_with_data: ModelRiskBridge,
) -> None:
    dists = reading.list_distributions("book.xlsx", sheet="Sheet1")
    assert all(d.ref.sheet == "Sheet1" for d in dists)


def test_get_cell(bridge_with_data: ModelRiskBridge) -> None:
    cell = reading.get_cell("book.xlsx", "Sheet1", "B1")
    assert cell.ref.cell == "B1"
    assert "VoseInput" in cell.formula


def test_read_range(bridge_with_data: ModelRiskBridge) -> None:
    rng = reading.read_range("book.xlsx", "Sheet1", "A1:A2")
    assert rng.range_ref == "A1:A2"


def test_find_hard_coded_inputs(bridge_with_data: ModelRiskBridge) -> None:
    candidates = reading.find_hard_coded_inputs("book.xlsx")
    refs = sorted(c["cell"] for c in candidates)
    # A1 and A2 are referenced from =VoseModPERT(A1,A2,100) and =SUM(A1:A2).
    assert "A1" in refs
    assert "A2" in refs


# ----------------------------------------------------------------------
# Simulation result tools (through ResultsReader)
# ----------------------------------------------------------------------


def test_get_simulation_results(bridge_with_data: ModelRiskBridge) -> None:
    results = reading.get_simulation_results("book.xlsx")
    assert len(results) == 2
    names = {r.output_name for r in results}
    assert names == {"Revenue", "Profit"}
    for r in results:
        assert r.iterations == 200
        assert 0.05 in r.percentiles
        assert 0.95 in r.percentiles
        assert r.min <= r.percentiles[0.05] <= r.percentiles[0.95] <= r.max


def test_get_simulation_results_filter(
    bridge_with_data: ModelRiskBridge,
) -> None:
    results = reading.get_simulation_results(
        "book.xlsx", output_names=["Profit"]
    )
    assert len(results) == 1
    assert results[0].output_name == "Profit"


def test_get_correlation_matrix(bridge_with_data: ModelRiskBridge) -> None:
    matrix = reading.get_correlation_matrix("book.xlsx")
    assert set(matrix.names) == {"Demand", "Revenue", "Profit"}
    assert len(matrix.pearson) == 3
    # Demand strongly correlates with Revenue (linear construction).
    demand_idx = matrix.names.index("Demand")
    revenue_idx = matrix.names.index("Revenue")
    assert matrix.pearson[demand_idx][revenue_idx] is not None
    assert matrix.pearson[demand_idx][revenue_idx] > 0.9


def test_get_sensitivity_ranking(bridge_with_data: ModelRiskBridge) -> None:
    ranking = reading.get_sensitivity_ranking("book.xlsx", "Profit")
    assert ranking.output_name == "Profit"
    assert len(ranking.entries) == 1  # one input (Demand)
    assert ranking.entries[0].input_name == "Demand"
    assert abs(ranking.entries[0].correlation) > 0.9


def test_get_sensitivity_ranking_unknown_output(
    bridge_with_data: ModelRiskBridge,
) -> None:
    with pytest.raises(SimulationFailedError):
        reading.get_sensitivity_ranking("book.xlsx", "DoesNotExist")


# ----------------------------------------------------------------------
# Bridge factory
# ----------------------------------------------------------------------


def test_set_bridge_for_testing_resets_to_none() -> None:
    reading.set_bridge_for_testing(None)
    # Subsequent call to get_bridge() would lazy-create a real one;
    # we don't trigger it here because that would attach to Excel.
    assert reading._bridge is None
