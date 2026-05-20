"""Phase 5 workflow-tool tests.

Mocks the bridge so audit_model + propose_distributions_for_inputs +
generate_executive_summary all run against synthetic data. The fixture
workbook contains a known set of methodology problems so we can pin
which rules fire.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import Any

import numpy as np
import pytest

from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.bridge.results import ResultsReader
from modelrisk_mcp.schemas.workbook import CellInfo, CellRef, WorkbookInfo
from modelrisk_mcp.tools import reading, workflows


def _cell(
    cell_ref: str,
    *,
    formula: str = "",
    value: Any = None,
    sheet: str = "Sheet1",
    workbook: str = "book.xlsx",
) -> CellInfo:
    return CellInfo(
        ref=CellRef(workbook=workbook, sheet=sheet, cell=cell_ref),
        formula=formula,
        value=value,
        cell_type="formula" if formula else ("number" if value is not None else "empty"),
    )


class FakeExcel:
    def __init__(self, cells: list[CellInfo]) -> None:
        self._cells = cells

    def list_workbooks(self) -> list[WorkbookInfo]:
        return [WorkbookInfo(name="book.xlsx", path="C:/book.xlsx", sheets=["Sheet1"])]

    def get_active_workbook(self) -> WorkbookInfo:
        return self.list_workbooks()[0]

    def iterate_cells(
        self, workbook: str, predicate: Any = None, *, sheet: str | None = None
    ) -> Iterator[CellInfo]:
        for c in self._cells:
            if c.ref.workbook != workbook:
                continue
            if sheet is not None and c.ref.sheet != sheet:
                continue
            if predicate is None or predicate(c):
                yield c

    def get_cell(self, workbook: str, sheet: str, cell: str) -> CellInfo:
        for c in self._cells:
            if (
                c.ref.workbook == workbook
                and c.ref.sheet == sheet
                and c.ref.cell == cell.upper()
            ):
                return c
        return _cell(cell, sheet=sheet, workbook=workbook)


class FakeSimVar:
    def __init__(self, name: str, samples: np.ndarray) -> None:
        self._n = name
        self._s = samples

    def GetName(self) -> str:  # noqa: N802
        return self._n

    def GetMean(self) -> float:  # noqa: N802
        return float(self._s.mean())

    def GetStDev(self) -> float:  # noqa: N802
        return float(self._s.std(ddof=1))

    def GetVariance(self) -> float:  # noqa: N802
        return float(self._s.var(ddof=1))

    def GetSkewness(self) -> float:  # noqa: N802
        return 0.0

    def GetKurtosis(self) -> float:  # noqa: N802
        return 3.0

    def GetPercentile(self, p: float) -> float:  # noqa: N802
        return float(np.quantile(self._s, p))

    def GetSamples(self) -> Any:  # noqa: N802
        return self._s.tolist()


class FakeResultsCom:
    def __init__(self, inputs: list[FakeSimVar], outputs: list[FakeSimVar]) -> None:
        self._i = inputs
        self._o = outputs

    def sim_outputs(self) -> list[FakeSimVar]:
        return list(self._o)

    def sim_inputs(self) -> list[FakeSimVar]:
        return list(self._i)


@pytest.fixture
def audit_fixture_cells() -> list[CellInfo]:
    """A small workbook with a known set of methodology problems so the
    audit detectors fire predictably."""
    return [
        # VOSE-006: hard-coded numeric inputs referenced by a formula.
        _cell("A1", value=100),
        _cell("A2", value=10),
        # VOSE-001: unknown Vose function (typo).
        _cell("B1", formula="=VoseModPert(A1,A2,200)"),
        # VOSE-002: distribution without VoseInput wrapper.
        _cell("B2", formula="=VoseNormal(0,1)"),
        # VOSE-005: arithmetic-before-input pattern.
        _cell("B3", formula="=2*VoseNormal(0,1)"),
        # Healthy cells (no findings expected).
        _cell("C1", formula='=VoseInput("Demand")+VoseModPERT(A1,A2,200,4)'),
        _cell("C2", formula='=VoseOutput("Profit")+C1-50'),
        # VOSE-003: fit without uncertainty.
        _cell("D1", formula="=VoseLognormalFit(A1:A10)"),
        # VOSE-004: VoseOutput cell with no distribution dependency.
        _cell("E1", formula='=VoseOutput("Constant")+42'),
    ]


@pytest.fixture
def bridge(audit_fixture_cells: list[CellInfo]) -> Iterator[ModelRiskBridge]:
    rng = np.random.default_rng(42)
    n = 200
    demand = rng.normal(100, 10, n)
    revenue = demand * 5 + rng.normal(0, 5, n)
    results_com = FakeResultsCom(
        inputs=[FakeSimVar("Demand", demand)],
        outputs=[FakeSimVar("Revenue", revenue)],
    )
    bridge = ModelRiskBridge(
        excel=FakeExcel(audit_fixture_cells),  # type: ignore[arg-type]
        results=ResultsReader(com=results_com),
    )
    reading.set_bridge_for_testing(bridge)
    yield bridge
    reading.set_bridge_for_testing(None)


# ----------------------------------------------------------------------
# propose_distributions_for_inputs
# ----------------------------------------------------------------------


class TestProposeDistributions:
    def test_three_point_keyword_match(self, bridge: ModelRiskBridge) -> None:
        result = workflows.propose_distributions_for_inputs(
            inputs=[
                {
                    "cell_ref": "B5",
                    "current_value": 100,
                    "description": "best case / most likely / worst case estimate of unit cost",
                }
            ]
        )
        assert len(result) == 1
        entry = result[0]
        assert entry["scenario_matched"] == "three_point_estimate"
        functions = [r["function"] for r in entry["recommendations"]]
        assert "VoseModPERT" in functions

    def test_count_keyword_match(self, bridge: ModelRiskBridge) -> None:
        result = workflows.propose_distributions_for_inputs(
            inputs=[{"description": "number of incidents per year"}]
        )
        assert result[0]["scenario_matched"] == "count"

    def test_unknown_falls_back(self, bridge: ModelRiskBridge) -> None:
        result = workflows.propose_distributions_for_inputs(
            inputs=[{"description": "whatever"}]
        )
        assert result[0]["scenario_matched"] == "unknown"


# ----------------------------------------------------------------------
# discover_inputs
# ----------------------------------------------------------------------


class TestDiscoverInputs:
    def test_returns_ranked_list(self, bridge: ModelRiskBridge) -> None:
        results = workflows.discover_inputs("book.xlsx")
        # A1 is 100 (round, +0.5 +0.5 = score 2.0), A2 is 10 (round, +0.5 = 1.5).
        assert len(results) >= 2
        cells = [r["cell"] for r in results]
        assert "A1" in cells
        assert "A2" in cells
        # Higher score sorts first.
        scores = [r["score"] for r in results]
        assert scores == sorted(scores, reverse=True)


# ----------------------------------------------------------------------
# audit_model
# ----------------------------------------------------------------------


class TestAuditModel:
    def test_detects_unknown_function(self, bridge: ModelRiskBridge) -> None:
        report = workflows.audit_model("book.xlsx")
        rule_ids = {f.rule_id for f in report.findings}
        assert "VOSE-001" in rule_ids
        # The unknown function should be reported with the suggestion.
        unknown_findings = [f for f in report.findings if f.rule_id == "VOSE-001"]
        assert any(
            f.cell is not None and f.cell.cell == "B1" for f in unknown_findings
        )
        assert any("VoseModPert" in f.message for f in unknown_findings)

    def test_detects_distribution_without_wrapper(
        self, bridge: ModelRiskBridge
    ) -> None:
        report = workflows.audit_model("book.xlsx")
        # B2 (=VoseNormal(0,1)) should trigger VOSE-002.
        b2_findings = [
            f for f in report.findings
            if f.rule_id == "VOSE-002"
            and f.cell is not None
            and f.cell.cell == "B2"
        ]
        assert len(b2_findings) == 1

    def test_detects_fit_without_uncertainty(
        self, bridge: ModelRiskBridge
    ) -> None:
        report = workflows.audit_model("book.xlsx")
        d1_findings = [
            f for f in report.findings
            if f.rule_id == "VOSE-003"
            and f.cell is not None
            and f.cell.cell == "D1"
        ]
        assert len(d1_findings) == 1

    def test_detects_hard_coded_inputs(self, bridge: ModelRiskBridge) -> None:
        report = workflows.audit_model("book.xlsx")
        rule_ids = {f.rule_id for f in report.findings}
        assert "VOSE-006" in rule_ids


# ----------------------------------------------------------------------
# generate_executive_summary
# ----------------------------------------------------------------------


class TestGenerateExecutiveSummary:
    def test_returns_markdown_with_table(self, bridge: ModelRiskBridge) -> None:
        result = workflows.generate_executive_summary("book.xlsx")
        md = result["markdown"]
        assert "Simulation summary" in md
        assert "| Output | Mean | P50 |" in md
        assert "Revenue" in md

    def test_contingency_section_when_deterministic_provided(
        self, bridge: ModelRiskBridge
    ) -> None:
        result = workflows.generate_executive_summary(
            "book.xlsx",
            deterministic_values={"Revenue": 500.0},
        )
        md = result["markdown"]
        assert "Contingency vs deterministic" in md

    def test_no_results_returns_explanatory_markdown(self) -> None:
        # Build a bridge with no results.
        empty_results = FakeResultsCom(inputs=[], outputs=[])
        bridge = ModelRiskBridge(
            excel=FakeExcel([]),  # type: ignore[arg-type]
            results=ResultsReader(com=empty_results),
        )
        reading.set_bridge_for_testing(bridge)
        try:
            result = workflows.generate_executive_summary("book.xlsx")
            assert "No simulation results" in result["markdown"]
        finally:
            reading.set_bridge_for_testing(None)
