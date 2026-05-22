"""MCP-wrapper tests for `tools/workflows.py`.

Workflow tools are higher-level than reading/building tools — they
compose multiple bridge calls and apply domain logic (distribution
selection, sensitivity ranking, markdown summary generation). What we
test here:

- `propose_distributions_for_inputs` is a pure function (no bridge);
  verify keyword matching and fallback behaviour.
- `discover_inputs` ranks bridge-supplied hard-coded inputs by a
  heuristic — verify ordering and that the bridge is queried correctly.
- `audit_model` delegates to `audit.engine.run_audit` — verify the
  delegation, not the audit logic itself (which has its own tests).
- `generate_executive_summary` produces markdown from bridge results —
  verify section presence and the no-results path.
"""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, patch

import pytest

from modelrisk_mcp.schemas.results import (
    AuditReport,
    SensitivityEntry,
    SensitivityRanking,
    SimulationResult,
)
from modelrisk_mcp.schemas.workbook import CellInfo, CellRef, WorkbookInfo, WorkbookSummary
from modelrisk_mcp.tools import reading, workflows


@pytest.fixture
def bridge() -> Iterator[MagicMock]:
    b = MagicMock()
    b.excel = MagicMock()
    reading.set_bridge_for_testing(b)  # type: ignore[arg-type]
    yield b
    reading.set_bridge_for_testing(None)


# ----------------------------------------------------------------------
# propose_distributions_for_inputs — pure function
# ----------------------------------------------------------------------


class TestProposeDistributions:
    # alpha.17 envelope sweep: workflow tools that returned bare lists
    # now wrap in `{"proposals": [...], "count": N}` — the tests
    # dereference the noun key but otherwise the semantics are unchanged.
    def test_returns_one_entry_per_input(self) -> None:
        out = workflows.propose_distributions_for_inputs(
            [
                {"description": "annual demand for widgets"},
                {"description": "time to failure of pump"},
            ]
        )["proposals"]
        assert len(out) == 2
        assert {e["description"] for e in out} == {
            "annual demand for widgets",
            "time to failure of pump",
        }

    def test_each_entry_has_recommendations(self) -> None:
        out = workflows.propose_distributions_for_inputs(
            [{"description": "monthly sales count"}]
        )["proposals"]
        assert "recommendations" in out[0]
        assert isinstance(out[0]["recommendations"], list)

    def test_unknown_description_falls_back(self) -> None:
        """Empty / gibberish descriptions hit the `unknown` scenario,
        which still returns the distribution-guide's catch-all
        recommendations (not an empty list)."""
        out = workflows.propose_distributions_for_inputs(
            [{"description": ""}]
        )["proposals"]
        assert out[0]["scenario_matched"] == "unknown"

    def test_preserves_cell_ref_and_current_value(self) -> None:
        out = workflows.propose_distributions_for_inputs(
            [
                {
                    "cell_ref": "S1!A1",
                    "current_value": 100.0,
                    "description": "cost",
                }
            ]
        )["proposals"]
        assert out[0]["cell_ref"] == "S1!A1"
        assert out[0]["current_value"] == 100.0


# ----------------------------------------------------------------------
# discover_inputs — bridge.find_hard_coded_inputs + heuristic ranking
# ----------------------------------------------------------------------


class TestDiscoverInputs:
    # alpha.17 envelope sweep: `discover_inputs` now returns
    # `{"candidates": [...], "count": N}` instead of a bare list.
    def test_returns_empty_when_no_candidates(
        self, bridge: MagicMock
    ) -> None:
        bridge.find_hard_coded_inputs.return_value = []
        bridge.excel.iterate_cells.return_value = iter([])
        result = workflows.discover_inputs("m.xlsx")
        assert result == {"candidates": [], "count": 0}

    def test_limit_respected(self, bridge: MagicMock) -> None:
        # 30 hard-coded cells, but limit=5 → only 5 returned.
        refs = [
            CellRef(workbook="m.xlsx", sheet="In", cell=f"A{i}")
            for i in range(1, 31)
        ]
        bridge.find_hard_coded_inputs.return_value = refs
        bridge.excel.iterate_cells.return_value = iter([])
        out = workflows.discover_inputs("m.xlsx", limit=5)["candidates"]
        assert len(out) == 5

    def test_higher_score_for_round_numbers(
        self, bridge: MagicMock
    ) -> None:
        refs = [
            CellRef(workbook="m.xlsx", sheet="In", cell="A1"),
            CellRef(workbook="m.xlsx", sheet="In", cell="A2"),
        ]
        bridge.find_hard_coded_inputs.return_value = refs
        # A1=1000 (multiple of 10, 100, 1000 → highest score)
        # A2=7 (not round → base score 1.0)
        bridge.excel.iterate_cells.return_value = iter([
            CellInfo(
                ref=CellRef(workbook="m.xlsx", sheet="In", cell="A1"),
                formula="", value=1000, cell_type="number",
            ),
            CellInfo(
                ref=CellRef(workbook="m.xlsx", sheet="In", cell="A2"),
                formula="", value=7, cell_type="number",
            ),
        ])
        out = workflows.discover_inputs("m.xlsx")["candidates"]
        # Sorted descending by score — A1 first.
        assert out[0]["cell"] == "A1"
        assert out[0]["score"] > out[1]["score"]

    def test_one_excluded_from_multiple_of_10_bonus(
        self, bridge: MagicMock
    ) -> None:
        """`1` is excluded from the `% 10 == 0` bonus — it's likely a
        flag, not a scenario assumption. Verify the exclusion fires for
        a value that would otherwise get the bonus."""
        refs = [
            CellRef(workbook="m.xlsx", sheet="In", cell="A1"),
            CellRef(workbook="m.xlsx", sheet="In", cell="A2"),
        ]
        bridge.find_hard_coded_inputs.return_value = refs
        bridge.excel.iterate_cells.return_value = iter([
            CellInfo(
                ref=CellRef(workbook="m.xlsx", sheet="In", cell="A1"),
                formula="", value=1, cell_type="number",
            ),
            CellInfo(
                ref=CellRef(workbook="m.xlsx", sheet="In", cell="A2"),
                formula="", value=20, cell_type="number",
            ),
        ])
        out = workflows.discover_inputs("m.xlsx")["candidates"]
        # A2 (=20) outranks A1 (=1) because A2 gets the multiple-of-10 bonus.
        assert out[0]["cell"] == "A2"

    def test_zero_excluded_from_all_round_number_bonuses(
        self, bridge: MagicMock
    ) -> None:
        """Regression: previously `value=0` triggered the % 100 and
        % 1000 bonuses because 0 % n == 0, so a cell holding 0 scored
        2.0 — identical to a cell holding 100 or 1000. Flag-shaped
        values must score lower than real scenario assumptions."""
        refs = [
            CellRef(workbook="m.xlsx", sheet="In", cell="A1"),
            CellRef(workbook="m.xlsx", sheet="In", cell="A2"),
        ]
        bridge.find_hard_coded_inputs.return_value = refs
        bridge.excel.iterate_cells.return_value = iter([
            CellInfo(
                ref=CellRef(workbook="m.xlsx", sheet="In", cell="A1"),
                formula="", value=0, cell_type="number",
            ),
            CellInfo(
                ref=CellRef(workbook="m.xlsx", sheet="In", cell="A2"),
                formula="", value=100, cell_type="number",
            ),
        ])
        out = workflows.discover_inputs("m.xlsx")["candidates"]
        # A2 (=100) MUST outrank A1 (=0).
        a1_entry = next(e for e in out if e["cell"] == "A1")
        a2_entry = next(e for e in out if e["cell"] == "A2")
        assert a2_entry["score"] > a1_entry["score"]
        assert out[0]["cell"] == "A2"

    def test_boolean_values_excluded(
        self, bridge: MagicMock
    ) -> None:
        """`True` / `False` from Excel come through as bools. They're
        `isinstance(x, int)` true in Python, so without the explicit
        bool guard a False cell would get the same scoring path as
        value=0 — wrong by category."""
        refs = [
            CellRef(workbook="m.xlsx", sheet="In", cell="A1"),
            CellRef(workbook="m.xlsx", sheet="In", cell="A2"),
        ]
        bridge.find_hard_coded_inputs.return_value = refs
        bridge.excel.iterate_cells.return_value = iter([
            CellInfo(
                ref=CellRef(workbook="m.xlsx", sheet="In", cell="A1"),
                formula="", value=False, cell_type="boolean",
            ),
            CellInfo(
                ref=CellRef(workbook="m.xlsx", sheet="In", cell="A2"),
                formula="", value=50, cell_type="number",
            ),
        ])
        out = workflows.discover_inputs("m.xlsx")["candidates"]
        # Both should be present, A2 should score higher.
        a1_entry = next(e for e in out if e["cell"] == "A1")
        a2_entry = next(e for e in out if e["cell"] == "A2")
        assert a1_entry["score"] == 1.0  # base; no bonuses for booleans
        assert a2_entry["score"] > a1_entry["score"]


# ----------------------------------------------------------------------
# audit_model — delegates to audit.engine.run_audit
# ----------------------------------------------------------------------


class TestAuditModel:
    def test_delegates_to_run_audit(self, bridge: MagicMock) -> None:
        fake_report = AuditReport(findings=[])
        with patch.object(workflows, "run_audit", return_value=fake_report) as p:
            result = workflows.audit_model("m.xlsx")
        p.assert_called_once_with(bridge, "m.xlsx")
        assert isinstance(result, AuditReport)
        assert result is fake_report


# ----------------------------------------------------------------------
# generate_executive_summary — markdown composition
# ----------------------------------------------------------------------


class TestExecutiveSummary:
    def test_no_results_message(self, bridge: MagicMock) -> None:
        bridge.get_simulation_results.return_value = []
        out = workflows.generate_executive_summary("m.xlsx")
        assert "No simulation results" in out["markdown"]

    def test_includes_per_output_table(self, bridge: MagicMock) -> None:
        bridge.get_simulation_results.return_value = [
            SimulationResult(
                output_name="profit",
                iterations=1000, mean=100.0, stdev=20.0, variance=400.0,
                skewness=0.1, kurtosis=2.9, min=50.0, max=180.0,
                percentiles={0.05: 60, 0.5: 99, 0.95: 145},
            )
        ]
        bridge.get_sensitivity_ranking.return_value = SensitivityRanking(
            output_name="profit", entries=[], iterations=1000,
        )
        md = workflows.generate_executive_summary("m.xlsx")["markdown"]
        assert "profit" in md
        assert "## Per-output statistics" in md
        assert "Mean" in md and "P50" in md

    def test_contingency_section_only_when_deterministic_provided(
        self, bridge: MagicMock
    ) -> None:
        bridge.get_simulation_results.return_value = [
            SimulationResult(
                output_name="cost",
                iterations=1000, mean=200.0, stdev=30.0, variance=900.0,
                skewness=0.0, kurtosis=3.0, min=100.0, max=300.0,
                percentiles={0.5: 200, 0.8: 230, 0.95: 250},
            )
        ]
        bridge.get_sensitivity_ranking.return_value = SensitivityRanking(
            output_name="cost", entries=[], iterations=1000,
        )

        # Without deterministic values: no contingency section.
        md_without = workflows.generate_executive_summary("m.xlsx")["markdown"]
        assert "Contingency vs deterministic" not in md_without

        # With deterministic values: section appears.
        md_with = workflows.generate_executive_summary(
            "m.xlsx", deterministic_values={"cost": 180.0}
        )["markdown"]
        assert "Contingency vs deterministic" in md_with

    def test_sensitivity_unavailable_is_graceful(
        self, bridge: MagicMock
    ) -> None:
        """If sensitivity-ranking raises for one output, summary
        continues with a parenthetical note rather than failing."""
        bridge.get_simulation_results.return_value = [
            SimulationResult(
                output_name="x",
                iterations=100, mean=1.0, stdev=0.5, variance=0.25,
                skewness=0, kurtosis=0, min=0, max=2,
                percentiles={0.5: 1.0},
            )
        ]
        bridge.get_sensitivity_ranking.side_effect = RuntimeError("no data")
        md = workflows.generate_executive_summary("m.xlsx")["markdown"]
        assert "sensitivity unavailable" in md

    def test_top_sensitivity_drivers_table_rendered(
        self, bridge: MagicMock
    ) -> None:
        bridge.get_simulation_results.return_value = [
            SimulationResult(
                output_name="margin",
                iterations=1000, mean=5.0, stdev=1.0, variance=1.0,
                skewness=0, kurtosis=0, min=2, max=8,
                percentiles={0.5: 5.0},
            )
        ]
        bridge.get_sensitivity_ranking.return_value = SensitivityRanking(
            output_name="margin",
            entries=[
                SensitivityEntry(
                    input_name="price", correlation=0.8,
                    regression_coefficient=0.7,
                ),
                SensitivityEntry(
                    input_name="volume", correlation=-0.3,
                    regression_coefficient=-0.2,
                ),
            ],
            iterations=1000,
        )
        md = workflows.generate_executive_summary("m.xlsx")["markdown"]
        assert "Top sensitivity drivers" in md
        assert "price" in md
        assert "volume" in md


# ----------------------------------------------------------------------
# diagnose_workbook — high-leverage first-call introspection
# ----------------------------------------------------------------------


class TestDiagnoseWorkbook:
    def _setup_happy_path(self, bridge: MagicMock, tmp_path: Path) -> Path:
        """Wire bridge fakes for a healthy workbook with one output."""
        from modelrisk_mcp.config import Settings

        wb_path = tmp_path / "m.xlsx"
        wb_path.write_text("")
        vmrs_path = tmp_path / "m.vmrs"
        vmrs_path.write_bytes(b"fake")

        bridge.excel.get_active_workbook.return_value = WorkbookInfo(
            name="m.xlsx", path=str(wb_path), sheets=["S1"], active_sheet="S1",
        )
        bridge.is_modelrisk_loaded.return_value = True
        bridge.get_workbook_summary.return_value = WorkbookSummary(
            workbook="m.xlsx", sheets=["S1"],
            input_count=3, output_count=1, distribution_count=2,
            formula_cell_count=10, numeric_cell_count=5,
            modelrisk_loaded=True,
        )
        # The diagnose tool reads _settings.writes_log_path off the bridge
        # to surface the audit-log location.
        bridge._settings = Settings()
        return vmrs_path

    def test_happy_path_no_issues(
        self, bridge: MagicMock, tmp_path: Path
    ) -> None:
        self._setup_happy_path(bridge, tmp_path)
        out = workflows.diagnose_workbook()
        assert out["excel_connected"] is True
        assert out["modelrisk_loaded"] is True
        assert out["active_workbook"] == "m.xlsx"
        assert out["input_count"] == 3
        assert out["output_count"] == 1
        assert out["vmrs_exists"] is True
        assert out["vmrs_modified"] is not None
        assert out["issues"] == []

    def test_no_output_cells_flagged(
        self, bridge: MagicMock, tmp_path: Path
    ) -> None:
        self._setup_happy_path(bridge, tmp_path)
        bridge.get_workbook_summary.return_value = WorkbookSummary(
            workbook="m.xlsx", sheets=["S1"],
            input_count=3, output_count=0, distribution_count=2,
            formula_cell_count=10, numeric_cell_count=5,
            modelrisk_loaded=True,
        )
        out = workflows.diagnose_workbook()
        assert any("VoseOutput" in i for i in out["issues"])

    def test_no_vmrs_flagged(self, bridge: MagicMock, tmp_path: Path) -> None:
        from modelrisk_mcp.config import Settings

        wb_path = tmp_path / "lonely.xlsx"
        wb_path.write_text("")
        bridge.excel.get_active_workbook.return_value = WorkbookInfo(
            name="lonely.xlsx", path=str(wb_path), sheets=["S1"],
        )
        bridge.is_modelrisk_loaded.return_value = True
        bridge.get_workbook_summary.return_value = WorkbookSummary(
            workbook="lonely.xlsx", sheets=["S1"],
            input_count=1, output_count=1, distribution_count=1,
            formula_cell_count=5, numeric_cell_count=2,
            modelrisk_loaded=True,
        )
        bridge._settings = Settings()
        out = workflows.diagnose_workbook()
        assert out["vmrs_exists"] is False
        assert any(".vmrs" in i for i in out["issues"])

    def test_excel_not_reachable_short_circuits(
        self, bridge: MagicMock
    ) -> None:
        """If Excel can't be reached, downstream checks should be skipped
        and the result should surface the issue cleanly."""
        from modelrisk_mcp.config import Settings

        bridge._settings = Settings()
        bridge.excel.get_active_workbook.side_effect = RuntimeError(
            "Excel not running"
        )
        out = workflows.diagnose_workbook()
        assert out["excel_connected"] is False
        assert any("Excel not reachable" in i for i in out["issues"])
        # Downstream bridge calls must not have been invoked.
        bridge.is_modelrisk_loaded.assert_not_called()
        bridge.get_workbook_summary.assert_not_called()

    def test_modelrisk_not_activated_flagged(
        self, bridge: MagicMock, tmp_path: Path
    ) -> None:
        self._setup_happy_path(bridge, tmp_path)
        bridge.is_modelrisk_loaded.return_value = False
        out = workflows.diagnose_workbook()
        assert out["modelrisk_loaded"] is False
        assert any("not activated" in i for i in out["issues"])

    def test_distributions_without_outputs(
        self, bridge: MagicMock, tmp_path: Path
    ) -> None:
        """A workbook with VoseOutputs but no distribution cells produces
        constant simulation results — flagged."""
        self._setup_happy_path(bridge, tmp_path)
        bridge.get_workbook_summary.return_value = WorkbookSummary(
            workbook="m.xlsx", sheets=["S1"],
            input_count=0, output_count=2, distribution_count=0,
            formula_cell_count=5, numeric_cell_count=10,
            modelrisk_loaded=True,
        )
        out = workflows.diagnose_workbook()
        assert any("no Vose distribution" in i for i in out["issues"])


# Quieten unused-imports for types declared via fixtures.
_ = Any
