"""MCP-wrapper tests for `tools/simulation.py`.

These exist because the bridge layer (`bridge/simulation.py`) is
exercised by `test_simulation_controller.py`. What this file guards is
the I/O between the MCP tool function and the bridge: argument shape,
defaults, return-schema translation, and the auto-pinning of the
resulting `.vmrs` as the active source.

A typo like `bridge.run_simulation(workbook=workbook_name, samples=samples)`
vs `bridge.run_simulation(workbook_name, samples)` would only surface
at end-user runtime without these tests.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import Any
from unittest.mock import MagicMock

import pytest

from modelrisk_mcp.bridge.simulation import SimulationOptions, SimulationRunResult
from modelrisk_mcp.tools import reading, simulation


@pytest.fixture
def mock_bridge() -> Iterator[MagicMock]:
    bridge = MagicMock()
    bridge.run_simulation.return_value = SimulationRunResult(
        workbook_name="model.xlsx",
        vmrs_path=r"C:\models\model.vmrs",
        iterations=1000,
        options=SimulationOptions(),
    )
    reading.set_bridge_for_testing(bridge)  # type: ignore[arg-type]
    yield bridge
    reading.set_bridge_for_testing(None)


class TestRunSimulationToolPassthrough:
    def test_defaults_match_spec(self, mock_bridge: MagicMock) -> None:
        """No args → samples=1000, seed=1, save_to=None, workbook=None.
        These defaults are part of the documented MCP surface."""
        result = simulation.run_simulation()
        assert isinstance(result, simulation.RunSimulationResult)
        mock_bridge.run_simulation.assert_called_once_with(
            workbook=None, samples=1000, seed=1, save_to=None
        )

    def test_workbook_name_forwarded_as_keyword(
        self, mock_bridge: MagicMock
    ) -> None:
        simulation.run_simulation(workbook_name="my.xlsx")
        kwargs = mock_bridge.run_simulation.call_args.kwargs
        # The bridge method's parameter is `workbook`, not `workbook_name` —
        # this is exactly the kind of rename that's only caught here.
        assert kwargs["workbook"] == "my.xlsx"

    def test_all_args_forwarded(self, mock_bridge: MagicMock) -> None:
        simulation.run_simulation(
            workbook_name="x.xlsx", samples=5000, seed=42, save_to=r"D:\x.vmrs"
        )
        mock_bridge.run_simulation.assert_called_once_with(
            workbook="x.xlsx", samples=5000, seed=42, save_to=r"D:\x.vmrs"
        )

    def test_response_contains_resolved_vmrs_path(
        self, mock_bridge: MagicMock
    ) -> None:
        result = simulation.run_simulation()
        assert result.vmrs_path == r"C:\models\model.vmrs"
        # alpha.11 removed `iterations` from the response shape; use
        # `samples` (the canonical name matching ModelRisk's UI).
        assert result.samples == 1000
        assert result.workbook_name == "model.xlsx"

    def test_response_echoes_request_samples_and_seed(
        self, mock_bridge: MagicMock
    ) -> None:
        """The MCP response includes the requested samples/seed so the
        caller can confirm what was actually run (the bridge result
        carries `iterations`, but the tool surfaces the request too)."""
        result = simulation.run_simulation(samples=2500, seed=99)
        assert result.samples == 2500
        assert result.seed == 99

    def test_next_step_hint_present(self, mock_bridge: MagicMock) -> None:
        """The response includes a `next_step` field — that's how the
        LLM client knows to call get_simulation_results next."""
        result = simulation.run_simulation()
        assert "get_simulation_results" in result.next_step.lower()

    def test_bridge_exception_propagates(self, mock_bridge: MagicMock) -> None:
        """Errors from the bridge are not swallowed by the tool wrapper —
        the MCP layer translates them to JSON-RPC errors itself."""
        from modelrisk_mcp.errors import SimulationFailedError

        mock_bridge.run_simulation.side_effect = SimulationFailedError("nope")
        with pytest.raises(SimulationFailedError, match="nope"):
            simulation.run_simulation()


class TestSamplesValidation:
    """Pydantic validates `samples` to ge=1, le=1_000_000. The tool
    function itself doesn't enforce — that happens at the MCP boundary
    via the Annotated[..., Field(ge=, le=)] metadata. We can't easily
    invoke that without going through FastMCP, so we just confirm the
    metadata is present (a regression here would silently allow
    nonsense like negative samples)."""

    def test_samples_field_constraints_present(self) -> None:
        from typing import get_type_hints

        hints = get_type_hints(simulation.run_simulation, include_extras=True)
        samples_anno = hints["samples"]
        # Annotated[int, Field(...)] — `__metadata__` carries the Field.
        meta = getattr(samples_anno, "__metadata__", ())
        assert meta, "samples parameter must use Annotated[int, Field(...)]"
        # The metadata entry should be a Pydantic FieldInfo carrying the
        # numeric constraints. Pydantic stores them in `.metadata` as a
        # list of constraint objects.
        field_info = meta[0]
        constraint_repr = repr(field_info)
        assert "ge=1" in constraint_repr, (
            f"samples Field missing ge=1 constraint: {constraint_repr}"
        )
        assert "le=1000000" in constraint_repr, (
            f"samples Field missing le=1_000_000 constraint: {constraint_repr}"
        )


class TestBridgeIntegration:
    """One end-to-end shape test that wires the real ModelRiskBridge.
    Confirms the tool calls the right method on the bridge — not just
    on a generic Mock."""

    def test_calls_real_bridge_method_name(self) -> None:
        from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge

        assert hasattr(ModelRiskBridge, "run_simulation"), (
            "ModelRiskBridge must expose run_simulation — the simulation "
            "tool wrapper depends on this exact method name."
        )

    def test_bridge_method_accepts_expected_kwargs(self) -> None:
        """`bridge.run_simulation(workbook=, samples=, seed=, save_to=)` —
        all four must be keyword-acceptable."""
        from inspect import signature

        from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge

        sig = signature(ModelRiskBridge.run_simulation)
        params = set(sig.parameters)
        assert {"workbook", "save_to"}.issubset(params), (
            f"bridge.run_simulation missing expected kwargs; has: {params}"
        )


class TestRunScenariosToolPassthrough:
    @pytest.fixture
    def scenario_bridge(self) -> Iterator[MagicMock]:
        from modelrisk_mcp.schemas.results import ScenarioSweepResult

        bridge = MagicMock()
        bridge.run_scenarios.return_value = ScenarioSweepResult(
            workbook_name="m.xlsx", sheet="S1", cell="B2",
            original_formula="=100",
            samples_per_scenario=1000,
        )
        reading.set_bridge_for_testing(bridge)  # type: ignore[arg-type]
        yield bridge
        reading.set_bridge_for_testing(None)

    def test_run_scenarios_passes_args_through(
        self, scenario_bridge: MagicMock
    ) -> None:
        simulation.run_scenarios(
            sheet="S1", cell="B2", values=[50.0, 75.0, 100.0],
        )
        scenario_bridge.run_scenarios.assert_called_once_with(
            "S1", "B2", [50.0, 75.0, 100.0],
            workbook=None, samples=1000, seed=1,
        )

    def test_run_scenarios_forwards_overrides(
        self, scenario_bridge: MagicMock
    ) -> None:
        simulation.run_scenarios(
            sheet="In", cell="A1", values=[1.0, 2.0],
            samples=500, seed=42, workbook_name="big.xlsx",
        )
        scenario_bridge.run_scenarios.assert_called_once_with(
            "In", "A1", [1.0, 2.0],
            workbook="big.xlsx", samples=500, seed=42,
        )


class TestRunScenariosBridge:
    """Integration of the run_scenarios bridge method with fake Excel +
    fake results reader. Verifies the original-formula restoration
    contract — the cell MUST end with its pre-call formula even when a
    scenario raises mid-sweep."""

    def _build_bridge(self) -> tuple[Any, Any]:
        """Returns (bridge, fake_excel) wired to a writable in-memory
        Excel and a stub results reader that always returns one output."""
        from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
        from modelrisk_mcp.schemas.results import SimulationResult
        from modelrisk_mcp.schemas.workbook import (
            CellInfo,
            CellRef,
            ModelRiskOutput,
        )

        class FakeExcel:
            def __init__(self) -> None:
                self.cells: dict[tuple[str, str, str], str] = {}
                self.write_log: list[tuple[str, str, str, str]] = []

            def get_active_workbook(self) -> Any:
                from modelrisk_mcp.schemas.workbook import WorkbookInfo
                return WorkbookInfo(name="m.xlsx", path="C:/m.xlsx", sheets=["S1"])

            def get_cell(self, wb: str, sh: str, c: str) -> CellInfo:
                key = (wb, sh, c)
                formula = self.cells.get(key, "")
                return CellInfo(
                    ref=CellRef(workbook=wb, sheet=sh, cell=c),
                    formula=formula, value=None,
                    cell_type="formula" if formula else "empty",
                )

            def write_cell(self, wb: str, sh: str, c: str, f: str) -> None:
                self.cells[(wb, sh, c)] = f
                self.write_log.append((wb, sh, c, f))

            def iterate_cells(self, *a: Any, **k: Any) -> Any:
                return iter([])

            def list_workbooks(self) -> list[Any]:
                return []

        fake_excel = FakeExcel()
        fake_excel.cells[("m.xlsx", "S1", "B2")] = "=VoseNormal(100,20)"

        # Stub the results reader so we don't need a real .vmrs.
        bridge = ModelRiskBridge(excel=fake_excel)  # type: ignore[arg-type]
        bridge._results.get_simulation_results = lambda *a, **k: [  # type: ignore[method-assign]
            SimulationResult(
                output_name="profit", iterations=100, mean=50.0, stdev=10.0,
                variance=100.0, skewness=0, kurtosis=0, min=20.0, max=80.0,
                percentiles={0.05: 30.0, 0.5: 50.0, 0.95: 70.0},
            )
        ]
        # Stub list_outputs so we don't need to scan cells.
        bridge.list_outputs = lambda wb: [  # type: ignore[method-assign]
            ModelRiskOutput(
                ref=CellRef(workbook=wb, sheet="S1", cell="C1"),
                name="profit", formula='=VoseOutput("profit")+B2',
                current_value=None,
            )
        ]
        # Stub run_simulation so we don't try to Application.Run.
        bridge.run_simulation = lambda **k: None  # type: ignore[method-assign]
        return bridge, fake_excel

    def test_original_formula_restored_after_sweep(self) -> None:
        bridge, fake = self._build_bridge()
        bridge.run_scenarios("S1", "B2", [50.0, 75.0, 100.0])
        # After sweep, the cell must be back to its original formula.
        assert fake.cells[("m.xlsx", "S1", "B2")] == "=VoseNormal(100,20)"

    def test_original_formula_restored_even_on_exception(self) -> None:
        bridge, fake = self._build_bridge()

        # Simulate the second scenario blowing up mid-sweep.
        call_count = [0]
        original_run = bridge.run_simulation

        def flaky_run(**k: Any) -> Any:
            call_count[0] += 1
            if call_count[0] == 2:
                raise RuntimeError("sim died")
            return original_run(**k)

        bridge.run_simulation = flaky_run  # type: ignore[method-assign]

        with pytest.raises(RuntimeError, match="sim died"):
            bridge.run_scenarios("S1", "B2", [50.0, 75.0, 100.0])

        # Cell must STILL be restored.
        assert fake.cells[("m.xlsx", "S1", "B2")] == "=VoseNormal(100,20)"

    def test_each_scenario_writes_its_value(self) -> None:
        bridge, fake = self._build_bridge()
        bridge.run_scenarios("S1", "B2", [50.0, 75.0])
        # write_log entries to B2: scenario writes + final restore.
        b2_writes = [w[3] for w in fake.write_log if w[2] == "B2"]
        assert "50.0" in b2_writes
        assert "75.0" in b2_writes
        # Last write must be the restore.
        assert b2_writes[-1] == "=VoseNormal(100,20)"

    def test_result_carries_p5_p50_p95(self) -> None:
        bridge, _ = self._build_bridge()
        result = bridge.run_scenarios("S1", "B2", [42.0])
        assert len(result.scenarios) == 1
        run = result.scenarios[0]
        assert run.scenario_value == 42.0
        assert len(run.outputs) == 1
        assert run.outputs[0].output_name == "profit"
        assert run.outputs[0].p5 == 30.0
        assert run.outputs[0].p50 == 50.0
        assert run.outputs[0].p95 == 70.0


# Quieten unused-import linter for typing helpers used only in fixtures.
_ = Any
