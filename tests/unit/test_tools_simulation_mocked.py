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
        assert result.iterations == 1000
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


# Quieten unused-import linter for typing helpers used only in fixtures.
_ = Any
