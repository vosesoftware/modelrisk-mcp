"""Phase 4 simulation integration tests.

Run an end-to-end Monte Carlo simulation against real Excel + ModelRisk
and confirm:
- set_simulation_settings lands on ModelRiskSimulationSettings
- run_simulation executes StartSimulation and blocks until it returns
- get_simulation_results returns sane percentiles after a small run
- stop_simulation raises with the documented message until the COM
  endpoint lands
- get_simulation_status reports idle outside a run
"""

from __future__ import annotations

from collections.abc import Iterator

import pytest

from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.errors import SimulationNotAvailableError
from modelrisk_mcp.tools import reading, simulation


@pytest.fixture
def bridge_for_sim(
    excel_bridge: ExcelBridge, modelrisk_loaded: bool
) -> Iterator[ModelRiskBridge]:
    if not modelrisk_loaded:
        pytest.skip("ModelRisk not loaded.")
    bridge = ModelRiskBridge(excel=excel_bridge)
    reading.set_bridge_for_testing(bridge)
    yield bridge
    reading.set_bridge_for_testing(None)


def test_stop_raises_until_com_lands(
    bridge_for_sim: ModelRiskBridge,
) -> None:
    """Documented stub behaviour until ModelRisk adds StopSimulation."""
    with pytest.raises(SimulationNotAvailableError):
        simulation.stop_simulation()


def test_status_idle_outside_run(bridge_for_sim: ModelRiskBridge) -> None:
    status = simulation.get_simulation_status()
    assert status.status == "idle"


def test_set_settings_returns_applied_dict(
    bridge_for_sim: ModelRiskBridge,
) -> None:
    response = simulation.set_simulation_settings(
        samples=100, hide_progress_window=True
    )
    assert response.applied == {
        "samples": 100,
        "hide_progress_window": True,
    }


def test_run_and_read_results_end_to_end(
    bridge_for_sim: ModelRiskBridge,
    excel_bridge: ExcelBridge,
) -> None:
    """Spec §13 Phase 4 acceptance: run a 1000-iteration sim end-to-end
    and assert percentiles are sane for a fixture model."""
    workbooks = excel_bridge.list_workbooks()
    if not workbooks:
        pytest.skip("No workbooks open in Excel.")
    wb = workbooks[0]
    outputs = bridge_for_sim.list_outputs(wb.name)
    if not outputs:
        pytest.skip(
            "Active workbook has no VoseOutput cells; nothing to read."
        )
    # Hide the progress window so the test doesn't interrupt the user.
    simulation.set_simulation_settings(hide_progress_window=True)
    run_response = simulation.run_simulation(iterations=1000, seed=42)
    assert run_response.succeeded is True
    results = reading.get_simulation_results(wb.name)
    assert len(results) > 0
    for r in results:
        assert r.iterations >= 1  # at least one iteration recorded
        assert r.min <= r.percentiles.get(0.05, r.min)
        assert r.percentiles.get(0.05, r.min) <= r.percentiles.get(0.95, r.max)
        assert r.percentiles.get(0.95, r.max) <= r.max
