"""Integration tests for the §7.1 reading tools.

Runs the tools through their MCP entry points against a real Excel.
The conftest fixture skips the suite if Excel isn't running. Each test
soft-skips additionally if a specific precondition (workbook open,
ModelRisk loaded, simulation run) isn't met — these tests want to
exercise real behaviour when the runner has it, not insist on a
specific environment.
"""

from __future__ import annotations

from collections.abc import Iterator

import pytest

from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.tools import reading


@pytest.fixture(autouse=True)
def _install_bridge(excel_bridge: ExcelBridge) -> Iterator[None]:
    bridge = ModelRiskBridge(excel_bridge)
    reading.set_bridge_for_testing(bridge)
    yield
    reading.set_bridge_for_testing(None)


def test_list_open_workbooks_returns_list(excel_bridge: ExcelBridge) -> None:
    """§13 Phase 2 acceptance: this is the headline reading tool."""
    result = reading.list_open_workbooks()
    assert isinstance(result, list)


def test_get_workbook_summary_runs(excel_bridge: ExcelBridge) -> None:
    workbooks = reading.list_open_workbooks()
    if not workbooks:
        pytest.skip("No workbooks open in Excel.")
    summary = reading.get_workbook_summary(workbooks[0].name)
    assert summary.workbook == workbooks[0].name
    assert summary.sheets == workbooks[0].sheets
    # Counts must be non-negative integers; we don't assert specific
    # values since we don't control the user's workbook content.
    assert summary.input_count >= 0
    assert summary.output_count >= 0
    assert summary.distribution_count >= 0


def test_list_modelrisk_inputs_runs(excel_bridge: ExcelBridge) -> None:
    workbooks = reading.list_open_workbooks()
    if not workbooks:
        pytest.skip("No workbooks open in Excel.")
    inputs = reading.list_modelrisk_inputs(workbooks[0].name)
    assert isinstance(inputs, list)
    for inp in inputs:
        assert inp.name
        assert inp.ref.cell


def test_list_distributions_runs(excel_bridge: ExcelBridge) -> None:
    workbooks = reading.list_open_workbooks()
    if not workbooks:
        pytest.skip("No workbooks open in Excel.")
    dists = reading.list_distributions(workbooks[0].name)
    assert isinstance(dists, list)
    for d in dists:
        # Every reported function name must be one the catalogue knows.
        bridge = reading.get_bridge()
        assert d.function_name in bridge.catalogue, (
            f"Reported function {d.function_name!r} is not in the catalogue."
        )


def test_find_hard_coded_inputs_runs(excel_bridge: ExcelBridge) -> None:
    workbooks = reading.list_open_workbooks()
    if not workbooks:
        pytest.skip("No workbooks open in Excel.")
    candidates = reading.find_hard_coded_inputs(workbooks[0].name)
    assert isinstance(candidates, list)
    for c in candidates:
        assert {"workbook", "sheet", "cell"} <= set(c.keys())


def test_get_simulation_results_skips_without_modelrisk(
    excel_bridge: ExcelBridge, modelrisk_loaded: bool
) -> None:
    if not modelrisk_loaded:
        pytest.skip("ModelRisk not loaded.")
    workbooks = reading.list_open_workbooks()
    if not workbooks:
        pytest.skip("No workbooks open in Excel.")
    # No assertion on contents — if no simulation has run, the list will
    # be empty. The test confirms the call doesn't raise.
    results = reading.get_simulation_results(workbooks[0].name)
    assert isinstance(results, list)
