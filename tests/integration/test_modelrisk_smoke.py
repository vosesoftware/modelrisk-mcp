"""Smoke tests for `ModelRiskBridge` — requires Excel + ModelRisk."""

from __future__ import annotations

from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge


def test_is_modelrisk_loaded(modelrisk_bridge: ModelRiskBridge) -> None:
    # Bridge is only injected when this is True, so this test is a no-op
    # confirmation that the fixture wiring works.
    assert modelrisk_bridge.is_modelrisk_loaded() is True


def test_lists_are_empty_or_lists(modelrisk_bridge: ModelRiskBridge) -> None:
    """If a workbook is open, lists are valid. If no workbook is open,
    the iterator yields nothing — both legal states."""
    bridge = modelrisk_bridge
    excel = bridge._excel
    workbooks = excel.list_workbooks()
    if not workbooks:
        # No workbooks open — nothing to inspect.
        return
    wb = workbooks[0].name
    inputs = bridge.list_inputs(wb)
    outputs = bridge.list_outputs(wb)
    dists = bridge.list_distributions(wb)
    assert isinstance(inputs, list)
    assert isinstance(outputs, list)
    assert isinstance(dists, list)
