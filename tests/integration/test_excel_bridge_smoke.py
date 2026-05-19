"""Smoke tests for `ExcelBridge` — the §13 Phase 1 acceptance test.

These run against a real Excel instance. The conftest fixture skips
the whole suite if Excel isn't running.
"""

from __future__ import annotations

import pytest

from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.errors import (
    CellReferenceError,
    WorkbookNotFoundError,
)


def test_list_workbooks_returns_list(excel_bridge: ExcelBridge) -> None:
    """Phase 1 acceptance criterion (spec §13)."""
    workbooks = excel_bridge.list_workbooks()
    assert isinstance(workbooks, list)
    # If the user has any workbook open, each should have a name and a
    # sheets list. If no workbooks are open, the list is empty — also valid.
    for wb in workbooks:
        assert isinstance(wb.name, str)
        assert wb.name
        assert isinstance(wb.sheets, list)


def test_get_active_workbook_or_raises(excel_bridge: ExcelBridge) -> None:
    """If no workbook is open, must raise WorkbookNotFoundError, not
    return None or leak a COM error."""
    if not excel_bridge.list_workbooks():
        with pytest.raises(WorkbookNotFoundError):
            excel_bridge.get_active_workbook()
        return
    wb = excel_bridge.get_active_workbook()
    assert wb.name
    assert wb.active_sheet


def test_get_cell_invalid_ref_raises(excel_bridge: ExcelBridge) -> None:
    workbooks = excel_bridge.list_workbooks()
    if not workbooks:
        pytest.skip("No workbooks open; cannot test cell access.")
    wb = workbooks[0]
    with pytest.raises(CellReferenceError):
        excel_bridge.get_cell(wb.name, wb.sheets[0], "NOT-A-CELL")


def test_unknown_workbook_raises(excel_bridge: ExcelBridge) -> None:
    with pytest.raises(WorkbookNotFoundError):
        excel_bridge.get_cell("definitely-not-a-real-workbook.xlsx", "Sheet1", "A1")


def test_reconnect_after_disconnect(excel_bridge: ExcelBridge) -> None:
    assert excel_bridge.is_connected()
    excel_bridge.disconnect()
    assert not excel_bridge.is_connected()
    # First call after disconnect should lazily re-attach.
    excel_bridge.list_workbooks()
    assert excel_bridge.is_connected()
