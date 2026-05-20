"""Phase 3 building-tool integration tests.

Each test exercises a real Excel — the conftest fixture skips the suite
if Excel isn't running. We use a scratch cell we choose deliberately
(`ZZ999`) so we never collide with anything the user might have in
their workbook.

Acceptance criteria from spec §13 Phase 3 covered here:
- Default dry_run=True doesn't mutate.
- dry_run=False writes through to Excel.
- The write lands in Excel's undo stack (verified by Excel.Undo).
- The audit log is created in %LOCALAPPDATA%\\VoseSoftware\\modelrisk-mcp.
- restore_cell round-trip recovers the pre-write state.
"""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

import pytest

from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.config import Settings
from modelrisk_mcp.safety import WriterMutex
from modelrisk_mcp.tools import building, reading, restore

# A cell we use as scratch — far enough from typical model content that
# we won't clobber anything important. Tests clean up after themselves.
SCRATCH_CELL = "ZZ999"


@pytest.fixture
def scratch_bridge(
    excel_bridge: ExcelBridge, tmp_path: Path
) -> Iterator[tuple[ModelRiskBridge, Path, str, str]]:
    """Build a ModelRiskBridge that writes its audit log into tmp_path
    and yields (bridge, audit_log, workbook_name, sheet)."""
    workbooks = excel_bridge.list_workbooks()
    if not workbooks:
        pytest.skip("No workbooks open in Excel.")
    wb = workbooks[0]
    sheet = wb.active_sheet or (wb.sheets[0] if wb.sheets else "Sheet1")
    audit_log = tmp_path / "writes.log"
    settings = Settings(log_dir=tmp_path, writes_log_name="writes.log")
    bridge = ModelRiskBridge(
        excel=excel_bridge,
        settings=settings,
        writer_mutex=WriterMutex(name="modelrisk-mcp-test-phase3-integration"),
    )
    reading.set_bridge_for_testing(bridge)
    # Save pre-state of the scratch cell so we can fully restore.
    pre = excel_bridge.get_cell(wb.name, sheet, SCRATCH_CELL)
    try:
        yield bridge, audit_log, wb.name, sheet
    finally:
        try:
            excel_bridge.write_cell(wb.name, sheet, SCRATCH_CELL, pre.formula)
        except Exception:
            pass
        reading.set_bridge_for_testing(None)


def test_dry_run_does_not_mutate(
    scratch_bridge: tuple[ModelRiskBridge, Path, str, str],
    excel_bridge: ExcelBridge,
) -> None:
    _, audit_log, workbook, sheet = scratch_bridge
    pre = excel_bridge.get_cell(workbook, sheet, SCRATCH_CELL)
    result = building.insert_distribution(
        workbook, sheet, SCRATCH_CELL,
        "VoseNormal",
        [{"name": "mu", "value": 0}, {"name": "sigma", "value": 1}],
    )
    assert result.written is False
    post = excel_bridge.get_cell(workbook, sheet, SCRATCH_CELL)
    assert post.formula == pre.formula
    assert not audit_log.exists() or audit_log.read_text(encoding="utf-8") == ""


def test_write_then_undo_round_trip(
    scratch_bridge: tuple[ModelRiskBridge, Path, str, str],
    excel_bridge: ExcelBridge,
) -> None:
    """The write must land in Excel's undo stack (spec §13 Phase 3
    acceptance)."""
    _, audit_log, workbook, sheet = scratch_bridge
    pre = excel_bridge.get_cell(workbook, sheet, SCRATCH_CELL)
    # Commit.
    building.insert_distribution(
        workbook, sheet, SCRATCH_CELL,
        "VoseNormal",
        [{"name": "mu", "value": 0}, {"name": "sigma", "value": 1}],
        dry_run=False,
    )
    written = excel_bridge.get_cell(workbook, sheet, SCRATCH_CELL)
    assert written.formula == "=VoseNormal(0,1)"
    # Excel.Undo restores the pre-state.
    excel_bridge.undo()
    after_undo = excel_bridge.get_cell(workbook, sheet, SCRATCH_CELL)
    assert after_undo.formula == pre.formula
    # Audit log captured the commit.
    assert audit_log.exists()
    contents = audit_log.read_text(encoding="utf-8").strip()
    assert "VoseNormal(0,1)" in contents


def test_restore_cell_round_trip(
    scratch_bridge: tuple[ModelRiskBridge, Path, str, str],
    excel_bridge: ExcelBridge,
) -> None:
    """restore_cell from audit log recovers the pre-write state — even
    after Excel's undo stack has been used or cleared."""
    _, _, workbook, sheet = scratch_bridge
    # Set a known starting formula.
    excel_bridge.write_cell(workbook, sheet, SCRATCH_CELL, "=VoseModPERT(1,2,3)")
    # Overwrite via the building tool.
    building.insert_distribution(
        workbook, sheet, SCRATCH_CELL,
        "VoseNormal",
        [{"name": "mu", "value": 0}, {"name": "sigma", "value": 1}],
        dry_run=False,
    )
    assert (
        excel_bridge.get_cell(workbook, sheet, SCRATCH_CELL).formula
        == "=VoseNormal(0,1)"
    )
    # Restore via the audit log.
    result = restore.restore_cell(workbook, sheet, SCRATCH_CELL)
    assert result.written is True
    assert result.formula == "=VoseModPERT(1,2,3)"
    assert (
        excel_bridge.get_cell(workbook, sheet, SCRATCH_CELL).formula
        == "=VoseModPERT(1,2,3)"
    )
