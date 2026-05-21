"""Pytest fixtures for the Excel-required integration suite.

Integration tests are gated: when Excel isn't running (or xlwings can't
attach to it), every test in this suite is skipped with a clear reason
instead of failing. CI runs them automatically only on release tags;
PRs trigger them manually.
"""

from __future__ import annotations

from collections.abc import Iterator

import pytest

from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.errors import ExcelNotRunningError


@pytest.fixture(scope="session")
def excel_bridge() -> Iterator[ExcelBridge]:
    """Attach to a running Excel. Skips all tests in this suite if Excel
    isn't running, since these tests cannot run without it."""
    bridge = ExcelBridge()
    try:
        bridge.connect()
    except ExcelNotRunningError as exc:
        pytest.skip(
            f"Excel is not running, skipping integration tests: {exc}",
            allow_module_level=True,
        )
    yield bridge
    bridge.disconnect()


@pytest.fixture(scope="session")
def modelrisk_loaded(excel_bridge: ExcelBridge) -> bool:
    """True if MRService.dll loads and activates (bundled key in v0.3+,
    or via MRSERVICE_ACTIVATION_KEY env override). Tests that read
    `.vmrs` files skip if this is False."""
    bridge = ModelRiskBridge(excel_bridge)
    return bridge.is_modelrisk_loaded()


@pytest.fixture(scope="session")
def modelrisk_bridge(
    excel_bridge: ExcelBridge, modelrisk_loaded: bool
) -> ModelRiskBridge:
    if not modelrisk_loaded:
        pytest.skip(
            "MRService.dll did not activate — install the ModelRisk SDK "
            "(MRService.dll resolvable via MRSERVICE_DLL or one of the "
            "standard install paths) before running ModelRisk-dependent "
            "integration tests."
        )
    return ModelRiskBridge(excel_bridge)
