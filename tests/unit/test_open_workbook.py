"""Unit tests for the open_workbook tool + ExcelBridge.open_workbook."""

from __future__ import annotations

import pytest

from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.errors import WorkbookNotFoundError
from modelrisk_mcp.schemas.workbook import WorkbookInfo
from modelrisk_mcp.tools import reading


def test_open_workbook_tool_delegates() -> None:
    class _Excel:
        def __init__(self) -> None:
            self.opened: str | None = None

        def open_workbook(self, path: str) -> WorkbookInfo:
            self.opened = path
            return WorkbookInfo(
                name="risk.xlsx", path=path, sheets=["Sheet1"], active_sheet="Sheet1"
            )

    class _Bridge:
        def __init__(self, excel: _Excel) -> None:
            self.excel = excel

    excel = _Excel()
    reading.set_bridge_for_testing(_Bridge(excel))  # type: ignore[arg-type]
    try:
        info = reading.open_workbook(r"C:\models\risk.xlsx")
        assert excel.opened == r"C:\models\risk.xlsx"
        assert info.name == "risk.xlsx"
        assert info.path == r"C:\models\risk.xlsx"
    finally:
        reading.set_bridge_for_testing(None)


def test_open_workbook_missing_file_raises(monkeypatch: pytest.MonkeyPatch) -> None:
    bridge = ExcelBridge()

    class _App:
        def __init__(self) -> None:
            self.books: list[object] = []

    # Bypass the real COM connect; the file-existence guard runs first.
    monkeypatch.setattr(bridge, "_ensure", lambda: _App())
    with pytest.raises(WorkbookNotFoundError, match="File not found"):
        bridge.open_workbook(r"C:\nope\does_not_exist_4f1c9.xlsx")


def test_close_workbook_tool_delegates() -> None:
    class _Excel:
        def __init__(self) -> None:
            self.closed: tuple[str, bool] | None = None

        def close_workbook(self, workbook: str, save: bool = False) -> dict[str, object]:
            self.closed = (workbook, save)
            return {"closed": workbook, "saved": save, "open_workbooks": []}

    class _Bridge:
        def __init__(self, excel: _Excel) -> None:
            self.excel = excel

    excel = _Excel()
    reading.set_bridge_for_testing(_Bridge(excel))  # type: ignore[arg-type]
    try:
        out = reading.close_workbook("risk.xlsx", save=True)
        assert excel.closed == ("risk.xlsx", True)
        assert out["closed"] == "risk.xlsx" and out["saved"] is True
    finally:
        reading.set_bridge_for_testing(None)


def test_close_workbook_not_open_raises(monkeypatch: pytest.MonkeyPatch) -> None:
    bridge = ExcelBridge()

    class _Books:
        def __getitem__(self, key: str) -> object:
            raise KeyError(key)

    class _App:
        def __init__(self) -> None:
            self.books = _Books()

    monkeypatch.setattr(bridge, "_ensure", lambda: _App())
    with pytest.raises(WorkbookNotFoundError):
        bridge.close_workbook("not_open.xlsx")


def test_open_workbook_returns_already_open(monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
    # An existing file whose basename matches an already-open book is returned
    # without re-opening.
    f = tmp_path / "already.xlsx"
    f.write_text("x")

    class _Book:
        def __init__(self) -> None:
            self.name = "already.xlsx"
            self.sheets: list[object] = []

    class _App:
        def __init__(self) -> None:
            self.books = [_Book()]

    bridge = ExcelBridge()
    monkeypatch.setattr(bridge, "_ensure", lambda: _App())
    monkeypatch.setattr(
        bridge,
        "_workbook_info",
        lambda book: WorkbookInfo(name=book.name, path=str(f), sheets=[], active_sheet=None),
    )
    info = bridge.open_workbook(str(f))
    assert info.name == "already.xlsx"
