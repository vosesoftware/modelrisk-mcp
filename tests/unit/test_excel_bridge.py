"""Tests for ExcelBridge focused on graceful degradation.

Targets the OneDrive path-resolution failure mode discovered in real-world
testing: xlwings raises when accessing `.fullname` on a workbook stored in
OneDrive for Business without the `ONEDRIVE_COMMERCIAL_WIN` env var. The
bridge must still return a usable WorkbookInfo (name-only, empty path) so
downstream name-based operations work."""

from __future__ import annotations

from types import SimpleNamespace
from typing import Any

import pytest

from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.errors import WorkbookNotFoundError


class _FakeSheet:
    def __init__(self, name: str) -> None:
        self.name = name


class _FakeSheets:
    def __init__(self, names: list[str], active: str | None) -> None:
        self._sheets = [_FakeSheet(n) for n in names]
        self.active = _FakeSheet(active) if active else None

    def __iter__(self) -> Any:
        return iter(self._sheets)


class _FakeBook:
    """Minimal stand-in for an xlwings Book."""

    def __init__(
        self,
        name: str,
        *,
        fullname: str | Exception = "",
        sheet_names: list[str] | Exception | None = None,
        active_sheet: str | None = "Sheet1",
    ) -> None:
        self.name = name
        self._fullname = fullname
        self._sheet_names = sheet_names if sheet_names is not None else ["Sheet1"]
        self._active_sheet = active_sheet

    @property
    def fullname(self) -> str:
        if isinstance(self._fullname, Exception):
            raise self._fullname
        return self._fullname

    @property
    def sheets(self) -> _FakeSheets:
        if isinstance(self._sheet_names, Exception):
            raise self._sheet_names
        return _FakeSheets(self._sheet_names, self._active_sheet)


class _FakeBooks:
    def __init__(self, books: list[_FakeBook], active: _FakeBook | None) -> None:
        self._books = books
        self.active = active
        self.count = len(books)

    def __iter__(self) -> Any:
        return iter(self._books)

    def __getitem__(self, name: str) -> _FakeBook:
        for b in self._books:
            if b.name == name:
                return b
        raise KeyError(name)


class _FakeApp:
    def __init__(self, books: _FakeBooks, *, api: Any = None) -> None:
        self.books = books
        self.api = api or SimpleNamespace()


def _make_bridge(app: _FakeApp) -> ExcelBridge:
    bridge = ExcelBridge()
    bridge._app = app  # type: ignore[assignment]
    bridge._xlwings = object()  # bypass connect()
    return bridge


class TestOneDriveFallback:
    def test_fullname_failure_yields_empty_path(self) -> None:
        """When `book.fullname` raises (xlwings OneDrive bug), the
        bridge should still return a WorkbookInfo with the name and an
        empty path — not propagate the exception."""
        book = _FakeBook(
            "OneDriveModel.xlsx",
            fullname=RuntimeError(
                "could not find path. Please set the environment "
                "variable ONEDRIVE_COMMERCIAL_WIN."
            ),
            sheet_names=["Inputs", "Outputs"],
            active_sheet="Inputs",
        )
        app = _FakeApp(_FakeBooks([book], active=book))
        bridge = _make_bridge(app)

        info = bridge.get_active_workbook()
        assert info.name == "OneDriveModel.xlsx"
        assert info.path == ""
        assert info.sheets == ["Inputs", "Outputs"]
        assert info.active_sheet == "Inputs"

    def test_list_workbooks_skips_unintrospectable_books(self) -> None:
        """One blown-up workbook shouldn't break enumeration of others."""
        good = _FakeBook("good.xlsx", fullname=r"C:\models\good.xlsx")
        bad = _FakeBook("bad.xlsx", sheet_names=RuntimeError("hosed"))
        # `bad` will still come through with empty sheets list now (we
        # gracefully degrade rather than skip). Just confirm `good` is
        # present and the call doesn't raise.
        app = _FakeApp(_FakeBooks([good, bad], active=good))
        bridge = _make_bridge(app)
        names = [w.name for w in bridge.list_workbooks()]
        assert "good.xlsx" in names

    def test_active_falls_back_to_com_api(self) -> None:
        """If xlwings' `books.active` raises, fall back to
        `app.api.ActiveWorkbook.Name` — the COM API doesn't touch
        FullName."""
        class _RaisingBooks:
            count = 1

            @property
            def active(self) -> Any:
                raise RuntimeError("xlwings stumbled")

            def __iter__(self) -> Any:
                return iter([])

        active_workbook = SimpleNamespace(Name="ComOnly.xlsx")
        api = SimpleNamespace(ActiveWorkbook=active_workbook)
        app = SimpleNamespace(books=_RaisingBooks(), api=api)

        bridge = ExcelBridge()
        bridge._app = app  # type: ignore[assignment]
        bridge._xlwings = object()
        # Skip stale-check (it accesses books.count, which is fine).

        info = bridge.get_active_workbook()
        assert info.name == "ComOnly.xlsx"
        assert info.path == ""
        assert info.sheets == []

    def test_no_active_workbook_raises(self) -> None:
        app = _FakeApp(_FakeBooks([], active=None))
        bridge = _make_bridge(app)
        with pytest.raises(WorkbookNotFoundError):
            bridge.get_active_workbook()


class TestPathDegradation:
    def test_fullname_succeeds_when_available(self) -> None:
        book = _FakeBook(
            "local.xlsx",
            fullname=r"C:\models\local.xlsx",
        )
        app = _FakeApp(_FakeBooks([book], active=book))
        bridge = _make_bridge(app)
        info = bridge.get_active_workbook()
        assert info.path == r"C:\models\local.xlsx"
