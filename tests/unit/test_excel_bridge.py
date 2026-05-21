"""Tests for ExcelBridge focused on graceful degradation.

Targets the OneDrive path-resolution failure mode discovered in real-world
testing: xlwings raises when accessing `.fullname` on a workbook stored in
OneDrive for Business without the `ONEDRIVE_COMMERCIAL_WIN` env var. The
bridge must still return a usable WorkbookInfo (name-only, empty path) so
downstream name-based operations work.

Also tests the `_as_2d` value-normalisation function — specifically the
tuple-vs-list collapse bug that silently broke list_modelrisk_inputs /
list_modelrisk_outputs / get_workbook_summary / run_simulation's input
registration when xlwings returned formulas as tuples (which it does for
the Windows COM `.formula` accessor)."""

from __future__ import annotations

from types import SimpleNamespace
from typing import Any

import pytest

from modelrisk_mcp.bridge.excel import ExcelBridge, _as_2d
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

    def test_unsaved_workbook_path_is_empty(self) -> None:
        """Unsaved workbooks return their NAME from Workbook.FullName
        (e.g. 'Book3') with no directory component. Previously the
        bridge propagated that string as `path` which misled downstream
        code that treated it as a filesystem location. Now we detect
        the missing path separator and report empty path."""
        book = _FakeBook("Book3", fullname="Book3", sheet_names=["Sheet1"])
        app = _FakeApp(_FakeBooks([book], active=book))
        bridge = _make_bridge(app)
        info = bridge.get_active_workbook()
        assert info.name == "Book3"
        assert info.path == "", (
            f"Unsaved workbook path should be empty, got {info.path!r}"
        )


class TestAs2dTupleHandling:
    """The list-scan collapse bug was here: xlwings on Windows returns
    `Range.formula` as a tuple of tuples (raw COM SAFEARRAY), while
    `Range.value` returns lists. Previously `_as_2d` only accepted
    `list`, so a tuple-shaped formula payload was treated as a scalar
    and the whole row's formulas were string-cast into one fake 'cell'.
    The regex would find the first `VoseInput("name")` in that string
    and we'd yield exactly one record instead of many — silently
    losing all-but-one input. This test class pins the fix."""

    def test_2d_list_passes_through_unchanged(self) -> None:
        out = _as_2d([["a", "b"], ["c", "d"]])
        assert out == [["a", "b"], ["c", "d"]]

    def test_2d_tuple_is_normalised_to_list_of_lists(self) -> None:
        """The bug-triggering case: tuples in tuples. We must produce
        a proper 2D list, not collapse to a 1x1 grid."""
        out = _as_2d((("a", "b"), ("c", "d")))
        assert out == [["a", "b"], ["c", "d"]]

    def test_1d_tuple_becomes_single_row(self) -> None:
        """xlwings often returns a single row as a 1D tuple. Must
        treat as one row, not as a scalar inside one cell."""
        out = _as_2d(("a", "b", "c"))
        assert out == [["a", "b", "c"]]

    def test_mixed_list_tuple_2d_works(self) -> None:
        """Defence in depth — accept any combo of list+tuple."""
        out = _as_2d([("a", "b"), ["c", "d"]])
        assert out == [["a", "b"], ["c", "d"]]

    def test_scalar_still_wraps_correctly(self) -> None:
        assert _as_2d("solo") == [["solo"]]
        assert _as_2d(42) == [[42]]
        assert _as_2d(None) == [[None]]

    def test_regression_three_voseinputs_in_one_row(self) -> None:
        """The exact production failure mode: a workbook with three
        cells in row 1, each containing a VoseInput-wrapped formula.
        xlwings returns formulas as a 1D tuple. _as_2d must produce
        [['=...A', '=...B', '=...C']] so the downstream iteration
        yields three cells, not one."""
        formulas = (
            '=VoseInput("A")+VoseNormal(0,1)',
            '=VoseInput("B")+VoseLognormal(50,15)',
            '=VoseInput("C")+VoseNormal(100,20)',
        )
        out = _as_2d(formulas)
        # 1 row, 3 columns.
        assert len(out) == 1
        assert len(out[0]) == 3
        assert all('VoseInput("' in cell for cell in out[0])
