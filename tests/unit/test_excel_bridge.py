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

from modelrisk_mcp.bridge.excel import (
    ExcelBridge,
    _as_2d,
    _classify_cell,
    _detect_excel_error,
)
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


# ----------------------------------------------------------------------
# save_workbook_as — must use SaveCopyAs, not SaveAs (bug #25)
# ----------------------------------------------------------------------


class _RecordingApi:
    """Stand-in for `book.api`: records which COM method was called
    on it. The whole point of the alpha.24 fix is that we must call
    `SaveCopyAs` (which leaves the open workbook's name untouched)
    not the implicit `SaveAs` from `book.save(path)` (which renames
    the open workbook in place)."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []

    def SaveCopyAs(self, *args: Any) -> None:  # noqa: N802
        self.calls.append(("SaveCopyAs", args))

    def SaveAs(self, *args: Any) -> None:  # noqa: N802
        self.calls.append(("SaveAs", args))


class _RecordingBook:
    def __init__(self) -> None:
        self.api = _RecordingApi()

    def save(self, path: str) -> None:
        # If anyone calls this, the bug is back. Fail loudly.
        raise AssertionError(
            f"book.save({path!r}) called — alpha.24 fix requires "
            f"book.api.SaveCopyAs instead. SaveAs (which book.save uses) "
            f"renames the open workbook in place; SaveCopyAs writes a "
            f"copy without touching it."
        )


class TestSaveWorkbookAsUsesSaveCopyAs:
    """Regression for bug #25 — save_workbook_as renamed the open
    workbook because it used `book.save(path)` (→ SaveAs). The fix
    routes through `book.api.SaveCopyAs(path)` which leaves the open
    workbook untouched."""

    def _make_bridge(self, book: _RecordingBook) -> ExcelBridge:
        bridge = ExcelBridge.__new__(ExcelBridge)  # bypass __init__
        bridge._app = None  # type: ignore[attr-defined]
        bridge._get_book = lambda _name: book  # type: ignore[method-assign]
        return bridge

    def test_uses_save_copy_as(self, tmp_path: Any) -> None:
        book = _RecordingBook()
        bridge = self._make_bridge(book)
        target = tmp_path / "saved.xlsx"
        result = bridge.save_workbook_as(
            "src.xlsx", str(target), overwrite=False,
        )
        # Resolves to absolute string.
        assert result == str(target.resolve())
        # SaveCopyAs was called, SaveAs was NOT.
        method_names = [name for name, _args in book.api.calls]
        assert method_names == ["SaveCopyAs"], (
            f"Expected SaveCopyAs only, got {method_names}. "
            f"The bug-#25 regression is back if SaveAs shows up here."
        )

    def test_overwrite_true_clears_existing_target_first(
        self, tmp_path: Any,
    ) -> None:
        """SaveCopyAs refuses to overwrite an existing file, so when
        overwrite=True the bridge must `unlink` first."""
        book = _RecordingBook()
        bridge = self._make_bridge(book)
        target = tmp_path / "saved.xlsx"
        target.write_bytes(b"stub")  # pre-existing
        bridge.save_workbook_as("src.xlsx", str(target), overwrite=True)
        # The pre-existing file was removed before SaveCopyAs (our
        # `_RecordingApi.SaveCopyAs` is a no-op so the target stays
        # absent — but the pre-write `target.exists()` is False after
        # the unlink).
        assert not target.exists()

    def test_overwrite_false_refuses_existing_target(
        self, tmp_path: Any,
    ) -> None:
        from modelrisk_mcp.errors import CellReferenceError

        book = _RecordingBook()
        bridge = self._make_bridge(book)
        target = tmp_path / "saved.xlsx"
        target.write_bytes(b"stub")
        with pytest.raises(CellReferenceError, match="Refusing to overwrite"):
            bridge.save_workbook_as(
                "src.xlsx", str(target), overwrite=False,
            )
        # No COM call should have happened.
        assert book.api.calls == []


# ----------------------------------------------------------------------
# Bug #34 (alpha.33): error-cell detection via Range.Text
# ----------------------------------------------------------------------


class _FakeCellApi:
    """Stand-in for `cell_obj.api` exposing only `.Text` (and optionally
    `.Value`). xlwings' Range.Text proxies straight through to the COM
    Range.Text property."""

    def __init__(self, text: Any = "") -> None:
        self.Text = text  # mirrors COM property name


class _FakeCell:
    def __init__(
        self,
        *,
        formula: str = "",
        value: Any = None,
        number_format: str = "",
        text: Any = "",
    ) -> None:
        self.formula = formula
        self.value = value
        self.number_format = number_format
        self.api = _FakeCellApi(text=text)


class TestDetectExcelError:
    """Bug #34: error cells (#DIV/0!, #REF!, etc.) returned value=None
    indistinguishable from empty cells. We now read Range.Text and
    match against the known Excel error literals."""

    def test_div_zero_detected(self) -> None:
        cell = _FakeCell(formula="=1/0", value=None, text="#DIV/0!")
        assert _detect_excel_error(cell, None) == "#DIV/0!"

    def test_ref_error_detected(self) -> None:
        cell = _FakeCell(formula="=#REF!", value=None, text="#REF!")
        assert _detect_excel_error(cell, None) == "#REF!"

    def test_name_error_detected(self) -> None:
        cell = _FakeCell(formula="=BadFn()", value=None, text="#NAME?")
        assert _detect_excel_error(cell, None) == "#NAME?"

    def test_na_error_detected(self) -> None:
        cell = _FakeCell(formula="=NA()", value=None, text="#N/A")
        assert _detect_excel_error(cell, None) == "#N/A"

    def test_value_error_detected(self) -> None:
        cell = _FakeCell(formula='=VALUE("nope")', value=None, text="#VALUE!")
        assert _detect_excel_error(cell, None) == "#VALUE!"

    def test_normal_number_is_not_an_error(self) -> None:
        cell = _FakeCell(formula="=1+1", value=2.0, text="2")
        assert _detect_excel_error(cell, 2.0) is None

    def test_text_label_is_not_an_error(self) -> None:
        cell = _FakeCell(formula="", value="Total", text="Total")
        assert _detect_excel_error(cell, "Total") is None

    def test_empty_cell_is_not_an_error(self) -> None:
        cell = _FakeCell(formula="", value=None, text="")
        assert _detect_excel_error(cell, None) is None

    def test_hash_prefix_string_that_is_not_an_excel_error(self) -> None:
        """A text cell whose displayed value happens to start with '#'
        (e.g. a hashtag) must NOT be flagged as an error — only the
        exact Excel error literals count."""
        cell = _FakeCell(formula="", value="#hashtag", text="#hashtag")
        assert _detect_excel_error(cell, "#hashtag") is None

    def test_text_with_surrounding_whitespace_still_matches(self) -> None:
        """Some custom formats can pad the displayed text; the detector
        strips before matching."""
        cell = _FakeCell(formula="=1/0", value=None, text="  #DIV/0!  ")
        assert _detect_excel_error(cell, None) == "#DIV/0!"

    def test_text_read_failure_returns_none(self) -> None:
        """If Range.Text raises (some COM versions / OLE-state issues),
        we should not crash — just return None and let the value/
        formula path handle the cell normally."""

        class _ExplodingApi:
            @property
            def Text(self) -> Any:  # noqa: N802
                raise RuntimeError("COM unavailable")

        cell = SimpleNamespace(api=_ExplodingApi())
        assert _detect_excel_error(cell, None) is None

    def test_non_string_text_returns_none(self) -> None:
        """If Text comes back as something other than a string (e.g.
        a COM Variant that didn't coerce), we don't crash."""
        cell = _FakeCell(formula="", value=None, text=42)
        assert _detect_excel_error(cell, None) is None


class TestClassifyCellWithError:
    """Bug #34: a cell with an error must classify as 'error' even if
    the formula starts with '=' (which would otherwise win)."""

    def test_error_overrides_formula_classification(self) -> None:
        # Error wins: a `=1/0` formula classifies as error, not formula.
        assert _classify_cell("=1/0", None, error="#DIV/0!") == "error"

    def test_no_error_classifies_as_formula(self) -> None:
        assert _classify_cell("=1+1", 2.0, error=None) == "formula"

    def test_no_error_no_formula_empty(self) -> None:
        assert _classify_cell("", None, error=None) == "empty"

    def test_no_error_no_formula_number(self) -> None:
        assert _classify_cell("", 3.14, error=None) == "number"

    def test_no_error_no_formula_text(self) -> None:
        assert _classify_cell("", "Label", error=None) == "text"
