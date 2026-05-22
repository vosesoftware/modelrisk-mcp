"""ExcelBridge — xlwings + pywin32 wrapper.

Only this module touches Excel/COM directly. Every other layer talks to
it via the typed methods below. See spec §8.1.

Design notes:
- Lazy connect on first use. The bridge attaches to a running Excel
  instance (`xw.apps.active`); it does not launch Excel itself in v0.1.
- All COM errors are caught and re-raised as typed exceptions from
  `modelrisk_mcp.errors`. Raw COM HRESULTs never leak.
- Workbook references are by name, never index — indices change when
  the user opens/closes books.
- Stale-connection recovery: if a method call fails because the Excel
  process was closed and re-opened, the next call transparently
  reconnects.
- Bulk operations (`read_range`, `iterate_cells`) use single COM calls
  and iterate the resulting Python arrays — per-cell COM round-trips
  are too slow for a 10k-cell workbook.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator
from typing import Any

from modelrisk_mcp.errors import (
    CellReferenceError,
    ExcelNotRunningError,
    WorkbookNotFoundError,
)
from modelrisk_mcp.schemas.workbook import (
    CellInfo,
    CellRef,
    RangeInfo,
    WorkbookInfo,
)


class ExcelBridge:
    """Wraps an attached Excel application.

    The bridge is not thread-safe; COM apartment threading rules mean a
    single bridge instance must be used from one thread. The MCP server
    enforces this by giving each request handler its own bridge.
    """

    def __init__(self, *, visible: bool = True) -> None:
        self._visible = visible
        self._app: Any | None = None
        # Lazy-imported so non-Windows test environments can import this
        # module without crashing.
        self._xlwings: Any | None = None

    # ------------------------------------------------------------------
    # Connection lifecycle
    # ------------------------------------------------------------------

    def connect(self) -> None:
        if self._app is not None:
            return
        if self._xlwings is None:
            try:
                import xlwings as xw
            except ImportError as exc:
                raise ExcelNotRunningError(
                    "xlwings is not installed; ExcelBridge cannot operate."
                ) from exc
            self._xlwings = xw
        try:
            self._app = self._xlwings.apps.active
        except Exception as exc:
            raise ExcelNotRunningError(
                "No running Excel instance found. Open Excel and load the "
                "workbook before calling MCP tools."
            ) from exc
        if self._app is None:
            raise ExcelNotRunningError(
                "No running Excel instance found. Open Excel and load the "
                "workbook before calling MCP tools."
            )

    def disconnect(self) -> None:
        # We never own the Excel process — disconnecting just drops the
        # handle. Excel keeps running.
        self._app = None

    def is_connected(self) -> bool:
        return self._app is not None

    def _ensure(self) -> Any:
        if self._app is None:
            self.connect()
            assert self._app is not None
        # Best-effort stale check: any attribute access that fails means
        # the Excel process is gone; reconnect transparently.
        try:
            _ = self._app.books.count  # cheap property; raises on stale handle
        except Exception:
            self._app = None
            self.connect()
            assert self._app is not None
        return self._app

    # ------------------------------------------------------------------
    # Workbook navigation
    # ------------------------------------------------------------------

    def list_workbooks(self) -> list[WorkbookInfo]:
        app = self._ensure()
        result: list[WorkbookInfo] = []
        for book in app.books:
            try:
                result.append(self._workbook_info(book))
            except Exception:
                # An individual workbook that can't be introspected (e.g.
                # protected, mid-recalc) shouldn't break enumeration.
                continue
        return result

    def get_active_workbook(self) -> WorkbookInfo:
        app = self._ensure()
        try:
            book = app.books.active
        except Exception as exc:
            # xlwings' `.active` accessor walks the books and can stumble
            # on OneDrive path resolution. Fall back to the COM API for
            # ActiveWorkbook directly — that one only touches Workbook.Name.
            try:
                api_book = app.api.ActiveWorkbook
            except Exception:
                raise WorkbookNotFoundError("No active workbook.") from exc
            if api_book is None:
                raise WorkbookNotFoundError("No active workbook.") from exc
            try:
                name = str(api_book.Name)
            except Exception:
                raise WorkbookNotFoundError("No active workbook.") from exc
            return WorkbookInfo(name=name, path="", sheets=[], active_sheet=None)
        if book is None:
            raise WorkbookNotFoundError("No active workbook.")
        return self._workbook_info(book)

    def _get_book(self, workbook: str) -> Any:
        app = self._ensure()
        try:
            return app.books[workbook]
        except Exception as exc:
            raise WorkbookNotFoundError(
                f"Workbook {workbook!r} is not open. Open it first or check the name."
            ) from exc

    def _workbook_info(self, book: Any) -> WorkbookInfo:
        # `book.fullname` triggers xlwings' OneDrive path resolution, which
        # raises if the `ONEDRIVE_COMMERCIAL_WIN` env var isn't set. The
        # name alone is what every downstream operation needs (Excel COM
        # indexes by Workbook.Name), so degrade gracefully: empty path,
        # full name still returned.
        try:
            sheets = [s.name for s in book.sheets]
        except Exception:
            sheets = []
        active = None
        try:
            active = book.sheets.active.name
        except Exception:
            pass
        try:
            path = getattr(book, "fullname", "") or ""
        except Exception:
            # OneDrive / SharePoint workbooks can blow up here. The COM
            # Workbook.FullName property would also raise — that's the
            # whole bug. Fall back to a name-only WorkbookInfo.
            path = ""
        # Unsaved workbooks: Workbook.FullName returns just the book's
        # name (e.g. "Book3") with no directory part. That's misleading
        # — downstream code that treats `path` as a filesystem location
        # would try to write `.vmrs` to a relative name in cwd. Detect
        # this by the absence of any path separator and report no path.
        if path and not any(sep in path for sep in ("\\", "/")):
            path = ""
        # `book.name` is a thin wrapper over Workbook.Name and should be
        # safe even when path resolution fails — but guard it anyway so a
        # broken book never tanks list_workbooks().
        try:
            name = book.name
        except Exception:
            try:
                name = str(book.api.Name)
            except Exception:
                name = ""
        return WorkbookInfo(
            name=name,
            path=path,
            sheets=sheets,
            active_sheet=active,
        )

    def _get_sheet(self, workbook: str, sheet: str) -> Any:
        book = self._get_book(workbook)
        try:
            return book.sheets[sheet]
        except Exception as exc:
            raise CellReferenceError(
                f"Sheet {sheet!r} not found in workbook {workbook!r}. "
                f"Available: {[s.name for s in book.sheets]}."
            ) from exc

    # ------------------------------------------------------------------
    # Cell / range reads
    # ------------------------------------------------------------------

    def get_cell(self, workbook: str, sheet: str, cell: str) -> CellInfo:
        sh = self._get_sheet(workbook, sheet)
        try:
            cell_obj = sh.range(cell)
        except Exception as exc:
            raise CellReferenceError(
                f"Invalid cell reference {cell!r} on {workbook}!{sheet}."
            ) from exc
        formula = cell_obj.formula or ""
        value = cell_obj.value
        number_format = ""
        try:
            number_format = cell_obj.number_format or ""
        except Exception:
            pass
        ref = CellRef(workbook=workbook, sheet=sheet, cell=cell)
        return CellInfo(
            ref=ref,
            formula=formula,
            value=value,
            number_format=number_format,
            cell_type=_classify_cell(formula, value),
        )

    def read_range(
        self, workbook: str, sheet: str, range_ref: str
    ) -> RangeInfo:
        sh = self._get_sheet(workbook, sheet)
        try:
            r = sh.range(range_ref)
        except Exception as exc:
            raise CellReferenceError(
                f"Invalid range reference {range_ref!r} on {workbook}!{sheet}."
            ) from exc
        # Single COM call each for values and formulas. xlwings returns a
        # 2D list for multi-cell ranges and a scalar for single cells —
        # normalise both shapes to 2D.
        raw_values = r.value
        raw_formulas = r.formula
        values = _as_2d(raw_values)
        formulas = [[str(f or "") for f in row] for row in _as_2d(raw_formulas)]
        return RangeInfo(
            workbook=workbook,
            sheet=sheet,
            range_ref=range_ref,
            values=values,
            formulas=formulas,
        )

    def iterate_cells(
        self,
        workbook: str,
        predicate: Callable[[CellInfo], bool] | None = None,
        *,
        sheet: str | None = None,
    ) -> Iterator[CellInfo]:
        """Iterate over every used cell in the workbook (or a specific sheet),
        yielding CellInfo for cells matching `predicate` (or all non-empty
        cells if `predicate` is None).

        Reads each sheet's `used_range` in a single COM call and iterates
        the Python arrays. Per-cell COM access is 1000x slower.
        """
        book = self._get_book(workbook)
        sheets_iter = [book.sheets[sheet]] if sheet else list(book.sheets)
        for sh in sheets_iter:
            used = sh.used_range
            if used is None or used.count == 0:
                continue
            try:
                values_2d = _as_2d(used.value)
                formulas_2d = _as_2d(used.formula)
            except Exception:
                continue
            first_row = used.row
            first_col = used.column
            for ri, (vrow, frow) in enumerate(
                zip(values_2d, formulas_2d, strict=False)
            ):
                for ci, (val, formula_str) in enumerate(
                    zip(vrow, frow, strict=False)
                ):
                    formula = str(formula_str or "")
                    if not formula and (val is None or val == ""):
                        continue
                    abs_row = first_row + ri
                    abs_col = first_col + ci
                    cell_ref = _coord_to_a1(abs_row, abs_col)
                    info = CellInfo(
                        ref=CellRef(
                            workbook=workbook,
                            sheet=sh.name,
                            cell=cell_ref,
                        ),
                        formula=formula,
                        value=val if not isinstance(val, list) else None,
                        cell_type=_classify_cell(formula, val),
                    )
                    if predicate is None or predicate(info):
                        yield info

    # ------------------------------------------------------------------
    # Writes
    # ------------------------------------------------------------------

    def write_cell(
        self, workbook: str, sheet: str, cell: str, formula: str
    ) -> None:
        sh = self._get_sheet(workbook, sheet)
        try:
            cell_obj = sh.range(cell)
        except Exception as exc:
            raise CellReferenceError(
                f"Invalid cell reference {cell!r} on {workbook}!{sheet}."
            ) from exc
        cell_obj.formula = formula

    def write_range(
        self,
        workbook: str,
        sheet: str,
        range_ref: str,
        formulas: list[list[str]],
    ) -> None:
        sh = self._get_sheet(workbook, sheet)
        try:
            r = sh.range(range_ref)
        except Exception as exc:
            raise CellReferenceError(
                f"Invalid range reference {range_ref!r} on {workbook}!{sheet}."
            ) from exc
        r.formula = formulas

    # ------------------------------------------------------------------
    # Named ranges
    # ------------------------------------------------------------------

    def set_named_range(
        self,
        workbook: str,
        name: str,
        range_ref: str,
    ) -> None:
        """Create or overwrite a workbook-level named range. `range_ref`
        is the A1 reference the name points to, e.g. 'Sheet1!$A$1:$A$10'.

        xlwings exposes `book.names.add(name, refers_to)`. We first
        attempt to remove an existing same-name entry so re-creation
        succeeds idempotently."""
        book = self._get_book(workbook)
        try:
            existing = book.names[name]
        except Exception:
            existing = None
        if existing is not None:
            try:
                existing.delete()
            except Exception:
                pass
        try:
            book.names.add(name, refers_to=f"={range_ref}")
        except Exception as exc:
            raise CellReferenceError(
                f"Could not create named range {name!r} pointing to "
                f"{range_ref!r}: {exc}"
            ) from exc

    def undo(self) -> None:
        """Trigger Excel's Undo. Used by integration tests to confirm
        the writes we make land in Excel's undo stack."""
        app = self._ensure()
        try:
            app.api.Undo()
        except Exception as exc:
            raise CellReferenceError(f"Excel.Undo() failed: {exc}") from exc

    def recalculate_workbook(self, workbook: str) -> None:
        """Force a full recalculation of `workbook`.

        Used by `restore_deterministic_state` to recover a workbook
        left with VoseOutput cells holding the last simulation sample
        (the "frozen sample" symptom of bug #20). A full recalc
        re-runs every formula in the book, which re-evaluates each
        VoseOutput's underlying expression to its deterministic
        value."""
        book = self._get_book(workbook)
        try:
            # `FullCalculate` rebuilds the dependency tree as well as
            # recomputing — more thorough than `Calculate`, and the
            # right tool when ModelRisk may have left dependency
            # tracking in an odd state.
            book.api.Application.CalculateFull()
        except Exception as exc:
            raise CellReferenceError(
                f"Excel.CalculateFull() failed on {workbook!r}: {exc}"
            ) from exc

    def save_workbook_as(
        self, workbook: str, path: str, *, overwrite: bool = False,
    ) -> str:
        """Save a COPY of `workbook` to `path` (absolute). Returns the
        resolved path that was actually written.

        Bug #25 (alpha.24): prior versions used `book.save(target)`
        which xlwings translates to `Workbook.SaveAs(target)`. That
        call doesn't "save a copy" — it renames the OPEN workbook to
        the new path and rebinds it in Excel's books collection.
        Subsequent tool calls referencing the original workbook name
        then fail with "Workbook 'X.xlsx' is not open" because Excel
        only knows the new name. Not the contract callers expect
        from a `save_as` operation in an MCP context.

        Fix: use the COM `SaveCopyAs` method, which writes the file
        without touching the open workbook's identity. The original
        stays open under its original name; the saved copy is an
        independent file on disk.

        Distinct from `Workbook.Save()` (which we never call
        implicitly per the §11 safety policy). This is the *explicit*
        save-a-copy: caller named a path; we honour it without
        changing the live workbook."""
        from pathlib import Path

        from modelrisk_mcp.errors import WorkbookNotFoundError

        target = Path(path).expanduser().resolve()
        if not overwrite and target.exists():
            raise CellReferenceError(
                f"Refusing to overwrite existing file {target!r}: pass "
                "overwrite=True to confirm."
            )
        if target.suffix.lower() not in {".xlsx", ".xlsm", ".xlsb", ".xls"}:
            raise CellReferenceError(
                f"Save target {target!r} doesn't have an Excel extension. "
                "Use .xlsx, .xlsm, .xlsb, or .xls."
            )
        book = self._get_book(workbook)
        # If overwrite is true and the target exists, SaveCopyAs will
        # refuse — clear the file first.
        if overwrite and target.exists():
            try:
                target.unlink()
            except OSError as exc:
                raise CellReferenceError(
                    f"Could not remove existing {target!r} prior to "
                    f"overwrite: {exc}"
                ) from exc
        try:
            book.api.SaveCopyAs(str(target))
        except Exception as exc:
            raise WorkbookNotFoundError(
                f"Excel refused to save {workbook!r} to {target!r}: {exc}"
            ) from exc
        return str(target)

    # ------------------------------------------------------------------
    # Add-in management
    # ------------------------------------------------------------------

    def list_com_addins(self) -> list[dict[str, Any]]:
        """Return Excel's COM Add-ins collection as plain dicts.

        Each entry: `{description, progid, connected, guid}`.
        """
        app = self._ensure()
        out: list[dict[str, Any]] = []
        try:
            for addin in app.api.COMAddIns:
                out.append(
                    {
                        "description": str(getattr(addin, "Description", "")),
                        "progid": str(getattr(addin, "ProgID", "")),
                        "guid": str(getattr(addin, "Guid", "")),
                        "connected": bool(getattr(addin, "Connect", False)),
                    }
                )
        except Exception:
            return []
        return out

    def list_excel_addins(self) -> list[dict[str, Any]]:
        """Return Excel's classic AddIns collection (XLL / XLA) as
        plain dicts. Each entry: `{name, installed, path}`."""
        app = self._ensure()
        out: list[dict[str, Any]] = []
        try:
            for addin in app.api.AddIns:
                out.append(
                    {
                        "name": str(getattr(addin, "Name", "")),
                        "installed": bool(getattr(addin, "Installed", False)),
                        "path": str(getattr(addin, "FullName", "")),
                    }
                )
        except Exception:
            return []
        return out

    def enable_com_addin(
        self, predicate: Callable[[dict[str, Any]], bool]
    ) -> list[str]:
        """Set `.Connect = True` on every COM add-in matching `predicate`.

        Returns the list of names (progid/description) actually flipped from
        off to on. No-op for add-ins already connected."""
        app = self._ensure()
        flipped: list[str] = []
        try:
            for addin in app.api.COMAddIns:
                info = {
                    "description": str(getattr(addin, "Description", "")),
                    "progid": str(getattr(addin, "ProgID", "")),
                    "guid": str(getattr(addin, "Guid", "")),
                    "connected": bool(getattr(addin, "Connect", False)),
                }
                if not predicate(info):
                    continue
                if info["connected"]:
                    continue
                try:
                    addin.Connect = True
                    name = str(info["description"] or info["progid"])
                    flipped.append(name)
                except Exception:
                    # Some add-ins refuse to be enabled from automation
                    # (signed-only policy, etc.). Best-effort.
                    continue
        except Exception:
            return flipped
        return flipped

    def enable_excel_addin(
        self, predicate: Callable[[dict[str, Any]], bool]
    ) -> list[str]:
        """Set `.Installed = True` on every classic add-in matching
        `predicate`. Returns the names of those flipped on."""
        app = self._ensure()
        flipped: list[str] = []
        try:
            for addin in app.api.AddIns:
                info = {
                    "name": str(getattr(addin, "Name", "")),
                    "installed": bool(getattr(addin, "Installed", False)),
                    "path": str(getattr(addin, "FullName", "")),
                }
                if not predicate(info):
                    continue
                if info["installed"]:
                    continue
                try:
                    addin.Installed = True
                    flipped.append(str(info["name"]))
                except Exception:
                    continue
        except Exception:
            return flipped
        return flipped


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------


def _classify_cell(formula: str, value: Any) -> str:
    if formula:
        return "formula"
    if value is None or value == "":
        return "empty"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, str):
        return "text"
    return "general"


def _as_2d(value: Any) -> list[list[Any]]:
    """Normalise a value as returned by xlwings' `.value` / `.formula`
    to a 2D list of cell values.

    xlwings on Windows returns:
    - a scalar for a single cell
    - a 1D LIST for `.value` on a single row or column
    - a 2D LIST for `.value` on a rectangular range
    - **a TUPLE (or tuple of tuples)** for `.formula` — because the COM
      property reads through `range.api.Formula` which returns raw COM
      VARIANT arrays. The xlwings value-side wraps them in lists; the
      formula-side does not.

    The tuple-vs-list distinction is the bug that silently collapsed
    every list-scan (list_modelrisk_inputs / get_workbook_summary /
    run_simulation's input registration). If a tuple lands here and we
    treat it as a scalar, the entire row's formulas get string-cast as
    a single "(formula_A, formula_B, formula_C)" tuple — the regex
    finds the first VoseInput inside that string and we yield one
    record instead of many.
    """
    if value is None:
        return [[None]]
    if isinstance(value, (list, tuple)):
        if not value:
            return [[]]
        if all(isinstance(row, (list, tuple)) for row in value):
            # 2D — coerce each inner row to a list for downstream uniformity.
            return [list(row) for row in value]
        # 1D row or column — wrap as a single row.
        # xlwings doesn't tell us which orientation, so treat as a row.
        return [list(value)]
    return [[value]]


_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def _coord_to_a1(row: int, col: int) -> str:
    """Convert (1-based row, 1-based column) to A1 like 'AB12'."""
    if col < 1:
        raise ValueError(f"Column must be 1-based, got {col}.")
    letters = ""
    c = col
    while c > 0:
        c, rem = divmod(c - 1, 26)
        letters = _LETTERS[rem] + letters
    return f"{letters}{row}"
