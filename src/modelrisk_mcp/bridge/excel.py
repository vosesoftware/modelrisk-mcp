"""ExcelBridge — xlwings + pywin32 wrapper.

Only this module touches Excel/COM directly. Every other layer talks to
it via the typed methods below. See spec §8.1.

Design notes:
- Lazy connect on first use. The bridge attaches to a running Excel
  instance (`xw.apps.active`); if none is running it starts an
  attachable Excel (`xw.App(add_book=True)`) and loads the ModelRisk
  add-in into it by registering the XLL. Auto-launch is on by default;
  disable with `MODELRISK_AUTO_LAUNCH=0`.
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

import os
from collections.abc import Callable, Iterator
from datetime import date, datetime
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

    def __init__(
        self, *, visible: bool = True, auto_launch: bool | None = None
    ) -> None:
        self._visible = visible
        self._app: Any | None = None
        # Lazy-imported so non-Windows test environments can import this
        # module without crashing.
        self._xlwings: Any | None = None
        # When no Excel is running, start ModelRisk via Vose's own
        # `modelrisk.exe` launcher (which opens Excel WITH the add-in
        # loaded natively — avoiding the COM-launch xlAutoOpen-skip of
        # bug #29). On by default; disable with MODELRISK_AUTO_LAUNCH=0.
        if auto_launch is None:
            import os

            auto_launch = os.environ.get(
                "MODELRISK_AUTO_LAUNCH", "1"
            ).strip().lower() not in ("0", "false", "no", "off")
        self._auto_launch = auto_launch

    # ------------------------------------------------------------------
    # Connection lifecycle
    # ------------------------------------------------------------------

    def _load_xlwings(self) -> None:
        if self._xlwings is None:
            try:
                import xlwings as xw
            except ImportError as exc:
                raise ExcelNotRunningError(
                    "xlwings is not installed; ExcelBridge cannot operate."
                ) from exc
            self._xlwings = xw

    def _attach_active(self) -> Any | None:
        """Return the active Excel app if one is running, else None.
        `xlwings.apps.active` either returns None or raises when no
        Excel is up — normalise both to None."""
        if self._xlwings is None:
            return None
        try:
            return self._xlwings.apps.active
        except Exception:
            return None

    def connect(self) -> None:
        if self._app is not None:
            return
        self._load_xlwings()
        app = self._attach_active()
        # No Excel running → optionally start ModelRisk for the user.
        if app is None and self._auto_launch:
            if self.launch_modelrisk():
                app = self._attach_active()
        if app is None:
            hint = (
                ""
                if self._auto_launch
                else " (auto-launch is disabled via MODELRISK_AUTO_LAUNCH)"
            )
            raise ExcelNotRunningError(
                "No running Excel instance found and ModelRisk could not "
                "be started automatically" + hint + ". Open Excel (or start "
                "ModelRisk from its shortcut) and load the workbook, then "
                "retry."
            )
        self._app = app

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

    def open_workbook(self, path: str) -> WorkbookInfo:
        """Open a workbook from disk in the running Excel and return its info.
        Excel keys workbooks by file name and won't open two with the same
        name — if one is already open, that existing workbook is returned
        rather than re-opened."""
        app = self._ensure()
        if not os.path.isfile(path):
            raise WorkbookNotFoundError(f"File not found: {path!r}.")
        name = os.path.basename(path)
        for book in app.books:
            try:
                if str(book.name).lower() == name.lower():
                    return self._workbook_info(book)
            except Exception:
                continue
        book = self._open_no_prompts(app, path)
        return self._workbook_info(book)

    @staticmethod
    def _open_no_prompts(app: Any, path: str) -> Any:
        """Workbooks.Open with the interactive prompts suppressed, so a headless
        open never hangs on a dialog: no Update-Links, read-only-recommended, or
        file-in-use prompts. ``update_links=False`` also means external links are
        NOT refreshed on open (cell values stay as last saved). DisplayAlerts and
        AskToUpdateLinks are toggled off around the call and restored after."""
        saved: dict[str, Any] = {}
        for attr in ("DisplayAlerts", "AskToUpdateLinks"):
            try:
                saved[attr] = getattr(app.api, attr)
                setattr(app.api, attr, False)
            except Exception:
                pass
        try:
            return app.books.open(
                path,
                update_links=False,
                ignore_read_only_recommended=True,
                notify=False,
                add_to_mru=False,
            )
        except Exception as exc:
            raise WorkbookNotFoundError(f"Excel could not open {path!r}: {exc}") from exc
        finally:
            for attr, val in saved.items():
                try:
                    setattr(app.api, attr, val)
                except Exception:
                    pass

    def close_workbook(self, workbook: str, save: bool = False) -> dict[str, Any]:
        """Close an open workbook. If ``save`` is False (default) UNSAVED CHANGES
        ARE DISCARDED; pass save=True to write them first. Raises
        WorkbookNotFoundError if the workbook isn't open. Returns the closed
        name, whether it was saved, and the names still open."""
        app = self._ensure()
        book = self._get_book(workbook)  # raises WorkbookNotFoundError if not open
        name = str(book.name)
        try:
            book.api.Close(SaveChanges=bool(save))
        except Exception as exc:
            raise WorkbookNotFoundError(f"Could not close {workbook!r}: {exc}") from exc
        remaining: list[str] = []
        for b in app.books:
            try:
                remaining.append(str(b.name))
            except Exception:
                continue
        return {"closed": name, "saved": bool(save), "open_workbooks": remaining}

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
        error = _detect_excel_error(cell_obj, value)
        number_format = ""
        try:
            number_format = cell_obj.number_format or ""
        except Exception:
            pass
        ref = CellRef(workbook=workbook, sheet=sheet, cell=cell)
        return CellInfo(
            ref=ref,
            formula=formula,
            value=_normalize_value(value),
            number_format=number_format,
            cell_type=_classify_cell(formula, value, error=error),
            error=error,
        )

    def get_range_shape(
        self, workbook: str, sheet: str, range_ref: str
    ) -> tuple[int, int]:
        """Return `(n_rows, n_cols)` of a range without reading values."""
        sh = self._get_sheet(workbook, sheet)
        try:
            r = sh.range(range_ref)
        except Exception as exc:
            raise CellReferenceError(
                f"Invalid range reference {range_ref!r} on {workbook}!{sheet}."
            ) from exc
        shape = r.shape
        return int(shape[0]), int(shape[1])

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
        # Bug #34 (alpha.33) / bug #35 (alpha.36): also surface error
        # cells. Same dual strategy as `iterate_cells`:
        # 1. `Range.Value2` returns integer CVErr codes on multi-cell
        #    ranges — robust across Excel versions (the alpha.36
        #    fallback after Text-returning-None was observed live).
        # 2. `Range.Text` as a secondary check for any cells Value2
        #    didn't classify.
        errors: list[list[str | None]] = []
        value2_2d: list[list[Any]] = []
        try:
            value2_2d = _as_2d(r.api.Value2)
            if len(value2_2d) != len(values):
                value2_2d = []
        except Exception:
            value2_2d = []
        text_2d: list[list[Any]] = []
        try:
            text_2d = _as_2d(r.api.Text)
            if len(text_2d) != len(values):
                text_2d = []
        except Exception:
            text_2d = []
        if value2_2d or text_2d:
            for ri in range(len(values)):
                err_row: list[str | None] = []
                v2_row = value2_2d[ri] if value2_2d and ri < len(value2_2d) else []
                tx_row = text_2d[ri] if text_2d and ri < len(text_2d) else []
                for ci in range(len(values[ri])):
                    err: str | None = None
                    if ci < len(v2_row):
                        err = _coerce_error_value(v2_row[ci])
                    if err is None and ci < len(tx_row):
                        t = tx_row[ci]
                        if (
                            isinstance(t, str)
                            and t.strip() in _EXCEL_ERROR_STRINGS
                        ):
                            err = t.strip()
                    err_row.append(err)
                errors.append(err_row)
        # Keep `errors` compact when nothing is errored — empty list
        # signals "no errors detected", same shape as before.
        if errors and not any(e for row in errors for e in row):
            errors = []
        return RangeInfo(
            workbook=workbook,
            sheet=sheet,
            range_ref=range_ref,
            values=values,
            formulas=formulas,
            errors=errors,
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
            # Bug #36 (alpha.37): cache `sh.name` once per sheet. Each
            # `sh.name` access is a COM round-trip (~150μs) — calling
            # it inside the per-cell loop turned a 10k-cell scan into
            # ~1.5s of pure COM-attribute overhead (profiled live;
            # `sh.name` was 81% of iterate_cells' total time before
            # this fix). One-liner cache → ~10x speedup on iteration.
            sheet_name = sh.name
            try:
                values_2d = _as_2d(used.value)
                formulas_2d = _as_2d(used.formula)
            except Exception:
                continue
            # Bug #34 (alpha.33) / bug #35 (alpha.36): bulk-detect
            # error cells per sheet so an audit scan flags them without
            # per-cell COM round-trips.
            #
            # alpha.33 used `Range.Text` for this. That works for
            # single cells but on multi-cell ranges some Excel versions
            # return `None` from the bulk `Range.Text` property — which
            # silently regressed VOSE-012 on real workbooks (caught by
            # round-10 live probe).
            #
            # alpha.36 fix: prefer `Range.Value2`, which on a bulk
            # range reliably returns a tuple-of-tuples with the COM
            # CVErr **integer code** in each errored cell's slot.
            # Mapping is stable across Excel versions. Fall back to
            # Text if Value2 is also unavailable.
            text_2d: list[list[Any]] = []
            value2_2d: list[list[Any]] = []
            try:
                value2_2d = _as_2d(used.api.Value2)
                if len(value2_2d) != len(values_2d):
                    value2_2d = []
            except Exception:
                value2_2d = []
            try:
                text_2d = _as_2d(used.api.Text)
                if len(text_2d) != len(values_2d):
                    text_2d = []
            except Exception:
                text_2d = []

            first_row = used.row
            first_col = used.column
            for ri, (vrow, frow) in enumerate(
                zip(values_2d, formulas_2d, strict=False)
            ):
                v2_row = (
                    value2_2d[ri]
                    if value2_2d and ri < len(value2_2d)
                    else []
                )
                tx_row = (
                    text_2d[ri] if text_2d and ri < len(text_2d) else []
                )
                for ci, (val, formula_str) in enumerate(
                    zip(vrow, frow, strict=False)
                ):
                    formula = str(formula_str or "")
                    # Prefer Value2 (robust on this Excel version); fall
                    # back to Text if Value2 didn't give us an answer.
                    cell_error: str | None = None
                    if ci < len(v2_row):
                        cell_error = _coerce_error_value(v2_row[ci])
                    if cell_error is None and ci < len(tx_row):
                        t = tx_row[ci]
                        if (
                            isinstance(t, str)
                            and t.strip() in _EXCEL_ERROR_STRINGS
                        ):
                            cell_error = t.strip()
                    if not formula and (val is None or val == ""):
                        # Empty cells: skip unless they're errored.
                        # (Some Vose paths can leave a cell that has
                        # no formula but evaluates to an error literal.)
                        if cell_error is None:
                            continue
                    abs_row = first_row + ri
                    abs_col = first_col + ci
                    cell_ref = _coord_to_a1(abs_row, abs_col)
                    info = CellInfo(
                        ref=CellRef(
                            workbook=workbook,
                            sheet=sheet_name,
                            cell=cell_ref,
                        ),
                        formula=formula,
                        value=(
                            _normalize_value(val)
                            if not isinstance(val, list)
                            else None
                        ),
                        cell_type=_classify_cell(
                            formula, val, error=cell_error,
                        ),
                        error=cell_error,
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

    # ------------------------------------------------------------------
    # ModelRisk add-in activation support (bug #38 — "add-in not loaded")
    # ------------------------------------------------------------------

    def evaluate(self, expr: str) -> Any:
        """Evaluate an Excel expression via `Application.Evaluate` and
        return the raw result (a number, string, or a COM CVErr integer
        for an error like `#NAME?`). Used to probe whether the ModelRisk
        add-in is live — a Vose function returns a number when the XLL is
        loaded and an error CVErr int when it isn't. Raises only if the COM
        call itself fails.

        ⚠️ `Application.Evaluate` parses the string with the user's LOCALE
        list/decimal separators — on a comma-decimal locale (Russian, German,
        …) `Foo(0,1)` is read as `Foo(0.1)` (one argument). To stay correct for
        MULTI-ARG expressions there, when `Application.Evaluate` returns an
        error we retry through a scratch cell's `.Formula`, which always uses
        US conventions (',' argument separator, '.' decimal) regardless of
        locale — the same parse a user gets typing the localized form. The
        common path (valid result first try) is unchanged and never touches
        the workbook. Keeping separator-free probes (see `_ADDIN_PROBE_EXPR`)
        avoids the retry entirely."""
        app = self._ensure()
        result = app.api.Evaluate(expr)
        if _coerce_error_value(result) is None:
            return result
        # Errored — could be a genuine error, or a comma-decimal-locale
        # mis-parse of a multi-arg expression. Retry via a scratch cell whose
        # `.Formula` is locale-invariant; keep the original result if we can't.
        ran, cell_value = self._evaluate_via_cell(app, expr)
        return cell_value if ran else result

    @staticmethod
    def _evaluate_via_cell(app: Any, expr: str) -> tuple[bool, Any]:
        """Evaluate ``=expr`` by writing it to the active sheet's far-corner
        cell via `.Formula` (locale-invariant) and reading the value back,
        restoring the cell afterwards. Returns ``(ran, value)`` — ``ran`` is
        False (and value None) if there's no usable worksheet or the write
        failed, so the caller falls back to the Application.Evaluate result."""
        try:
            wb = app.api.ActiveWorkbook
            if wb is None:
                return False, None
            ws = wb.ActiveSheet
            cell = ws.Cells(ws.Rows.Count, ws.Columns.Count)  # bottom-right corner
            prev = cell.Formula
            try:
                cell.Formula = "=" + expr
                value = cell.Value
            finally:
                cell.Formula = prev if prev not in (None, "") else ""
            return True, value
        except Exception:
            return False, None

    def register_xll(self, path: str) -> bool:
        """`Application.RegisterXLL(path)` — re-runs the XLL's
        `xlAutoOpen`, which is what actually wires up its commands and
        UDFs. Idempotent. Returns True on success, False on failure."""
        app = self._ensure()
        try:
            app.api.RegisterXLL(str(path))
            return True
        except Exception:
            return False

    def register_modelrisk_xlls(self) -> list[str]:
        """RegisterXLL every *installed* ModelRisk `.xll` in the AddIns
        collection. Fixes the "installed but commands/UDFs unreachable"
        state that occurs when Excel was started programmatically (its
        normal startup skips `xlAutoOpen`). Returns the XLL names
        re-registered."""
        app = self._ensure()
        done: list[str] = []
        try:
            addins = app.api.AddIns
            for i in range(1, int(addins.Count) + 1):
                try:
                    addin = addins(i)
                    name = str(getattr(addin, "Name", "") or "")
                    if not name.lower().endswith(".xll"):
                        continue
                    if "modelrisk" not in name.lower():
                        continue
                    if not bool(getattr(addin, "Installed", False)):
                        continue
                    if self.register_xll(str(addin.FullName)):
                        done.append(name)
                except Exception:
                    continue
        except Exception:
            return done
        return done

    def launch_modelrisk(self) -> bool:
        """Start a fresh, attachable Excel and load the ModelRisk add-in
        into it. Returns True if Excel is up and attachable afterwards.

        Why `xw.App(add_book=True)` rather than Vose's `modelrisk.exe`
        launcher: verified on a real install, `modelrisk.exe` opens
        Excel on the Start screen with NO workbook — and a workbook-less
        Excel is absent from the COM Running Object Table, so it can
        never be attached to (`xw.apps.count` stays 0). Adding a blank
        workbook makes the instance attachable immediately.

        The trade-off is that an Excel started this way skips the XLL's
        `xlAutoOpen` (bug #29), so the add-in isn't auto-loaded. We then
        register the ModelRisk XLL from disk to bring it live — the same
        step the activation ladder uses, confirmed to make Vose
        functions resolve against a real ModelRisk install."""
        self._load_xlwings()
        if self._xlwings is None:
            return False
        try:
            app = self._xlwings.App(visible=self._visible, add_book=True)
        except Exception:
            return False
        self._app = app
        # Load the add-in into the freshly-started Excel (best-effort):
        # an already-installed XLL just needs re-registering; otherwise
        # register it from its install path on disk.
        try:
            if not self.register_modelrisk_xlls():
                for path in self.find_modelrisk_xll_paths():
                    if self.register_xll(path):
                        break
        except Exception:
            pass
        return self._attach_active() is not None or self._app is not None

    def find_modelrisk_xll_paths(self) -> list[str]:
        """Best-effort search of the standard Vose Software install
        directories for a `ModelRisk*.xll` on disk. Used as the last
        resort when the add-in isn't even in the AddIns collection
        (ModelRisk set to NOT 'Start with Excel' and never opened this
        session). Returns existing absolute paths, 64-bit-named first."""
        import os
        from pathlib import Path

        roots: list[Path] = []
        for env in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
            base = os.environ.get(env)
            if base:
                roots.append(Path(base) / "Vose Software")
        found: list[str] = []
        seen: set[str] = set()
        for root in roots:
            if not root.exists():
                continue
            try:
                for p in root.rglob("*.xll"):
                    if "modelrisk" not in p.name.lower():
                        continue
                    key = str(p).lower()
                    if key in seen:
                        continue
                    seen.add(key)
                    found.append(str(p))
            except Exception:
                continue
        # Prefer the 64-bit build (commonly suffixed 64) when present.
        found.sort(key=lambda s: (("64" not in Path(s).stem), s))
        return found


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------


def _normalize_value(value: Any) -> float | str | bool | None:
    """Coerce an xlwings cell value into the type union accepted by
    `CellInfo.value` (`float | str | bool | None`).

    Bug #35 (alpha.4-followup): Excel date-formatted cells come back from
    xlwings as `datetime.datetime` (or `datetime.date`). `CellInfo.value`
    doesn't permit those, so Pydantic refuses to construct the model and
    `iterate_cells` blows up on any workbook with a date in its used
    range — which then takes the audit down with it.

    Fix: stringify date/datetime as ISO 8601 at the bridge boundary. The
    audit rules don't reason about dates (they tokenise formulas), so
    losing the typed datetime here is harmless; surfacing the value as
    text keeps it visible to the LLM. Other unsupported types (e.g.
    Decimal) fall through to `str(...)` for the same reason.

    `bool` must be checked *before* `int/float` because `True` is an
    `int` in Python and we want `cell_type="boolean"` semantics
    preserved upstream by `_classify_cell` (which sees the original
    value, not the coerced one).
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        return value
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    # Last-ditch fallback for anything else xlwings might hand us
    # (Decimal, custom COM scalars, etc.). Stringify rather than crash.
    return str(value)


def _classify_cell(formula: str, value: Any, *, error: str | None = None) -> str:
    """Classify a cell by what it actually contains.

    Bug #27 (alpha.25): xlwings' `Range.Formula` accessor returns the
    cell's text content even for non-formula cells (e.g. a cell
    containing the label "Total Revenue" comes back with
    formula="Total Revenue"). The prior check `if formula:` then
    flagged every text and numeric cell as a formula. That inflated
    `formula_cell_count` and silently broke `find_hard_coded_inputs`
    on any workbook with text labels — including the typical
    "convert this Excel model to ModelRisk" use case.

    Fix: a cell only counts as a formula if its `.Formula` actually
    starts with `=`. Everything else gets classified by value type.

    Bug #34 (alpha.33): an Excel error (`#DIV/0!`, `#REF!`, ...) is
    a distinct cell state and gets its own classification. Detected
    from the caller (it requires a COM Text read we don't have here).
    """
    if error is not None:
        return "error"
    if formula and formula.lstrip().startswith("="):
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


# The set of Excel error literals — what Range.Text returns when a
# cell evaluates to an error. Covers classic errors (#REF!, #DIV/0!,
# #VALUE!, #NAME?, #NULL!, #NUM!, #N/A) plus the newer dynamic-array
# / external-data errors introduced from Excel 2019 onwards.
_EXCEL_ERROR_STRINGS = frozenset({
    "#DIV/0!",
    "#N/A",
    "#NAME?",
    "#NULL!",
    "#NUM!",
    "#REF!",
    "#VALUE!",
    "#GETTING_DATA",
    "#SPILL!",
    "#CALC!",
    "#FIELD!",
    "#UNKNOWN!",
    "#BLOCKED!",
    "#BUSY!",
    "#CONNECT!",
    "#EXTERNAL!",
    "#PYTHON!",
})


# Bug #35 (alpha.36): on at least some Excel versions, `Range.Text` on
# a MULTI-cell range returns `None` (single-cell Text works fine). The
# alpha.33 `iterate_cells` path relied on bulk Text — which then
# silently produced no error info for any sheet, regressing VOSE-012
# on real workbooks.
#
# Fix: also detect errors via `Range.Value2`, which on a multi-cell
# range reliably returns a tuple-of-tuples containing the COM CVErr
# **integer code** (HRESULTs in the 0x800A07D0-0x800A07FB range) for
# each errored cell. Mapping is stable across Excel versions because
# it's the Office COM facility's error codes — the lower 16 bits are
# the well-known xlCVError constants. This map was probed empirically
# against Excel 365 (each row was confirmed: typed the formula in,
# called Range.api.Value2, recorded the int).
_EXCEL_ERROR_CODE_TO_STRING: dict[int, str] = {
    -2146826281: "#DIV/0!",   # 0x800A07D7, xlErrDiv0   (2007)
    -2146826288: "#NULL!",    # 0x800A07D0, xlErrNull   (2000)
    -2146826273: "#VALUE!",   # 0x800A07DF, xlErrValue  (2015)
    -2146826265: "#REF!",     # 0x800A07E7, xlErrRef    (2023)
    -2146826259: "#NAME?",    # 0x800A07ED, xlErrName   (2029)
    -2146826252: "#NUM!",     # 0x800A07F4, xlErrNum    (2036)
    -2146826246: "#N/A",      # 0x800A07FA, xlErrNA     (2042)
    -2146826245: "#GETTING_DATA",  # 0x800A07FB, xlErrGettingData (2043)
    # Newer dynamic-array errors fall outside this canonical set on
    # some Excel versions; if Value2 returns a code we don't recognise
    # we fall through to the Text-based detector instead.
}


def _coerce_error_value(value: Any) -> str | None:
    """If `value` is one of Excel's COM CVErr integer codes, return the
    corresponding error literal; else None. Used by both the bulk
    iterate_cells path and the per-cell get_cell path as a robust
    backstop for cases where Range.Text returns None."""
    if isinstance(value, int) and not isinstance(value, bool):
        return _EXCEL_ERROR_CODE_TO_STRING.get(value)
    return None


def _detect_excel_error(cell_obj: Any, value: Any) -> str | None:
    """Return the Excel error string (`"#DIV/0!"` etc.) if this cell
    evaluates to an error, else `None`.

    Detection strategy (in order):

    1. `Range.Text`. For a single cell this reliably returns the
       displayed error literal regardless of number format.
    2. `Range.Value2` integer CVErr code lookup (bug #35 alpha.36
       fallback). Some Excel versions return `None` for `Range.Text`
       on certain cells but always report the integer error code.
       The map is stable across Excel versions because it's the
       Office COM facility's error codes.

    A single extra COM round-trip per `get_cell` is a fine tax to
    surface a real category of workbook problem to the LLM (bug #34).
    """
    try:
        text = cell_obj.api.Text
    except Exception:
        text = None
    if isinstance(text, str):
        stripped = text.strip()
        if stripped in _EXCEL_ERROR_STRINGS:
            return stripped
    # Fallback: integer CVErr code via api.Value2 (bug #35).
    try:
        val2 = cell_obj.api.Value2
    except Exception:
        val2 = None
    return _coerce_error_value(val2)


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
