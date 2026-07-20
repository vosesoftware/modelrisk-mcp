"""In-session results fallback — read per-iteration simulation samples
via `VoseSimValue` scratch cells when MRService.dll is unavailable.

Live-verified semantics (2026-07-20 spikes, real Excel + ModelRisk XLL):

- The `VoseSim*` STATISTIC worksheet functions (`VoseSimMean`, …) are
  computed once, during the end-of-simulation pass, and cached per
  calling cell. Cells written AFTER a run always return the text
  "No simulation results" — so they can't serve as a post-hoc reader.
- `VoseSimValue(ref, k)` reads the live per-iteration sample store
  directly and works ANY time after a run in the same Excel session
  (1-based `k`; out-of-range returns the "No simulation results" text).
- The samples it returns are sample-for-sample identical to what
  MRService.dll reads from the `.vmrs` saved from the same run
  (verified < 1e-12 over a full run).

So when the DLL is missing (ModelRisk installers up to 9.1.x didn't
ship it), we can still serve every results-reading tool by writing a
scratch column of `=VoseSimValue(<output>,ROW())` formulas, bulk-reading
it, and deleting the scratch sheet again. Downstream statistics,
percentiles, correlations and sensitivity are computed in numpy from
the samples — the same pipeline the `.vmrs` path feeds.

Constraints (inherent, documented to the caller in error messages):
- Session-bound: results live in the Excel process that ran the sim.
  A closed Excel or a machine restart loses them; the `.vmrs` file is
  the durable artifact and needs MRService.dll to re-open.
- Latest run only: a new simulation replaces the store.
- Requires scratch-cell writes: in `--read-only` mode the fallback is
  refused (the DLL path is the write-free reader).

⚠ Scratch placement is load-bearing: the scratch column lives on an
EXISTING sheet (first free column beyond the used range) and is cleaned
with clear-contents. Deleting a worksheet is a STRUCTURAL change that
invalidates ModelRisk's in-session store — live-verified: after a
scratch-sheet delete the store first degrades to returning 0.0 for
every in-range k (silent garbage!) and then to "No simulation results".
Plain cell writes/overwrites/clears leave the store intact.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from modelrisk_mcp.config import read_only_active
from modelrisk_mcp.errors import (
    ReadOnlyModeError,
    SimulationFailedError,
    WorkbookNotFoundError,
)

if TYPE_CHECKING:
    from modelrisk_mcp.bridge.excel import ExcelBridge

# Extraction cap. One column write + one bulk read stays fast up to
# ~100k rows; runs bigger than this get their first 100k iterations
# (plenty for every statistic we serve). Excel's row limit is ~1.048M
# anyway.
_MAX_EXTRACT = 100_000

# Columns of clearance between the sheet's used range and the scratch
# column, so a stray formula fill next to user data is impossible.
_SCRATCH_GAP = 2

# Iteration-count detection cap — matches the simulation soft cap.
_MAX_DETECT = 10_000_000

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")


def _qualified_ref(sheet: str, cell: str) -> str:
    """`('My Sheet', 'B2')` → `'My Sheet'!$B$2`.

    The absolute form is REQUIRED: the formula string is assigned to a
    whole scratch column at once, and Excel treats relative refs in a
    bulk `.Formula` assignment as relative PER CELL — an unanchored B2
    would silently become B3, B4, … down the fill."""
    m = _CELL_RE.match(cell.strip())
    if not m:
        raise SimulationFailedError(
            f"Cell reference {cell!r} is not a plain A1-style single-cell "
            "reference; cannot build a VoseSimValue probe for it."
        )
    col, row = m.groups()
    quoted = sheet.replace("'", "''")
    return f"'{quoted}'!${col.upper()}${row}"


def _is_sample(value: Any) -> bool:
    """A live sample comes back as a float; a dead / out-of-range read
    is the text "No simulation results" (or an error type)."""
    return isinstance(value, (int, float)) and not isinstance(value, bool)


@dataclass
class _ScratchArea:
    """One scratch column on an existing sheet. `max_row` tracks the
    deepest row written so the cleanup clears exactly what we used."""

    sheet: Any
    column: int
    max_row: int


class InSessionSampleReader:
    """Extracts per-iteration samples from the live Excel session via
    `VoseSimValue` scratch formulas. One instance per ModelRiskBridge;
    stateless between calls (every call creates and removes its own
    scratch sheet)."""

    def __init__(self, excel: ExcelBridge) -> None:
        self._excel = excel
        self._scratch_last: _ScratchArea | None = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def read_many(
        self,
        workbook: str,
        targets: list[tuple[str, str, str]],
        *,
        max_n: int = _MAX_EXTRACT,
        n_hint: int | None = None,
    ) -> dict[str, tuple[float, ...]]:
        """Read samples for several variables in one scratch session.

        `targets` is a list of `(name, sheet, cell)` — the cell is the
        VoseInput / VoseOutput wrapper cell the store is keyed on.
        Returns `{name: samples}` for every target that has data; a
        target whose cell yields no samples is silently omitted (the
        caller decides whether that's an error).

        Raises SimulationFailedError when the session holds NO results
        at all, ReadOnlyModeError in read-only mode."""
        if read_only_active():
            raise ReadOnlyModeError(
                "Results reading without MRService.dll works by writing "
                "temporary VoseSimValue scratch formulas, which read-only "
                "mode blocks. Update ModelRisk from vosesoftware.com (the "
                "current release includes MRService.dll — the write-free "
                "results reader), or lift read-only mode."
            )
        if not targets:
            return {}
        max_n = max(1, min(max_n, _MAX_EXTRACT))
        app, book = self._book(workbook)

        prev_saved = None
        try:
            prev_saved = bool(book.api.Saved)
        except Exception:
            pass

        # ⚠ VoseSimValue resolves the results store against the ACTIVE
        # workbook (live-verified: with another book active, a fully
        # qualified reference into the sim'd book still returns "No
        # simulation results"). The post-sim active book is not
        # deterministic (ModelRisk simulates on a copy of the workbook
        # and restores focus best-effort), so activate the target book
        # for the extraction and restore focus afterwards.
        prev_active_book = None
        try:
            prev_active_book = str(app.api.ActiveWorkbook.Name)
        except Exception:
            pass
        if prev_active_book != workbook:
            try:
                book.activate()
            except Exception:
                pass

        # ⚠ NEVER add/delete a worksheet here: structural changes
        # invalidate the in-session store (module docstring). The
        # scratch is a spare column on the first target's own sheet,
        # cleared afterwards.
        scratch = self._scratch_area(book, targets[0][1])
        try:
            n: int | None = None
            out: dict[str, tuple[float, ...]] = {}
            for name, sheet, cellref in targets:
                ref = _qualified_ref(sheet, cellref)
                if n is None:
                    n = self._detect_iterations(book, scratch, ref, n_hint)
                    if n == 0:
                        # First target has no store — the run may not have
                        # recorded it, or there are no results at all.
                        # Try the next target before concluding.
                        n = None
                        continue
                samples = self._fill_and_read(scratch, ref, min(n, max_n))
                if samples:
                    out[name] = samples
            if not out:
                raise SimulationFailedError(
                    "No in-session simulation results are available for "
                    f"{[t[0] for t in targets]!r} in {workbook!r}. "
                    "In-session reading only works in the Excel session "
                    "that ran the simulation, for the most recent run — "
                    "run the simulation again (run_simulation), or update "
                    "ModelRisk from vosesoftware.com so MRService.dll can "
                    "read the saved .vmrs file directly."
                )
            return out
        finally:
            self._clear_scratch(book, prev_saved)
            if prev_active_book and prev_active_book != workbook:
                try:
                    app.books[prev_active_book].activate()
                except Exception:
                    pass

    # ------------------------------------------------------------------
    # COM-touching internals (kept small and overridable for tests)
    # ------------------------------------------------------------------

    def _book(self, workbook: str) -> tuple[Any, Any]:
        if not self._excel.is_connected():
            self._excel.connect()
        app = self._excel._app  # bridge-internal access, same family
        if app is None:
            raise WorkbookNotFoundError(
                "Excel is not connected; cannot read in-session results."
            )
        for b in app.books:
            if b.name == workbook:
                return app, b
        raise WorkbookNotFoundError(
            f"Workbook {workbook!r} is not open in the attached Excel — "
            "in-session results can only be read from the live session "
            "that ran the simulation."
        )

    def _scratch_area(self, book: Any, sheet_name: str) -> _ScratchArea:
        """Pick a scratch column on `sheet_name`: a safety gap past the
        sheet's used range, so it can't touch user data."""
        sheet = book.sheets[sheet_name]
        try:
            last_col = int(sheet.api.UsedRange.Column) + int(
                sheet.api.UsedRange.Columns.Count
            ) - 1
        except Exception:
            last_col = 1
        col = min(last_col + 1 + _SCRATCH_GAP, 16_300)
        area = _ScratchArea(sheet=sheet, column=col, max_row=0)
        self._scratch_last = area
        return area

    def _clear_scratch(self, book: Any, prev_saved: bool | None) -> None:
        area = getattr(self, "_scratch_last", None)
        if area is not None and area.max_row > 0:
            try:
                area.sheet.range(
                    (1, area.column), (area.max_row, area.column)
                ).clear_contents()
            except Exception:
                pass
        self._scratch_last = None
        if prev_saved is not None:
            try:
                # A pure read shouldn't flip the workbook to "unsaved".
                book.api.Saved = prev_saved
            except Exception:
                pass

    def _probe(self, scratch: _ScratchArea, ref: str, k: int) -> bool:
        """One-cell probe: does iteration `k` exist in the store?"""
        cell = scratch.sheet.range((1, scratch.column))
        scratch.max_row = max(scratch.max_row, 1)
        cell.formula = f"=VoseSimValue({ref},{k})"
        return _is_sample(cell.value)

    def _fill_and_read(
        self, scratch: _ScratchArea, ref: str, n: int
    ) -> tuple[float, ...]:
        """One column fill + one bulk read. ROW() gives 1..n in the
        scratch column; the target ref is absolute so it doesn't slide."""
        rng = scratch.sheet.range((1, scratch.column), (n, scratch.column))
        scratch.max_row = max(scratch.max_row, n)
        rng.formula = f"=VoseSimValue({ref},ROW())"
        raw = rng.value
        if n == 1:
            raw = [raw]
        out: list[float] = []
        for v in raw:
            item = v[0] if isinstance(v, (list, tuple)) else v
            if not _is_sample(item):
                break  # first dead cell = end of the store
            out.append(float(item))
        return tuple(out)

    # ------------------------------------------------------------------
    # Iteration-count detection
    # ------------------------------------------------------------------

    def _detect_iterations(
        self, book: Any, scratch: _ScratchArea, ref: str,
        n_hint: int | None,
    ) -> int:
        """How many iterations does the store hold for `ref`?

        Order: validate the caller's hint, then the workbook-persisted
        `SimOpt_Samples` defined name (written by VoseSetSimulOptions12),
        then an exponential + binary search. A candidate is confirmed
        by `probe(n) and not probe(n+1)` so a stale hint can never
        truncate or over-read."""
        if not self._probe(scratch, ref, 1):
            return 0
        for candidate in (n_hint, self._simopt_samples(book)):
            if (
                candidate
                and candidate > 0
                and self._probe(scratch, ref, candidate)
                and not self._probe(scratch, ref, candidate + 1)
            ):
                return int(candidate)
        # Exponential climb, then binary search the boundary.
        lo = 1  # known-valid
        hi = 1024
        while hi <= _MAX_DETECT and self._probe(scratch, ref, hi):
            lo = hi
            hi *= 8
        if hi > _MAX_DETECT:
            return _MAX_DETECT
        while hi - lo > 1:
            mid = (lo + hi) // 2
            if self._probe(scratch, ref, mid):
                lo = mid
            else:
                hi = mid
        return lo

    @staticmethod
    def _simopt_samples(book: Any) -> int | None:
        """The XLL persists sim options as SimOpt_* defined names; the
        iteration count lands in SimOpt_Samples as a literal (`=1000`).
        Anything non-literal is ignored — the caller re-validates by
        probe anyway."""
        try:
            refers = str(book.api.Names("SimOpt_Samples").RefersTo)
        except Exception:
            return None
        body = refers.lstrip("=").strip().strip('"')
        if body.isdigit():
            return int(body)
        return None


__all__ = ["InSessionSampleReader"]
