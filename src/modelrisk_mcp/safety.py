"""Safety mechanisms — see spec §11.

Implements:
- Tokenised "is this a Vose formula?" detector (replaces the bad substring
  check on the literal string "Vose"; matches by formula head against the
  catalogue).
- WriterMutex: a Windows named mutex so two MCP server instances cannot
  drive the same Excel concurrently.
- Bulk-write guard: enforce a >50-cell threshold for tools that need
  explicit `confirm_bulk=True` (time-series/copula tools are exempt and
  opt out by setting `exempt=True`).
- Writes-log append helper.
"""

from __future__ import annotations

import json
import os
import re
import sys
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import TYPE_CHECKING

from modelrisk_mcp.errors import ConcurrentWriterError

if TYPE_CHECKING:
    from modelrisk_mcp.bridge.catalogue import FunctionCatalogue


# Matches every IDENT(  in a formula. The IDENT is the "head" of a call
# expression; comparing it to the catalogue tells us if the cell contains
# any Vose function call (including wrapped inputs/outputs).
_CALL_HEAD_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*\(")

# Excel built-ins that are always treated as "non-Vose" no matter what
# else is in the formula — used by tests but useful for documentation too.
EXCEL_BUILTIN_FUNCTIONS_SAMPLE = frozenset({
    "SUM", "IF", "AND", "OR", "NOT", "VLOOKUP", "INDEX", "MATCH",
    "AVERAGE", "MIN", "MAX", "COUNT",
})


def extract_call_heads(formula: str) -> list[str]:
    """Return the list of IDENT(...) heads in a formula, in source order.

    Strips Excel string literals first so a function name inside a quoted
    string doesn't get counted. (e.g. `=VoseInput("VoseTotal")+...` has
    `VoseTotal` inside a string and only `VoseInput` outside.)
    """
    stripped = _strip_excel_strings(formula)
    return _CALL_HEAD_RE.findall(stripped)


def _strip_excel_strings(formula: str) -> str:
    # Excel string literals use double quotes, with "" as the escape for
    # an embedded quote. We replace each string with a placeholder of the
    # same length so column positions stay roughly intact for diagnostics.
    out: list[str] = []
    i = 0
    n = len(formula)
    while i < n:
        ch = formula[i]
        if ch != '"':
            out.append(ch)
            i += 1
            continue
        # entering a string literal
        out.append(" ")
        i += 1
        while i < n:
            if formula[i] == '"':
                if i + 1 < n and formula[i + 1] == '"':
                    out.append("  ")
                    i += 2
                    continue
                out.append(" ")
                i += 1
                break
            out.append(" ")
            i += 1
    return "".join(out)


def is_vose_formula(formula: str, catalogue: FunctionCatalogue) -> bool:
    """True iff any call expression in the formula has a head that is a
    known Vose function (per the catalogue). Wrapper-only formulas like
    `=VoseInput("x")+B12` still count as Vose."""
    if not formula.strip():
        return False
    heads = extract_call_heads(formula)
    return any(head in catalogue for head in heads)


def has_only_known_functions(formula: str, catalogue: FunctionCatalogue) -> bool:
    """Stricter than `is_vose_formula`: returns True iff every call in the
    formula is in the catalogue. Used to validate formulas we want to
    parse strictly. Returns True for a formula containing no calls at all
    (e.g. `=B12*1.1`)."""
    heads = extract_call_heads(formula)
    return all(head in catalogue for head in heads)


# ----------------------------------------------------------------------
# Bulk-write guard
# ----------------------------------------------------------------------

BULK_WRITE_THRESHOLD: int = 50


@dataclass(frozen=True)
class BulkWriteGuardResult:
    cell_count: int
    requires_confirmation: bool


def check_bulk_write(
    cell_count: int, *, confirm_bulk: bool = False, exempt: bool = False
) -> BulkWriteGuardResult:
    """Decide whether a tool call writing `cell_count` cells is allowed.

    Raises `PermissionError` if the count is over the threshold and the
    caller hasn't passed `confirm_bulk=True`. Time-series and copula
    tools that inherently write contiguous ranges pass `exempt=True`.
    """
    if exempt or cell_count <= BULK_WRITE_THRESHOLD or confirm_bulk:
        return BulkWriteGuardResult(
            cell_count=cell_count,
            requires_confirmation=False,
        )
    raise PermissionError(
        f"Refusing to write {cell_count} cells in one call (threshold: "
        f"{BULK_WRITE_THRESHOLD}). Pass `confirm_bulk=True` to override."
    )


# ----------------------------------------------------------------------
# Writes log (spec §11.6)
# ----------------------------------------------------------------------


def _default_writes_log_path() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "VoseSoftware" / "modelrisk-mcp" / "writes.log"


def append_write_log(
    *,
    cell: str,
    before_formula: str,
    before_value: object,
    after_formula: str,
    log_path: Path | None = None,
) -> Path:
    """Append a JSONL record describing one cell write. Returns the log
    path so callers can verify it landed."""
    path = log_path or _default_writes_log_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    record = {
        "ts": datetime.now(UTC).isoformat(),
        "cell": cell,
        "before_formula": before_formula,
        "before_value": before_value,
        "after_formula": after_formula,
    }
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")
    return path


# ----------------------------------------------------------------------
# Writer mutex (spec §11.8)
# ----------------------------------------------------------------------

_MUTEX_NAME: str = "modelrisk-mcp-excel-writer"


class WriterMutex:
    """Process-global mutex preventing two MCP server instances from
    concurrently writing to Excel.

    On Windows uses a named Win32 mutex via pywin32. On non-Windows (where
    this whole project doesn't run, but tests may), falls back to a
    file-based lock under %LOCALAPPDATA% so unit tests are exercisable.
    """

    def __init__(self, name: str = _MUTEX_NAME) -> None:
        self._name = name
        self._handle: object | None = None
        self._fallback_path: Path | None = None

    def acquire(self, timeout_ms: int = 0) -> bool:
        if sys.platform == "win32":
            return self._acquire_win32(timeout_ms)
        return self._acquire_filelock_fallback()

    def release(self) -> None:
        if self._handle is None and self._fallback_path is None:
            return
        if sys.platform == "win32":
            self._release_win32()
        else:
            self._release_filelock_fallback()

    @contextmanager
    def held(self, timeout_ms: int = 0) -> Iterator[None]:
        if not self.acquire(timeout_ms):
            raise ConcurrentWriterError(
                "Another MCP server instance holds the Excel writer mutex. "
                "Close the other client or wait for it to finish."
            )
        try:
            yield
        finally:
            self.release()

    # ------------- platform-specific helpers -----------------------------

    def _acquire_win32(self, timeout_ms: int) -> bool:
        try:
            import win32event
        except ImportError:
            # pywin32 not installed — fall through to the filelock path.
            return self._acquire_filelock_fallback()
        # bInitialOwner=False so we don't accidentally own it on creation.
        handle = win32event.CreateMutex(None, False, self._name)
        # WAIT_OBJECT_0 = 0, WAIT_TIMEOUT = 0x102
        wait_result = win32event.WaitForSingleObject(handle, timeout_ms)
        if wait_result != 0:
            # Either timed out or abandoned — treat as "not acquired".
            try:
                import win32api

                win32api.CloseHandle(handle)
            except ImportError:
                pass
            return False
        self._handle = handle
        return True

    def _release_win32(self) -> None:
        if self._handle is None:
            return
        try:
            import win32api
            import win32event
        except ImportError:
            self._handle = None
            return
        try:
            win32event.ReleaseMutex(self._handle)
        finally:
            win32api.CloseHandle(self._handle)
            self._handle = None

    def _acquire_filelock_fallback(self) -> bool:
        """Best-effort cross-process lock via O_EXCL file creation. Used
        on non-Windows for tests. Not as robust as a named mutex but
        good enough — the test suite only needs in-process semantics."""
        base = os.environ.get("LOCALAPPDATA") or str(
            Path.home() / "AppData" / "Local"
        )
        path = Path(base) / "VoseSoftware" / "modelrisk-mcp" / f"{self._name}.lock"
        path.parent.mkdir(parents=True, exist_ok=True)
        try:
            fd = os.open(str(path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
        except FileExistsError:
            return False
        self._fallback_path = path
        return True

    def _release_filelock_fallback(self) -> None:
        if self._fallback_path is not None and self._fallback_path.exists():
            try:
                self._fallback_path.unlink()
            except OSError:
                pass
        self._fallback_path = None
