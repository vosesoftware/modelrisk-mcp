"""ModelRisk's simulation-results files: read the header, rebuild a `.vmrs`.

A run lives in two files:

- the engine's own run file, `f<hWndExcel hex><book>.dmr`, in its
  simulation folder (normally `C:\\Vose Software\\ModelRisk\\SimulationStorage`).
  ModelRisk rewrites it on every run of that workbook.
- the `.vmrs` a save produces from it: every byte XORed with '-' (0x2D),
  then gzipped. That is all `compress_one_file(dmr, vmrs, bencrypt=true)`
  does (ModelRiskCloude/zlib.cpp:48), and `SaveWorkbookResults(sc, path)`
  is that one call on the container's run file.

Both start with the same UTF-16LE text header, at fixed offsets
(`<OUTPUTS>` sits at byte 2380 in every file seen):

    [SpreadsheetName]: <Book.xlsx>
    [NIters]:10000
    [StartDate]:<20261003233209>
    [EndDate]:<20261003233224>

`run_simulation` uses the header to prove that a saved file holds the run it
just made, and the rebuild to recover when ModelRisk's save hands back
another run (see `SimulationController._save_results`).
"""

from __future__ import annotations

import gzip
import os
import re
import zlib
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

_XOR_TABLE = bytes(i ^ 0x2D for i in range(256))
# The header keys end well inside the first 3 KB; read a generous prefix.
_HEADER_PROBE_BYTES = 64 * 1024
_CHUNK = 1 << 20

_BOOK_RE = re.compile(r"\[SpreadsheetName\]:\s*<([^>\r\n]*)>")
_ITERS_RE = re.compile(r"\[NIters\]:\s*(\d+)")
_START_RE = re.compile(r"\[StartDate\]:\s*<(\d{14})>")
_END_RE = re.compile(r"\[EndDate\]:\s*<(\d{14})>")


@dataclass(frozen=True)
class ResultsHeader:
    """The fields of a results header that identify a run."""

    spreadsheet_name: str
    iterations: int | None
    start_date: datetime | None
    end_date: datetime | None


def read_vmrs_header(path: str | os.PathLike[str]) -> ResultsHeader:
    """Header of a saved `.vmrs`. Raises ValueError if the file is not one."""
    try:
        with gzip.open(path, "rb") as fh:
            raw = fh.read(_HEADER_PROBE_BYTES)
    except (gzip.BadGzipFile, EOFError, zlib.error) as exc:
        raise ValueError(f"not a gzipped results file: {exc}") from exc
    return _parse(raw.translate(_XOR_TABLE))


def read_dmr_header(path: str | os.PathLike[str]) -> ResultsHeader:
    """Header of ModelRisk's run file. Raises ValueError if the file is not one."""
    with open(path, "rb") as fh:
        raw = fh.read(_HEADER_PROBE_BYTES)
    return _parse(raw)


def write_vmrs_from_dmr(
    dmr_path: str | os.PathLike[str], vmrs_path: str | os.PathLike[str]
) -> None:
    """Write the `.vmrs` ModelRisk's own save would write from this run file.

    The file appears at `vmrs_path` only once it is complete."""
    target = Path(vmrs_path)
    part = target.with_name(target.name + ".part")
    try:
        with open(dmr_path, "rb") as src, open(part, "wb") as raw_out:
            # filename="" and mtime=0: no name or timestamp in the gzip
            # header, as zlib's gzopen writes it. Level 6 is zlib's default.
            with gzip.GzipFile(
                filename="", mode="wb", fileobj=raw_out, compresslevel=6, mtime=0
            ) as out:
                while chunk := src.read(_CHUNK):
                    out.write(chunk.translate(_XOR_TABLE))
        os.replace(part, target)
    finally:
        part.unlink(missing_ok=True)


def _parse(raw: bytes) -> ResultsHeader:
    text = raw.decode("utf-16-le", errors="ignore")
    book = _BOOK_RE.search(text)
    if book is None or "[HeaderSize]" not in text:
        raise ValueError("no ModelRisk results header")
    iters = _ITERS_RE.search(text)
    return ResultsHeader(
        spreadsheet_name=book.group(1).strip(),
        iterations=int(iters.group(1)) if iters else None,
        start_date=_stamp(_START_RE.search(text)),
        end_date=_stamp(_END_RE.search(text)),
    )


def _stamp(match: re.Match[str] | None) -> datetime | None:
    if match is None:
        return None
    try:
        return datetime.strptime(match.group(1), "%Y%m%d%H%M%S")
    except ValueError:
        return None


__all__ = [
    "ResultsHeader",
    "read_dmr_header",
    "read_vmrs_header",
    "write_vmrs_from_dmr",
]
