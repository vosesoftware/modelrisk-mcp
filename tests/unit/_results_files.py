"""Fake ModelRisk results files for the run_simulation tests.

`dmr_bytes` is a run file with the header fields that identify a run, laid
out like the engine's (UTF-16LE, CRLF). `vmrs_bytes` is what ModelRisk's
save makes of a run file: every byte XORed with 0x2D, then gzipped.
"""

from __future__ import annotations

import gzip
from datetime import datetime
from pathlib import Path


def dmr_bytes(book: str, *, start: datetime, iterations: int = 1000) -> bytes:
    stamp = start.strftime("%Y%m%d%H%M%S")
    header = (
        "[VERSION]: 5\r\n[IS_FREE_VER]: 0\r\n<HEADER>\r\n"
        f"[SpreadsheetName]: <{book}>\r\n[AUT]: <Vose Software>\r\n"
        f"[MD]: <{stamp}>\r\n[HeaderSize]:4096\r\n[NSims]:1\r\n[TSims]:1\r\n"
        f"[NIters]:{iterations}\r\n[FixedSeed]:1\r\n[NOuts]:1\r\n"
        f"[StartDate]:<{stamp}>\r\n[EndDate]:<{stamp}>\r\n<OUTPUTS>\r\n"
    )
    # Stand-in for the sample data: every byte value, so the XOR is exercised.
    return header.encode("utf-16-le") + bytes(range(256)) * 16


def vmrs_bytes(dmr: bytes) -> bytes:
    return gzip.compress(bytes(b ^ 0x2D for b in dmr))


def write_dmr(
    path: Path, book: str, *, start: datetime, iterations: int = 1000
) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(dmr_bytes(book, start=start, iterations=iterations))
    return path


def write_vmrs(
    path: Path, book: str, *, start: datetime, iterations: int = 1000
) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(
        vmrs_bytes(dmr_bytes(book, start=start, iterations=iterations))
    )
    return path
