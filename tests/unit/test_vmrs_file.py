"""Tests for `bridge/vmrs_file.py`: reading a results header and rebuilding
a `.vmrs` from ModelRisk's run file the way its own save does."""

from __future__ import annotations

import gzip
from datetime import datetime
from pathlib import Path

import pytest

from modelrisk_mcp.bridge.vmrs_file import (
    read_dmr_header,
    read_vmrs_header,
    write_vmrs_from_dmr,
)
from tests.unit._results_files import dmr_bytes, write_dmr, write_vmrs

_START = datetime(2026, 10, 3, 23, 43, 1)


def test_saved_file_header_identifies_the_run(tmp_path: Path) -> None:
    path = write_vmrs(
        tmp_path / "a.vmrs", "OG-05 Prospect.xlsx", start=_START, iterations=10000
    )
    header = read_vmrs_header(path)
    assert header.spreadsheet_name == "OG-05 Prospect.xlsx"
    assert header.iterations == 10000
    assert header.start_date == _START
    assert header.end_date == _START


def test_run_file_header_reads_without_unpacking(tmp_path: Path) -> None:
    path = write_dmr(tmp_path / "fABCbook.xlsx.dmr", "book.xlsx", start=_START)
    header = read_dmr_header(path)
    assert header.spreadsheet_name == "book.xlsx"
    assert header.iterations == 1000
    assert header.start_date == _START


def test_rebuild_writes_what_modelrisks_save_writes(tmp_path: Path) -> None:
    """ModelRisk's save is compress_one_file(dmr, vmrs, bencrypt=true):
    XOR every byte with '-' and gzip. Unpacking the rebuilt file must give
    the run file back byte for byte."""
    dmr = write_dmr(tmp_path / "run.dmr", "book.xlsx", start=_START)
    target = tmp_path / "book.vmrs"

    write_vmrs_from_dmr(dmr, target)

    unpacked = bytes(b ^ 0x2D for b in gzip.decompress(target.read_bytes()))
    assert unpacked == dmr.read_bytes()
    assert read_vmrs_header(target).spreadsheet_name == "book.xlsx"


def test_rebuild_replaces_a_wrong_file_and_leaves_no_part_file(
    tmp_path: Path,
) -> None:
    dmr = write_dmr(tmp_path / "run.dmr", "book.xlsx", start=_START)
    target = write_vmrs(tmp_path / "book.vmrs", "other.xlsx", start=_START)

    write_vmrs_from_dmr(dmr, target)

    assert read_vmrs_header(target).spreadsheet_name == "book.xlsx"
    assert sorted(p.name for p in tmp_path.iterdir()) == ["book.vmrs", "run.dmr"]


def test_files_that_are_not_results_are_rejected(tmp_path: Path) -> None:
    text = tmp_path / "notes.vmrs"
    text.write_text("not gzip", encoding="utf-8")
    with pytest.raises(ValueError):
        read_vmrs_header(text)

    no_header = tmp_path / "empty.vmrs"
    no_header.write_bytes(gzip.compress(b"\x00" * 64))
    with pytest.raises(ValueError):
        read_vmrs_header(no_header)

    # A run file read as a saved file: not gzipped.
    dmr = tmp_path / "run.dmr"
    dmr.write_bytes(dmr_bytes("book.xlsx", start=_START))
    with pytest.raises(ValueError):
        read_vmrs_header(dmr)
