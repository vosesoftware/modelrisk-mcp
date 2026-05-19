"""Tests for `modelrisk_mcp.safety`.

Covers:
- Tokenised "is this a Vose formula?" detector (§11.5)
- Bulk-write guard (§11.3)
- Writes log appender (§11.6)
- Writer mutex (§11.8) — in-process semantics only; full cross-process
  semantics need integration tests.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from modelrisk_mcp.bridge.catalogue import load_catalogue
from modelrisk_mcp.errors import ConcurrentWriterError
from modelrisk_mcp.safety import (
    BULK_WRITE_THRESHOLD,
    WriterMutex,
    append_write_log,
    check_bulk_write,
    extract_call_heads,
    has_only_known_functions,
    is_vose_formula,
)


@pytest.fixture(scope="module")
def cat():
    return load_catalogue()


class TestExtractCallHeads:
    def test_simple(self) -> None:
        assert extract_call_heads("=VoseNormal(0,1)") == ["VoseNormal"]

    def test_wrapped(self) -> None:
        heads = extract_call_heads('=VoseInput("Demand")+VoseModPERT(1,2,3)')
        assert heads == ["VoseInput", "VoseModPERT"]

    def test_string_literal_ignored(self) -> None:
        # A function-name-looking token *inside* a quoted string must not
        # count as a call head.
        heads = extract_call_heads('=VoseInput("VoseTotal")+B12')
        assert heads == ["VoseInput"]

    def test_escaped_quote_in_string(self) -> None:
        # "" is the Excel escape for a literal quote.
        heads = extract_call_heads(
            '=VoseInput("with ""quotes""")+VoseNormal(0,1)'
        )
        assert heads == ["VoseInput", "VoseNormal"]

    def test_plain_excel_function(self) -> None:
        assert extract_call_heads("=SUM(A1:A10)") == ["SUM"]

    def test_nested(self) -> None:
        heads = extract_call_heads("=IF(B1>0,VoseNormal(0,1),0)")
        assert heads == ["IF", "VoseNormal"]


class TestIsVoseFormula:
    def test_plain_vose_call(self, cat) -> None:
        assert is_vose_formula("=VoseNormal(0,1)", cat) is True

    def test_wrapped_formula(self, cat) -> None:
        assert (
            is_vose_formula('=VoseInput("x")+VoseModPERT(1,2,3)', cat) is True
        )

    def test_plain_excel_formula(self, cat) -> None:
        assert is_vose_formula("=SUM(A1:A10)", cat) is False

    def test_named_range_called_vose_total_not_treated_as_vose(
        self, cat,
    ) -> None:
        """Regression: the bad substring check would incorrectly flag a
        formula referencing a named range like `VoseTotal` as Vose. The
        tokenised check must not be fooled."""
        assert is_vose_formula("=SUM(VoseTotal)", cat) is False

    def test_vose_function_inside_string_not_treated_as_vose(
        self, cat,
    ) -> None:
        # If "VoseNormal" only appears inside a string literal, it's not a call.
        assert (
            is_vose_formula('=CONCAT("VoseNormal(0,1)","")', cat)
            is False
        )

    def test_empty_formula(self, cat) -> None:
        assert is_vose_formula("", cat) is False
        assert is_vose_formula("   ", cat) is False


class TestHasOnlyKnownFunctions:
    def test_pure_vose(self, cat) -> None:
        assert (
            has_only_known_functions("=VoseInput(\"x\")+VoseNormal(0,1)", cat)
            is True
        )

    def test_mixed_with_sum(self, cat) -> None:
        # SUM is not in the Vose catalogue, so this is False.
        assert (
            has_only_known_functions("=SUM(A1:A10)+VoseNormal(0,1)", cat)
            is False
        )

    def test_no_calls(self, cat) -> None:
        assert has_only_known_functions("=B12*1.1", cat) is True


class TestBulkWriteGuard:
    def test_under_threshold_ok(self) -> None:
        r = check_bulk_write(10)
        assert r.requires_confirmation is False

    def test_at_threshold_ok(self) -> None:
        r = check_bulk_write(BULK_WRITE_THRESHOLD)
        assert r.requires_confirmation is False

    def test_over_threshold_requires_confirmation(self) -> None:
        with pytest.raises(PermissionError):
            check_bulk_write(BULK_WRITE_THRESHOLD + 1)

    def test_over_threshold_with_confirm_ok(self) -> None:
        r = check_bulk_write(BULK_WRITE_THRESHOLD + 1, confirm_bulk=True)
        assert r.cell_count == BULK_WRITE_THRESHOLD + 1

    def test_exempt_bypasses_check(self) -> None:
        # Time-series / copula tools pass `exempt=True`.
        r = check_bulk_write(1000, exempt=True)
        assert r.requires_confirmation is False


class TestAppendWriteLog:
    def test_appends_jsonl(self, tmp_path: Path) -> None:
        log = tmp_path / "writes.log"
        append_write_log(
            cell="Sheet1!B12",
            before_formula="",
            before_value=42,
            after_formula="=VoseNormal(0,1)",
            log_path=log,
        )
        append_write_log(
            cell="Sheet1!B13",
            before_formula="=B12*1.1",
            before_value=46.2,
            after_formula="=VoseModPERT(1,2,3)",
            log_path=log,
        )
        lines = log.read_text(encoding="utf-8").strip().splitlines()
        assert len(lines) == 2
        record = json.loads(lines[0])
        assert record["cell"] == "Sheet1!B12"
        assert record["after_formula"] == "=VoseNormal(0,1)"
        assert "ts" in record


class TestWriterMutex:
    """Windows named mutexes are reentrant per *thread*, not per process.
    So same-thread contention can't be verified — we use a worker thread
    to confirm that a foreign thread (and by extension, a foreign
    process) is blocked."""

    def test_acquire_release_roundtrip(self) -> None:
        m = WriterMutex(name="modelrisk-mcp-test-roundtrip")
        try:
            assert m.acquire(timeout_ms=1000) is True
        finally:
            m.release()

    def test_other_thread_blocks_while_held(self) -> None:
        import threading

        m_main = WriterMutex(name="modelrisk-mcp-test-contended-thread")
        m_main.acquire(timeout_ms=500)

        result: dict[str, bool] = {}

        def worker() -> None:
            m_worker = WriterMutex(name="modelrisk-mcp-test-contended-thread")
            try:
                result["acquired"] = m_worker.acquire(timeout_ms=100)
            finally:
                m_worker.release()

        t = threading.Thread(target=worker)
        t.start()
        t.join(timeout=5)
        m_main.release()

        assert result.get("acquired") is False, (
            "Worker thread acquired the mutex while the main thread held it; "
            "expected acquire() to return False after the 100ms timeout."
        )

    def test_context_manager_raises_on_contention(self) -> None:
        import threading

        m_main = WriterMutex(name="modelrisk-mcp-test-ctxmgr")
        m_main.acquire(timeout_ms=500)

        result: dict[str, bool] = {"raised": False}

        def worker() -> None:
            m_worker = WriterMutex(name="modelrisk-mcp-test-ctxmgr")
            try:
                with m_worker.held(timeout_ms=100):
                    pass
            except ConcurrentWriterError:
                result["raised"] = True

        t = threading.Thread(target=worker)
        t.start()
        t.join(timeout=5)
        m_main.release()

        assert result["raised"] is True

    def test_context_manager_releases_on_exit(self) -> None:
        name = "modelrisk-mcp-test-release"
        with WriterMutex(name=name).held(timeout_ms=500):
            pass
        # After release, a fresh acquire should succeed.
        m = WriterMutex(name=name)
        assert m.acquire(timeout_ms=500) is True
        m.release()
