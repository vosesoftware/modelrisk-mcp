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
    count_call_args,
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


class TestCountCallArgs:
    """`count_call_args` underpins VOSE-013 (alpha.35). Must correctly
    count top-level argument commas while skipping:
    - commas inside string literals
    - commas inside nested function calls
    - commas inside array literals `{...}`
    And must distinguish empty parens from a single-arg call."""

    def test_three_args(self) -> None:
        assert count_call_args("=VosePERT(10, 20, 30)", "VosePERT") == [3]

    def test_two_args(self) -> None:
        assert count_call_args("=VoseNormal(0, 1)", "VoseNormal") == [2]

    def test_one_arg(self) -> None:
        assert count_call_args("=VoseExpon(1.5)", "VoseExpon") == [1]

    def test_zero_args(self) -> None:
        assert count_call_args("=VoseOutput()", "VoseOutput") == [0]

    def test_single_string_arg(self) -> None:
        """A single string-literal arg must count as 1, not 0 (the bug
        the alpha.35 prototype hit when it pre-stripped strings)."""
        assert count_call_args('=VoseOutput("name")', "VoseOutput") == [1]

    def test_comma_inside_string_does_not_split(self) -> None:
        assert count_call_args(
            '=VoseNormal("label,with,commas", 1)', "VoseNormal"
        ) == [2]

    def test_escaped_quote_inside_string(self) -> None:
        """Excel doubles `""` to embed a quote in a string literal."""
        assert count_call_args(
            '=VoseOutput("say ""hi"", world", "unit")', "VoseOutput"
        ) == [2]

    def test_comma_in_nested_call_does_not_split(self) -> None:
        assert count_call_args(
            "=VoseNormal(MAX(A1,B1), 5)", "VoseNormal"
        ) == [2]

    def test_array_literal_treated_as_single_arg(self) -> None:
        assert count_call_args(
            "=VoseDiscrete({0,1,2}, {0.3,0.4,0.3})", "VoseDiscrete"
        ) == [2]

    def test_multiple_calls(self) -> None:
        assert count_call_args(
            "=VosePERT(1,2,3) + VosePERT(4,5)", "VosePERT"
        ) == [3, 2]

    def test_nested_same_function(self) -> None:
        """Two calls to the same function, one inside the other.
        Both should be counted, the outer one with 3 (its outer args),
        the inner one with 3 (its inner args)."""
        assert count_call_args(
            "=VosePERT(1, VosePERT(2,3,4), 5)", "VosePERT"
        ) == [3, 3]

    def test_no_match_returns_empty(self) -> None:
        assert count_call_args("=SUM(A1:A10)", "VosePERT") == []

    def test_function_name_inside_string_is_not_a_call(self) -> None:
        """A function name that only appears inside a string literal
        (not as an actual call) must not be counted as a call site."""
        assert (
            count_call_args(
                '=VoseInput("the VosePERT one") + 1', "VosePERT"
            )
            == []
        )

    def test_lots_of_args(self) -> None:
        assert count_call_args(
            "=VosePERT(1,2,3,4,5,6,7,8)", "VosePERT"
        ) == [8]

    def test_whitespace_only_inside_parens(self) -> None:
        """Excel treats `VoseOutput( )` as a no-arg call."""
        assert count_call_args("=VoseOutput(   )", "VoseOutput") == [0]

    def test_unterminated_call_skipped(self) -> None:
        """A malformed formula with unclosed paren is some other rule's
        problem — this counter just skips it rather than crashing."""
        # The first VosePERT is unclosed; should yield no count for it.
        # Other rules / Excel itself will flag the syntax error.
        assert count_call_args("=VosePERT(1,2,3", "VosePERT") == []


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
