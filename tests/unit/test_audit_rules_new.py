"""Tests for the 5 new detectors added in v0.3.0-alpha.4.

Each detector is tested with one positive case (rule fires) and one
negative case (similar but valid formula doesn't trigger). For numeric
threshold rules (VOSE-011) we also include a boundary check.

The bridge isn't needed for these detectors — they read only from
`ctx.cells` and `ctx.catalogue` — so we hand-roll a minimal RuleContext
in each test rather than wiring a fake ModelRiskBridge."""

from __future__ import annotations

from typing import Any
from unittest.mock import MagicMock

import pytest

from modelrisk_mcp.audit.engine import RuleContext, RuleSpec
from modelrisk_mcp.audit.rules import (
    detect_arg_count_mismatch,
    detect_cell_evaluates_to_error,
    detect_duplicate_output_names,
    detect_high_volatility_normal_positive_mean,
    detect_input_wrapper_without_distribution,
    detect_risk_event_degenerate_probability,
    detect_voseoutput_missing_name,
)
from modelrisk_mcp.bridge.catalogue import load_catalogue
from modelrisk_mcp.schemas.workbook import CellInfo, CellRef


def _cell(workbook: str, sheet: str, ref: str, formula: str) -> CellInfo:
    return CellInfo(
        ref=CellRef(workbook=workbook, sheet=sheet, cell=ref),
        formula=formula,
        value=None,
        cell_type="formula" if formula else "empty",
    )


@pytest.fixture
def catalogue() -> Any:
    return load_catalogue()


def _make_ctx(
    cells: list[CellInfo],
    rule_name: str,
    severity: str = "warning",
    catalogue: Any = None,
) -> RuleContext:
    return RuleContext(
        bridge=MagicMock(),
        catalogue=catalogue or load_catalogue(),
        workbook="m.xlsx",
        cells=cells,
        rule=RuleSpec(
            id="VOSE-XXX",
            name=rule_name,
            severity=severity,
            enabled=True,
            description="test",
            suggested_fix_template="apply the fix",
        ),
    )


class TestRiskEventDegenerateProbability:
    def test_fires_on_probability_zero(self) -> None:
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=VoseRiskEvent(0, VoseLognormal(100,50))")],
            "risk_event_degenerate_probability",
        )
        findings = list(detect_risk_event_degenerate_probability(ctx))
        assert len(findings) == 1
        assert "= 0" in findings[0].message

    def test_fires_on_probability_one(self) -> None:
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=VoseRiskEvent(1, VoseNormal(0,1))")],
            "risk_event_degenerate_probability",
        )
        findings = list(detect_risk_event_degenerate_probability(ctx))
        assert len(findings) == 1
        assert "= 1" in findings[0].message

    def test_silent_on_valid_probability(self) -> None:
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=VoseRiskEvent(0.3, VoseNormal(0,1))")],
            "risk_event_degenerate_probability",
        )
        assert list(detect_risk_event_degenerate_probability(ctx)) == []

    def test_silent_on_cell_reference_probability(self) -> None:
        """We don't flag VoseRiskEvent(B5, ...) — we can't know B5's
        value statically. False positives are worse than false negatives
        here."""
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=VoseRiskEvent(B5, VoseNormal(0,1))")],
            "risk_event_degenerate_probability",
        )
        assert list(detect_risk_event_degenerate_probability(ctx)) == []


class TestVoseoutputMissingName:
    def test_fires_on_no_arg(self) -> None:
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=VoseOutput()+B1")],
            "voseoutput_missing_name",
            severity="error",
        )
        findings = list(detect_voseoutput_missing_name(ctx))
        assert len(findings) == 1

    def test_fires_on_empty_string(self) -> None:
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", '=VoseOutput("")+B1')],
            "voseoutput_missing_name",
            severity="error",
        )
        findings = list(detect_voseoutput_missing_name(ctx))
        assert len(findings) == 1

    def test_silent_on_named_output(self) -> None:
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", '=VoseOutput("Profit")+B1')],
            "voseoutput_missing_name",
            severity="error",
        )
        assert list(detect_voseoutput_missing_name(ctx)) == []


class TestDuplicateOutputNames:
    def test_fires_when_two_cells_share_name(self) -> None:
        ctx = _make_ctx(
            [
                _cell("m.xlsx", "S1", "A1", '=VoseOutput("Profit")+B1'),
                _cell("m.xlsx", "S2", "C5", '=VoseOutput("Profit")+D5'),
            ],
            "duplicate_output_names",
            severity="error",
        )
        findings = list(detect_duplicate_output_names(ctx))
        # Emits one finding per occurrence so the user sees all cells.
        assert len(findings) == 2
        assert all("Profit" in f.message for f in findings)

    def test_silent_on_unique_names(self) -> None:
        ctx = _make_ctx(
            [
                _cell("m.xlsx", "S1", "A1", '=VoseOutput("Profit")+B1'),
                _cell("m.xlsx", "S1", "A2", '=VoseOutput("Cost")+B2'),
            ],
            "duplicate_output_names",
            severity="error",
        )
        assert list(detect_duplicate_output_names(ctx)) == []


class TestInputWrapperWithoutDistribution:
    def test_fires_on_wrapper_with_no_distribution(self) -> None:
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", '=VoseInput("X")+B1')],
            "input_wrapper_without_distribution",
        )
        findings = list(detect_input_wrapper_without_distribution(ctx))
        assert len(findings) == 1

    def test_silent_when_distribution_present(self) -> None:
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", '=VoseInput("X")+VoseNormal(0,1)')],
            "input_wrapper_without_distribution",
        )
        assert list(detect_input_wrapper_without_distribution(ctx)) == []

    def test_silent_on_constants_without_wrapper(self) -> None:
        """A bare constant doesn't trigger this rule — VOSE-006 covers
        hard-coded inputs separately."""
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=42")],
            "input_wrapper_without_distribution",
        )
        assert list(detect_input_wrapper_without_distribution(ctx)) == []


class TestHighVolatilityNormalPositiveMean:
    def test_fires_when_sigma_exceeds_half_mu(self) -> None:
        # mu=100, sigma=60 → sigma > mu/2 (50). ~16% of samples negative.
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=VoseNormal(100, 60)")],
            "high_volatility_normal_positive_mean",
        )
        findings = list(detect_high_volatility_normal_positive_mean(ctx))
        assert len(findings) == 1
        assert "100" in findings[0].message
        assert "60" in findings[0].message

    def test_silent_at_boundary(self) -> None:
        """mu=100, sigma=50 exactly: sigma == mu/2. Not flagged (only
        strictly greater triggers — the inequality matches the rule
        description)."""
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=VoseNormal(100, 50)")],
            "high_volatility_normal_positive_mean",
        )
        assert list(detect_high_volatility_normal_positive_mean(ctx)) == []

    def test_silent_on_low_volatility(self) -> None:
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=VoseNormal(100, 10)")],
            "high_volatility_normal_positive_mean",
        )
        assert list(detect_high_volatility_normal_positive_mean(ctx)) == []

    def test_silent_on_negative_mean(self) -> None:
        """The rule is about positive quantities. If mu < 0 the user is
        modelling something where negatives are expected (deltas, P&L)."""
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=VoseNormal(-50, 100)")],
            "high_volatility_normal_positive_mean",
        )
        assert list(detect_high_volatility_normal_positive_mean(ctx)) == []

    def test_silent_on_cell_reference_args(self) -> None:
        """Static analysis can't evaluate VoseNormal(A1, B1)."""
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=VoseNormal(A1, B1)")],
            "high_volatility_normal_positive_mean",
        )
        assert list(detect_high_volatility_normal_positive_mean(ctx)) == []


# ----------------------------------------------------------------------
# VOSE-012 — cell evaluates to an Excel error (alpha.33)
# ----------------------------------------------------------------------


def _errored_cell(
    workbook: str, sheet: str, ref: str, formula: str, error: str,
) -> CellInfo:
    """Helper: a cell whose formula evaluated to `error` (e.g. '#DIV/0!')."""
    return CellInfo(
        ref=CellRef(workbook=workbook, sheet=sheet, cell=ref),
        formula=formula,
        value=None,
        cell_type="error",
        error=error,
    )


class TestCellEvaluatesToError:
    """VOSE-012 (alpha.33): error cells (`#DIV/0!` etc.) get flagged
    after bug-#34's CellInfo.error field is populated by the bridge.
    Vose calls inside errored cells get a sharper diagnostic message."""

    def test_fires_on_vose_call_evaluating_to_error(self) -> None:
        """The high-value case: a distribution call whose argument
        resolved to #DIV/0!. Without alpha.33 this was invisible."""
        ctx = _make_ctx(
            [
                _errored_cell(
                    "m.xlsx", "S1", "A1",
                    "=VosePERT(10, #DIV/0!, 30)",
                    "#DIV/0!",
                ),
            ],
            "cell_evaluates_to_error",
            severity="error",
        )
        findings = list(detect_cell_evaluates_to_error(ctx))
        assert len(findings) == 1
        assert findings[0].cell.cell == "A1"
        # The message must mention VosePERT and the error literal.
        assert "VosePERT" in findings[0].message
        assert "#DIV/0!" in findings[0].message
        # Distribution-specific phrasing should be present.
        assert "simulation" in findings[0].message.lower()

    def test_fires_on_non_vose_error_with_generic_message(self) -> None:
        """A vanilla errored cell (no Vose call) still fires, but the
        message uses the generic 'trace the formula' phrasing instead
        of the simulation-specific text."""
        ctx = _make_ctx(
            [
                _errored_cell(
                    "m.xlsx", "S1", "B5",
                    "=B4/B3",
                    "#DIV/0!",
                ),
            ],
            "cell_evaluates_to_error",
            severity="error",
        )
        findings = list(detect_cell_evaluates_to_error(ctx))
        assert len(findings) == 1
        assert "trace" in findings[0].message.lower()
        assert "#DIV/0!" in findings[0].message
        # No Vose call mentioned.
        assert "Vose" not in findings[0].message

    def test_silent_on_clean_cells(self) -> None:
        """Healthy cells (value present, no error) must not fire."""
        ctx = _make_ctx(
            [
                _cell("m.xlsx", "S1", "A1", "=1+1"),
                _cell("m.xlsx", "S1", "A2", "=VoseNormal(0, 1)"),
            ],
            "cell_evaluates_to_error",
            severity="error",
        )
        assert list(detect_cell_evaluates_to_error(ctx)) == []

    def test_fires_once_per_errored_cell(self) -> None:
        """One finding per errored cell — no duplicates."""
        ctx = _make_ctx(
            [
                _errored_cell("m.xlsx", "S1", "A1", "=1/0", "#DIV/0!"),
                _errored_cell("m.xlsx", "S1", "A2", "=BogusFn()", "#NAME?"),
                _errored_cell("m.xlsx", "S1", "A3", "=#REF!+1", "#REF!"),
            ],
            "cell_evaluates_to_error",
            severity="error",
        )
        findings = list(detect_cell_evaluates_to_error(ctx))
        assert len(findings) == 3
        # Each error literal should appear once.
        refs = {f.cell.cell for f in findings}
        assert refs == {"A1", "A2", "A3"}

    def test_severity_inherits_from_rule_spec(self) -> None:
        """The severity is read from the rule spec, not hard-coded."""
        ctx = _make_ctx(
            [_errored_cell("m.xlsx", "S1", "A1", "=1/0", "#DIV/0!")],
            "cell_evaluates_to_error",
            severity="warning",
        )
        findings = list(detect_cell_evaluates_to_error(ctx))
        assert findings[0].severity == "warning"


# ----------------------------------------------------------------------
# VOSE-013 — arg-count mismatch against catalogue (alpha.35)
# ----------------------------------------------------------------------


class TestArgCountMismatch:
    """VOSE-013: Vose call has too few or too many arguments versus the
    catalogue's declared signature. Classic LLM hallucination case
    (`VosePERT(min, max)` — missing mode) that VOSE-001 doesn't catch
    because the function name is real."""

    def test_fires_on_too_few_args_pert(self) -> None:
        """`VosePERT` requires 3 args (min, mode, max). Two is wrong."""
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=VosePERT(10, 30)")],
            "arg_count_mismatch",
            severity="error",
        )
        findings = list(detect_arg_count_mismatch(ctx))
        assert len(findings) == 1
        assert "VosePERT" in findings[0].message
        assert "2 arguments" in findings[0].message
        assert "too few" in findings[0].message

    def test_fires_on_too_few_args_normal(self) -> None:
        """`VoseNormal` requires 2 args (mu, sigma). One is wrong."""
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "B2", "=VoseNormal(100)")],
            "arg_count_mismatch",
            severity="error",
        )
        findings = list(detect_arg_count_mismatch(ctx))
        assert len(findings) == 1
        assert "VoseNormal" in findings[0].message
        assert "1 argument" in findings[0].message
        # Singular: "1 argument" not "1 arguments".

    def test_fires_on_too_many_args(self) -> None:
        """`VosePERT` has up to 6 params (min, mode, max, U, ext1, ext2).
        Eight is beyond catalogue and almost certainly a hallucination."""
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "C3", "=VosePERT(1,2,3,4,5,6,7,8)")],
            "arg_count_mismatch",
            severity="error",
        )
        findings = list(detect_arg_count_mismatch(ctx))
        assert len(findings) == 1
        assert "8 arguments" in findings[0].message
        assert "too many" in findings[0].message

    def test_silent_on_correct_arity(self) -> None:
        """Healthy calls don't fire."""
        ctx = _make_ctx(
            [
                _cell("m.xlsx", "S1", "A1", "=VosePERT(10, 20, 30)"),
                _cell("m.xlsx", "S1", "A2", "=VoseNormal(0, 1)"),
                _cell("m.xlsx", "S1", "A3", "=VoseExpon(1.5)"),
            ],
            "arg_count_mismatch",
            severity="error",
        )
        assert list(detect_arg_count_mismatch(ctx)) == []

    def test_silent_on_optional_trailing_args(self) -> None:
        """Most Vose distributions accept optional U-object + extension
        args. A call with 4 args to VosePERT (min, mode, max, u_obj)
        should NOT fire."""
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=VosePERT(10, 20, 30, A5)")],
            "arg_count_mismatch",
            severity="error",
        )
        assert list(detect_arg_count_mismatch(ctx)) == []

    def test_silent_on_voseinput_voseoutput_wrappers(self) -> None:
        """VoseInput/VoseOutput catalogue rows show total=1 but Excel
        happily accepts VoseOutput() with no args (VOSE-008's job).
        VOSE-013 must NOT double-fire here."""
        ctx = _make_ctx(
            [
                _cell("m.xlsx", "S1", "A1", "=VoseInput(B1)+VoseNormal(0,1)"),
                _cell("m.xlsx", "S1", "A2", "=VoseOutput()"),
                _cell("m.xlsx", "S1", "A3", "=VoseOutput(\"Total\")"),
            ],
            "arg_count_mismatch",
            severity="error",
        )
        # No findings from VOSE-013 — only VOSE-008 handles wrappers.
        wrapper_findings = [
            f for f in detect_arg_count_mismatch(ctx)
            if any(w in f.message for w in ("VoseInput", "VoseOutput"))
        ]
        assert wrapper_findings == []

    def test_silent_on_unknown_function(self) -> None:
        """Unknown Vose-prefixed functions are VOSE-001's job; VOSE-013
        must skip them so we don't double-fire."""
        ctx = _make_ctx(
            [_cell("m.xlsx", "S1", "A1", "=VoseNomral(10, 5)")],
            "arg_count_mismatch",
            severity="error",
        )
        assert list(detect_arg_count_mismatch(ctx)) == []

    def test_silent_on_non_vose_calls(self) -> None:
        """Excel built-ins like SUM/IF/VLOOKUP must never trigger."""
        ctx = _make_ctx(
            [
                _cell("m.xlsx", "S1", "A1", "=SUM(A1:A10)"),
                _cell("m.xlsx", "S1", "A2", "=IF(B1>0, 1)"),
                _cell("m.xlsx", "S1", "A3", "=VLOOKUP(A,table,2,FALSE)"),
            ],
            "arg_count_mismatch",
            severity="error",
        )
        assert list(detect_arg_count_mismatch(ctx)) == []

    def test_fires_on_nested_vose_call_with_wrong_arity(self) -> None:
        """A nested broken call must still be caught."""
        ctx = _make_ctx(
            [
                _cell(
                    "m.xlsx", "S1", "A1",
                    "=VoseInput(\"x\") + VoseLognormal(100)",
                ),
            ],
            "arg_count_mismatch",
            severity="error",
        )
        findings = list(detect_arg_count_mismatch(ctx))
        assert len(findings) == 1
        assert "VoseLognormal" in findings[0].message

    def test_suggested_fix_includes_catalogue_signature(self) -> None:
        """The suggested-fix template should expose the expected arity
        and required-param names so the LLM can self-repair. Uses the
        real YAML template (not the test fixture's placeholder) so we
        verify the actual `{function_name}` / `{min_args}` / etc.
        placeholder substitution."""
        real_template = (
            "Check the function's signature in the catalogue. "
            "{function_name} expects {min_args}-{max_args} args; "
            "this call has {actual_args}. Required params: "
            "{required_params}."
        )
        ctx = RuleContext(
            bridge=MagicMock(),
            catalogue=load_catalogue(),
            workbook="m.xlsx",
            cells=[_cell("m.xlsx", "S1", "A1", "=VosePERT(10, 30)")],
            rule=RuleSpec(
                id="VOSE-013",
                name="arg_count_mismatch",
                severity="error",
                enabled=True,
                description="test",
                suggested_fix_template=real_template,
            ),
        )
        findings = list(detect_arg_count_mismatch(ctx))
        fix = findings[0].suggested_fix or ""
        assert "VosePERT" in fix
        assert "3" in fix  # min args (required count)
        # Required params should be enumerated.
        assert "min" in fix and "mode" in fix and "max" in fix

    def test_one_finding_per_unique_call(self) -> None:
        """Two cells calling the same function with the same wrong arity
        should each get their own finding (one per cell), not one
        deduplicated finding."""
        ctx = _make_ctx(
            [
                _cell("m.xlsx", "S1", "A1", "=VosePERT(1, 2)"),
                _cell("m.xlsx", "S1", "A2", "=VosePERT(3, 4)"),
            ],
            "arg_count_mismatch",
            severity="error",
        )
        findings = list(detect_arg_count_mismatch(ctx))
        assert len(findings) == 2
        refs = {f.cell.cell for f in findings}
        assert refs == {"A1", "A2"}
