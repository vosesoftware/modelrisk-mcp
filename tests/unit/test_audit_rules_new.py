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
