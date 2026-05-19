"""Tests for `modelrisk_mcp.bridge.formulas`.

These are the core hallucination-firewall tests — every formula written
to Excel goes through these functions, so we exercise the validator
heavily.
"""

from __future__ import annotations

import pytest

from modelrisk_mcp.bridge.catalogue import load_catalogue
from modelrisk_mcp.bridge.formulas import (
    build_aggregate_mc,
    build_copula,
    build_distribution_formula,
    build_input_wrapper,
    build_output_wrapper,
    build_risk_event,
    build_time_series,
    render_value,
)
from modelrisk_mcp.errors import ParameterMismatchError, UnknownFunctionError


@pytest.fixture(scope="module")
def cat():
    return load_catalogue()


# ----------------------------------------------------------------------
# render_value
# ----------------------------------------------------------------------


class TestRenderValue:
    def test_int(self) -> None:
        assert render_value(42) == "42"

    def test_float_whole(self) -> None:
        assert render_value(1.0) == "1"

    def test_float_fractional(self) -> None:
        assert render_value(0.5) == "0.5"

    def test_bool_true(self) -> None:
        assert render_value(True) == "TRUE"

    def test_bool_false(self) -> None:
        assert render_value(False) == "FALSE"

    def test_array_1d(self) -> None:
        assert render_value([1, 2, 3]) == "{1,2,3}"

    def test_array_2d(self) -> None:
        assert render_value([[1, 2], [3, 4]]) == "{1,2;3,4}"

    def test_array_empty_rejected(self) -> None:
        with pytest.raises(ParameterMismatchError):
            render_value([])

    def test_array_mixed_dim_rejected(self) -> None:
        with pytest.raises(ParameterMismatchError):
            render_value([1, [2, 3]])

    def test_cell_ref(self) -> None:
        assert render_value("B12") == "B12"

    def test_cell_ref_with_sheet(self) -> None:
        assert render_value("Sheet1!A1") == "Sheet1!A1"

    def test_range_ref(self) -> None:
        assert render_value("A1:A10") == "A1:A10"

    def test_formula_fragment_function_call(self) -> None:
        assert render_value("VoseNormal(0,1)") == "VoseNormal(0,1)"

    def test_leading_equals_stripped(self) -> None:
        assert render_value("=B12") == "B12"

    def test_plain_string_quoted(self) -> None:
        assert render_value("hello world") == '"hello world"'

    def test_string_with_embedded_quote_escaped(self) -> None:
        assert render_value('he said "hi"') == '"he said ""hi"""'

    def test_unsupported_type_rejected(self) -> None:
        with pytest.raises(ParameterMismatchError):
            render_value({"not": "supported"})  # type: ignore[arg-type]


# ----------------------------------------------------------------------
# build_distribution_formula — core acceptance test
# ----------------------------------------------------------------------


class TestBuildDistributionFormula:
    def test_unknown_function_raises_with_suggestion(self, cat) -> None:
        # Phase 1 acceptance criterion (spec §13).
        with pytest.raises(UnknownFunctionError) as exc:
            build_distribution_formula("VoseFoo", {}, cat)
        assert "Did you mean" in str(exc.value)

    def test_vose_normal_with_required_args(self, cat) -> None:
        formula = build_distribution_formula(
            "VoseNormal", {"mu": 0, "sigma": 1}, cat
        )
        assert formula == "=VoseNormal(0,1)"

    def test_vose_modpert_with_cell_refs(self, cat) -> None:
        # VoseModPERT's signature in the IDL marks all of min/mode/max/gamma
        # as required parameters (no [optional] flag), so the formula
        # builder demands all four.
        formula = build_distribution_formula(
            "VoseModPERT",
            {"min": "B1", "mode": "B2", "max": "B3", "gamma": 4},
            cat,
        )
        assert formula == "=VoseModPERT(B1,B2,B3,4)"

    def test_optional_omitted(self, cat) -> None:
        # u, extended1, extended2 are optional and should be dropped.
        formula = build_distribution_formula(
            "VoseNormal", {"mu": 1, "sigma": 0.5}, cat
        )
        assert formula == "=VoseNormal(1,0.5)"

    def test_optional_supplied(self, cat) -> None:
        formula = build_distribution_formula(
            "VoseNormal", {"mu": 0, "sigma": 1, "u": 0.95}, cat
        )
        assert formula == "=VoseNormal(0,1,0.95)"

    def test_optional_skipped_middle_emits_empty_slot(self, cat) -> None:
        # If a later optional is supplied without an earlier one, the
        # earlier slot must be empty so positional argument matching
        # works in Excel.
        formula = build_distribution_formula(
            "VoseNormal", {"mu": 0, "sigma": 1, "extended1": 7}, cat
        )
        # u was skipped → empty slot
        assert formula == "=VoseNormal(0,1,,7)"

    def test_missing_required_raises(self, cat) -> None:
        with pytest.raises(ParameterMismatchError) as exc:
            build_distribution_formula("VoseNormal", {"mu": 0}, cat)
        assert "sigma" in str(exc.value)

    def test_unknown_param_raises(self, cat) -> None:
        with pytest.raises(ParameterMismatchError):
            build_distribution_formula(
                "VoseNormal", {"mu": 0, "sigma": 1, "bogus": 7}, cat
            )

    def test_positional_args_accepted(self, cat) -> None:
        formula = build_distribution_formula("VoseNormal", [0, 1], cat)
        assert formula == "=VoseNormal(0,1)"

    def test_positional_too_many_rejected(self, cat) -> None:
        with pytest.raises(ParameterMismatchError):
            build_distribution_formula(
                "VoseNormal", [0, 1, 0.5, 7, 8, 9, 99], cat
            )

    def test_array_param_for_discrete(self, cat) -> None:
        formula = build_distribution_formula(
            "VoseDiscrete",
            {"values": [1, 2, 3], "probabilities": [0.2, 0.5, 0.3]},
            cat,
        )
        assert formula == "=VoseDiscrete({1,2,3},{0.2,0.5,0.3})"


# ----------------------------------------------------------------------
# Wrappers
# ----------------------------------------------------------------------


class TestWrappers:
    def test_input_wrapper(self) -> None:
        assert (
            build_input_wrapper("Demand", "=VoseNormal(100,10)")
            == '=VoseInput("Demand")+VoseNormal(100,10)'
        )

    def test_input_wrapper_strips_leading_equals(self) -> None:
        assert (
            build_input_wrapper("Demand", "VoseNormal(100,10)")
            == '=VoseInput("Demand")+VoseNormal(100,10)'
        )

    def test_output_wrapper(self) -> None:
        assert (
            build_output_wrapper("Profit", "=B1-B2")
            == '=VoseOutput("Profit")+B1-B2'
        )

    def test_wrapper_name_with_quotes_escaped(self) -> None:
        result = build_input_wrapper('Quoted "Name"', "=B1")
        assert '"Quoted ""Name"""' in result

    def test_empty_wrapper_name_rejected(self) -> None:
        with pytest.raises(ParameterMismatchError):
            build_input_wrapper("", "=B1")


# ----------------------------------------------------------------------
# Composite builders
# ----------------------------------------------------------------------


class TestCompositeBuilders:
    def test_aggregate_mc_minimal(self, cat) -> None:
        formula = build_aggregate_mc("A1", "B1", cat)
        assert formula == "=VoseAggregateMC(A1,B1)"

    def test_aggregate_mc_with_limits(self, cat) -> None:
        formula = build_aggregate_mc("A1", "B1", cat, min_limit=0, max_limit=1000)
        assert formula == "=VoseAggregateMC(A1,B1,0,1000)"

    def test_risk_event_requires_object_impact(self, cat) -> None:
        # VoseNormal is a continuous distribution, not an object — should reject.
        with pytest.raises(ParameterMismatchError) as exc:
            build_risk_event(0.1, "VoseNormal", {"mu": 5, "sigma": 1}, cat)
        assert "object" in str(exc.value).lower()

    def test_risk_event_with_object(self, cat) -> None:
        # VoseLognormalObject is in the 'object' category — should accept.
        if "VoseLognormalObject" not in cat:
            pytest.skip("VoseLognormalObject not in catalogue")
        formula = build_risk_event(
            0.05, "VoseLognormalObject", {"mu": 10, "sigma": 2}, cat
        )
        assert "=VoseRiskEvent(0.05," in formula
        assert "VoseLognormalObject(10,2)" in formula

    def test_time_series_requires_time_series_category(self, cat) -> None:
        with pytest.raises(ParameterMismatchError) as exc:
            build_time_series("VoseNormal", {"mu": 0, "sigma": 1}, cat)
        assert "time-series" in str(exc.value).lower()

    def test_time_series_accepts_time_function(self, cat) -> None:
        # VoseTimeGBM is in the time-series category.
        formula = build_time_series(
            "VoseTimeGBM",
            {"OutputSize": 10, "mu": 0.05, "sigma": 0.2},
            cat,
        )
        assert formula.startswith("=VoseTimeGBM(")

    def test_copula_requires_copula_category(self, cat) -> None:
        with pytest.raises(ParameterMismatchError):
            build_copula("VoseNormal", {"mu": 0, "sigma": 1}, cat)

    def test_copula_accepts_copula_function(self, cat) -> None:
        formula = build_copula(
            "VoseCopulaMultiNormal", {"cov_matrix": "A1:C3"}, cat
        )
        assert formula == "=VoseCopulaMultiNormal(A1:C3)"
