"""Unit tests for `bridge/name_parser.py`.

The parser must handle both the string-literal form and the
cell-reference form of VoseInput/VoseOutput name arguments. The
cell-reference form is what real ModelRisk workbooks use most often
(labels in headers, referenced from VoseInput cells); previously the
scanner missed all of them.
"""

from __future__ import annotations

import pytest

from modelrisk_mcp.bridge.name_parser import (
    CellRefName,
    LiteralName,
    extract_vose_first_arg,
)

# ----------------------------------------------------------------------
# String-literal form
# ----------------------------------------------------------------------


class TestLiteralForm:
    def test_basic_literal(self) -> None:
        result = extract_vose_first_arg(
            '=VoseInput("WidgetCost")+VoseNormal(0,1)', "VoseInput",
        )
        assert isinstance(result, LiteralName)
        assert result.name == "WidgetCost"

    def test_literal_with_spaces(self) -> None:
        result = extract_vose_first_arg(
            '=VoseOutput("Total Cost")+B1', "VoseOutput",
        )
        assert isinstance(result, LiteralName)
        assert result.name == "Total Cost"

    def test_literal_with_doubled_quotes(self) -> None:
        """Excel escapes a literal " inside a string by doubling it."""
        result = extract_vose_first_arg(
            '=VoseInput("Cost ""per unit""")', "VoseInput",
        )
        assert isinstance(result, LiteralName)
        assert result.name == 'Cost "per unit"'

    def test_literal_with_special_chars(self) -> None:
        """Question marks, parentheses, slashes — all legal in ModelRisk
        names. The parser should pass them through verbatim."""
        result = extract_vose_first_arg(
            '=VoseInput("Conservatives get in? (1=yes)")', "VoseInput",
        )
        assert isinstance(result, LiteralName)
        assert result.name == "Conservatives get in? (1=yes)"

    def test_literal_with_leading_whitespace(self) -> None:
        result = extract_vose_first_arg(
            '=VoseInput(   "X")', "VoseInput",
        )
        assert isinstance(result, LiteralName)
        assert result.name == "X"

    def test_unterminated_literal_returns_none(self) -> None:
        result = extract_vose_first_arg(
            '=VoseInput("unterminated', "VoseInput",
        )
        assert result is None


# ----------------------------------------------------------------------
# Cell-reference form
# ----------------------------------------------------------------------


class TestCellRefForm:
    def test_same_sheet_ref(self) -> None:
        result = extract_vose_first_arg(
            "=VoseInput(A5)+VoseNormal(0,1)", "VoseInput",
        )
        assert isinstance(result, CellRefName)
        assert result.sheet == ""
        assert result.cell == "A5"

    def test_absolute_ref(self) -> None:
        result = extract_vose_first_arg(
            "=VoseInput($A$5)", "VoseInput",
        )
        assert isinstance(result, CellRefName)
        assert result.cell == "A5"

    def test_qualified_sheet_ref(self) -> None:
        result = extract_vose_first_arg(
            "=VoseInput(Sheet1!A5)", "VoseInput",
        )
        assert isinstance(result, CellRefName)
        assert result.sheet == "Sheet1"
        assert result.cell == "A5"

    def test_quoted_sheet_with_spaces(self) -> None:
        result = extract_vose_first_arg(
            "=VoseInput('Sheet with spaces'!B12)", "VoseInput",
        )
        assert isinstance(result, CellRefName)
        assert result.sheet == "Sheet with spaces"
        assert result.cell == "B12"

    def test_multicolumn_ref(self) -> None:
        result = extract_vose_first_arg(
            "=VoseInput(AB123)", "VoseInput",
        )
        assert isinstance(result, CellRefName)
        assert result.cell == "AB123"

    def test_voseoutput_with_cell_ref(self) -> None:
        result = extract_vose_first_arg(
            "=VoseOutput(B$2)+A1*C1", "VoseOutput",
        )
        assert isinstance(result, CellRefName)
        assert result.cell == "B2"


# ----------------------------------------------------------------------
# Unrecognised forms — be conservative
# ----------------------------------------------------------------------


class TestUnrecognised:
    def test_no_wrapper_returns_none(self) -> None:
        assert extract_vose_first_arg("=A1+B1", "VoseInput") is None

    def test_empty_formula(self) -> None:
        assert extract_vose_first_arg("", "VoseInput") is None

    def test_function_call_as_arg_returns_none(self) -> None:
        """We don't try to evaluate Excel expressions. A function call
        as the name arg is unrecognised — caller skips this cell."""
        result = extract_vose_first_arg(
            "=VoseInput(CONCAT(A1,B1))", "VoseInput",
        )
        assert result is None

    def test_arithmetic_as_arg_returns_none(self) -> None:
        result = extract_vose_first_arg(
            "=VoseInput(A1+B1)", "VoseInput",
        )
        assert result is None

    def test_range_as_arg_returns_none(self) -> None:
        result = extract_vose_first_arg(
            "=VoseInput(A1:A5)", "VoseInput",
        )
        assert result is None


# ----------------------------------------------------------------------
# Wrapper-name matching is exact
# ----------------------------------------------------------------------


class TestWrapperMatching:
    def test_voseinput_not_matched_when_asked_for_voseoutput(self) -> None:
        assert extract_vose_first_arg(
            '=VoseInput("X")', "VoseOutput"
        ) is None

    def test_voseoutput_not_matched_when_asked_for_voseinput(self) -> None:
        assert extract_vose_first_arg(
            '=VoseOutput("X")', "VoseInput"
        ) is None

    def test_wrapper_inside_other_function(self) -> None:
        """VoseInput nested inside another expression still extracts."""
        result = extract_vose_first_arg(
            '=IF(A1>0, VoseInput("X")+B1, 0)', "VoseInput",
        )
        assert isinstance(result, LiteralName)
        assert result.name == "X"


_ = pytest
