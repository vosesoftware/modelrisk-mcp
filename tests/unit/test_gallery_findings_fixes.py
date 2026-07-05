"""Regression tests for the demo-gallery live-test findings (0.3.9):

F1  text cells are not live formulas (scanner phantoms)
F3  cross-sheet data ranges are honoured, not double-qualified
F4  array functions are CSE-entered (covered in test_functional_tools)
F7  writes into merged non-anchor cells raise instead of silently no-op
F8  the phantom 'OutputSize' parameter is stripped from the catalogue
F10 the default percentile set includes 0.80 (true P80 reads)
    + _params_dict accepts positional {value}-only entries
"""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from modelrisk_mcp.bridge.catalogue import load_catalogue
from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.bridge.modelrisk import _is_live_formula, _offset_a1
from modelrisk_mcp.bridge.results import _DEFAULT_PERCENTILES
from modelrisk_mcp.errors import CellReferenceError
from modelrisk_mcp.tools.analysis import _qualify_range
from modelrisk_mcp.tools.building import _params_dict


class TestF1LiveFormula:
    def test_text_mentioning_vose_is_not_live(self) -> None:
        assert _is_live_formula('wrap as VoseOutput("NPV")') is False
        assert _is_live_formula("VosePERT(1,2,3)") is False

    def test_real_formulas_are_live(self) -> None:
        assert _is_live_formula('=VoseOutput("NPV")+B10') is True
        assert _is_live_formula("  =SUM(A1:A9)") is True

    def test_empty_and_none(self) -> None:
        assert _is_live_formula("") is False
        assert _is_live_formula(None) is False


class TestF3QualifyRange:
    def test_plain_range_gets_sheet(self) -> None:
        assert _qualify_range("Data", "B5:B369") == "'Data'!B5:B369"

    def test_cross_sheet_range_kept_verbatim(self) -> None:
        # Previously became 'Model'!Data!D5:D44 → opaque COM exception.
        assert _qualify_range("Model", "Data!D5:D44") == "Data!D5:D44"


class TestF7MergedGuard:
    def _cell(self, merged: bool, addr: str, anchor: str) -> MagicMock:
        c = MagicMock()
        c.api.MergeCells = merged
        c.api.Address = addr
        c.api.MergeArea.Cells.return_value = MagicMock(Address=anchor)
        c.api.MergeArea.Address = "$C$11:$H$11"
        return c

    def test_non_anchor_merged_cell_raises(self) -> None:
        cell = self._cell(True, "$D$11", "$C$11")
        with pytest.raises(CellReferenceError, match="merged"):
            ExcelBridge._guard_merged_cell(cell, "wb", "Model", "D11")

    def test_anchor_of_merge_is_allowed(self) -> None:
        cell = self._cell(True, "$C$11", "$C$11")
        ExcelBridge._guard_merged_cell(cell, "wb", "Model", "C11")

    def test_unmerged_cell_is_allowed(self) -> None:
        cell = self._cell(False, "$B$30", "$B$30")
        ExcelBridge._guard_merged_cell(cell, "wb", "Model", "B30")


class TestF8CatalogueOutputSize:
    def test_time_gbm_has_no_output_size(self) -> None:
        cat = load_catalogue()
        spec = cat.require("VoseTimeGBM")
        names = [p.name for p in spec.parameters]
        assert "OutputSize" not in names
        assert names[0] == "mu"  # real signature starts at mu

    def test_copula_fit_has_no_output_size(self) -> None:
        cat = load_catalogue()
        spec = cat.require("VoseCopulaMultiNormalFit")
        assert all(p.name != "OutputSize" for p in spec.parameters)


class TestF10Percentiles:
    def test_default_set_includes_p80(self) -> None:
        assert 0.80 in _DEFAULT_PERCENTILES


class TestParamsDictPositional:
    def test_all_unnamed_becomes_positional_list(self) -> None:
        assert _params_dict([{"value": 1}, {"value": 2}]) == [1, 2]

    def test_named_becomes_dict(self) -> None:
        assert _params_dict([{"name": "mu", "value": 0.05}]) == {"mu": 0.05}


class TestOffsetA1:
    def test_offsets(self) -> None:
        assert _offset_a1("H2", 0, 0) == "H2"
        assert _offset_a1("H2", 2, 0) == "H4"
        assert _offset_a1("A1", 0, 27) == "AB1"
