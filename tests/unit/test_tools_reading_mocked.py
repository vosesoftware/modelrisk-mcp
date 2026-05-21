"""MCP-wrapper tests for `tools/reading.py`.

The 14 tools in this module are thin wrappers over either
`ModelRiskBridge` or `ExcelBridge`. The bridge logic itself has its
own tests (`test_modelrisk_mocked.py`, `test_excel_bridge.py`,
`test_mrservice.py`); what this file guards is:

- Tool fn → bridge method name matches exactly.
- Arguments pass through with their correct keyword/positional shape.
- Optional args default to `None` when omitted (not to empty
  collections or sentinels, which the bridge layer treats differently).
- Return types are the right Pydantic schemas.

A regression like `bridge.list_inputs(workbook_name=workbook_name)` vs
the bridge's actual parameter `workbook` would only surface at end-user
runtime without these tests.
"""

from __future__ import annotations

from collections.abc import Iterator
from unittest.mock import MagicMock

import pytest

from modelrisk_mcp.schemas.results import (
    CorrelationMatrix,
    SensitivityEntry,
    SensitivityRanking,
    SimulationResult,
)
from modelrisk_mcp.schemas.workbook import (
    CellInfo,
    CellRef,
    DistributionCell,
    ModelRiskOutput,
    RangeInfo,
    WorkbookInfo,
    WorkbookSummary,
)
from modelrisk_mcp.tools import reading

# ----------------------------------------------------------------------
# Fixture — a fully-spec'd MagicMock for the bridge, with `.excel`
# and `.results` sub-mocks. Each tool function reads through this.
# ----------------------------------------------------------------------


@pytest.fixture
def bridge() -> Iterator[MagicMock]:
    b = MagicMock()
    # Pre-populate plausible return shapes so tools that use the return
    # value (rather than just passing it through) don't choke.
    b.excel = MagicMock()
    b.results = MagicMock()
    reading.set_bridge_for_testing(b)  # type: ignore[arg-type]
    yield b
    reading.set_bridge_for_testing(None)


# ----------------------------------------------------------------------
# Workbook navigation
# ----------------------------------------------------------------------


class TestWorkbookNavigation:
    def test_list_open_workbooks_delegates_to_excel(
        self, bridge: MagicMock
    ) -> None:
        bridge.excel.list_workbooks.return_value = [
            WorkbookInfo(name="a.xlsx", path="C:/a.xlsx", sheets=["S1"]),
        ]
        result = reading.list_open_workbooks()
        bridge.excel.list_workbooks.assert_called_once_with()
        assert result[0].name == "a.xlsx"

    def test_get_active_workbook_delegates_to_excel(
        self, bridge: MagicMock
    ) -> None:
        bridge.excel.get_active_workbook.return_value = WorkbookInfo(
            name="active.xlsx", path="", sheets=[]
        )
        result = reading.get_active_workbook()
        bridge.excel.get_active_workbook.assert_called_once_with()
        assert result.name == "active.xlsx"


# ----------------------------------------------------------------------
# Workbook summary / listings
# ----------------------------------------------------------------------


class TestWorkbookListings:
    def test_get_workbook_summary_uses_bridge_method(
        self, bridge: MagicMock
    ) -> None:
        bridge.get_workbook_summary.return_value = WorkbookSummary(
            workbook="m.xlsx",
            sheets=["S1"],
            input_count=2,
            output_count=1,
            distribution_count=3,
            formula_cell_count=5,
            numeric_cell_count=10,
            modelrisk_loaded=True,
        )
        result = reading.get_workbook_summary("m.xlsx")
        bridge.get_workbook_summary.assert_called_once_with("m.xlsx")
        assert isinstance(result, WorkbookSummary)
        assert result.input_count == 2

    def test_list_modelrisk_inputs_calls_list_inputs_not_list_modelrisk_inputs(
        self, bridge: MagicMock
    ) -> None:
        """The MCP tool is `list_modelrisk_inputs` (brand-prefixed) but
        the bridge method is `list_inputs`. Easy to fat-finger."""
        bridge.list_inputs.return_value = []
        reading.list_modelrisk_inputs("m.xlsx")
        bridge.list_inputs.assert_called_once_with("m.xlsx")

    def test_list_modelrisk_outputs_calls_list_outputs(
        self, bridge: MagicMock
    ) -> None:
        bridge.list_outputs.return_value = [
            ModelRiskOutput(
                ref=CellRef(workbook="m.xlsx", sheet="S1", cell="A1"),
                name="profit",
                formula='=VoseOutput("profit")+B1',
                current_value=100.0,
            )
        ]
        result = reading.list_modelrisk_outputs("m.xlsx")
        bridge.list_outputs.assert_called_once_with("m.xlsx")
        assert result[0].name == "profit"

    def test_list_distributions_forwards_sheet_kwarg(
        self, bridge: MagicMock
    ) -> None:
        bridge.list_distributions.return_value = []
        reading.list_distributions("m.xlsx", sheet="Inputs")
        bridge.list_distributions.assert_called_once_with(
            "m.xlsx", sheet="Inputs"
        )

    def test_list_distributions_sheet_defaults_to_none(
        self, bridge: MagicMock
    ) -> None:
        bridge.list_distributions.return_value = []
        reading.list_distributions("m.xlsx")
        # The bridge accepts `sheet=None` to mean "all sheets" — verify
        # the wrapper preserves that contract rather than passing "" or
        # omitting the kwarg entirely.
        kwargs = bridge.list_distributions.call_args.kwargs
        assert kwargs.get("sheet") is None


# ----------------------------------------------------------------------
# Cell / range reads
# ----------------------------------------------------------------------


class TestCellReads:
    def test_get_cell_delegates_to_excel(self, bridge: MagicMock) -> None:
        bridge.excel.get_cell.return_value = CellInfo(
            ref=CellRef(workbook="m.xlsx", sheet="S1", cell="B12"),
            formula="=42",
            value=42,
            cell_type="formula",
        )
        result = reading.get_cell("m.xlsx", "S1", "B12")
        bridge.excel.get_cell.assert_called_once_with("m.xlsx", "S1", "B12")
        assert result.formula == "=42"

    def test_read_range_delegates_to_excel(self, bridge: MagicMock) -> None:
        bridge.excel.read_range.return_value = RangeInfo(
            workbook="m.xlsx",
            sheet="S1",
            range_ref="A1:B2",
            values=[[1, 2], [3, 4]],
            formulas=[["", ""], ["", ""]],
        )
        result = reading.read_range("m.xlsx", "S1", "A1:B2")
        bridge.excel.read_range.assert_called_once_with(
            "m.xlsx", "S1", "A1:B2"
        )
        assert result.values == [[1, 2], [3, 4]]


# ----------------------------------------------------------------------
# Simulation results
# ----------------------------------------------------------------------


class TestSimulationResultsReads:
    def test_get_simulation_results_default_names_is_none(
        self, bridge: MagicMock
    ) -> None:
        """Default `output_names=None` must reach the bridge — passing
        `[]` instead would change behaviour (some bridges treat empty
        list as 'filter to nothing')."""
        bridge.get_simulation_results.return_value = []
        reading.get_simulation_results("m.xlsx")
        args = bridge.get_simulation_results.call_args.args
        assert args == ("m.xlsx", None)

    def test_get_simulation_results_forwards_names(
        self, bridge: MagicMock
    ) -> None:
        bridge.get_simulation_results.return_value = [
            SimulationResult(
                output_name="x", iterations=100, mean=1.0, stdev=0.5,
                variance=0.25, skewness=0, kurtosis=0, min=0, max=2,
                percentiles={0.5: 1.0},
            )
        ]
        reading.get_simulation_results("m.xlsx", ["x", "y"])
        bridge.get_simulation_results.assert_called_once_with(
            "m.xlsx", ["x", "y"]
        )

    def test_get_correlation_matrix_forwards_names(
        self, bridge: MagicMock
    ) -> None:
        bridge.get_correlation_matrix.return_value = CorrelationMatrix(
            names=["a", "b"], pearson=[[1.0, 0.5], [0.5, 1.0]], iterations=100,
        )
        reading.get_correlation_matrix("m.xlsx", ["a", "b"])
        bridge.get_correlation_matrix.assert_called_once_with(
            "m.xlsx", ["a", "b"]
        )

    def test_get_sensitivity_ranking_passes_output_name_only(
        self, bridge: MagicMock
    ) -> None:
        """The bridge.get_sensitivity_ranking signature in v0.3 takes
        only the output name (workbook is resolved inside the bridge).
        The tool wrapper currently passes workbook_name but the bridge
        ignores it — locking that shape in here."""
        bridge.get_sensitivity_ranking.return_value = SensitivityRanking(
            output_name="profit",
            entries=[
                SensitivityEntry(
                    input_name="cost", correlation=-0.6,
                    regression_coefficient=-0.5,
                )
            ],
            iterations=1000,
        )
        result = reading.get_sensitivity_ranking("m.xlsx", "profit")
        bridge.get_sensitivity_ranking.assert_called_once_with("profit")
        assert result.output_name == "profit"


# ----------------------------------------------------------------------
# Hard-coded input discovery
# ----------------------------------------------------------------------


class TestFindHardCodedInputs:
    def test_returns_plain_dicts_with_three_keys(
        self, bridge: MagicMock
    ) -> None:
        """The tool flattens CellRef → {workbook, sheet, cell} dicts so
        the MCP JSON is obvious. Verify shape, not just count."""
        bridge.find_hard_coded_inputs.return_value = [
            CellRef(workbook="m.xlsx", sheet="In", cell="A1"),
            CellRef(workbook="m.xlsx", sheet="In", cell="A2"),
        ]
        result = reading.find_hard_coded_inputs("m.xlsx")
        assert result == [
            {"workbook": "m.xlsx", "sheet": "In", "cell": "A1"},
            {"workbook": "m.xlsx", "sheet": "In", "cell": "A2"},
        ]

    def test_empty_list_passes_through(self, bridge: MagicMock) -> None:
        bridge.find_hard_coded_inputs.return_value = []
        assert reading.find_hard_coded_inputs("m.xlsx") == []


# ----------------------------------------------------------------------
# .vmrs control tools (new in v0.3.0-alpha.2)
# ----------------------------------------------------------------------


class TestVmrsPinning:
    def test_set_active_vmrs_with_path(self, bridge: MagicMock) -> None:
        result = reading.set_active_vmrs(r"C:\models\m.vmrs")
        bridge.results.set_active_vmrs.assert_called_once_with(
            r"C:\models\m.vmrs"
        )
        assert result == {"active_vmrs": r"C:\models\m.vmrs"}

    def test_set_active_vmrs_with_empty_string_clears(
        self, bridge: MagicMock
    ) -> None:
        """Empty string → caller wants to clear the pin. Must reach the
        bridge as `None`, not as `""` (different semantics inside the
        ResultsReader)."""
        result = reading.set_active_vmrs("")
        bridge.results.set_active_vmrs.assert_called_once_with(None)
        assert result == {"active_vmrs": ""}

    def test_set_active_vmrs_strips_whitespace(
        self, bridge: MagicMock
    ) -> None:
        reading.set_active_vmrs("   ")
        bridge.results.set_active_vmrs.assert_called_once_with(None)

    def test_read_vmrs_pins_then_reads(self, bridge: MagicMock) -> None:
        """read_vmrs is a convenience: set_active_vmrs + get_simulation_results
        in one call. Verify both happen in order."""
        bridge.results.get_simulation_results.return_value = []
        reading.read_vmrs(r"D:\x.vmrs", output_names=["out1"])
        # Order matters: pin BEFORE the read.
        calls = bridge.results.method_calls
        first_call = calls[0]
        second_call = calls[1]
        assert first_call[0] == "set_active_vmrs"
        assert first_call[1] == (r"D:\x.vmrs",)
        assert second_call[0] == "get_simulation_results"
        # Second arg to get_simulation_results is the output_names list.
        assert second_call[1] == (None, ["out1"])

    def test_read_vmrs_default_output_names_is_none(
        self, bridge: MagicMock
    ) -> None:
        bridge.results.get_simulation_results.return_value = []
        reading.read_vmrs(r"D:\x.vmrs")
        second_call_args = bridge.results.method_calls[1][1]
        assert second_call_args == (None, None)


# ----------------------------------------------------------------------
# Variable enumeration + raw samples (v0.3.0-alpha.4)
# ----------------------------------------------------------------------


class TestVmrsVariableEnumeration:
    def test_list_vmrs_variables_default_active(
        self, bridge: MagicMock
    ) -> None:
        bridge.list_vmrs_variables.return_value = [
            {"name": "profit", "var_id": 0, "kind": "output", "iterations": 1000},
            {"name": "demand", "var_id": 1, "kind": "input", "iterations": 1000},
        ]
        result = reading.list_vmrs_variables()
        bridge.list_vmrs_variables.assert_called_once_with(None)
        assert len(result) == 2
        assert result[0]["name"] == "profit"

    def test_list_vmrs_variables_with_workbook_name(
        self, bridge: MagicMock
    ) -> None:
        bridge.list_vmrs_variables.return_value = []
        reading.list_vmrs_variables(workbook_name="m.xlsx")
        bridge.list_vmrs_variables.assert_called_once_with("m.xlsx")


class TestGetSamples:
    def test_get_samples_passes_max_n_as_keyword(
        self, bridge: MagicMock
    ) -> None:
        """Bridge signature is `get_samples(name, workbook, *, max_n=)` —
        the wrapper must keep max_n keyword-only."""
        bridge.get_samples.return_value = [1.0, 2.0, 3.0]
        reading.get_samples("profit")
        bridge.get_samples.assert_called_once_with("profit", None, max_n=10_000)

    def test_get_samples_forwards_all_args(self, bridge: MagicMock) -> None:
        bridge.get_samples.return_value = [1.0]
        reading.get_samples("profit", max_n=500, workbook_name="m.xlsx")
        bridge.get_samples.assert_called_once_with(
            "profit", "m.xlsx", max_n=500
        )

    def test_get_samples_returns_list_of_floats(
        self, bridge: MagicMock
    ) -> None:
        bridge.get_samples.return_value = [1.5, 2.5, 3.5]
        result = reading.get_samples("profit")
        assert result == [1.5, 2.5, 3.5]


# Quieten unused-imports for types referenced via fixtures only.
_ = DistributionCell
