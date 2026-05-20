"""Reading tools (spec §7.1) — every entry is read-only and idempotent.

Each tool is registered against the shared FastMCP instance from
`modelrisk_mcp.server`. Descriptions start with "ModelRisk: " for brand
visibility in the MCP client UI (spec §7).

Tools obtain a `ModelRiskBridge` via `get_bridge()`. The bridge is
process-global and lazy — the first call attaches to a running Excel.
Tests inject a fake via `set_bridge_for_testing()` to avoid touching
COM.
"""

from __future__ import annotations

from typing import Annotated

from pydantic import Field

from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.schemas.results import (
    CorrelationMatrix,
    SensitivityRanking,
    SimulationResult,
)
from modelrisk_mcp.schemas.workbook import (
    CellInfo,
    DistributionCell,
    ModelRiskInput,
    ModelRiskOutput,
    RangeInfo,
    WorkbookInfo,
    WorkbookSummary,
)
from modelrisk_mcp.server import mcp

_bridge: ModelRiskBridge | None = None


def get_bridge() -> ModelRiskBridge:
    """Return the process-global bridge, lazily attaching to Excel on
    the first call."""
    global _bridge
    if _bridge is None:
        _bridge = ModelRiskBridge(ExcelBridge())
    return _bridge


def set_bridge_for_testing(bridge: ModelRiskBridge | None) -> None:
    """Inject a fake bridge (typically a Mock or a ModelRiskBridge over
    a FakeExcelBridge). Pass None to reset to lazy production behaviour."""
    global _bridge
    _bridge = bridge


# ----------------------------------------------------------------------
# §7.1 tools
# ----------------------------------------------------------------------


@mcp.tool(description="ModelRisk: List all Excel workbooks currently open.")
def list_open_workbooks() -> list[WorkbookInfo]:
    return get_bridge().excel.list_workbooks()


@mcp.tool(description="ModelRisk: Get the name and path of the active workbook.")
def get_active_workbook() -> WorkbookInfo:
    return get_bridge().excel.get_active_workbook()


@mcp.tool(
    description=(
        "ModelRisk: Aggregated summary of a workbook — sheet names plus "
        "counts of VoseInput, VoseOutput, distribution, formula, and "
        "numeric cells. One-shot alternative to running the individual "
        "list tools."
    )
)
def get_workbook_summary(
    workbook_name: Annotated[str, Field(description="Workbook file name (e.g. 'model.xlsx').")],
) -> WorkbookSummary:
    return get_bridge().get_workbook_summary(workbook_name)


@mcp.tool(
    description=(
        "ModelRisk: List every cell wrapped with VoseInput() in the workbook. "
        "Returns each cell's reference, the input name, the full formula, "
        "and the current calculated value."
    )
)
def list_modelrisk_inputs(
    workbook_name: str,
) -> list[ModelRiskInput]:
    return get_bridge().list_inputs(workbook_name)


@mcp.tool(
    description=(
        "ModelRisk: List every cell wrapped with VoseOutput() in the workbook."
    )
)
def list_modelrisk_outputs(
    workbook_name: str,
) -> list[ModelRiskOutput]:
    return get_bridge().list_outputs(workbook_name)


@mcp.tool(
    description=(
        "ModelRisk: List every cell containing a Vose distribution / aggregate "
        "/ copula / time-series / fitting function. Includes flags for "
        "whether each cell is also wrapped with VoseInput or VoseOutput. "
        "Optional `sheet` restricts the scan to one sheet."
    )
)
def list_distributions(
    workbook_name: str,
    sheet: str | None = None,
) -> list[DistributionCell]:
    return get_bridge().list_distributions(workbook_name, sheet=sheet)


@mcp.tool(
    description=(
        "ModelRisk: Read a single cell's formula, value, and number format."
    )
)
def get_cell(
    workbook_name: str,
    sheet: str,
    cell: Annotated[str, Field(description="A1-style cell reference like 'B12'.")],
) -> CellInfo:
    return get_bridge().excel.get_cell(workbook_name, sheet, cell)


@mcp.tool(
    description=(
        "ModelRisk: Read a contiguous range as a 2D array of values and "
        "formulas. Use ranges like 'A1:C10'."
    )
)
def read_range(
    workbook_name: str,
    sheet: str,
    range_ref: Annotated[str, Field(description="A1-style range like 'A1:C10'.")],
) -> RangeInfo:
    return get_bridge().excel.read_range(workbook_name, sheet, range_ref)


@mcp.tool(
    description=(
        "ModelRisk: Read simulation result statistics for the workbook's "
        "outputs. Returns per-output mean, stdev, variance, skewness, "
        "kurtosis, min, max, and P5/P10/P25/P50/P75/P90/P95 percentiles. "
        "Requires a simulation to have been run; pass output_names to "
        "filter."
    )
)
def get_simulation_results(
    workbook_name: str,
    output_names: list[str] | None = None,
) -> list[SimulationResult]:
    return get_bridge().get_simulation_results(workbook_name, output_names)


@mcp.tool(
    description=(
        "ModelRisk: Pearson and Spearman rank correlation between the named "
        "simulation inputs and outputs. Computed from the per-iteration "
        "samples ModelRisk records. Pass a name list to restrict; otherwise "
        "all variables are included."
    )
)
def get_correlation_matrix(
    workbook_name: str,
    name_list: list[str] | None = None,
) -> CorrelationMatrix:
    return get_bridge().get_correlation_matrix(workbook_name, name_list)


@mcp.tool(
    description=(
        "ModelRisk: Tornado / sensitivity ranking for a single output. "
        "Returns each input ranked by its Spearman rank correlation with "
        "the output, plus the standardised regression coefficient."
    )
)
def get_sensitivity_ranking(
    workbook_name: str,
    output_name: str,
) -> SensitivityRanking:
    _ = workbook_name
    return get_bridge().get_sensitivity_ranking(output_name)


@mcp.tool(
    description=(
        "ModelRisk: Heuristic discovery of cells that look like "
        "deterministic numeric inputs — i.e. a plain number, referenced "
        "by at least one formula. These are candidates for replacing "
        "with a distribution + VoseInput wrapper."
    )
)
def find_hard_coded_inputs(
    workbook_name: str,
) -> list[dict[str, str]]:
    refs = get_bridge().find_hard_coded_inputs(workbook_name)
    # Return as plain dicts so the MCP JSON serialization is obvious.
    return [
        {"workbook": r.workbook, "sheet": r.sheet, "cell": r.cell} for r in refs
    ]


@mcp.tool(
    description=(
        "ModelRisk: Pin a specific `.vmrs` file as the source of simulation "
        "results. Pass the absolute path of the file; subsequent calls to "
        "get_simulation_results / get_correlation_matrix / "
        "get_sensitivity_ranking will read from it instead of trying to "
        "locate a sibling file next to the workbook. Pass an empty string "
        "to clear the override."
    )
)
def set_active_vmrs(
    path: Annotated[
        str,
        Field(description="Absolute path to a .vmrs file, or '' to clear."),
    ],
) -> dict[str, str]:
    p = path.strip() or None
    get_bridge().results.set_active_vmrs(p)
    return {"active_vmrs": p or ""}


@mcp.tool(
    description=(
        "ModelRisk: Read simulation results directly from a `.vmrs` file. "
        "Convenience wrapper for `set_active_vmrs` + `get_simulation_results` "
        "that doesn't need an open workbook. Pass `output_names` to filter; "
        "leave empty to attempt enumeration of all known outputs."
    )
)
def read_vmrs(
    path: Annotated[
        str, Field(description="Absolute path to a .vmrs file.")
    ],
    output_names: list[str] | None = None,
) -> list[SimulationResult]:
    reader = get_bridge().results
    reader.set_active_vmrs(path)
    return reader.get_simulation_results(None, output_names)


__all__ = [
    "find_hard_coded_inputs",
    "get_active_workbook",
    "get_bridge",
    "get_cell",
    "get_correlation_matrix",
    "get_sensitivity_ranking",
    "get_simulation_results",
    "get_workbook_summary",
    "list_distributions",
    "list_modelrisk_inputs",
    "list_modelrisk_outputs",
    "list_open_workbooks",
    "read_range",
    "read_vmrs",
    "set_active_vmrs",
    "set_bridge_for_testing",
]
