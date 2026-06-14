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

from typing import Annotated, Any

from pydantic import Field

from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.schemas.results import (
    CorrelationMatrix,
    SensitivityRanking,
)
from modelrisk_mcp.schemas.workbook import (
    CellInfo,
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
def list_open_workbooks() -> dict[str, Any]:
    # Envelope (vs bare `list[WorkbookInfo]`) avoids FastMCP's
    # list-element expansion which serialises each element as its own
    # MCP content block — see the module docstring for the full
    # backstory (alpha.17 sweep).
    workbooks = get_bridge().excel.list_workbooks()
    return {"workbooks": workbooks, "count": len(workbooks)}


@mcp.tool(description="ModelRisk: Get the name and path of the active workbook.")
def get_active_workbook() -> WorkbookInfo:
    return get_bridge().excel.get_active_workbook()


@mcp.tool(
    description=(
        "ModelRisk: Open a workbook (.xlsx/.xlsm) from disk in the running "
        "Excel so the other tools can act on it. Pass an absolute file path. "
        "If a workbook with the same file name is already open, returns that "
        "one (Excel won't open two with the same name). Requires Excel running."
    )
)
def open_workbook(
    path: Annotated[
        str, Field(description="Absolute path to the workbook file, e.g. r'C:\\models\\risk.xlsx'.")
    ],
) -> WorkbookInfo:
    return get_bridge().excel.open_workbook(path)


@mcp.tool(
    description=(
        "ModelRisk: Close an open workbook by file name. By DEFAULT unsaved "
        "changes are DISCARDED (save=False) — pass save=True to write them "
        "first. Returns the closed name and the workbooks still open. Raises if "
        "the named workbook isn't open."
    )
)
def close_workbook(
    workbook_name: Annotated[
        str, Field(description="File name of an open workbook, e.g. 'risk.xlsx'.")
    ],
    save: Annotated[
        bool,
        Field(description="Save before closing. False (default) discards unsaved changes."),
    ] = False,
) -> dict[str, Any]:
    return get_bridge().excel.close_workbook(workbook_name, save)


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
) -> dict[str, Any]:
    inputs = get_bridge().list_inputs(workbook_name)
    return {"inputs": inputs, "count": len(inputs)}


@mcp.tool(
    description=(
        "ModelRisk: List every cell wrapped with VoseOutput() in the workbook."
    )
)
def list_modelrisk_outputs(
    workbook_name: str,
) -> dict[str, Any]:
    outputs = get_bridge().list_outputs(workbook_name)
    return {"outputs": outputs, "count": len(outputs)}


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
) -> dict[str, Any]:
    distributions = get_bridge().list_distributions(workbook_name, sheet=sheet)
    return {"distributions": distributions, "count": len(distributions)}


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
) -> dict[str, Any]:
    results = get_bridge().get_simulation_results(workbook_name, output_names)
    return {"results": results, "count": len(results)}


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
) -> dict[str, Any]:
    refs = get_bridge().find_hard_coded_inputs(workbook_name)
    candidates = [
        {"workbook": r.workbook, "sheet": r.sheet, "cell": r.cell} for r in refs
    ]
    return {"candidates": candidates, "count": len(candidates)}


@mcp.tool(
    description=(
        "ModelRisk: List every variable in the active simulation results "
        "(`.vmrs`) that's also declared as a VoseInput or VoseOutput in "
        "the workbook. Each entry: `{name, kind, var_id, iterations}`. "
        "Use this before `get_samples` or `read_vmrs` when you don't "
        "already know which outputs / inputs exist in the file."
    )
)
def list_vmrs_variables(
    workbook_name: Annotated[
        str | None,
        Field(description="Workbook name. Omit for the active workbook."),
    ] = None,
) -> dict[str, Any]:
    variables = get_bridge().list_vmrs_variables(workbook_name)
    return {"variables": variables, "count": len(variables)}


@mcp.tool(
    description=(
        "ModelRisk: Return raw per-iteration sample values for a single "
        "output or input. Useful for custom histograms, arbitrary "
        "percentiles, downstream analysis. Caps at 10 000 samples by "
        "default to keep the MCP response small; raise `max_n` if you "
        "need more (a 100 000-iteration sim returns ~100 KB of JSON at "
        "max_n=100000)."
    )
)
def get_samples(
    output_name: Annotated[str, Field(description="VoseInput or VoseOutput name.")],
    max_n: Annotated[
        int,
        Field(
            ge=1,
            le=1_000_000,
            description="Maximum samples to return (default 10 000).",
        ),
    ] = 10_000,
    workbook_name: Annotated[
        str | None,
        Field(description="Workbook name. Omit for the active workbook."),
    ] = None,
) -> dict[str, Any]:
    # NOTE: wrapping in a dict (vs returning the bare list) is deliberate.
    # FastMCP serializes a bare `list[float]` return type as one MCP text
    # content-block per element, which for 10 000 samples produced
    # ~30x the JSON size and forced every consumer to unwrap
    # `[{"type":"text","text":"<float>"}, ...]`. Returning a dict gives
    # FastMCP one structured payload to JSON-encode in a single block.
    samples = get_bridge().get_samples(output_name, workbook_name, max_n=max_n)
    return {
        "output_name": output_name,
        "sample_count": len(samples),
        "samples": list(samples),
    }


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
) -> dict[str, Any]:
    reader = get_bridge().results
    reader.set_active_vmrs(path)
    results = reader.get_simulation_results(None, output_names)
    return {"results": results, "count": len(results)}


__all__ = [
    "close_workbook",
    "find_hard_coded_inputs",
    "get_active_workbook",
    "get_bridge",
    "get_cell",
    "get_correlation_matrix",
    "get_samples",
    "get_sensitivity_ranking",
    "get_simulation_results",
    "get_workbook_summary",
    "list_distributions",
    "list_modelrisk_inputs",
    "list_modelrisk_outputs",
    "list_open_workbooks",
    "list_vmrs_variables",
    "open_workbook",
    "read_range",
    "read_vmrs",
    "set_active_vmrs",
    "set_bridge_for_testing",
]
