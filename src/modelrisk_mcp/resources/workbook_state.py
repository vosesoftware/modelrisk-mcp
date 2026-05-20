"""Live workbook state resources."""

from __future__ import annotations

import json

from modelrisk_mcp.server import mcp
from modelrisk_mcp.tools.reading import get_bridge


@mcp.resource(
    uri="modelrisk://workbook/current",
    name="modelrisk-workbook-current",
    description=(
        "ModelRisk: live structured summary of the active workbook — "
        "sheet names, counts of inputs/outputs/distributions/formulas/"
        "numerics. Lighter than calling list_open_workbooks + "
        "get_workbook_summary separately."
    ),
    mime_type="application/json",
)
def workbook_current_resource() -> str:
    bridge = get_bridge()
    active = bridge.excel.get_active_workbook()
    summary = bridge.get_workbook_summary(active.name)
    return summary.model_dump_json(indent=2)


@mcp.resource(
    uri="modelrisk://workbook/current/sheet/{name}",
    name="modelrisk-workbook-current-sheet",
    description=(
        "ModelRisk: live structured summary of one sheet in the active "
        "workbook. Includes the list of distribution cells on that "
        "sheet so the LLM can spot-check without listing the whole "
        "workbook."
    ),
    mime_type="application/json",
)
def workbook_sheet_resource(name: str) -> str:
    bridge = get_bridge()
    active = bridge.excel.get_active_workbook()
    if name not in active.sheets:
        return json.dumps(
            {
                "error": f"Sheet {name!r} not found in active workbook "
                f"{active.name!r}. Available: {active.sheets}",
            },
            indent=2,
        )
    distributions = bridge.list_distributions(active.name, sheet=name)
    payload = {
        "workbook": active.name,
        "sheet": name,
        "distribution_cells": [
            {
                "cell": d.ref.cell,
                "function_name": d.function_name,
                "has_input_wrapper": d.has_input_wrapper,
                "has_output_wrapper": d.has_output_wrapper,
                "formula": d.formula,
            }
            for d in distributions
        ],
    }
    return json.dumps(payload, indent=2)
