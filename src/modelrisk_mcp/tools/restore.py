"""The `restore_cell` MCP tool (spec §11.9) plus the workbook-level
`restore_deterministic_state` recovery tool (bug #21).

`restore_cell` reads the writes audit log, finds entries matching a
specific cell, and rewrites the pre-write formula. Lets the user roll
back changes even when Excel's undo stack has been cleared.

`restore_deterministic_state` recalculates the workbook to clear any
VoseOutput cells that are stuck on sample values from a previous
simulation — the recovery path for the bug-#20 frozen-sample state.
"""

from __future__ import annotations

from datetime import datetime
from typing import Annotated, Any

from pydantic import Field

from modelrisk_mcp.schemas.distributions import InsertResult
from modelrisk_mcp.schemas.workbook import CellRef
from modelrisk_mcp.server import mcp
from modelrisk_mcp.tools.reading import get_bridge


@mcp.tool(
    description=(
        "ModelRisk: Restore a cell to its pre-write state from the audit "
        "log. Reads %LOCALAPPDATA%\\VoseSoftware\\modelrisk-mcp\\writes.log "
        "and rewrites the oldest captured before-formula for the cell. "
        "Pass `since` (ISO timestamp) to restrict the window."
    )
)
def restore_cell(
    workbook: str,
    sheet: str,
    cell: str,
    since: Annotated[
        str | None,
        Field(
            description=(
                "Optional ISO-8601 timestamp. Restore the oldest write "
                "captured at or after this time."
            )
        ),
    ] = None,
) -> InsertResult:
    ref = CellRef(workbook=workbook, sheet=sheet, cell=cell)
    parsed: datetime | None = None
    if since is not None:
        try:
            parsed = datetime.fromisoformat(since)
        except ValueError as exc:
            raise ValueError(
                f"`since` must be ISO-8601 (got {since!r}): {exc}"
            ) from exc
    return get_bridge().restore_cell(ref, since=parsed)


@mcp.tool(
    description=(
        "ModelRisk: Recover a workbook that's been left in a 'frozen "
        "sample' state — VoseOutput cells stuck on per-iteration sample "
        "values instead of their deterministic baseline. Triggers a "
        "full Excel recalculation (Application.CalculateFull) which "
        "re-evaluates every formula and restores the deterministic "
        "values. Use this after `run_simulation` raises a "
        "post-condition error, or whenever `list_modelrisk_outputs` "
        "shows nonsense `current_value`s that look like a single sample "
        "draw rather than the model's deterministic answer."
    )
)
def restore_deterministic_state(
    workbook_name: Annotated[
        str | None,
        Field(
            description=(
                "Workbook to recalculate. Omit for the active workbook."
            )
        ),
    ] = None,
) -> dict[str, Any]:
    return get_bridge().restore_deterministic_state(workbook_name)
