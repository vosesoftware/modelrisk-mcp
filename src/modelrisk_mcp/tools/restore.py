"""The `restore_cell` MCP tool (spec §11.9).

Reads the writes audit log, finds entries matching a specific cell,
and rewrites the pre-write formula. Lets the user roll back changes
even when Excel's undo stack has been cleared.
"""

from __future__ import annotations

from datetime import datetime
from typing import Annotated

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
