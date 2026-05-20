"""Simulation tools (§7.4) — drive a Monte Carlo run from the MCP client.

Implementation strategy:

The ATL's IModelRiskSimulation::StartSimulation and
IModelRiskSimulationResults::SaveResultsToFile are both reachable from
outside Excel via plain `Application.Run` against the XLL command
surface — no in-process ATL Dispatch required. See
`bridge/simulation.py` for the full protocol reference.

A single tool, `run_simulation`, hides the two underlying Application.Run
calls and auto-pins the produced .vmrs so the caller's next
get_simulation_results call finds it.
"""

from __future__ import annotations

from typing import Annotated

from pydantic import BaseModel, Field

from modelrisk_mcp.server import mcp
from modelrisk_mcp.tools.reading import get_bridge


class RunSimulationResult(BaseModel):
    """Compact JSON shape returned by the run_simulation tool."""

    workbook_name: str
    vmrs_path: str
    iterations: int
    samples: int
    seed: int
    next_step: str = Field(
        description=(
            "Suggested follow-up call for the MCP client — typically "
            "`get_simulation_results` to pull the per-output statistics."
        )
    )


@mcp.tool(
    description=(
        "ModelRisk: Run a Monte Carlo simulation on the active (or named) "
        "workbook and save the results to a `.vmrs` file. Defaults to "
        "1000 iterations with a fixed seed for reproducibility, and "
        "saves the .vmrs next to the workbook as `<book>.vmrs`. The "
        "simulation is run via the same XLL commands ModelRisk's own "
        "ribbon uses (VoseStartSimulCustom12 + VoseGetDataSZ12 with the "
        "SaveResultsToFile session), so behaviour matches what you'd "
        "see clicking 'Simulate' manually. Blocks until the simulation "
        "completes. After this returns, call get_simulation_results — "
        "the produced .vmrs is automatically pinned as the active "
        "results source."
    )
)
def run_simulation(
    workbook_name: Annotated[
        str | None,
        Field(
            description=(
                "Workbook file name (e.g. 'model.xlsx'). Omit for the "
                "active workbook."
            )
        ),
    ] = None,
    samples: Annotated[
        int, Field(ge=1, le=1_000_000, description="Iteration count.")
    ] = 1000,
    seed: Annotated[
        int,
        Field(description="Random seed for reproducibility (fixed seed)."),
    ] = 1,
    save_to: Annotated[
        str | None,
        Field(
            description=(
                "Absolute path to write the .vmrs. Default: next to the "
                "workbook as `<book_stem>.vmrs`. For OneDrive-hosted "
                "workbooks (where path resolution can fail) the default "
                "falls back to the user's Desktop folder."
            )
        ),
    ] = None,
) -> RunSimulationResult:
    bridge = get_bridge()
    result = bridge.run_simulation(
        workbook=workbook_name,
        samples=samples,
        seed=seed,
        save_to=save_to,
    )
    return RunSimulationResult(
        workbook_name=result.workbook_name,
        vmrs_path=result.vmrs_path,
        iterations=result.iterations,
        samples=samples,
        seed=seed,
        next_step=(
            "Call get_simulation_results to read per-output statistics. "
            "The produced .vmrs is already pinned as the active source."
        ),
    )


__all__ = ["RunSimulationResult", "run_simulation"]
