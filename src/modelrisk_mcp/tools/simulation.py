"""Simulation control tools (spec §7.3).

The COM dance is 3-step (Settings → StartSimulation → Results). The
tools here cover the first two steps; results are read via the §7.1
tools (`get_simulation_results`, `get_correlation_matrix`,
`get_sensitivity_ranking`) which were wired in Phase 2.

`stop_simulation` and `get_simulation_status` ship with the documented
limitations from spec §7.3 — both endpoints are not yet exposed on
ModelRisk's COM surface (tracked in §14 item 1). The tools are still
*registered* so the LLM can describe the limit accurately when asked.
"""

from __future__ import annotations

from typing import Any

from modelrisk_mcp.errors import ModelRiskNotLoadedError
from modelrisk_mcp.schemas.results import (
    SimulationRunResponse,
    SimulationSettingsRequest,
    SimulationSettingsResponse,
    SimulationStatus,
)
from modelrisk_mcp.server import mcp
from modelrisk_mcp.tools.reading import get_bridge


def _ensure_modelrisk_or_raise() -> None:
    """Auto-activate ModelRisk before any simulation-control call.

    Tries to enable the add-in transparently. If it still isn't
    reachable, raises a typed error with a diagnostic the LLM can
    surface to the user."""
    bridge = get_bridge()
    diag = bridge.ensure_modelrisk_active()
    if diag["modelrisk_dispatchable"]:
        return
    com_seen = diag.get("com_addins_seen") or []
    xll_seen = diag.get("excel_addins_seen") or []
    raise ModelRiskNotLoadedError(
        "ModelRisk's COM surface is unreachable. Auto-activation didn't "
        "find a ModelRisk-named add-in to enable. COM add-ins seen: "
        f"{com_seen!r}. Excel add-ins seen: {xll_seen!r}. Confirm "
        "ModelRisk is installed and that Excel's bitness (32/64) matches "
        "the installed ModelRiskAtl.dll."
    )


@mcp.tool(
    description=(
        "ModelRisk: Write simulation settings onto "
        "ModelRiskSimulationSettings before a run. Every field is "
        "optional — only the ones explicitly passed are written. Pass "
        "`samples` for iteration count, `seed` (with `use_fixed_seed=True`) "
        "for a reproducible run, `simulations` for the number of separate "
        "simulation runs to perform back-to-back, and the various UX flags "
        "(`hide_progress_window`, `refresh_excel`, etc.) to tune Excel's "
        "behaviour during the run."
    )
)
def set_simulation_settings(
    samples: int | None = None,
    simulations: int | None = None,
    use_fixed_seed: bool | None = None,
    seed: float | None = None,
    multiple_seed_type: int | None = None,
    hide_progress_window: bool | None = None,
    refresh_excel: bool | None = None,
    refresh_rate: int | None = None,
    stop_on_output_error: bool | None = None,
) -> SimulationSettingsResponse:
    _ensure_modelrisk_or_raise()
    request = SimulationSettingsRequest(
        samples=samples,
        simulations=simulations,
        use_fixed_seed=use_fixed_seed,
        seed=seed,
        multiple_seed_type=multiple_seed_type,
        hide_progress_window=hide_progress_window,
        refresh_excel=refresh_excel,
        refresh_rate=refresh_rate,
        stop_on_output_error=stop_on_output_error,
    )
    return get_bridge().simulation.apply_settings(request)


@mcp.tool(
    description=(
        "ModelRisk: Run a Monte Carlo simulation. Optionally sets "
        "`samples` and `seed` first (a convenience over calling "
        "set_simulation_settings separately), then triggers ModelRisk's "
        "StartSimulation(). Blocks until the run completes. Read results "
        "via get_simulation_results."
    )
)
def run_simulation(
    iterations: int | None = None,
    seed: float | None = None,
) -> SimulationRunResponse:
    _ensure_modelrisk_or_raise()
    return get_bridge().simulation.run(iterations=iterations, seed=seed)


@mcp.tool(
    description=(
        "ModelRisk: Make sure the ModelRisk add-in is loaded inside the "
        "running Excel session. Scans Excel's COM and classic add-in "
        "collections, enables anything named ModelRisk or Vose, then "
        "retries the COM Dispatch. Returns a diagnostic showing which "
        "add-ins were enabled and whether the surface is now "
        "reachable. Most simulation tools call this transparently — "
        "use it directly only to debug 'COM unreachable' errors."
    )
)
def ensure_modelrisk_active() -> dict[str, Any]:
    return get_bridge().ensure_modelrisk_active()


@mcp.tool(
    description=(
        "ModelRisk: Cancel a running simulation. Currently raises "
        "SimulationNotAvailableError — ModelRisk's COM surface doesn't "
        "yet expose StopSimulation. In current ModelRisk versions, press "
        "Cancel in the ModelRisk progress dialog instead."
    )
)
def stop_simulation() -> dict[str, str]:
    get_bridge().simulation.stop()  # raises SimulationNotAvailableError
    return {}  # pragma: no cover — unreachable until the COM surface lands


@mcp.tool(
    description=(
        "ModelRisk: Report whether a simulation is currently running. "
        "Coarse-grained while ModelRisk doesn't expose a SimulationStatus "
        "property — reflects in-process state only. Returns 'running' "
        "while run_simulation is in flight in this server, 'idle' "
        "otherwise."
    )
)
def get_simulation_status() -> SimulationStatus:
    return get_bridge().simulation.status()
