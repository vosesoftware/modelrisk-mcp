"""Higher-level workflow tools (spec §7.4).

These tools compose lower-level reading / building / simulation tools
into "intent-shaped" operations the LLM can call once and get a
methodology-aware result.
"""

from __future__ import annotations

import re
from importlib import resources
from pathlib import Path
from typing import Annotated, Any

import yaml
from pydantic import Field

from modelrisk_mcp.audit.engine import run_audit
from modelrisk_mcp.bridge.charts import DistributionChartResult, TornadoChartResult
from modelrisk_mcp.bridge.formulas import (
    build_distribution_formula,
    build_input_wrapper,
    build_output_wrapper,
)
from modelrisk_mcp.bridge.reports import DriversReportResult, ExecutiveReportResult
from modelrisk_mcp.errors import ModelRiskComputationError
from modelrisk_mcp.schemas.analysis import (
    BuiltInput,
    FitAndWireResult,
    ModelFromBriefResult,
    RiskModelPlan,
    WiredColumnResult,
)
from modelrisk_mcp.schemas.results import AuditReport, SimulationResult
from modelrisk_mcp.schemas.workbook import CellRef
from modelrisk_mcp.server import mcp
from modelrisk_mcp.tools.reading import get_bridge

# ----------------------------------------------------------------------
# Distribution selection guide loader
# ----------------------------------------------------------------------


def _load_distribution_guide() -> dict[str, Any]:
    text = (
        resources.files("modelrisk_mcp.data")
        .joinpath("distributions.yaml")
        .read_text(encoding="utf-8")
    )
    return yaml.safe_load(text) or {}


def _pick_scenario(description: str) -> tuple[str, list[dict[str, str]]]:
    guide = _load_distribution_guide()
    scenarios = guide.get("scenarios", {})
    haystack = (description or "").lower()
    # Skip the catch-all on first pass.
    for scenario_name, entry in scenarios.items():
        if scenario_name == "unknown":
            continue
        keywords = entry.get("keywords", []) or []
        if any(kw.lower() in haystack for kw in keywords):
            return scenario_name, entry.get("recommendations", [])
    fallback = scenarios.get("unknown", {})
    return "unknown", fallback.get("recommendations", [])


# ----------------------------------------------------------------------
# Tools
# ----------------------------------------------------------------------


@mcp.tool(
    description=(
        "ModelRisk: Propose distribution families for a list of "
        "uncertain inputs. Each input gets a ranked list of "
        "recommendations from the methodology-grounded selection guide. "
        "The tool does NOT write to Excel — it returns suggestions for "
        "the LLM to walk through with the user before committing via "
        "replace_constant_with_distribution."
    )
)
def propose_distributions_for_inputs(
    inputs: Annotated[
        list[dict[str, Any]],
        Field(
            description=(
                "Each entry: {cell_ref?, current_value?, description}. "
                "`description` is the natural-language description of "
                "the uncertain quantity (e.g. 'unit cost of widget X')."
            )
        ),
    ],
) -> dict[str, Any]:
    proposals: list[dict[str, Any]] = []
    for entry in inputs:
        description = str(entry.get("description", "") or "")
        scenario_name, recs = _pick_scenario(description)
        proposals.append(
            {
                "cell_ref": entry.get("cell_ref"),
                "current_value": entry.get("current_value"),
                "description": description,
                "scenario_matched": scenario_name,
                "recommendations": recs,
            }
        )
    return {"proposals": proposals, "count": len(proposals)}


@mcp.tool(
    description=(
        "ModelRisk: Discover candidate input cells — numeric cells "
        "referenced by formulas — and rank them by how likely they are "
        "to be uncertain model inputs (vs. constants like 12 months "
        "per year). The ranking weighs reference count and number "
        "magnitude. Pair with propose_distributions_for_inputs."
    )
)
def discover_inputs(
    workbook_name: str,
    limit: int = 25,
) -> dict[str, Any]:
    bridge = get_bridge()
    refs: list[CellRef] = bridge.find_hard_coded_inputs(workbook_name)
    # Build a small score per cell: weight by reference count (we
    # already filtered to "referenced") and by a "round number bonus"
    # for cells whose value looks like a scenario assumption.
    cells_by_ref = {
        f"{c.ref.sheet}!{c.ref.cell}": c
        for c in bridge.excel.iterate_cells(workbook_name)
    }
    scored: list[tuple[float, dict[str, Any]]] = []
    for ref in refs:
        info = cells_by_ref.get(f"{ref.sheet}!{ref.cell}")
        value = info.value if info else None
        score = 1.0
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            # "Round-ish" numbers (multiples of 10 / 100 / 1000) score
            # higher; flag-shaped values (0 / 1) are excluded from ALL
            # the round-number bonuses, not just multiple-of-10 — the
            # earlier version checked the exclusion only on the first
            # bonus, so value=0 still picked up the % 100 == 0 and
            # % 1000 == 0 bonuses because 0 % n == 0 trivially.
            if value not in (0, 1):
                if value % 10 == 0:
                    score += 0.5
                if value % 100 == 0:
                    score += 0.5
                if value % 1000 == 0:
                    score += 0.5
        scored.append(
            (
                score,
                {
                    "workbook": ref.workbook,
                    "sheet": ref.sheet,
                    "cell": ref.cell,
                    "current_value": value,
                    "score": score,
                },
            )
        )
    scored.sort(key=lambda kv: kv[0], reverse=True)
    candidates = [entry for _, entry in scored[:limit]]
    return {"candidates": candidates, "count": len(candidates)}


@mcp.tool(
    description=(
        "ModelRisk: One-call blueprint for turning a deterministic workbook "
        "into a Monte Carlo risk model. Reports what's already there (declared "
        "outputs, existing distributions), the ranked hard-coded cells that "
        "look like uncertain inputs, and an ordered, state-aware checklist of "
        "next actions (wrap outputs, fit/propose distributions, correlate, "
        "audit, simulate, interpret). Read-only — it plans, it doesn't modify. "
        "Run this first when asked to 'add uncertainty' or 'make this a risk "
        "model'."
    )
)
def plan_risk_model(workbook_name: str) -> RiskModelPlan:
    bridge = get_bridge()
    outputs = bridge.list_outputs(workbook_name)
    distributions = bridge.list_distributions(workbook_name)
    candidates = discover_inputs(workbook_name, limit=15)["candidates"]

    has_outputs = len(outputs) > 0
    has_candidates = len(candidates) > 0
    if not has_outputs and not has_candidates:
        readiness = "empty"
    elif not has_outputs:
        readiness = "needs-outputs"
    elif not has_candidates and not distributions:
        readiness = "needs-inputs"
    else:
        readiness = "ready"

    steps: list[str] = []
    if not has_outputs:
        steps.append(
            "Declare the result cell(s) you care about with wrap_with_output — "
            "nothing can be simulated until at least one output exists."
        )
    if has_candidates:
        steps.append(
            "For each uncertain input candidate, fit_and_rank_distributions "
            "(if you have data) or propose_distributions_for_inputs (from a "
            "range), then replace_constant_with_distribution."
        )
    steps.append(
        "Correlate inputs that move together: compute_correlation_matrix on "
        "historical data, then create_copula."
    )
    steps.append("Run audit_model and fix any errors before simulating.")
    steps.append("run_simulation, then read the tail with get_tail_risk.")

    return RiskModelPlan(
        workbook=workbook_name,
        output_count=len(outputs),
        outputs=[o.name for o in outputs],
        distribution_count=len(distributions),
        input_candidate_count=len(candidates),
        input_candidates=candidates,
        readiness=readiness,
        steps=steps,
    )


@mcp.tool(
    description=(
        "ModelRisk: Run the model audit against the workbook. Each "
        "rule's detector lives in modelrisk_mcp.audit.rules; the rule "
        "set is editable in data/audit_rules.yaml. Returns an "
        "AuditReport with severity-tagged findings (error/warning/"
        "info) and suggested fixes."
    )
)
def audit_model(workbook_name: str) -> AuditReport:
    bridge = get_bridge()
    return run_audit(bridge, workbook_name)


@mcp.tool(
    description=(
        "ModelRisk: One-call workbook health check. Returns everything an "
        "MCP client typically wants at the start of a session: whether "
        "Excel is reachable, whether the ModelRisk SDK is activated, "
        "the active workbook's name + sheets, counts of inputs / outputs "
        "/ distributions, whether a sibling `.vmrs` exists and when it "
        "was last modified, and the audit-log location. Use this as the "
        "first call instead of orchestrating 4-5 individual reading "
        "tools."
    )
)
def diagnose_workbook(
    workbook_name: Annotated[
        str | None,
        Field(description="Workbook name. Omit for the active workbook."),
    ] = None,
) -> dict[str, Any]:
    from modelrisk_mcp.bridge.mrservice import find_latest_vmrs

    bridge = get_bridge()
    out: dict[str, Any] = {
        "excel_connected": False,
        "modelrisk_loaded": False,
        "addin_functional": False,
        "active_workbook": None,
        "workbook_path": "",
        "sheets": [],
        "input_count": 0,
        "output_count": 0,
        "distribution_count": 0,
        "formula_cell_count": 0,
        "vmrs_path": None,
        "vmrs_exists": False,
        "vmrs_modified": None,
        "audit_log_path": str(bridge._settings.writes_log_path),
        "issues": [],
    }
    issues: list[str] = []

    # 1. Excel reachability + workbook resolution.
    #
    # Bug #30 (alpha.28): prior versions assigned
    # `active_workbook = active.name` and `workbook_path = active.path`
    # regardless of whether the caller passed an explicit
    # `workbook_name`. Result: calling `diagnose_workbook("foo.xlsx")`
    # while "bar.xlsx" was active in Excel would report
    # active_workbook="bar.xlsx" and workbook_path=<bar's path>
    # alongside foo's input/output counts — misleading by
    # construction. The downstream `.vmrs` lookup also used the wrong
    # path and would silently find bar's vmrs instead of foo's.
    #
    # Fix: when `workbook_name` is supplied, look up THAT book's path
    # from `list_workbooks` and report it. The `active_workbook` field
    # always reflects Excel's actually-active book (informational),
    # while `workbook_path` describes the diagnosed workbook.
    try:
        active = bridge.excel.get_active_workbook()
        out["excel_connected"] = True
        wb_name = workbook_name or active.name
        out["active_workbook"] = active.name
        # Default to active.path; override below if a specific
        # workbook was named.
        out["workbook_path"] = active.path
        if workbook_name and workbook_name != active.name:
            try:
                target = next(
                    (b for b in bridge.excel.list_workbooks()
                     if b.name == workbook_name),
                    None,
                )
                if target is None:
                    issues.append(
                        f"Workbook {workbook_name!r} is not currently "
                        f"open. The diagnose result will reflect the "
                        f"empty fallback values."
                    )
                else:
                    out["workbook_path"] = target.path
            except Exception:
                # If list_workbooks fails, keep the active.path
                # default; the input counts below will still
                # be sourced from the requested workbook name.
                pass
    except Exception as exc:
        issues.append(f"Excel not reachable: {exc!s}")
        out["issues"] = issues
        return out

    # 2. MRService.dll activation (results-reading) — independent of the
    #    in-Excel add-in below.
    try:
        out["modelrisk_loaded"] = bridge.is_modelrisk_loaded()
        if not out["modelrisk_loaded"]:
            # Surface the SPECIFIC activation error (e.g. "bundled key
            # rejected — your ModelRisk SDK may be too new/old; set
            # MRSERVICE_ACTIVATION_KEY") rather than a generic line, and make
            # clear this only affects READING .vmrs results — sims still run.
            detail = ""
            try:
                bridge.mrservice.ensure_ready()
            except Exception as exc:
                detail = str(exc).strip().splitlines()[0]
            issues.append(
                "MRService.dll not activated — affects READING saved .vmrs "
                "simulation results only; running simulations is unaffected. "
                + (detail or "Set MRSERVICE_ACTIVATION_KEY to your ModelRisk key.")
            )
    except Exception as exc:
        issues.append(f"MRService check failed: {exc!s}")

    # 2b. ModelRisk add-in liveness (building + simulation). Bug #38:
    #     reported separately from MRService because they fail
    #     independently — the add-in can be dead (Vose functions return
    #     #NAME?, sims can't run) while the results DLL is fine. Probe
    #     only; never mutates the add-in state from a diagnostic.
    out["addin_functional"] = False
    try:
        out["addin_functional"] = bridge.probe_addin_functional()
        if not out["addin_functional"]:
            issues.append(
                "ModelRisk add-in not detected as live — a Vose probe "
                "function didn't return a number. If typing a Vose function "
                "into a cell (e.g. =VosePoisson(5)) DOES work, the add-in is "
                "loaded and this is a detection problem worth reporting; "
                "otherwise click the ModelRisk ribbon tab (or start ModelRisk "
                "via its shortcut) to load it. run_simulation will also try to "
                "auto-activate it. Enabling ModelRisk's 'Start with Excel' "
                "setting loads it automatically each session."
            )
    except Exception as exc:
        issues.append(f"ModelRisk add-in probe failed: {exc!s}")

    # 3. Workbook content summary
    try:
        summary = bridge.get_workbook_summary(wb_name)
        out["sheets"] = summary.sheets
        out["input_count"] = summary.input_count
        out["output_count"] = summary.output_count
        out["distribution_count"] = summary.distribution_count
        out["formula_cell_count"] = summary.formula_cell_count
        if summary.output_count == 0:
            issues.append(
                "Workbook has no VoseOutput cells. run_simulation will "
                "fail until at least one output is declared."
            )
        if summary.distribution_count == 0 and summary.output_count > 0:
            issues.append(
                "Workbook has VoseOutput(s) but no Vose distribution cells. "
                "Simulation will produce constant results."
            )
    except Exception as exc:
        issues.append(f"Workbook summary failed: {exc!s}")

    # 4. Sibling .vmrs status
    if out["workbook_path"]:
        try:
            vmrs = find_latest_vmrs(out["workbook_path"])
            out["vmrs_path"] = vmrs
            if vmrs:
                out["vmrs_exists"] = True
                out["vmrs_modified"] = _format_mtime(Path(vmrs))
        except Exception:
            pass
    if not out["vmrs_exists"]:
        issues.append(
            "No sibling .vmrs file found next to the workbook. Call "
            "run_simulation to produce one, or set_active_vmrs to point "
            "at a specific file elsewhere."
        )

    out["issues"] = issues
    return out


def _fmt_num(value: float) -> str:
    """Format a number for an executive-audience markdown table.

    Switched from the prior `.3g` (which kicks into scientific
    notation past 1e4 — e.g. 63300 → '6.33e+04') to a thousands-
    separated decimal with two decimal places. Falls back to
    scientific only for extreme magnitudes (>=1e9 or below 1e-2
    in absolute value) where the decimal form is unreadable.

    Negative values keep their minus sign — this is the unsigned
    formatter for absolute readouts; `_fmt_signed` is the variant
    for deltas where a leading +/- is useful."""
    if value != value:  # NaN
        return "n/a"
    abs_v = abs(value)
    if abs_v >= 1e9 or (0 < abs_v < 1e-2):
        return f"{value:.3g}"
    return f"{value:,.2f}"


def _fmt_signed(value: float) -> str:
    """Like `_fmt_num` but always prepends +/-. Used for delta
    columns (P50-Deterministic etc.) where the sign is the headline
    information."""
    if value != value:  # NaN
        return "n/a"
    abs_v = abs(value)
    if abs_v >= 1e9 or (0 < abs_v < 1e-2):
        return f"{value:+.3g}"
    return f"{value:+,.2f}"


def _format_mtime(path: Path) -> str | None:
    from datetime import datetime

    try:
        return datetime.fromtimestamp(path.stat().st_mtime).isoformat()
    except OSError:
        return None


@mcp.tool(
    description=(
        "ModelRisk: Build a single-sheet drivers report — a sensitivity "
        "analysis presented for a decision-maker. Drops onto a new "
        "sheet: title band; auto-generated KEY FINDINGS in plain "
        "English ('The dominant driver of NPV is widget cost, r = "
        "-0.65; higher widget cost lowers NPV'); a prominent tornado "
        "chart; a driver-ranking table with correlation + |r| + "
        "approximate variance share; a HOW TO READ THIS CHART panel "
        "for stakeholders who don't know what Spearman correlation "
        "means; and tiered RECOMMENDED ACTIONS (focus / monitor / "
        "deprioritise) grouping inputs by strength. Use this when the "
        "user asks for an uncertainty-drivers report rather than "
        "the broader executive dashboard."
    )
)
def build_drivers_report(
    output_name: Annotated[
        str,
        Field(
            description=(
                "The output to analyze drivers for (e.g. 'NPV', "
                "'TotalCost'). Each call produces one sheet for one "
                "output. Call multiple times for multiple outputs."
            )
        ),
    ],
    title: Annotated[
        str | None,
        Field(
            description=(
                "Report title. Default: 'Uncertainty Drivers — <output>'."
            )
        ),
    ] = None,
    subtitle: Annotated[
        str | None,
        Field(
            description=(
                "Subtitle. Default: 'Sensitivity analysis · N "
                "iterations · <date>'."
            )
        ),
    ] = None,
    sheet_name: Annotated[
        str,
        Field(
            description=(
                "Target sheet name. Default 'Drivers_Report'. "
                "Replaced if it already exists."
            )
        ),
    ] = "Drivers_Report",
    workbook_name: Annotated[
        str | None,
        Field(description="Workbook name. Omit for the active workbook."),
    ] = None,
) -> dict[str, Any]:
    result: DriversReportResult = get_bridge().build_drivers_report(
        output_name,
        workbook=workbook_name,
        title=title,
        subtitle=subtitle,
        sheet_name=sheet_name,
    )
    return {
        "sheet_name": result.sheet_name,
        "output_name": result.output_name,
        "drivers_analyzed": result.drivers_analyzed,
        "top_driver": result.top_driver,
        "top_correlation": result.top_correlation,
        "concentration": result.concentration,
        "headline_finding": result.headline_finding,
    }


@mcp.tool(
    description=(
        "ModelRisk: Build a single-sheet executive report for a "
        "decision-maker. Drops a curated dashboard onto a new sheet "
        "with: title band, headline numbers (mean / P5 / P50 / P95 / "
        "stdev — colored by volatility), histogram + cumulative chart "
        "of the primary output, tornado of top N sensitivity drivers, "
        "a stats table for the primary plus any secondary outputs, and "
        "auto-generated risk callouts framed in plain English ('90% "
        "confident X lands between A and B', 'tail risk Y% above mean', "
        "'primary driver is Z'). Idempotent — re-running replaces the "
        "sheet. Use this when the user asks for a decision-maker-"
        "facing summary rather than raw stats."
    )
)
def build_executive_report(
    primary_output: Annotated[
        str,
        Field(
            description=(
                "The single output the report focuses on (e.g. 'NPV', "
                "'TotalCost'). Headline numbers and the histogram + "
                "tornado are about this output."
            )
        ),
    ],
    title: Annotated[
        str | None,
        Field(
            description=(
                "Report title shown in the top band. Default: "
                "'Simulation Report — <primary_output>'."
            )
        ),
    ] = None,
    subtitle: Annotated[
        str | None,
        Field(
            description=(
                "Subtitle shown beneath the title. Default: "
                "'<N> iterations · <today's date>'."
            )
        ),
    ] = None,
    secondary_outputs: Annotated[
        list[str] | None,
        Field(
            description=(
                "Additional outputs to include in the stats table. The "
                "primary output is always first; these appear below."
            )
        ),
    ] = None,
    contingency_percentile: Annotated[
        float,
        Field(
            ge=0.5,
            le=0.99,
            description=(
                "The 'high-side' percentile to highlight in the "
                "headline. Default 0.90 (P90)."
            ),
        ),
    ] = 0.90,
    top_drivers: Annotated[
        int,
        Field(
            ge=1,
            le=20,
            description="How many inputs to show in the tornado mini-chart.",
        ),
    ] = 5,
    sheet_name: Annotated[
        str,
        Field(
            description=(
                "Target sheet name. Default 'Executive_Report'. "
                "Replaced if it already exists."
            )
        ),
    ] = "Executive_Report",
    workbook_name: Annotated[
        str | None,
        Field(description="Workbook name. Omit for the active workbook."),
    ] = None,
) -> dict[str, Any]:
    result: ExecutiveReportResult = get_bridge().build_executive_report(
        primary_output,
        workbook=workbook_name,
        title=title,
        subtitle=subtitle,
        secondary_outputs=secondary_outputs,
        contingency_percentile=contingency_percentile,
        top_drivers=top_drivers,
        sheet_name=sheet_name,
    )
    return {
        "sheet_name": result.sheet_name,
        "primary_output": result.primary_output,
        "secondary_outputs": list(result.secondary_outputs),
        "chart_count": result.chart_count,
        "callout_count": result.callout_count,
        "headline_summary": result.headline_summary,
    }


@mcp.tool(
    description=(
        "ModelRisk: Render a tornado chart of input sensitivity for a "
        "single output as a new sheet in the workbook. The sheet has "
        "a sorted data table (Spearman rank correlation + regression "
        "coefficient per input) plus a native Excel BarClustered chart "
        "with the largest-magnitude input at the top. Idempotent — if "
        "a sheet with the target name already exists, it's replaced. "
        "Useful when the user wants the visualization persisted in the "
        "workbook, not just returned over MCP."
    )
)
def create_tornado_chart(
    output_name: Annotated[
        str, Field(description="VoseOutput name to analyze.")
    ],
    workbook_name: Annotated[
        str | None,
        Field(description="Workbook name. Omit for the active workbook."),
    ] = None,
    sheet_name: Annotated[
        str | None,
        Field(
            description=(
                "Target sheet name. Default: `Tornado_<output_name>` "
                "(truncated to Excel's 31-char limit)."
            )
        ),
    ] = None,
) -> dict[str, Any]:
    result: TornadoChartResult = get_bridge().create_tornado_chart(
        output_name, workbook_name, sheet_name=sheet_name,
    )
    return {
        "sheet_name": result.sheet_name,
        "chart_name": result.chart_name,
        "output_name": result.output_name,
        "input_count": result.input_count,
        "top_input": result.top_input,
        "top_correlation": result.top_correlation,
    }


def _distribution_chart_dict(result: DistributionChartResult) -> dict[str, Any]:
    return {
        "sheet_name": result.sheet_name,
        "chart_name": result.chart_name,
        "output_name": result.output_name,
        "chart_kind": result.chart_kind,
        "sample_count": result.sample_count,
        "bin_count": result.bin_count,
        "mean": result.mean,
        "p10": result.p10,
        "p50": result.p50,
        "p90": result.p90,
    }


@mcp.tool(
    description=(
        "ModelRisk: Render a histogram of one output's simulation "
        "result distribution as a new sheet in the workbook. The sheet "
        "has a binned data table (bin centre / frequency / cumulative %) "
        "plus a native Excel chart: frequency columns with the "
        "cumulative-probability curve overlaid on a secondary % axis and "
        "the central-80% (P10-P90) band highlighted — the same view as "
        "ModelRisk's Results Viewer, persisted into the workbook. "
        "Requires a completed simulation (reads samples from the active "
        ".vmrs). Idempotent — a sheet with the target name is replaced."
    )
)
def create_histogram_chart(
    output_name: Annotated[
        str, Field(description="VoseOutput name to chart.")
    ],
    workbook_name: Annotated[
        str | None,
        Field(description="Workbook name. Omit for the active workbook."),
    ] = None,
    sheet_name: Annotated[
        str | None,
        Field(
            description=(
                "Target sheet name. Default: `Histogram_<output_name>` "
                "(truncated to Excel's 31-char limit)."
            )
        ),
    ] = None,
) -> dict[str, Any]:
    result: DistributionChartResult = get_bridge().create_histogram_chart(
        output_name, workbook_name, sheet_name=sheet_name,
    )
    return _distribution_chart_dict(result)


@mcp.tool(
    description=(
        "ModelRisk: Render the ascending cumulative-probability curve "
        "(CDF) of one output's simulation result distribution as a new "
        "sheet in the workbook. The sheet has a binned data table plus a "
        "native Excel line chart of cumulative probability (0-100%) "
        "against the output value — the 'what's the chance the output is "
        "below X' view. Requires a completed simulation (reads samples "
        "from the active .vmrs). Idempotent — a sheet with the target "
        "name is replaced."
    )
)
def create_cdf_chart(
    output_name: Annotated[
        str, Field(description="VoseOutput name to chart.")
    ],
    workbook_name: Annotated[
        str | None,
        Field(description="Workbook name. Omit for the active workbook."),
    ] = None,
    sheet_name: Annotated[
        str | None,
        Field(
            description=(
                "Target sheet name. Default: `CDF_<output_name>` "
                "(truncated to Excel's 31-char limit)."
            )
        ),
    ] = None,
) -> dict[str, Any]:
    result: DistributionChartResult = get_bridge().create_cdf_chart(
        output_name, workbook_name, sheet_name=sheet_name,
    )
    return _distribution_chart_dict(result)


@mcp.tool(
    description=(
        "ModelRisk: Generate an executive-audience summary of the most "
        "recent simulation results for a workbook. Returns markdown "
        "ready to paste into a deck/report — covers deterministic vs "
        "P50 vs mean comparisons, P80 contingency, and the top "
        "sensitivity drivers."
    )
)
def generate_executive_summary(
    workbook_name: str,
    deterministic_values: Annotated[
        dict[str, float] | None,
        Field(
            description=(
                "Optional map of output name → its deterministic "
                "(unsimulated) value, so the summary can quote the "
                "uplift/contingency. If omitted, the summary skips that "
                "comparison."
            )
        ),
    ] = None,
) -> dict[str, str]:
    bridge = get_bridge()
    results = bridge.get_simulation_results()
    lines: list[str] = []
    lines.append(f"# Simulation summary — `{workbook_name}`")
    lines.append("")
    if not results:
        lines.append(
            "_No simulation results available. Run a simulation first._"
        )
        return {"markdown": "\n".join(lines)}
    lines.append(
        f"_Based on {results[0].iterations} iterations across "
        f"{len(results)} output(s)._"
    )
    lines.append("")
    lines.append("## Per-output statistics")
    lines.append("")
    lines.append("| Output | Mean | P50 | P5 | P95 | StDev |")
    lines.append("|---|---:|---:|---:|---:|---:|")
    # alpha.29: switched from `.3g` (scientific notation kicks in
    # past 1e4) to `_fmt_num` which uses thousands-separated decimal
    # for normal-range values and only falls back to scientific for
    # extreme magnitudes. Reads as "63,300" instead of "6.33e+04" —
    # corporate-summary appropriate.
    for r in results:
        p50 = r.percentiles.get(0.50, r.mean)
        p5 = r.percentiles.get(0.05, r.min)
        p95 = r.percentiles.get(0.95, r.max)
        lines.append(
            f"| {r.output_name} | {_fmt_num(r.mean)} | {_fmt_num(p50)} | "
            f"{_fmt_num(p5)} | {_fmt_num(p95)} | {_fmt_num(r.stdev)} |"
        )
    if deterministic_values:
        lines.append("")
        lines.append("## Contingency vs deterministic")
        lines.append("")
        lines.append(
            "| Output | Deterministic | P50 | P80 | P50-Det | P80-Det |"
        )
        lines.append("|---|---:|---:|---:|---:|---:|")
        for r in results:
            det = deterministic_values.get(r.output_name)
            if det is None:
                continue
            p50 = r.percentiles.get(0.50, r.mean)
            p80 = r.percentiles.get(0.80, r.percentiles.get(0.95, r.max))
            lines.append(
                f"| {r.output_name} | {_fmt_num(det)} | {_fmt_num(p50)} | "
                f"{_fmt_num(p80)} | {_fmt_signed(p50 - det)} | "
                f"{_fmt_signed(p80 - det)} |"
            )
    lines.append("")
    lines.append("## Top sensitivity drivers")
    lines.append("")
    for r in results[:3]:
        try:
            ranking = bridge.get_sensitivity_ranking(r.output_name)
        except Exception as exc:
            lines.append(
                f"- {r.output_name}: sensitivity unavailable ({exc!s})"
            )
            continue
        if not ranking.entries:
            lines.append(f"- {r.output_name}: no inputs identified")
            continue
        top = ranking.entries[:5]
        lines.append(f"### {r.output_name}")
        lines.append("")
        lines.append("| Input | Rank correlation |")
        lines.append("|---|---:|")
        for e in top:
            lines.append(f"| {e.input_name} | {e.correlation:+.3f} |")
        lines.append("")
    return {"markdown": "\n".join(lines).rstrip() + "\n"}


# ----------------------------------------------------------------------
# Atomic multi-cell build helpers (staged change-set with rollback)
# ----------------------------------------------------------------------

# The repo has per-cell restore (restore_cell) but no transaction
# primitive. A model build writes many cells; if step N fails we must
# undo steps 1..N-1 or the workbook is left half-built. _ChangeSet wraps
# safe_write_cell (mutex + audit log) and records each cell's prior
# formula so rollback can restore them in reverse order.


class _ChangeSet:
    def __init__(self, bridge: Any) -> None:
        self._bridge = bridge
        self.applied: list[tuple[CellRef, str]] = []

    def write(
        self, ref: CellRef, formula: str, *, allow_overwrite_non_vose: bool = False
    ) -> Any:
        res = self._bridge.safe_write_cell(
            ref, formula, allow_overwrite_non_vose=allow_overwrite_non_vose
        )
        self.applied.append((ref, res.previous_formula or ""))
        return res

    def rollback(self) -> None:
        for ref, prev in reversed(self.applied):
            try:
                self._bridge.excel.write_cell(
                    ref.workbook, ref.sheet, ref.cell, prev
                )
            except Exception:
                pass


_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_BLOCK_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+):\$?([A-Za-z]{1,3})\$?(\d+)$")

_DEFAULT_MARGINAL_FAMILIES = [
    "Normal", "Lognormal", "Gamma", "Weibull", "Expon",
    "Logistic", "Beta", "LogLogistic", "Pareto", "Gumbel",
]

# Copula families for the wiring flow (Multi* handle any n>=2).
_COPULA_FAMILIES_WF = [
    "CopulaMultiNormal", "CopulaMultiT", "CopulaMultiClayton",
    "CopulaMultiFrank", "CopulaMultiGumbel",
]


def _col_num(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _num_col(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _split_cell(cell: str) -> tuple[str, int]:
    m = _CELL_RE.match(cell.strip())
    if not m:
        raise ModelRiskComputationError(f"Not a single-cell reference: {cell!r}.")
    return m.group(1).upper(), int(m.group(2))


def _block_columns(data_range: str) -> list[str]:
    """Slice a rectangular block like 'A2:C500' into per-column A1 ranges
    ['A2:A500', 'B2:B500', 'C2:C500']."""
    m = _BLOCK_RE.match(data_range.strip())
    if not m:
        raise ModelRiskComputationError(
            f"data_range {data_range!r} must be a rectangular block like 'A2:C500'."
        )
    c1, r1, c2, r2 = _col_num(m.group(1)), int(m.group(2)), _col_num(m.group(3)), int(m.group(4))
    lo_c, hi_c = sorted((c1, c2))
    lo_r, hi_r = sorted((r1, r2))
    return [f"{_num_col(c)}{lo_r}:{_num_col(c)}{hi_r}" for c in range(lo_c, hi_c + 1)]


def _coerce_params(parameters: list[dict[str, Any]]) -> dict[str, Any] | list[Any]:
    """Accept [{'name','value'}...] (→ dict) or [{'value'}...] (→ positional
    list), mirroring insert_distribution's parameter convention."""
    if all("name" in p for p in parameters):
        return {p["name"]: p["value"] for p in parameters}
    return [p.get("value") for p in parameters]


@mcp.tool(
    description=(
        "ModelRisk: Fit BOTH the marginal distributions and the copula "
        "dependence from a data block, then wire the whole correlated, "
        "simulation-ready model into the workbook in one reversible pass. "
        "For each variable column it fits+ranks the best marginal "
        "(AIC/SIC/HQIC); across the columns it fits the best copula "
        "(fit_copula_to_data) and writes the correlated-U block at "
        "`copula_anchor`; each marginal is wired to its copula U so the "
        "inputs are dependent, not independent — capturing the tail "
        "co-movement a single correlation coefficient discards. Optionally "
        "runs a validating simulation. Defaults to dry_run=True (returns "
        "the exact planned formulas without writing). On any mid-build "
        "failure the whole change-set is rolled back. This is the "
        "data→model step no advisory agent can perform."
    )
)
def fit_all_data_and_wire(
    workbook: Annotated[str, Field(description="Workbook file name.")],
    sheet: Annotated[str, Field(description="Sheet holding the data and target cells.")],
    data_range: Annotated[
        str,
        Field(description="Rectangular data block, one column per variable, e.g. 'A2:C500'."),
    ],
    columns: Annotated[
        list[dict[str, str]],
        Field(
            description=(
                "One entry per data column, in column order: "
                "{'input_name': 'Demand', 'target_cell': 'F2'}."
            )
        ),
    ],
    copula_anchor: Annotated[
        str | None,
        Field(description="Top cell for the correlated-U block. Omit to skip correlation."),
    ] = None,
    criterion: Annotated[
        str, Field(description="Fit criterion: 'SIC' (default), 'AIC', or 'HQIC'.")
    ] = "SIC",
    uncertainty: Annotated[
        bool, Field(description="Fit with parameter uncertainty. Default False.")
    ] = False,
    run: Annotated[
        bool, Field(description="Run a validating simulation after wiring. Default False.")
    ] = False,
    samples: Annotated[int, Field(ge=1, le=1_000_000)] = 1000,
    seed: Annotated[int, Field()] = 1,
    dry_run: Annotated[
        bool, Field(description="Preview the planned formulas without writing. Default True.")
    ] = True,
) -> FitAndWireResult:
    crit = criterion.upper()
    if crit not in ("AIC", "SIC", "HQIC"):
        raise ModelRiskComputationError(f"Unknown criterion {criterion!r}.")
    key = crit.lower()
    bridge = get_bridge()
    unc = "TRUE" if uncertainty else "FALSE"
    col_ranges = _block_columns(data_range)
    if len(columns) != len(col_ranges):
        raise ModelRiskComputationError(
            f"{len(columns)} column specs but data_range spans {len(col_ranges)} columns."
        )
    steps: list[str] = []

    # 1. Fit each marginal.
    marginals: list[dict[str, Any]] = []
    for spec, col_rng in zip(columns, col_ranges, strict=True):
        qualified = f"'{sheet}'!{col_rng}"
        scored, _skipped, _ = bridge.fit_and_rank(
            qualified, list(_DEFAULT_MARGINAL_FAMILIES),
            workbook=workbook, uncertainty=uncertainty,
        )
        scored.sort(key=lambda d: d[key])
        best = scored[0]["family"] if scored else None
        marginals.append(
            {"name": spec["input_name"], "cell": spec["target_cell"],
             "range": qualified, "family": best}
        )
        steps.append(f"Fitted {spec['input_name']}: best marginal = {best or 'none'}")

    # 2. Fit the copula across all columns (if requested and >=2 fittable).
    copula_family: str | None = None
    n_fit = sum(1 for m in marginals if m["family"])
    if copula_anchor and n_fit >= 2:
        qualified_block = f"'{sheet}'!{data_range}"
        scored_c, _sk, _ = bridge.fit_and_rank_copulas(
            qualified_block, list(_COPULA_FAMILIES_WF),
            workbook=workbook, uncertainty=uncertainty,
        )
        scored_c.sort(key=lambda d: d[key])
        copula_family = scored_c[0]["family"] if scored_c else None
        steps.append(f"Fitted copula: best = {copula_family or 'none'}")
    elif copula_anchor:
        steps.append("Correlation skipped: fewer than 2 fittable marginals.")

    # 3. Build the formulas (copula block first, then wired marginals).
    anchor_col = anchor_row = None
    if copula_family and copula_anchor:
        anchor_col, anchor_row = _split_cell(copula_anchor)

    col_results: list[WiredColumnResult] = []
    plan: list[tuple[CellRef, str, bool]] = []  # (ref, formula, allow_overwrite)
    if copula_family and copula_anchor:
        cop_ref = CellRef(workbook=workbook, sheet=sheet, cell=copula_anchor)
        cop_formula = (
            f"=Vose{copula_family}Fit('{sheet}'!{data_range},FALSE,{unc})"
        )
        plan.append((cop_ref, cop_formula, True))

    for i, m in enumerate(marginals):
        u_ref = None
        if copula_family and anchor_col is not None and anchor_row is not None:
            u_ref = f"'{sheet}'!{anchor_col}{anchor_row + i}"
        if m["family"]:
            inner = f"Vose{m['family']}Fit({m['range']},{unc}"
            inner += f",{u_ref})" if u_ref else ")"
            formula = build_input_wrapper(m["name"], inner)
            ref = CellRef(workbook=workbook, sheet=sheet, cell=m["cell"])
            plan.append((ref, formula, True))
            col_results.append(
                WiredColumnResult(
                    input_name=m["name"], target_cell=m["cell"],
                    best_family=m["family"], formula=formula, written=False,
                )
            )
        else:
            col_results.append(
                WiredColumnResult(
                    input_name=m["name"], target_cell=m["cell"],
                    best_family=None, formula="", written=False,
                    skipped_reason="no marginal family fitted",
                )
            )

    # 4. Commit (unless dry_run), with atomic rollback.
    rolled_back = False
    simulated = False
    if not dry_run:
        cs = _ChangeSet(bridge)
        try:
            for ref, formula, allow in plan:
                cs.write(ref, formula, allow_overwrite_non_vose=allow)
            for cr in col_results:
                if cr.formula:
                    cr.written = True
            steps.append(f"Wrote {len(cs.applied)} cell(s).")
            if run:
                bridge.run_simulation(workbook=workbook, samples=samples, seed=seed)
                simulated = True
                steps.append(f"Ran validating simulation ({samples} iterations).")
        except Exception as exc:
            cs.rollback()
            rolled_back = True
            steps.append(f"Build failed ({exc!r}); rolled back all writes.")
            for cr in col_results:
                cr.written = False

    note = (
        "Correlated data→model wired and simulated."
        if simulated
        else "Preview only — pass dry_run=False to write."
        if dry_run
        else "Model wired; pass run=True to validate by simulation."
    )
    if rolled_back:
        note = "Build failed and was fully rolled back — workbook unchanged."
    return FitAndWireResult(
        workbook=workbook, sheet=sheet, columns=col_results,
        copula_family=copula_family,
        copula_anchor=copula_anchor if copula_family else None,
        dry_run=dry_run, simulated=simulated, achieved_correlation=None,
        rolled_back=rolled_back, steps=steps, note=note,
    )


@mcp.tool(
    description=(
        "ModelRisk: Turn a deterministic workbook into a simulation-ready "
        "Monte Carlo model in one atomic, reversible pass. Given the "
        "output cells to track and the uncertain inputs to add (each with "
        "a Vose distribution family + parameters you choose from the "
        "brief), it wraps the outputs with VoseOutput, replaces the input "
        "cells with VoseInput-wrapped distributions, optionally runs a "
        "validating simulation, and returns the headline percentiles. "
        "Every write goes through the audit-logged safe-write path and is "
        "tracked in a change-set: if any step fails, the ENTIRE build is "
        "rolled back so the workbook is never left half-converted. "
        "Defaults to dry_run=True. This end-to-end build+simulate is "
        "exactly what an advisory agent cannot do."
    )
)
def build_model_from_brief(
    workbook: Annotated[str, Field(description="Workbook file name.")],
    sheet: Annotated[str, Field(description="Sheet holding the cells.")],
    inputs: Annotated[
        list[dict[str, Any]],
        Field(
            description=(
                "Uncertain inputs to create, each: {'cell': 'B4', "
                "'input_name': 'Demand', 'function_name': 'VoseModPERT', "
                "'parameters': [{'value': 100}, {'value': 150}, {'value': 250}]}."
            )
        ),
    ],
    outputs: Annotated[
        list[dict[str, str]] | None,
        Field(
            description=(
                "Output cells to wrap: [{'cell': 'B12', 'output_name': 'NPV'}]. "
                "Omit if outputs are already wrapped."
            )
        ),
    ] = None,
    run: Annotated[
        bool, Field(description="Run a validating simulation after building. Default True.")
    ] = True,
    samples: Annotated[int, Field(ge=1, le=1_000_000)] = 1000,
    seed: Annotated[int, Field()] = 1,
    dry_run: Annotated[
        bool, Field(description="Preview the planned build without writing. Default True.")
    ] = True,
) -> ModelFromBriefResult:
    bridge = get_bridge()
    steps: list[str] = []
    outs = outputs or []

    # Build the planned change-set (formulas first, so a bad spec fails
    # before any write).
    plan: list[tuple[CellRef, str, bool]] = []
    built_inputs: list[BuiltInput] = []
    for spec in inputs:
        inner = build_distribution_formula(
            spec["function_name"], _coerce_params(spec["parameters"]), bridge.catalogue
        )
        formula = build_input_wrapper(spec["input_name"], inner)
        ref = CellRef(workbook=workbook, sheet=sheet, cell=spec["cell"])
        plan.append((ref, formula, True))
        built_inputs.append(
            BuiltInput(cell=spec["cell"], input_name=spec["input_name"],
                       formula=formula, source="proposed")
        )
    output_plan: list[tuple[CellRef, str]] = []
    for spec in outs:
        ref = CellRef(workbook=workbook, sheet=sheet, cell=spec["cell"])
        current = bridge.excel.get_cell(workbook, sheet, spec["cell"])
        inner = current.formula or (
            f"={current.value}" if current.value is not None else "=0"
        )
        formula = build_output_wrapper(spec["output_name"], inner)
        output_plan.append((ref, formula))
    steps.append(
        f"Planned {len(built_inputs)} input(s) and {len(outs)} output wrap(s)."
    )

    headline: dict[str, dict[str, float]] = {}
    rolled_back = False
    simulated = False
    change_set_size = 0

    if not dry_run:
        cs = _ChangeSet(bridge)
        try:
            for ref, formula in output_plan:
                cs.write(ref, formula, allow_overwrite_non_vose=True)
            for ref, formula, allow in plan:
                cs.write(ref, formula, allow_overwrite_non_vose=allow)
            change_set_size = len(cs.applied)
            steps.append(f"Wrote {change_set_size} cell(s).")
            if run:
                bridge.run_simulation(workbook=workbook, samples=samples, seed=seed)
                simulated = True
                names = [s["output_name"] for s in outs] or None
                results = bridge.get_simulation_results(workbook, names)
                for r in results:
                    headline[r.output_name] = {
                        "mean": r.mean,
                        "p10": r.percentiles.get(0.1, r.mean),
                        "p50": r.percentiles.get(0.5, r.mean),
                        "p90": r.percentiles.get(0.9, r.mean),
                    }
                steps.append(f"Ran validating simulation ({samples} iterations).")
        except Exception as exc:
            cs.rollback()
            rolled_back = True
            change_set_size = 0
            steps.append(f"Build failed ({exc!r}); rolled back all writes.")

    if rolled_back:
        note = "Build failed and was fully rolled back — workbook unchanged."
    elif dry_run:
        note = "Preview only — pass dry_run=False to build the model."
    elif simulated:
        note = "Model built and validated by simulation."
    else:
        note = "Model built; pass run=True to validate by simulation."

    return ModelFromBriefResult(
        workbook=workbook, dry_run=dry_run,
        outputs_wrapped=[s["output_name"] for s in outs],
        inputs_built=built_inputs, correlated=False, simulated=simulated,
        headline=headline, rolled_back=rolled_back,
        change_set_size=change_set_size, steps=steps, note=note,
    )


__all__ = [
    "audit_model",
    "build_drivers_report",
    "build_executive_report",
    "build_model_from_brief",
    "create_tornado_chart",
    "diagnose_workbook",
    "discover_inputs",
    "fit_all_data_and_wire",
    "generate_executive_summary",
    "propose_distributions_for_inputs",
]


# Quieten unused-import linter — types referenced via Pydantic generics
_ = SimulationResult
