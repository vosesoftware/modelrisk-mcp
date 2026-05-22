"""Higher-level workflow tools (spec §7.4).

These tools compose lower-level reading / building / simulation tools
into "intent-shaped" operations the LLM can call once and get a
methodology-aware result.
"""

from __future__ import annotations

from importlib import resources
from pathlib import Path
from typing import Annotated, Any

import yaml
from pydantic import Field

from modelrisk_mcp.audit.engine import run_audit
from modelrisk_mcp.bridge.charts import TornadoChartResult
from modelrisk_mcp.bridge.reports import ExecutiveReportResult
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
) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for entry in inputs:
        description = str(entry.get("description", "") or "")
        scenario_name, recs = _pick_scenario(description)
        out.append(
            {
                "cell_ref": entry.get("cell_ref"),
                "current_value": entry.get("current_value"),
                "description": description,
                "scenario_matched": scenario_name,
                "recommendations": recs,
            }
        )
    return out


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
) -> list[dict[str, Any]]:
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
    return [entry for _, entry in scored[:limit]]


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

    # 1. Excel reachability + active workbook
    try:
        active = bridge.excel.get_active_workbook()
        out["excel_connected"] = True
        wb_name = workbook_name or active.name
        out["active_workbook"] = active.name
        out["workbook_path"] = active.path
    except Exception as exc:
        issues.append(f"Excel not reachable: {exc!s}")
        out["issues"] = issues
        return out

    # 2. MRService.dll activation
    try:
        out["modelrisk_loaded"] = bridge.is_modelrisk_loaded()
        if not out["modelrisk_loaded"]:
            issues.append(
                "MRService.dll not activated. Set MRSERVICE_ACTIVATION_KEY "
                "or rely on the bundled key."
            )
    except Exception as exc:
        issues.append(f"MRService check failed: {exc!s}")

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


def _format_mtime(path: Path) -> str | None:
    from datetime import datetime

    try:
        return datetime.fromtimestamp(path.stat().st_mtime).isoformat()
    except OSError:
        return None


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
    for r in results:
        p50 = r.percentiles.get(0.50, r.mean)
        p5 = r.percentiles.get(0.05, r.min)
        p95 = r.percentiles.get(0.95, r.max)
        lines.append(
            f"| {r.output_name} | {r.mean:.3g} | {p50:.3g} | "
            f"{p5:.3g} | {p95:.3g} | {r.stdev:.3g} |"
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
                f"| {r.output_name} | {det:.3g} | {p50:.3g} | "
                f"{p80:.3g} | {p50 - det:+.3g} | {p80 - det:+.3g} |"
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


__all__ = [
    "audit_model",
    "build_executive_report",
    "create_tornado_chart",
    "diagnose_workbook",
    "discover_inputs",
    "generate_executive_summary",
    "propose_distributions_for_inputs",
]


# Quieten unused-import linter — types referenced via Pydantic generics
_ = SimulationResult
