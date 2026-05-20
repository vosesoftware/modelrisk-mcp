"""Higher-level workflow tools (spec §7.4).

These tools compose lower-level reading / building / simulation tools
into "intent-shaped" operations the LLM can call once and get a
methodology-aware result.
"""

from __future__ import annotations

from importlib import resources
from typing import Annotated, Any

import yaml
from pydantic import Field

from modelrisk_mcp.audit.engine import run_audit
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
        if isinstance(value, (int, float)):
            # "Round-ish" numbers (multiples of 10/100/1000) score higher;
            # exact 0/1 score lower (likely flags).
            if value not in (0, 1) and value % 10 == 0:
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
    "discover_inputs",
    "generate_executive_summary",
    "propose_distributions_for_inputs",
]


# Quieten unused-import linter — types referenced via Pydantic generics
_ = SimulationResult
