"""Phase 0 spike: enumerate the ModelRisk COM surface.

ModelRisk registers four COM coclasses (see spec §8.0), each instantiable
by ProgID via win32com.client.Dispatch. This script probes each one,
reports which methods/properties exist, and writes the findings to
docs/com-surface.md.

Run from a Windows machine that has Excel + ModelRisk installed:

    uv run python scripts/spike_com_surface.py
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path

OUT_PATH = Path(__file__).resolve().parent.parent / "docs" / "com-surface.md"


@dataclass(frozen=True)
class CoclassProbe:
    label: str
    progid: str
    expected_methods: tuple[str, ...] = ()
    expected_properties: tuple[str, ...] = ()


PROBES: tuple[CoclassProbe, ...] = (
    CoclassProbe(
        label="ModelRisk (distribution functions)",
        progid="ModelRisk",
        expected_methods=("Normal", "ModPERT", "Bernoulli", "AggregateMC"),
    ),
    CoclassProbe(
        label="ModelRisk.ModelRiskSimulation",
        progid="ModelRisk.ModelRiskSimulation",
        expected_methods=("StartSimulation",),
    ),
    CoclassProbe(
        label="ModelRisk.ModelRiskSimulationSettings",
        progid="ModelRisk.ModelRiskSimulationSettings",
        expected_properties=(
            "Samples",
            "Simulations",
            "UseFixedSeed",
            "Seed",
            "MultipleSeedType",
            "RefreshExcel",
            "RefreshRate",
            "StopOnOutputError",
            "ShowResultsAtEnd",
            "HideProgressWindow",
        ),
    ),
    CoclassProbe(
        label="ModelRisk.ModelRiskSimulationResults",
        progid="ModelRisk.ModelRiskSimulationResults",
        expected_methods=(
            "SimVariables",
            "SimInputs",
            "SimOutputs",
            "GetVariablesCount",
            "GetVarSamples",
            "GetVarName",
            "AddChart",
            "ReportAllVariables",
            "SaveResultsToFile",
            "LoadResultsFromFile",
        ),
    ),
)

ISIMVARIABLE_EXPECTED_METHODS = (
    "GetName",
    "GetRangeName",
    "GetMean",
    "GetVariance",
    "GetStDev",
    "GetSkewness",
    "GetKurtosis",
    "GetCofV",
    "GetPercentile",
    "GetProbability",
    "GetSamples",
    # Pending dev confirmation per spec §8.4:
    "GetCorrelation",
    "GetTornado",
)


@dataclass
class ProbeResult:
    label: str
    progid: str
    dispatched: bool
    error: str | None = None
    method_results: dict[str, bool] = field(default_factory=dict)
    property_results: dict[str, bool] = field(default_factory=dict)


def _probe_endpoint(obj: object, name: str) -> bool:
    return hasattr(obj, name)


def _probe_coclass(com_module: object, probe: CoclassProbe) -> ProbeResult:
    dispatch = com_module.Dispatch  # type: ignore[attr-defined]
    try:
        obj = dispatch(probe.progid)
    except Exception as exc:
        return ProbeResult(probe.label, probe.progid, dispatched=False, error=str(exc))

    result = ProbeResult(probe.label, probe.progid, dispatched=True)
    for m in probe.expected_methods:
        result.method_results[m] = _probe_endpoint(obj, m)
    for p in probe.expected_properties:
        result.property_results[p] = _probe_endpoint(obj, p)
    return result


def _probe_sim_variable(com_module: object) -> dict[str, bool] | str:
    """Try to obtain one ISimVariable and probe its accessor surface.

    Best-effort: requires that a simulation has already been run (or that
    SimOutputs/SimInputs contains at least one variable). If the collection
    is empty, returns a diagnostic string instead.
    """
    dispatch = com_module.Dispatch  # type: ignore[attr-defined]
    try:
        results = dispatch("ModelRisk.ModelRiskSimulationResults")
    except Exception as exc:
        return f"could not Dispatch ModelRiskSimulationResults: {exc}"

    for collection_name in ("SimOutputs", "SimInputs"):
        try:
            collection = getattr(results, collection_name)()
            count = int(collection.Count)
        except Exception:
            continue
        if count >= 1:
            try:
                var = collection.Item(1)
            except Exception as exc:
                return f"{collection_name}.Item(1) failed: {exc}"
            return {name: _probe_endpoint(var, name) for name in ISIMVARIABLE_EXPECTED_METHODS}

    return (
        "no ISimVariable could be obtained (SimOutputs and SimInputs are both empty). "
        "Run a simulation in Excel first, then re-run this spike script."
    )


def _format_endpoint_table(title: str, results: dict[str, bool]) -> str:
    if not results:
        return ""
    lines = [f"### {title}", "", "| Endpoint | Available |", "| --- | --- |"]
    for name, ok in results.items():
        lines.append(f"| `{name}` | {'YES' if ok else 'NO'} |")
    lines.append("")
    return "\n".join(lines)


def _format_probe(result: ProbeResult) -> str:
    sections = [f"## {result.label}", "", f"ProgID: `{result.progid}`", ""]
    if not result.dispatched:
        sections.append(f"**Could not Dispatch:** {result.error}")
        return "\n".join(sections)
    sections.append("Dispatch: **OK**")
    sections.append("")
    if result.method_results:
        sections.append(_format_endpoint_table("Methods", result.method_results))
    if result.property_results:
        sections.append(_format_endpoint_table("Properties", result.property_results))
    return "\n".join(sections)


def main() -> int:
    try:
        import win32com.client as com  # type: ignore[import-not-found]
    except ImportError:
        print("pywin32 is not installed; run `uv sync` first.", file=sys.stderr)
        return 2

    out: list[str] = [
        "# ModelRisk COM surface",
        "",
        f"Probed at {datetime.now(UTC).isoformat()}.",
        "",
        "This file is the canonical record of which ModelRisk COM endpoints "
        "are exposed by the installed ModelRisk build. It is overwritten by "
        "`scripts/spike_com_surface.py`.",
        "",
    ]

    overall_ok = True
    for probe in PROBES:
        result = _probe_coclass(com, probe)
        if not result.dispatched:
            overall_ok = False
        out.append(_format_probe(result))
        out.append("")

    out.append("## ISimVariable accessor surface")
    out.append("")
    out.append(
        "Probed by Dispatching `ModelRisk.ModelRiskSimulationResults`, "
        "calling `SimOutputs()` (or `SimInputs()`) and inspecting `.Item(1)`."
    )
    out.append("")
    sv = _probe_sim_variable(com)
    if isinstance(sv, str):
        out.append(f"**Skipped:** {sv}")
    else:
        out.append(_format_endpoint_table("ISimVariable methods", sv))

    out.append("## Disposition")
    out.append("")
    out.append(
        "- Endpoints marked `YES` are wired into the MCP tool surface.\n"
        "- Endpoints marked `NO` are either ticketed for ModelRisk core, "
        "or — for `GetCorrelation` / `GetTornado` — pending developer "
        "confirmation per spec §8.4. The Python-via-numpy fallback uses "
        "`GetSamples()` for those.\n"
    )

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text("\n".join(out), encoding="utf-8")
    print(f"Wrote {OUT_PATH}")
    return 0 if overall_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
