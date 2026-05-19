"""Phase 0 spike: enumerate the ModelRisk COM surface.

Discovers which methods and properties are actually exposed on the
ModelRisk add-in's Application object and its Results sub-object, then
writes the findings to docs/com-surface.md. This becomes the canonical
record of what's available vs what is still ticketed for the ModelRisk
core team.

Run from a Windows machine that has Excel + ModelRisk installed:

    uv run python scripts/spike_com_surface.py
"""

from __future__ import annotations

import sys
from datetime import UTC, datetime
from pathlib import Path

OUT_PATH = Path(__file__).resolve().parent.parent / "docs" / "com-surface.md"

EXPECTED_APPLICATION_METHODS = (
    "Simulate",
    "StopSimulation",
)
EXPECTED_APPLICATION_PROPERTIES = (
    "SimulationStatus",
    "Settings",
)
EXPECTED_SETTINGS_PROPERTIES = (
    "Iterations",
    "Seed",
    "SamplingMethod",
)
EXPECTED_RESULTS_METHODS = (
    "GetMean",
    "GetPercentile",
    "GetCorrelation",
    "GetTornado",
    "GetIterations",
)


def _probe(obj: object, names: tuple[str, ...]) -> dict[str, bool]:
    return {n: hasattr(obj, n) for n in names}


def _format_table(title: str, results: dict[str, bool]) -> str:
    lines = [f"### {title}", "", "| Endpoint | Available |", "| --- | --- |"]
    for name, ok in results.items():
        lines.append(f"| `{name}` | {'YES' if ok else 'NO'} |")
    lines.append("")
    return "\n".join(lines)


def main() -> int:
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print("pywin32 is not installed; run `uv sync` first.", file=sys.stderr)
        return 2

    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        try:
            excel = win32com.client.Dispatch("Excel.Application")
        except Exception as exc:
            print(f"Could not connect to Excel: {exc}", file=sys.stderr)
            return 2

    sections: list[str] = []
    sections.append(f"# ModelRisk COM surface\n\nProbed at {datetime.now(UTC).isoformat()}.\n")

    modelrisk_app: object | None = None
    for candidate in ("ModelRisk.Application", "ModelRisk"):
        try:
            modelrisk_app = win32com.client.Dispatch(candidate)
            sections.append(f"ModelRisk COM root resolved via `{candidate}`.\n")
            break
        except Exception:
            continue

    if modelrisk_app is None:
        sections.append(
            "**Could not resolve ModelRisk COM root.** Excel started, but no "
            "ModelRisk COM object was found. Confirm ModelRisk is installed "
            "and loaded, then re-run.\n"
        )
        OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
        OUT_PATH.write_text("\n".join(sections), encoding="utf-8")
        return 1

    sections.append(_format_table(
        "Application methods",
        _probe(modelrisk_app, EXPECTED_APPLICATION_METHODS),
    ))
    sections.append(_format_table(
        "Application properties",
        _probe(modelrisk_app, EXPECTED_APPLICATION_PROPERTIES),
    ))

    settings = getattr(modelrisk_app, "Settings", None)
    if settings is not None:
        sections.append(_format_table(
            "Application.Settings properties",
            _probe(settings, EXPECTED_SETTINGS_PROPERTIES),
        ))
    else:
        sections.append("### Application.Settings properties\n\nNot reachable.\n")

    results = getattr(modelrisk_app, "Results", None)
    if results is not None:
        sections.append(_format_table(
            "Application.Results methods",
            _probe(results, EXPECTED_RESULTS_METHODS),
        ))
    else:
        sections.append("### Application.Results methods\n\nNot reachable.\n")

    sections.append(
        "## Disposition\n\n"
        "- Endpoints marked `YES` are wired into the MCP tool surface.\n"
        "- Endpoints marked `NO` ship as `SimulationNotAvailableError` stubs in v0.1 "
        "and are tracked for the ModelRisk core team.\n"
    )

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text("\n".join(sections), encoding="utf-8")
    print(f"Wrote {OUT_PATH}")

    _ = excel  # keep ref alive
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
