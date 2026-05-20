"""Audit engine — runs the §7.4 audit_model tool's checks.

Each rule lives in `rules.py` as a small function taking a `RuleContext`
and yielding `AuditFinding`s. Which rules run, and at what severity, is
driven by `data/audit_rules.yaml` so the rule set is editable without
touching code.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable
from dataclasses import dataclass
from importlib import resources

import yaml

from modelrisk_mcp.bridge.catalogue import FunctionCatalogue
from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.schemas.results import AuditFinding, AuditReport
from modelrisk_mcp.schemas.workbook import CellInfo


@dataclass(frozen=True)
class RuleSpec:
    id: str
    name: str
    severity: str
    enabled: bool
    description: str
    suggested_fix_template: str


@dataclass
class RuleContext:
    """Bag of state every detector can lean on."""

    bridge: ModelRiskBridge
    catalogue: FunctionCatalogue
    workbook: str
    cells: list[CellInfo]
    rule: RuleSpec


Detector = Callable[[RuleContext], Iterable[AuditFinding]]


def load_rules() -> list[RuleSpec]:
    text = (
        resources.files("modelrisk_mcp.data")
        .joinpath("audit_rules.yaml")
        .read_text(encoding="utf-8")
    )
    raw = yaml.safe_load(text) or {}
    rules_raw = raw.get("rules", [])
    return [
        RuleSpec(
            id=r["id"],
            name=r["name"],
            severity=r.get("severity", "warning"),
            enabled=r.get("enabled", True),
            description=r.get("description", "").strip(),
            suggested_fix_template=r.get("suggested_fix_template", "").strip(),
        )
        for r in rules_raw
    ]


def run_audit(bridge: ModelRiskBridge, workbook: str) -> AuditReport:
    from modelrisk_mcp.audit.rules import RULES_BY_NAME

    rules = [r for r in load_rules() if r.enabled]
    cells = list(bridge.excel.iterate_cells(workbook))
    findings: list[AuditFinding] = []
    for spec in rules:
        detector = RULES_BY_NAME.get(spec.name)
        if detector is None:
            continue
        ctx = RuleContext(
            bridge=bridge,
            catalogue=bridge.catalogue,
            workbook=workbook,
            cells=cells,
            rule=spec,
        )
        for finding in detector(ctx):
            findings.append(finding)
    return AuditReport(findings=findings)
