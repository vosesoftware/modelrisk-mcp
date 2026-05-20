"""Audit rule set resource."""

from __future__ import annotations

from importlib import resources

from modelrisk_mcp.server import mcp


@mcp.resource(
    uri="modelrisk://audit-rules",
    name="modelrisk-audit-rules",
    description=(
        "ModelRisk: the active audit-rule set as YAML. Lists every "
        "rule's id, name, severity, enabled state, description, and "
        "suggested-fix template. Editable in data/audit_rules.yaml "
        "without code changes."
    ),
    mime_type="application/yaml",
)
def audit_rules_resource() -> str:
    return (
        resources.files("modelrisk_mcp.data")
        .joinpath("audit_rules.yaml")
        .read_text(encoding="utf-8")
    )
