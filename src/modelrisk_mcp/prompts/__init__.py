"""Prompt registrations (spec §7.6).

Importing this package side-effects every `@mcp.prompt(...)`
registration. Each prompt file exports `template: str` and
`description: str` per the spec.
"""

from modelrisk_mcp.prompts import (
    add_uncertainty,
    audit_model,
    build_model,
    import_legacy_model,
    interpret_results,
)

__all__ = [
    "add_uncertainty",
    "audit_model",
    "build_model",
    "import_legacy_model",
    "interpret_results",
]
