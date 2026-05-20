"""Resource registrations (spec §7.5).

Importing this package side-effects every `@mcp.resource(...)`
registration into the shared FastMCP instance.
"""

from modelrisk_mcp.resources import (
    audit_rules,
    distribution_guide,
    function_reference,
    methodology,
    workbook_state,
)

__all__ = [
    "audit_rules",
    "distribution_guide",
    "function_reference",
    "methodology",
    "workbook_state",
]
