"""MCP tool registrations.

Importing this package side-effects every `@mcp.tool(...)` registration
into the shared FastMCP instance from `modelrisk_mcp.server`.
"""

from modelrisk_mcp.tools import reading

__all__ = ["reading"]
