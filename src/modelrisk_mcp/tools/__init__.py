"""MCP tool registrations.

Importing this package side-effects every `@mcp.tool(...)` registration
into the shared FastMCP instance from `modelrisk_mcp.server`. Order
matters: the reading module is imported first because building/restore
depend on its `get_bridge` factory.
"""

from modelrisk_mcp.tools import building, reading, restore, simulation, workflows

__all__ = ["building", "reading", "restore", "simulation", "workflows"]
