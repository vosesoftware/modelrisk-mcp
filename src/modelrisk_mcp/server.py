"""MCP server entrypoint (mcp 2.0 SDK).

Constructs the `mcp` instance and triggers tool registration by importing
`modelrisk_mcp.tools`. The tools/* modules attach themselves via the
`@mcp.tool(...)` decorator side-effect.

Ported from the 1.x SDK's FastMCP to 2.0's MCPServer (same high-level
decorator surface; the 1.x `mcp.server.fastmcp` module no longer exists,
which is what broke 0.3.x installs the moment a co-installed package
pulled mcp>=2 into the shared environment).
"""

from __future__ import annotations

from mcp.server import MCPServer

from modelrisk_mcp import __version__
from modelrisk_mcp.config import Settings

settings = Settings()

mcp = MCPServer(
    name="modelrisk-mcp",
    version=__version__,
    instructions=(
        "ModelRisk MCP exposes read and write access to Vose Software's "
        "ModelRisk Excel add-in. Use the reading tools to inspect an "
        "existing model and its simulation results, and (in later "
        "phases) the building tools to insert distributions, fit "
        "distributions to data, build aggregates, copulas, and "
        "time-series, and run simulations."
    ),
)


# Importing the tools + resources + prompts packages side-effects every
# @mcp.tool / @mcp.resource / @mcp.prompt registration. MUST come after
# `mcp` is constructed.
from modelrisk_mcp import prompts, resources, tools  # noqa: E402, F401

__all__ = ["__version__", "mcp", "settings"]
