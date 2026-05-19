from mcp.server.fastmcp import FastMCP

from modelrisk_mcp import __version__
from modelrisk_mcp.config import Settings

settings = Settings()

mcp = FastMCP(
    name="modelrisk-mcp",
    instructions=(
        "ModelRisk MCP exposes read and write access to Vose Software's ModelRisk "
        "Excel add-in. In v0.0.1 the tool surface is empty; tools, resources, and "
        "prompts land across Phases 2 through 5."
    ),
)


@mcp.resource("modelrisk://server/about")
def about() -> str:
    return (
        f"ModelRisk MCP Server v{__version__}\n"
        f"Read-only mode: {settings.read_only}\n"
        "Tool surface: empty (Phase 0 scaffold)."
    )
