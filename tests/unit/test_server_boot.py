import pytest

from modelrisk_mcp import __version__
from modelrisk_mcp.server import mcp


def test_version_is_set() -> None:
    assert __version__ == "0.0.1"


def test_server_name() -> None:
    assert mcp.name == "modelrisk-mcp"


@pytest.mark.asyncio
async def test_tools_list_is_empty() -> None:
    tools = await mcp.list_tools()
    assert tools == []


@pytest.mark.asyncio
async def test_about_resource_is_registered() -> None:
    resources = await mcp.list_resources()
    uris = {str(r.uri) for r in resources}
    assert "modelrisk://server/about" in uris
