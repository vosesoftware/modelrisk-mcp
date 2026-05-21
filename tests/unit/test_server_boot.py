import pytest

from modelrisk_mcp import __version__
from modelrisk_mcp.server import mcp


def test_version_is_set() -> None:
    assert __version__ == "0.3.0a10"


def test_server_name() -> None:
    assert mcp.name == "modelrisk-mcp"


@pytest.mark.asyncio
async def test_phase2_reading_tools_registered() -> None:
    """Phase 2 acceptance (§13): all 12 reading tools from §7.1 must be
    callable from MCP Inspector — i.e. they must appear in tools/list."""
    tools = await mcp.list_tools()
    names = {t.name for t in tools}
    expected = {
        "list_open_workbooks",
        "get_active_workbook",
        "get_workbook_summary",
        "list_modelrisk_inputs",
        "list_modelrisk_outputs",
        "list_distributions",
        "get_cell",
        "read_range",
        "get_simulation_results",
        "get_correlation_matrix",
        "get_sensitivity_ranking",
        "find_hard_coded_inputs",
    }
    missing = expected - names
    assert not missing, f"Phase 2 tools missing: {missing}"


@pytest.mark.asyncio
async def test_tool_descriptions_have_brand_prefix() -> None:
    """Spec §7: every tool's description starts with 'ModelRisk: '."""
    tools = await mcp.list_tools()
    for t in tools:
        assert t.description is not None
        assert t.description.startswith("ModelRisk: "), (
            f"Tool {t.name!r} missing brand prefix: {t.description!r}"
        )
