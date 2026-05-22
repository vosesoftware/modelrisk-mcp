import pytest

from modelrisk_mcp import __version__
from modelrisk_mcp.server import mcp


def test_version_is_set() -> None:
    assert __version__ == "0.3.0a30"


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


def test_no_tool_returns_bare_list() -> None:
    """alpha.17 envelope-sweep guard: no tool may declare a bare
    `list[...]` return type.

    FastMCP serialises a bare list-typed return as one MCP content
    block per element, which makes the LLM see concatenated objects
    instead of a single array (bugs #1, #2, #15). The fix is to wrap
    every list response in a `{"<noun>": [...], "count": N}` dict so
    FastMCP emits one structured payload. This test scans the source
    of every tool function and fails if any uses `-> list[`."""
    import inspect

    from modelrisk_mcp.tools import (
        building,
        reading,
        restore,
        simulation,
        workflows,
    )

    offenders: list[str] = []
    for module in (building, reading, restore, simulation, workflows):
        for name, obj in inspect.getmembers(module, inspect.isfunction):
            if name.startswith("_") or name in {"get_bridge", "set_bridge_for_testing"}:
                continue
            # Only check things registered as MCP tools — helpers exempt.
            if not hasattr(obj, "__wrapped__") and not hasattr(obj, "fn"):
                # Heuristic: every @mcp.tool registers via @mcp.tool(...)
                # which decorates without adding __wrapped__. Easier to
                # just check the source string.
                pass
            try:
                src = inspect.getsource(obj)
            except (OSError, TypeError):
                continue
            if "-> list[" in src and "@mcp.tool" in src:
                offenders.append(f"{module.__name__}.{name}")
    assert not offenders, (
        f"Tools must wrap list returns in a dict envelope "
        f"(see alpha.17 sweep). Offenders: {offenders}"
    )
