"""Phase 5 resource tests — verify every modelrisk:// URI is registered
and returns non-empty content."""

from __future__ import annotations

import json

import pytest

from modelrisk_mcp.server import mcp


@pytest.mark.asyncio
async def test_seven_resources_registered() -> None:
    """Spec §7.5 lists 7 URIs total: 5 static + 2 templated.
    list_resources only enumerates static ones (templated are
    enumerated by list_resource_templates), so we check both."""
    static = await mcp.list_resources()
    templated = await mcp.list_resource_templates()
    static_uris = {str(r.uri) for r in static}
    templated_uris = {t.uriTemplate for t in templated}
    # Static resources.
    assert "modelrisk://functions" in static_uris
    assert "modelrisk://distributions" in static_uris
    assert "modelrisk://methodology" in static_uris
    assert "modelrisk://knowledge" in static_uris
    assert "modelrisk://workbook/current" in static_uris
    assert "modelrisk://audit-rules" in static_uris
    # Templated resources.
    assert "modelrisk://functions/{name}" in templated_uris
    assert "modelrisk://workbook/current/sheet/{name}" in templated_uris


@pytest.mark.asyncio
async def test_resources_have_brand_prefix() -> None:
    static = await mcp.list_resources()
    for r in static:
        assert r.description is not None
        assert r.description.startswith("ModelRisk: "), (
            f"Resource {r.uri} description missing prefix: {r.description}"
        )


@pytest.mark.asyncio
async def test_function_catalogue_resource_returns_json() -> None:
    from modelrisk_mcp.resources.function_reference import (
        function_catalogue_resource,
    )

    content = function_catalogue_resource()
    data = json.loads(content)
    assert "VoseNormal" in data
    assert "VoseModPERT" in data
    # Schema is consistent with the unit tests for the loader.
    entry = data["VoseNormal"]
    assert "category" in entry
    assert "parameters" in entry
    assert "returns" in entry


@pytest.mark.asyncio
async def test_function_entry_resource_returns_known_entry() -> None:
    from modelrisk_mcp.resources.function_reference import (
        function_entry_resource,
    )

    payload = json.loads(function_entry_resource("VoseModPERT"))
    assert payload["name"] == "VoseModPERT"
    assert payload["category"] == "continuous"


def test_knowledge_base_resource_returns_cited_markdown() -> None:
    """The curated knowledge base ships as a data file, is served as
    markdown, and must carry its ModelRisk Help attribution so the
    distillation never loses its provenance."""
    from modelrisk_mcp.resources.knowledge_base import (
        knowledge_base_resource,
    )

    md = knowledge_base_resource()
    assert len(md) > 1000
    # Provenance is non-negotiable for a distilled resource.
    assert "ModelRisk Help" in md
    # The six distilled topics are present.
    for topic in (
        "two forms of uncertainty",
        "Selecting the appropriate distribution",
        "expert opinion",
        "correlation",
        "over time",
        "Fitting distributions to data",
        "aggregation method",
        "Interpreting sensitivity",
        "family quick-reference",
    ):
        assert topic.lower() in md.lower(), f"knowledge base missing: {topic}"


@pytest.mark.asyncio
async def test_function_entry_unknown_raises_with_suggestion() -> None:
    from modelrisk_mcp.errors import UnknownFunctionError
    from modelrisk_mcp.resources.function_reference import (
        function_entry_resource,
    )

    with pytest.raises(UnknownFunctionError) as exc:
        function_entry_resource("VoseFoo")
    assert "Did you mean" in str(exc.value)


def test_methodology_resource_returns_markdown() -> None:
    from modelrisk_mcp.resources.methodology import methodology_resource

    content = methodology_resource()
    assert content.startswith("# Vose ModelRisk methodology")
    assert "VoseRiskEvent" in content


def test_distribution_guide_resource_returns_yaml() -> None:
    from modelrisk_mcp.resources.distribution_guide import (
        distribution_guide_resource,
    )

    content = distribution_guide_resource()
    assert "scenarios:" in content
    assert "VoseModPERT" in content


def test_audit_rules_resource_returns_yaml() -> None:
    from modelrisk_mcp.resources.audit_rules import audit_rules_resource

    content = audit_rules_resource()
    assert "VOSE-001" in content
    assert "VOSE-006" in content
