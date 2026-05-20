"""Phase 5 prompt tests — verify every prompt is registered with the
expected name + non-trivial content."""

from __future__ import annotations

import pytest

from modelrisk_mcp.server import mcp

EXPECTED_PROMPT_NAMES: set[str] = {
    "build-risk-model",
    "audit-model",
    "interpret-results",
    "add-uncertainty",
    "import-legacy-model",
}


@pytest.mark.asyncio
async def test_five_prompts_registered() -> None:
    listed = await mcp.list_prompts()
    names = {p.name for p in listed}
    missing = EXPECTED_PROMPT_NAMES - names
    assert not missing, f"Phase 5 prompts missing: {missing}"


@pytest.mark.asyncio
async def test_prompts_have_descriptions() -> None:
    listed = await mcp.list_prompts()
    for p in listed:
        if p.name not in EXPECTED_PROMPT_NAMES:
            continue
        assert p.description is not None
        # Descriptions should be substantive.
        assert len(p.description) > 50


@pytest.mark.asyncio
async def test_get_build_risk_model_returns_template() -> None:
    result = await mcp.get_prompt("build-risk-model")
    messages = result.messages
    assert len(messages) >= 1
    # The template mentions key tools the LLM should invoke.
    body = " ".join(str(m.content) for m in messages)
    assert "get_active_workbook" in body
    assert "propose_distributions_for_inputs" in body
    assert "run_simulation" in body


@pytest.mark.asyncio
async def test_get_audit_model_prompt_returns_template() -> None:
    result = await mcp.get_prompt("audit-model")
    body = " ".join(str(m.content) for m in result.messages)
    assert "audit_model" in body
    assert "restore_cell" in body


@pytest.mark.asyncio
async def test_get_import_legacy_model_prompt_returns_template() -> None:
    result = await mcp.get_prompt("import-legacy-model")
    body = " ".join(str(m.content) for m in result.messages)
    # Mentions vendor prefixes it knows how to map.
    assert "Risk" in body
    assert "Vose" in body
