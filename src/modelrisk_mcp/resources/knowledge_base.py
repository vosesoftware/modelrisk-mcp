"""Risk-analysis knowledge base resource.

A curated, attributed distillation of risk-modelling guidance from the
ModelRisk Help (Vose Software), loaded into the LLM's context so model
construction and critique are grounded in the authoritative source
rather than general intuition. The substance lives in the shipped data
file `data/knowledge_base.md` (single source of truth, also mirrored to
docs for GitHub readers)."""

from __future__ import annotations

from importlib import resources

from modelrisk_mcp.server import mcp


@mcp.resource(
    uri="modelrisk://knowledge",
    name="modelrisk-knowledge-base",
    description=(
        "ModelRisk: curated risk-analysis knowledge base distilled from "
        "the ModelRisk Help (Vose Software) — why to quantify "
        "uncertainty, the two forms of uncertainty, selecting "
        "distributions (the five properties), eliciting expert opinion, "
        "modeling correlation, and modeling over time. Consult when "
        "building or critiquing a model for authoritative methodology."
    ),
    mime_type="text/markdown",
)
def knowledge_base_resource() -> str:
    return (
        resources.files("modelrisk_mcp.data")
        .joinpath("knowledge_base.md")
        .read_text(encoding="utf-8")
    )
