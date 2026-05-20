"""Distribution selection guide resource."""

from __future__ import annotations

from importlib import resources

from modelrisk_mcp.server import mcp


@mcp.resource(
    uri="modelrisk://distributions",
    name="modelrisk-distribution-guide",
    description=(
        "ModelRisk: methodology-grounded guide for choosing a "
        "distribution family given a description of the uncertain "
        "quantity. Drives propose_distributions_for_inputs."
    ),
    mime_type="application/yaml",
)
def distribution_guide_resource() -> str:
    return (
        resources.files("modelrisk_mcp.data")
        .joinpath("distributions.yaml")
        .read_text(encoding="utf-8")
    )
