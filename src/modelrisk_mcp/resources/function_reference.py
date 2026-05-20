"""Resources for the function catalogue."""

from __future__ import annotations

import json

from modelrisk_mcp.bridge.catalogue import load_catalogue
from modelrisk_mcp.errors import UnknownFunctionError
from modelrisk_mcp.server import mcp


@mcp.resource(
    uri="modelrisk://functions",
    name="modelrisk-function-catalogue",
    description=(
        "ModelRisk: full ModelRisk function catalogue (1400+ entries) "
        "as JSON. The LLM grounds every formula-writing tool against "
        "this list to prevent hallucinated function names."
    ),
    mime_type="application/json",
)
def function_catalogue_resource() -> str:
    cat = load_catalogue()
    payload = {
        name: {
            "category": spec.category,
            "returns": spec.returns,
            "description": spec.description,
            "parameters": [
                {
                    "name": p.name,
                    "type": p.type,
                    "required": p.required,
                    **({"default": p.default} if p.default is not None else {}),
                }
                for p in spec.parameters
            ],
        }
        for name, spec in cat.by_name.items()
    }
    return json.dumps(payload, indent=2, sort_keys=True)


@mcp.resource(
    uri="modelrisk://functions/{name}",
    name="modelrisk-function-entry",
    description=(
        "ModelRisk: single catalogue entry for one Vose function — "
        "category, parameters, return type, and description."
    ),
    mime_type="application/json",
)
def function_entry_resource(name: str) -> str:
    cat = load_catalogue()
    spec = cat.get(name)
    if spec is None:
        suggestions = cat.suggest(name)
        raise UnknownFunctionError(
            f"Function {name!r} not in catalogue."
            + (f" Did you mean: {', '.join(suggestions)}?" if suggestions else "")
        )
    payload = {
        "name": spec.name,
        "category": spec.category,
        "returns": spec.returns,
        "description": spec.description,
        "parameters": [
            {
                "name": p.name,
                "type": p.type,
                "required": p.required,
                **({"default": p.default} if p.default is not None else {}),
            }
            for p in spec.parameters
        ],
    }
    return json.dumps(payload, indent=2)
