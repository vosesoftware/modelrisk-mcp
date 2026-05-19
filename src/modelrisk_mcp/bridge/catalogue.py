"""Function catalogue loader.

Loads `functions.json` once at import time, exposes immutable lookup
operations, and provides the "did you mean?" close-match suggester that
the formula builder uses to make unknown-function errors actionable.
"""

from __future__ import annotations

import json
from collections.abc import Iterator
from dataclasses import dataclass, field
from difflib import get_close_matches
from functools import lru_cache
from importlib import resources
from typing import Any, Literal

from modelrisk_mcp.errors import CatalogueError, UnknownFunctionError

Category = Literal[
    "continuous",
    "discrete",
    "time-series",
    "aggregate",
    "copula",
    "fitting",
    "property",
    "object",
    "utility",
]

ReturnType = Literal["number", "array", "object"]
ParamType = Literal["number", "array", "boolean", "string", "object"]


@dataclass(frozen=True)
class ParamSpec:
    name: str
    type: ParamType
    required: bool
    default: Any | None = None


@dataclass(frozen=True)
class FunctionSpec:
    name: str
    category: Category
    parameters: tuple[ParamSpec, ...]
    returns: ReturnType
    description: str

    @property
    def required_param_names(self) -> tuple[str, ...]:
        return tuple(p.name for p in self.parameters if p.required)

    @property
    def all_param_names(self) -> tuple[str, ...]:
        return tuple(p.name for p in self.parameters)


def _empty_function_map() -> dict[str, FunctionSpec]:
    return {}


@dataclass(frozen=True)
class FunctionCatalogue:
    """Immutable view over the parsed functions.json file."""

    by_name: dict[str, FunctionSpec] = field(default_factory=_empty_function_map)

    def __contains__(self, name: object) -> bool:
        return isinstance(name, str) and name in self.by_name

    def __len__(self) -> int:
        return len(self.by_name)

    def __iter__(self) -> Iterator[FunctionSpec]:
        return iter(self.by_name.values())

    def get(self, name: str) -> FunctionSpec | None:
        return self.by_name.get(name)

    def require(self, name: str) -> FunctionSpec:
        spec = self.by_name.get(name)
        if spec is None:
            suggestions = self.suggest(name)
            hint = (
                f" Did you mean: {', '.join(suggestions)}?" if suggestions else ""
            )
            raise UnknownFunctionError(
                f"Function {name!r} not found in ModelRisk function catalogue."
                + hint
            )
        return spec

    def suggest(self, name: str, n: int = 3) -> list[str]:
        return get_close_matches(name, list(self.by_name.keys()), n=n, cutoff=0.6)

    def filter(self, category: Category) -> list[FunctionSpec]:
        return [s for s in self.by_name.values() if s.category == category]


def _spec_from_entry(name: str, raw: dict[str, Any]) -> FunctionSpec:
    try:
        params = tuple(
            ParamSpec(
                name=p["name"],
                type=p["type"],
                required=p["required"],
                default=p.get("default"),
            )
            for p in raw["parameters"]
        )
        return FunctionSpec(
            name=name,
            category=raw["category"],
            parameters=params,
            returns=raw["returns"],
            description=raw["description"],
        )
    except KeyError as exc:
        raise CatalogueError(f"Catalogue entry {name!r} is missing key {exc}") from exc


@lru_cache(maxsize=1)
def load_catalogue() -> FunctionCatalogue:
    """Load functions.json from the packaged data directory.

    Cached: the catalogue is large (1400+ entries) and immutable, so we
    parse it once per process.
    """
    text = (
        resources.files("modelrisk_mcp.data")
        .joinpath("functions.json")
        .read_text(encoding="utf-8")
    )
    raw = json.loads(text)
    if not isinstance(raw, dict):
        raise CatalogueError("functions.json root must be an object")
    by_name = {name: _spec_from_entry(name, entry) for name, entry in raw.items()}
    return FunctionCatalogue(by_name=by_name)
