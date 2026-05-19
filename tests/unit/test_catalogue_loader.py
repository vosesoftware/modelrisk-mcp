"""Tests for `modelrisk_mcp.bridge.catalogue`."""

from __future__ import annotations

import pytest

from modelrisk_mcp.bridge.catalogue import (
    FunctionCatalogue,
    FunctionSpec,
    ParamSpec,
    load_catalogue,
)
from modelrisk_mcp.errors import CatalogueError, UnknownFunctionError


@pytest.fixture(scope="module")
def cat() -> FunctionCatalogue:
    return load_catalogue()


def test_catalogue_is_cached(cat: FunctionCatalogue) -> None:
    cat_again = load_catalogue()
    assert cat is cat_again


def test_get_returns_none_for_missing(cat: FunctionCatalogue) -> None:
    assert cat.get("DefinitelyNotAVoseFunction") is None


def test_require_raises_unknown_function_error(cat: FunctionCatalogue) -> None:
    with pytest.raises(UnknownFunctionError) as exc:
        cat.require("VoseFoo")
    # Acceptance criterion from spec §13 Phase 1: close-match suggestion.
    msg = str(exc.value)
    assert "VoseFoo" in msg
    assert "Did you mean" in msg


def test_require_suggests_real_distribution(cat: FunctionCatalogue) -> None:
    with pytest.raises(UnknownFunctionError) as exc:
        cat.require("VoseModPert")  # capitalised wrong vs VoseModPERT
    assert "VoseModPERT" in str(exc.value)


def test_contains_membership(cat: FunctionCatalogue) -> None:
    assert "VoseNormal" in cat
    assert "Banana" not in cat
    assert 42 not in cat  # __contains__ accepts non-str safely


def test_iter_yields_specs(cat: FunctionCatalogue) -> None:
    first = next(iter(cat))
    assert isinstance(first, FunctionSpec)
    assert first.name.startswith("Vose")


def test_filter_by_category(cat: FunctionCatalogue) -> None:
    copulas = cat.filter("copula")
    assert len(copulas) > 0
    for spec in copulas:
        assert spec.category == "copula"
        assert spec.name.startswith("Vose")


def test_spec_required_and_all_param_names(cat: FunctionCatalogue) -> None:
    spec = cat.require("VoseModPERT")
    assert "min" in spec.required_param_names
    assert "mode" in spec.required_param_names
    assert "max" in spec.required_param_names
    # Optional ones present in all_param_names but not required.
    assert "u" in spec.all_param_names
    assert "u" not in spec.required_param_names


def test_load_catalogue_rejects_malformed_entry(tmp_path, monkeypatch):
    """Internal robustness check: if functions.json gains a malformed
    entry, the loader raises CatalogueError on its next call."""
    # We don't actually corrupt the packaged file — we exercise the
    # `_spec_from_entry` helper directly.
    from modelrisk_mcp.bridge.catalogue import _spec_from_entry

    with pytest.raises(CatalogueError):
        _spec_from_entry("VoseBroken", {"category": "continuous"})


def test_param_spec_is_immutable() -> None:
    p = ParamSpec(name="mu", type="number", required=True)
    with pytest.raises(Exception):  # noqa: B017 — frozen dataclass raises FrozenInstanceError
        p.name = "sigma"  # type: ignore[misc]
