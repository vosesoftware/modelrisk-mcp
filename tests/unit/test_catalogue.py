"""Sanity tests for the generated functions.json catalogue.

These pin a small set of canonical entries so a regression in
`scripts/extract_catalogue.py` (or in the upstream ModelRisk source) is
caught loudly when the catalogue is regenerated.
"""

from __future__ import annotations

import json
from importlib import resources
from pathlib import Path

import pytest

CATALOGUE_PATH = (
    Path(__file__).resolve().parents[2]
    / "src"
    / "modelrisk_mcp"
    / "data"
    / "functions.json"
)


@pytest.fixture(scope="module")
def catalogue() -> dict[str, dict]:
    with CATALOGUE_PATH.open(encoding="utf-8") as f:
        return json.load(f)


def test_catalogue_loads(catalogue: dict[str, dict]) -> None:
    assert len(catalogue) > 1000, "catalogue suspiciously small"


def test_every_entry_has_required_keys(catalogue: dict[str, dict]) -> None:
    required = {"category", "parameters", "returns", "description"}
    for name, entry in catalogue.items():
        missing = required - entry.keys()
        assert not missing, f"{name} missing keys: {missing}"
        assert isinstance(entry["parameters"], list), name
        for p in entry["parameters"]:
            assert set(p.keys()) == {"name", "type", "required"}, (name, p)
            assert p["type"] in {"number", "array", "boolean", "string", "object"}, (
                name,
                p,
            )
            assert isinstance(p["required"], bool), (name, p)


def test_all_names_have_vose_prefix(catalogue: dict[str, dict]) -> None:
    for name in catalogue:
        assert name.startswith("Vose"), name


def test_categories_are_in_spec_enum(catalogue: dict[str, dict]) -> None:
    allowed = {
        "continuous",
        "discrete",
        "time-series",
        "aggregate",
        "copula",
        "fitting",
        "property",
        "object",
        "utility",
    }
    seen = {entry["category"] for entry in catalogue.values()}
    extras = seen - allowed
    assert not extras, f"unexpected categories: {extras}"


@pytest.mark.parametrize(
    ("name", "expected_category"),
    [
        ("VoseModPERT", "continuous"),
        ("VoseNormal", "continuous"),
        ("VosePoisson", "discrete"),
        ("VoseBernoulli", "discrete"),
        ("VoseBinomial", "discrete"),
        ("VoseAggregateMC", "aggregate"),
        ("VoseCopulaMultiNormal", "copula"),
        ("VoseTimeGBM", "time-series"),
        ("VoseBetaFit", "fitting"),
        ("VoseShift", "utility"),
    ],
)
def test_canonical_entries_present_and_categorised(
    catalogue: dict[str, dict], name: str, expected_category: str
) -> None:
    assert name in catalogue, f"{name} missing from catalogue"
    assert catalogue[name]["category"] == expected_category, (
        f"{name}: expected {expected_category!r}, got {catalogue[name]['category']!r}"
    )


def test_vose_mod_pert_signature(catalogue: dict[str, dict]) -> None:
    entry = catalogue["VoseModPERT"]
    names = [p["name"] for p in entry["parameters"]]
    assert names[:4] == ["min", "mode", "max", "gamma"]
    required = [p["required"] for p in entry["parameters"]]
    # min, mode, max, gamma all required; u and extended* optional
    assert required[:4] == [True, True, True, True]
    assert any(not r for r in required[4:]), "expected at least one optional param"


def test_aggregate_mc_optional_limits(catalogue: dict[str, dict]) -> None:
    """The .h file marks Min/Max/DistributionShift as 'Not applied if omitted'.
    Confirm the extractor's heuristic translates this to required=False."""
    entry = catalogue["VoseAggregateMC"]
    by_name = {p["name"]: p for p in entry["parameters"]}
    assert by_name["n"]["required"] is True
    assert by_name["distribution"]["required"] is True
    assert by_name["MinLimit"]["required"] is False
    assert by_name["MaxLimit"]["required"] is False
    assert by_name["DistributionShift"]["required"] is False


def test_no_out_retval_leakage(catalogue: dict[str, dict]) -> None:
    """Regression: earlier versions leaked the IDL return-value param ([out,retval]
    VARIANT *res) as two extra parameters named '[out' and 'res'."""
    forbidden = {"[out", "res", "retval"}
    for name, entry in catalogue.items():
        for p in entry["parameters"]:
            assert p["name"] not in forbidden, f"{name}: leaked param {p['name']!r}"


def test_catalogue_is_bundled_in_package() -> None:
    """The catalogue must be readable via importlib.resources so the packaged
    wheel doesn't lose it when shipped."""
    text = (
        resources.files("modelrisk_mcp.data")
        .joinpath("functions.json")
        .read_text(encoding="utf-8")
    )
    data = json.loads(text)
    assert "VoseNormal" in data
