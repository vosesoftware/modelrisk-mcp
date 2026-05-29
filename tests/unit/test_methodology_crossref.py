"""Guard against drift between the methodology resource and the audit
rule set.

The methodology resource (`modelrisk://methodology`) cross-references
the audit rules that enforce each principle. If a rule is renamed,
removed, or added without updating the methodology text, the
cross-references silently rot — knowledge that drifts from the code is
worse than no knowledge. These tests make that impossible to merge."""

from __future__ import annotations

import re

import yaml

from modelrisk_mcp.resources.methodology import _METHODOLOGY

# Match any rule id (VOSE-013, SS-002, …) so new rule families are
# covered by the coverage guard automatically.
_RULE_REF_RE = re.compile(r"\b[A-Z]{2,5}-\d{2,3}\b")


def _audit_rule_ids() -> set[str]:
    from importlib import resources

    text = (
        resources.files("modelrisk_mcp.data")
        .joinpath("audit_rules.yaml")
        .read_text(encoding="utf-8")
    )
    raw = yaml.safe_load(text) or {}
    return {r["id"] for r in raw.get("rules", [])}


def _referenced_ids() -> set[str]:
    return set(_RULE_REF_RE.findall(_METHODOLOGY))


def test_every_referenced_rule_exists() -> None:
    """No dangling references: every VOSE-### cited in the methodology
    must be a real rule."""
    referenced = _referenced_ids()
    actual = _audit_rule_ids()
    dangling = referenced - actual
    assert not dangling, (
        f"Methodology resource references rule(s) that don't exist in "
        f"audit_rules.yaml: {sorted(dangling)}"
    )


def test_every_rule_is_referenced() -> None:
    """Full coverage: every audit rule must be cited somewhere in the
    methodology resource, so adding a rule forces documenting which
    principle (or correctness class) it belongs to."""
    referenced = _referenced_ids()
    actual = _audit_rule_ids()
    orphaned = actual - referenced
    assert not orphaned, (
        f"Audit rule(s) not referenced anywhere in the methodology "
        f"resource — document which principle they enforce: "
        f"{sorted(orphaned)}"
    )


def test_principles_are_numbered_one_to_eight() -> None:
    """The 8 core principles are part of the resource's contract (the
    README and docs say '8 core principles')."""
    headings = re.findall(r"^## (\d+)\.", _METHODOLOGY, flags=re.MULTILINE)
    assert headings == [str(i) for i in range(1, 9)], (
        f"Expected principles numbered 1-8, found headings {headings}"
    )


def test_each_principle_has_why_and_failure_mode() -> None:
    """Every core principle carries the deepened structure: a *why* and
    a *failure mode*, not just the bare instruction."""
    # Split on principle headings, drop the preamble before principle 1.
    sections = re.split(r"^## \d+\. ", _METHODOLOGY, flags=re.MULTILINE)[1:]
    assert len(sections) == 8
    for i, body in enumerate(sections, start=1):
        assert "**Why.**" in body, f"Principle {i} missing **Why.**"
        assert "**Failure mode.**" in body, (
            f"Principle {i} missing **Failure mode.**"
        )
        assert "**Enforced by.**" in body, (
            f"Principle {i} missing **Enforced by.**"
        )
