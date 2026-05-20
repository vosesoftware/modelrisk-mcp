"""Individual audit-rule detectors.

Each function takes a `RuleContext` and yields zero or more
`AuditFinding`s. The mapping at the bottom binds rule names (matching
`audit_rules.yaml`) to detector functions.

Adding a rule:
1. Add the entry to `audit_rules.yaml`.
2. Implement the detector here.
3. Register it in `RULES_BY_NAME`.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterable

from modelrisk_mcp.audit.engine import RuleContext
from modelrisk_mcp.safety import extract_call_heads
from modelrisk_mcp.schemas.results import AuditFinding

_VOSE_NAME_RE = re.compile(r"\bVose[A-Za-z0-9_]+")
_VOSE_INPUT_RE = re.compile(r'VoseInput\(\s*"[^"]*"\s*\)')
_VOSE_OUTPUT_RE = re.compile(r'VoseOutput\(\s*"[^"]*"\s*\)')


def _format_fix(template: str, **kwargs: object) -> str:
    try:
        return template.format(**kwargs)
    except (KeyError, IndexError):
        return template


def detect_unknown_vose_function(ctx: RuleContext) -> Iterable[AuditFinding]:
    """VOSE-001 — formula contains a Vose-prefixed call that's not in
    the catalogue (typo or hand-edit)."""
    catalogue_size = len(ctx.catalogue)
    for cell in ctx.cells:
        if not cell.formula:
            continue
        for head in extract_call_heads(cell.formula):
            if not head.startswith("Vose"):
                continue
            if head in ctx.catalogue:
                continue
            yield AuditFinding(
                severity=ctx.rule.severity,  # type: ignore[arg-type]
                cell=cell.ref,
                rule_id=ctx.rule.id,
                message=(
                    f"Cell {cell.ref.a1} calls {head!r}, which is not in "
                    f"the ModelRisk catalogue."
                ),
                suggested_fix=_format_fix(
                    ctx.rule.suggested_fix_template,
                    function_name=head,
                    catalogue_size=catalogue_size,
                ),
            )


def detect_distribution_without_input_wrapper(
    ctx: RuleContext,
) -> Iterable[AuditFinding]:
    """VOSE-002 — distribution function used but no VoseInput wrapper."""
    for cell in ctx.cells:
        if not cell.formula:
            continue
        if _VOSE_INPUT_RE.search(cell.formula):
            continue
        # Skip object-category functions: they're meant to be passed to
        # VoseAggregateMC / VoseRiskEvent etc., not wrapped directly.
        head = _first_distribution_head(cell.formula, ctx)
        if head is None:
            continue
        spec = ctx.catalogue.get(head)
        if spec is None or spec.category == "object":
            continue
        yield AuditFinding(
            severity=ctx.rule.severity,  # type: ignore[arg-type]
            cell=cell.ref,
            rule_id=ctx.rule.id,
            message=(
                f"Cell {cell.ref.a1} uses {head} without a VoseInput "
                f"wrapper; it won't appear in the Results Viewer."
            ),
            suggested_fix=_format_fix(ctx.rule.suggested_fix_template),
        )


def detect_fit_without_uncertainty(ctx: RuleContext) -> Iterable[AuditFinding]:
    """VOSE-003 — Vose*Fit(...) without uncertainty=TRUE."""
    for cell in ctx.cells:
        if not cell.formula:
            continue
        for head in extract_call_heads(cell.formula):
            spec = ctx.catalogue.get(head)
            if spec is None or spec.category != "fitting":
                continue
            # Crude check: look for "TRUE" anywhere in the cell's
            # uncertainty position. The fit functions all have
            # `uncertainty` as the second positional parameter; for the
            # v0.1 audit it's enough to check that the formula mentions
            # TRUE — false positives possible but bias is conservative.
            if re.search(r"\bTRUE\b", cell.formula, re.IGNORECASE):
                continue
            yield AuditFinding(
                severity=ctx.rule.severity,  # type: ignore[arg-type]
                cell=cell.ref,
                rule_id=ctx.rule.id,
                message=(
                    f"Cell {cell.ref.a1} uses {head} without "
                    f"uncertainty=TRUE."
                ),
                suggested_fix=_format_fix(ctx.rule.suggested_fix_template),
            )


def detect_output_no_distribution_reference(
    ctx: RuleContext,
) -> Iterable[AuditFinding]:
    """VOSE-004 — VoseOutput cell whose formula doesn't reference any
    distribution-bearing cell transitively (best-effort, no full graph
    walk; we just look for any other Vose* call in the workbook that the
    output formula could plausibly reach via a cell reference)."""
    output_cells = [
        c for c in ctx.cells if c.formula and _VOSE_OUTPUT_RE.search(c.formula)
    ]
    if not output_cells:
        return
    # Build the set of cells that themselves contain a Vose distribution.
    distribution_refs: set[str] = set()
    for c in ctx.cells:
        if not c.formula:
            continue
        if any(
            h in ctx.catalogue
            and ctx.catalogue.get(h)
            and ctx.catalogue.get(h).category  # type: ignore[union-attr]
            in {"continuous", "discrete", "aggregate", "time-series", "copula"}
            for h in extract_call_heads(c.formula)
        ):
            distribution_refs.add(f"{c.ref.sheet}!{c.ref.cell}")
            distribution_refs.add(c.ref.cell)
    ref_token_re = re.compile(r"\$?[A-Z]+\$?\d+")
    for out in output_cells:
        # Skip if the output formula directly contains any Vose
        # distribution call (then it's by construction stochastic).
        heads = [
            h
            for h in extract_call_heads(out.formula)
            if (spec := ctx.catalogue.get(h)) is not None
            and spec.category
            in {"continuous", "discrete", "aggregate", "time-series", "copula"}
        ]
        if heads:
            continue
        referenced = set(ref_token_re.findall(out.formula))
        if referenced & distribution_refs:
            continue
        yield AuditFinding(
            severity=ctx.rule.severity,  # type: ignore[arg-type]
            cell=out.ref,
            rule_id=ctx.rule.id,
            message=(
                f"Output cell {out.ref.a1} doesn't reference any "
                f"distribution cell directly. Iteration values will be "
                f"deterministic."
            ),
            suggested_fix=_format_fix(ctx.rule.suggested_fix_template),
        )


def detect_arithmetic_before_input_wrapper(
    ctx: RuleContext,
) -> Iterable[AuditFinding]:
    """VOSE-005 — formula like `=2*VoseNormal(0,1)` with no VoseInput
    wrapper. The arithmetic-then-distribution pattern hides the cell
    from the Results Viewer."""
    for cell in ctx.cells:
        if not cell.formula:
            continue
        if _VOSE_INPUT_RE.search(cell.formula):
            continue
        body = cell.formula.lstrip("=").lstrip()
        # Heuristic: contains arithmetic operator AND a Vose distribution call.
        has_arith = bool(re.search(r"[*/+\-^]", body))
        if not has_arith:
            continue
        head = _first_distribution_head(cell.formula, ctx)
        if head is None:
            continue
        spec = ctx.catalogue.get(head)
        if spec is None or spec.category == "object":
            continue
        # Skip if the formula starts directly with the Vose call (no
        # arithmetic prefix) — those are caught by VOSE-002.
        first_token = body.split("(", 1)[0]
        if first_token == head:
            continue
        yield AuditFinding(
            severity=ctx.rule.severity,  # type: ignore[arg-type]
            cell=cell.ref,
            rule_id=ctx.rule.id,
            message=(
                f"Cell {cell.ref.a1} has arithmetic before/around "
                f"{head}; consider wrapping with VoseInput or refactoring."
            ),
            suggested_fix=_format_fix(ctx.rule.suggested_fix_template),
        )


def detect_hard_coded_inputs(ctx: RuleContext) -> Iterable[AuditFinding]:
    """VOSE-006 — numeric cells referenced by formulas (candidates for
    distribution replacement). Uses the existing ModelRiskBridge
    finder."""
    candidates = ctx.bridge.find_hard_coded_inputs(ctx.workbook)
    for ref in candidates[:25]:  # cap to keep the audit report tidy
        yield AuditFinding(
            severity=ctx.rule.severity,  # type: ignore[arg-type]
            cell=ref,
            rule_id=ctx.rule.id,
            message=(
                f"Cell {ref.a1} is a plain number referenced by at least "
                f"one formula. Candidate for distribution replacement."
            ),
            suggested_fix=_format_fix(ctx.rule.suggested_fix_template),
        )


def _first_distribution_head(
    formula: str, ctx: RuleContext
) -> str | None:
    for head in extract_call_heads(formula):
        spec = ctx.catalogue.get(head)
        if spec is None:
            continue
        if spec.category in {
            "continuous",
            "discrete",
            "aggregate",
            "time-series",
            "copula",
            "fitting",
            "object",
        }:
            return head
    return None


Detector = Callable[[RuleContext], Iterable[AuditFinding]]

RULES_BY_NAME: dict[str, Detector] = {
    "unknown_vose_function": detect_unknown_vose_function,
    "distribution_without_input_wrapper": (
        detect_distribution_without_input_wrapper
    ),
    "fit_without_uncertainty": detect_fit_without_uncertainty,
    "output_cell_no_distribution_reference": (
        detect_output_no_distribution_reference
    ),
    "arithmetic_before_input_wrapper": detect_arithmetic_before_input_wrapper,
    "hard_coded_inputs_present": detect_hard_coded_inputs,
}
