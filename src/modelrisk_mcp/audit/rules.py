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
from modelrisk_mcp.bridge.name_parser import extract_vose_first_arg
from modelrisk_mcp.safety import extract_call_heads
from modelrisk_mcp.schemas.results import AuditFinding
from modelrisk_mcp.schemas.workbook import CellRef

_VOSE_NAME_RE = re.compile(r"\bVose[A-Za-z0-9_]+")

# Bug #26 (alpha.25): replaced regex-only matchers with calls to
# `extract_vose_first_arg`. The old `_VOSE_INPUT_RE` only matched the
# string-literal form `VoseInput("Name")` — false-positive warnings
# fired on every distribution cell in workbooks using the cell-ref
# form `VoseInput(Cell)`, which is what most real ModelRisk models
# use. Same root cause as bug #13; the audit just didn't get updated
# when the scanner did.


def _has_input_wrapper(formula: str) -> bool:
    """True if the formula contains a VoseInput(...) wrapper in any
    form (string-literal name OR cell-reference name)."""
    return extract_vose_first_arg(formula, "VoseInput") is not None


def _has_output_wrapper(formula: str) -> bool:
    """True if the formula contains a VoseOutput(...) wrapper in any
    form. Mirror of `_has_input_wrapper`."""
    return extract_vose_first_arg(formula, "VoseOutput") is not None


# Legacy regexes kept around for the rules that genuinely want the
# string-literal form (e.g. VOSE-007 checking for `VoseOutput("")`).
_VOSE_INPUT_RE = re.compile(r'VoseInput\(\s*"[^"]*"\s*\)')
_VOSE_OUTPUT_RE = re.compile(r'VoseOutput\(\s*"[^"]*"\s*\)')

# Tighter capture: distinguishes "named output" from "VoseOutput()" /
# "VoseOutput("")". Used by detect_voseoutput_missing_name.
_VOSE_OUTPUT_NAME_RE = re.compile(
    r'VoseOutput\(\s*(?:"((?:[^"\\]|\\.|"")*)"\s*)?\)'
)

# Captures the literal first argument of VoseRiskEvent. Detects integer
# and float literals only — references and expressions don't match.
_VOSE_RISK_EVENT_PROB_RE = re.compile(
    r'VoseRiskEvent\(\s*([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)\s*[,)]'
)

# Captures VoseNormal(mu, sigma) with both args as numeric literals.
_VOSE_NORMAL_LITERAL_RE = re.compile(
    r'VoseNormal\(\s*([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)\s*,'
    r'\s*([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)\s*\)'
)


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
        # Bug #26 (alpha.25): also detect cell-ref-form VoseInput.
        if _has_input_wrapper(cell.formula):
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
    # Bug #26: also recognise cell-ref-form VoseOutput wrappers.
    output_cells = [
        c for c in ctx.cells if c.formula and _has_output_wrapper(c.formula)
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
        # Bug #26 (alpha.25): also detect cell-ref-form VoseInput.
        if _has_input_wrapper(cell.formula):
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


def detect_risk_event_degenerate_probability(
    ctx: RuleContext,
) -> Iterable[AuditFinding]:
    """VOSE-007 — VoseRiskEvent with literal probability of 0 or 1."""
    for cell in ctx.cells:
        if not cell.formula or "VoseRiskEvent" not in cell.formula:
            continue
        m = _VOSE_RISK_EVENT_PROB_RE.search(cell.formula)
        if not m:
            continue
        try:
            prob = float(m.group(1))
        except ValueError:
            continue
        if prob not in (0.0, 1.0):
            continue
        yield AuditFinding(
            severity=ctx.rule.severity,  # type: ignore[arg-type]
            cell=cell.ref,
            rule_id=ctx.rule.id,
            message=(
                f"Cell {cell.ref.a1} has VoseRiskEvent with probability "
                f"= {int(prob)}. The wrapper is degenerate at this value."
            ),
            suggested_fix=_format_fix(ctx.rule.suggested_fix_template),
        )


def detect_voseoutput_missing_name(
    ctx: RuleContext,
) -> Iterable[AuditFinding]:
    """VOSE-008 — VoseOutput() with no name or empty-string name."""
    for cell in ctx.cells:
        if not cell.formula or "VoseOutput" not in cell.formula:
            continue
        m = _VOSE_OUTPUT_NAME_RE.search(cell.formula)
        if m is None:
            continue
        captured = m.group(1)
        # Match groups: missing capture → None; empty literal → "".
        if captured is None or captured == "":
            yield AuditFinding(
                severity=ctx.rule.severity,  # type: ignore[arg-type]
                cell=cell.ref,
                rule_id=ctx.rule.id,
                message=(
                    f"Cell {cell.ref.a1} uses VoseOutput without a name; "
                    "Results Viewer / get_simulation_results cannot "
                    "reference it."
                ),
                suggested_fix=_format_fix(ctx.rule.suggested_fix_template),
            )


def detect_duplicate_output_names(
    ctx: RuleContext,
) -> Iterable[AuditFinding]:
    """VOSE-009 — same VoseOutput("name") declared on multiple cells."""
    named_output_re = re.compile(
        r'VoseOutput\(\s*"((?:[^"\\]|\\.|"")*)"\s*\)'
    )
    by_name: dict[str, list[CellRef]] = {}
    for cell in ctx.cells:
        if not cell.formula:
            continue
        m = named_output_re.search(cell.formula)
        if not m:
            continue
        name = m.group(1)
        if not name:
            continue
        by_name.setdefault(name, []).append(cell.ref)
    for name, refs in by_name.items():
        if len(refs) < 2:
            continue
        # Emit one finding per duplicated occurrence so the user sees
        # every offending cell in the report.
        for ref in refs:
            yield AuditFinding(
                severity=ctx.rule.severity,  # type: ignore[arg-type]
                cell=ref,
                rule_id=ctx.rule.id,
                message=(
                    f"Cell {ref.a1} declares VoseOutput({name!r}); the "
                    f"same name appears on {len(refs)} cells total."
                ),
                suggested_fix=_format_fix(ctx.rule.suggested_fix_template),
            )


def detect_input_wrapper_without_distribution(
    ctx: RuleContext,
) -> Iterable[AuditFinding]:
    """VOSE-010 — VoseInput wrapper but no Vose distribution function
    in the cell's formula."""
    for cell in ctx.cells:
        if not cell.formula or "VoseInput" not in cell.formula:
            continue
        # Bug #26: cell-ref form counts too.
        if not _has_input_wrapper(cell.formula):
            continue
        head = _first_distribution_head(cell.formula, ctx)
        if head is not None:
            continue
        yield AuditFinding(
            severity=ctx.rule.severity,  # type: ignore[arg-type]
            cell=cell.ref,
            rule_id=ctx.rule.id,
            message=(
                f"Cell {cell.ref.a1} is wrapped with VoseInput but has "
                "no Vose distribution function; the input won't vary "
                "across iterations."
            ),
            suggested_fix=_format_fix(ctx.rule.suggested_fix_template),
        )


def detect_high_volatility_normal_positive_mean(
    ctx: RuleContext,
) -> Iterable[AuditFinding]:
    """VOSE-011 — VoseNormal(mu, sigma) with mu > 0 and sigma > mu/2.

    Generates negatives ~16% of the time; user probably wants a
    positive-only distribution (lognormal, gamma)."""
    for cell in ctx.cells:
        if not cell.formula or "VoseNormal" not in cell.formula:
            continue
        m = _VOSE_NORMAL_LITERAL_RE.search(cell.formula)
        if not m:
            continue
        try:
            mu = float(m.group(1))
            sigma = float(m.group(2))
        except ValueError:
            continue
        if mu <= 0 or sigma <= mu / 2:
            continue
        yield AuditFinding(
            severity=ctx.rule.severity,  # type: ignore[arg-type]
            cell=cell.ref,
            rule_id=ctx.rule.id,
            message=(
                f"Cell {cell.ref.a1} has VoseNormal({mu:g}, {sigma:g}) "
                f"— sigma > mu/2, so ~16% of samples will be negative."
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


def detect_cell_evaluates_to_error(ctx: RuleContext) -> Iterable[AuditFinding]:
    """VOSE-012 — a cell evaluates to an Excel error (`#DIV/0!`, etc.).

    Powered by the bug-#34 fix (alpha.33) which surfaces error cells via
    `CellInfo.error`. Before that, error cells looked identical to empty
    cells through the audit and this rule couldn't fire. With the fix:

    - All errored cells get flagged (severity: error).
    - The message is sharper when the cell's formula contains a Vose
      call — that's the high-value case (a broken distribution will
      poison every simulation iteration).
    """
    for cell in ctx.cells:
        if cell.error is None:
            continue
        # Detect whether the formula contains a Vose call so we can
        # produce a sharper message. `extract_call_heads` is the
        # canonical "what functions does this formula call?" helper.
        vose_calls = [
            head for head in extract_call_heads(cell.formula or "")
            if head.startswith("Vose")
        ]
        if vose_calls:
            head = vose_calls[0]
            message = (
                f"Cell {cell.ref.a1} contains {head}(...) but evaluates "
                f"to {cell.error}. The distribution call is broken — "
                f"the simulation will produce error samples from this "
                f"cell on every iteration."
            )
        else:
            message = (
                f"Cell {cell.ref.a1} evaluates to {cell.error}. Trace "
                f"the formula back to find the root cause."
            )
        yield AuditFinding(
            severity=ctx.rule.severity,  # type: ignore[arg-type]
            cell=cell.ref,
            rule_id=ctx.rule.id,
            message=message,
            suggested_fix=_format_fix(
                ctx.rule.suggested_fix_template,
                error=cell.error,
            ),
        )


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
    "risk_event_degenerate_probability": (
        detect_risk_event_degenerate_probability
    ),
    "voseoutput_missing_name": detect_voseoutput_missing_name,
    "duplicate_output_names": detect_duplicate_output_names,
    "input_wrapper_without_distribution": (
        detect_input_wrapper_without_distribution
    ),
    "high_volatility_normal_positive_mean": (
        detect_high_volatility_normal_positive_mean
    ),
    "cell_evaluates_to_error": detect_cell_evaluates_to_error,
}
