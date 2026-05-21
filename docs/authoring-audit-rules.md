# Authoring audit rules

The audit set is editable. Adding a rule is a three-file change:
1. `src/modelrisk_mcp/data/audit_rules.yaml` — metadata + the human-readable text.
2. `src/modelrisk_mcp/audit/rules.py` — the Python detector.
3. `tests/unit/test_audit.py` (or a new file) — at least one positive + one negative case.

The two files are deliberately decoupled. The YAML controls *what* runs, with what severity, and what the report says. The Python implements *how* the rule detects the problem. You can disable a rule without touching code by setting `enabled: false` on its YAML entry.

## Rule lifecycle

When `audit_model` is invoked, the engine in `audit/engine.py`:
1. Loads `audit_rules.yaml`.
2. For each `enabled: true` entry, looks up the matching detector in `RULES_BY_NAME` in `audit/rules.py`.
3. Calls the detector with a `RuleContext` holding the workbook's cells, the function catalogue, and a back-reference to the bridge for tools that need more than per-cell info (e.g. `find_hard_coded_inputs`).
4. Collects the yielded `AuditFinding` objects into the report.

Detectors must be cheap. They run on every cell of the workbook; a `re.search` per cell is fine, but a per-cell COM round-trip is not.

## Step 1 — YAML entry

Add a block under `rules:`. Required fields:

```yaml
- id: VOSE-XXX           # next free identifier
  name: snake_case_name  # must match RULES_BY_NAME key in rules.py
  severity: error|warning|info
  enabled: true
  description: >
    What the rule detects, in 1-3 sentences. Renders as the
    explanation in the audit report.
  suggested_fix_template: >
    What the user should do about it. May reference {function_name}
    or {catalogue_size} (the detector supplies these via _format_fix).
```

Pick `severity` from:
- **`error`** — almost certainly wrong; should be fixed before running. Examples: unknown Vose function, VoseOutput missing name.
- **`warning`** — methodology concern; the model will run but the result may be misleading. Examples: fit without `uncertainty=TRUE`, high-volatility normal with positive mean.
- **`info`** — observation worth flagging but not a defect. Examples: hard-coded numeric cells that could become distributions.

## Step 2 — Python detector

```python
def detect_my_new_rule(ctx: RuleContext) -> Iterable[AuditFinding]:
    """VOSE-XXX — one-line summary of what this rule catches."""
    for cell in ctx.cells:
        if not cell.formula:
            continue
        # ... your detection logic ...
        if not _is_problem(cell.formula):
            continue
        yield AuditFinding(
            severity=ctx.rule.severity,  # type: ignore[arg-type]
            cell=cell.ref,
            rule_id=ctx.rule.id,
            message=f"Cell {cell.ref.a1} has the problem because ...",
            suggested_fix=_format_fix(ctx.rule.suggested_fix_template),
        )
```

Register it at the bottom of `rules.py`:

```python
RULES_BY_NAME: dict[str, Detector] = {
    # ... existing rules ...
    "my_new_rule": detect_my_new_rule,
}
```

### Patterns that work well

- **Per-cell regex match.** Most rules look at one cell's formula in isolation. The existing rules use compiled module-level regex constants.
- **Cross-cell aggregation.** If the rule needs to find duplicates or relationships (e.g. VOSE-009 `duplicate_output_names`), build a dict-of-lists as you walk `ctx.cells`, then emit findings in a second pass.
- **Bridge calls for workbook-wide computation.** For things like "is this numeric cell referenced by at least one formula", call `ctx.bridge.find_hard_coded_inputs(ctx.workbook)` — the bridge already does the work.

### Patterns to avoid

- **Iterating `ctx.cells` more than twice per rule.** Once to gather state, once to emit. The audit runs every rule against every cell; quadratic behaviour adds up.
- **COM round-trips inside the detector.** All cell data is already in `ctx.cells` from a single sweep.
- **Hardcoded function names.** Look up the function in `ctx.catalogue` instead — that respects user overrides via `data/optional_overrides.yaml`.

### What goes in `message` vs `suggested_fix`

- **`message`** describes the specific instance: which cell, what value, why it's flagged. The audit report uses this verbatim.
- **`suggested_fix`** is generic guidance pulled from `suggested_fix_template`. Use the same template for every cell that triggers a given rule.

## Step 3 — Tests

At minimum:

```python
def test_my_new_rule_fires_on_problem_formula(audit_ctx):
    audit_ctx.add_cell("S1", "A1", "=VoseSomethingBad(0)")
    findings = list(detect_my_new_rule(audit_ctx))
    assert len(findings) == 1
    assert findings[0].rule_id == "VOSE-XXX"

def test_my_new_rule_silent_on_good_formula(audit_ctx):
    audit_ctx.add_cell("S1", "A1", "=VoseSomethingGood(0)")
    findings = list(detect_my_new_rule(audit_ctx))
    assert findings == []
```

If the rule's detection involves numeric thresholds (like VOSE-011's `sigma > mu/2`), include boundary cases.

## Worked example — VOSE-007 (`risk_event_degenerate_probability`)

YAML:
```yaml
- id: VOSE-007
  name: risk_event_degenerate_probability
  severity: warning
  enabled: true
  description: >
    A VoseRiskEvent(probability, ...) has a literal probability of 0
    or 1. p=0 means the event never fires; p=1 means it always fires.
  suggested_fix_template: >
    For p=0: delete the VoseRiskEvent wrapper. For p=1: replace with
    the impact distribution directly. For uncertain p, use Beta.
```

Detector:
```python
_VOSE_RISK_EVENT_PROB_RE = re.compile(
    r'VoseRiskEvent\(\s*([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)\s*[,)]'
)

def detect_risk_event_degenerate_probability(ctx):
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
            severity=ctx.rule.severity,
            cell=cell.ref,
            rule_id=ctx.rule.id,
            message=f"Cell {cell.ref.a1} has VoseRiskEvent with probability = {int(prob)}.",
            suggested_fix=_format_fix(ctx.rule.suggested_fix_template),
        )
```

The detector:
1. Cheap pre-filter on the substring `"VoseRiskEvent"` before running regex.
2. Regex captures only literal numeric values as the first arg — references like `=VoseRiskEvent(A1, VoseNormal(...))` don't match (that's correct; we can't know whether A1 is 0 or 1).
3. `try/except ValueError` covers exotic float formats the regex permits but Python's `float()` rejects.
4. Only fires for the two degenerate cases `0` and `1`.

## When NOT to add a rule

- **Style preferences** (e.g. "use parentheses around exponents") — those belong in a separate linter, not the methodology audit.
- **Things only fixable by re-architecting the workbook.** Rules should flag local mistakes the LLM or user can fix one cell at a time.
- **Anything that requires running a simulation.** Audits are static-analysis; if you need iteration data, the right tool is `get_sensitivity_ranking` or a custom analysis on `get_samples` output, not an audit rule.
