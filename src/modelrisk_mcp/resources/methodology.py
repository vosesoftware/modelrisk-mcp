"""Vose methodology principles resource."""

from __future__ import annotations

from modelrisk_mcp.server import mcp

_METHODOLOGY: str = """\
# Vose ModelRisk methodology — core principles

These principles guide every decision the LLM makes when building or
auditing a risk model with this MCP server. Each is paired with the
audit rule(s) that enforce it, so a principle is never just advice —
it's a check `audit_model` runs against the live workbook. Rule ids
(`VOSE-0NN`) are defined in `data/audit_rules.yaml`.

## 1. The model expresses uncertainty, not just a point estimate

Every input that could plausibly take a range of values *must* be
modelled as a distribution, not a single number.

**Why.** A Monte Carlo result is only as honest as its inputs. A cell
left as a fixed number contributes zero variance — the simulation
treats it as known with certainty.

**Failure mode.** Total output uncertainty is understated by exactly
the amount the frozen input could have swung. The decision-maker sees
a falsely narrow range and over-trusts it.

**Enforced by.** VOSE-006 (hard_coded_inputs_present) flags round
numbers that formulas depend on — candidates for a distribution.
VOSE-010 (input_wrapper_without_distribution) catches a cell tagged
as an input that still holds a constant, so it never varies.

## 2. Distributions reflect *real* uncertainty about parameters

When fitting a distribution to data, set `uncertainty=TRUE`.

**Why.** A fit to a finite sample gives you a *best estimate* of the
parameters, not their true values. Sampling only from the best-fit
parameters pretends you know them exactly — you don't.

**Failure mode.** The simulation is over-confident: it captures the
natural variability of the quantity but discards the (often larger)
uncertainty about the distribution's own parameters. Tails come out
too thin.

**Enforced by.** VOSE-003 (fit_without_uncertainty) flags any
`Vose*Fit(...)` call that omits `uncertainty=TRUE`.

## 3. Discrete events use VoseRiskEvent

For "the bad thing might happen" scenarios, use `VoseRiskEvent(prob,
impact)` — not `probability * impact`.

**Why.** A risk event is bimodal: most iterations it contributes
nothing; occasionally it fires and contributes the full impact draw.
Multiplying probability by impact smears that lump across every
iteration.

**Failure mode.** The smeared version reproduces the *mean* but
destroys the *shape* — it shows a small loss every year instead of
"usually nothing, rarely catastrophic." Every tail metric (P95, P99,
TVaR) is wrong, which is exactly the metric a risk event exists to
inform.

**Enforced by.** VOSE-007 (risk_event_degenerate_probability) flags a
`VoseRiskEvent` with a literal probability of 0 (never fires — delete
it) or 1 (always fires — use the impact directly).

## 4. Aggregates use VoseAggregateMC

For sums of a (possibly random) number of i.i.d. severity draws — e.g.
total annual loss = number-of-incidents x per-incident-loss — use
`VoseAggregateMC`, not a hand-rolled loop.

**Why.** The count of events is itself uncertain. The aggregate
functions perform the frequency-severity convolution correctly; a
fixed-length sum of severities silently assumes the frequency is
known.

**Failure mode.** Hand-rolling with a fixed number of severity cells
freezes the frequency and understates the spread of the annual total
— the same error as principle 1, one level up.

**Enforced by.** No automatic rule — apply by judgement. (Audit
coverage of aggregate structure is candidate future work.)

## 5. Correlate inputs that are correlated in the real world

Two inputs that move together in reality (e.g. unit cost and demand
during a recession) must be correlated in the model via a copula
(`VoseCopulaMultiNormal` or similar).

**Why.** Independent draws cancel out; correlated draws reinforce.
Whether dependence is present materially changes the spread — and
sometimes the sign — of the output.

**Failure mode.** Treating correlated inputs as independent produces
artificially tight outputs: the model says the portfolio is safer
than it is, because it assumes the bad cases won't coincide.

**Enforced by.** No automatic rule — apply by judgement. Use
`get_correlation_matrix` after a run to confirm the dependence you
intended actually landed.

## 6. Output cells are marked with VoseOutput

Only cells wrapped with `VoseOutput("name")` appear in the Results
Viewer with per-iteration history, percentiles, and sensitivity
analysis.

**Why.** The simulation only records the cells you mark. An unmarked
cell still recalculates every iteration, but its values are thrown
away — there's nothing to interrogate afterwards.

**Failure mode.** You run 10,000 iterations and then can't read the
distribution of the very number you cared about, because it was never
recorded.

**Enforced by.** VOSE-008 (voseoutput_missing_name) flags an unnamed
output; VOSE-009 (duplicate_output_names) flags two outputs sharing a
name (the by-name result lookup becomes ambiguous); VOSE-004
(output_cell_no_distribution_reference) flags an output that doesn't
depend on any random input — it will be constant across iterations,
which is almost never intended.

## 7. Input cells are marked with VoseInput

Symmetric to outputs. `VoseInput("name")` makes the cell trackable for
correlation and tornado analysis.

**Why.** Sensitivity ranking and correlation require named inputs to
attribute output variance to. An unmarked distribution still varies,
but the Results Viewer can't tell you how much it drives the outcome.

**Failure mode.** The tornado is missing drivers; you can't see which
assumption to refine first.

**Enforced by.** VOSE-002 (distribution_without_input_wrapper) flags a
bare distribution cell with no wrapper; VOSE-005
(arithmetic_before_input_wrapper) flags `=2*VoseNormal(...)`-style
cells where the arithmetic sits outside the wrapper, hiding the input
from the Results Viewer.

## 8. Don't simulate constants

If a value can't possibly vary across iterations (days-per-year, unit
conversion factors), keep it as a plain number.

**Why.** Wrapping a genuine constant as an input wastes samples and
clutters the Results Viewer and the tornado with a driver that
contributes zero variance.

**Failure mode.** Noise in the inputs list and sensitivity output;
harder to see the assumptions that actually matter.

**Enforced by.** VOSE-010 (input_wrapper_without_distribution) — the
same rule as principle 1, read from the other direction: a wrapper
around a constant.

---

## What the audit also checks (beyond the core principles)

Three classes of problem the audit catches that are about mechanical
correctness or distribution choice rather than the principles above:

- **Distribution selection.** VOSE-011
  (high_volatility_normal_positive_mean): a `VoseNormal(mu, sigma)`
  with `mu > 0` and `sigma > mu/2` puts roughly 16% of its mass below
  zero — usually wrong for a price, quantity, or duration that can't
  be negative. Prefer `VoseLognormal` / `VoseGamma`, or a truncated
  normal. Family-selection guidance lives in the
  `modelrisk://distributions` resource.
- **Formula correctness.** VOSE-001 (unknown_vose_function — a typo'd
  `VoseNomral`) and VOSE-013 (arg_count_mismatch — e.g. `VosePERT`
  called with two args instead of three). A model that won't
  calculate is not a model.
- **Error propagation.** VOSE-012 (cell_evaluates_to_error): a cell
  resolving to `#DIV/0!`, `#REF!`, etc. — especially inside a Vose
  call — poisons every iteration of the run.

## Spreadsheet integrity (the SS-* rules)

A separate family from the Monte-Carlo-methodology rules above:
general spreadsheet hygiene, drawn from the established
spreadsheet-error / model-control discipline (O'Beirne, *Spreadsheet
Check and Control*; Rees, *Principles of Financial Modelling*; the
EuSpRIG literature). A model can be methodologically perfect and still
wrong if its deterministic scaffolding is broken.

- **SS-001 (magic_number_in_formula)** — a parameter-like constant
  (a decimal such as 1.21 or 0.85) buried in a formula instead of
  living in its own labelled input cell, where it would be visible,
  auditable, and changeable in one place.
- **SS-002 (number_stored_as_text)** — a numeric value held as text in
  a cell a formula references; text numbers are silently skipped by
  SUM and arithmetic, corrupting totals with no error shown.
- **SS-003 (overly_complex_formula)** — a single formula doing too
  much; dense formulas are the hardest to review and the easiest to
  get subtly wrong. Break the calculation into one step per cell.
- **SS-004 (inconsistent_formula_in_block)** — a cell in the middle of
  a filled run whose pattern differs from both neighbours when those
  neighbours agree; the classic "overtyped one cell after filling a
  row/column" error, and the most dangerous because it's invisible.
"""


@mcp.resource(
    uri="modelrisk://methodology",
    name="modelrisk-methodology",
    description=(
        "ModelRisk: the 8 core methodology principles the LLM follows "
        "when building, fitting, or auditing models. Loaded into "
        "context whenever the user starts a /build-risk-model or "
        "/audit-model conversation."
    ),
    mime_type="text/markdown",
)
def methodology_resource() -> str:
    return _METHODOLOGY
