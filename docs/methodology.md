# Monte Carlo Methodology

The principles behind every model this server builds. This is the human-readable companion to the `modelrisk://methodology` resource (which Claude loads into context at build/audit time) — same principles, expanded for people, with the reasoning at more length.

The guiding idea: **a Monte Carlo model is only as honest as its weakest assumption.** Each principle below closes off one common way to fool yourself, and each is paired with the audit rule that catches it — so the methodology isn't just advice, it's enforced by `audit_model`.

> **On provenance.** The principles here are established, broadly-held risk-analysis practice — the kind found in standard references on quantitative risk modelling. They are stated conservatively and are not proprietary claims. Where a principle could be deepened with material from Vose Software's own published canon (the Vose book, *Risk Analysis: A Quantitative Guide*, and the ModelRisk documentation), that is noted and left for sourced expansion rather than invented here. See [Sources & further reading](#sources--further-reading).

---

## The eight core principles

### 1. Express uncertainty, not a point estimate

Every input that could plausibly take a range of values should be a distribution, not a fixed number. A cell left as a single "best guess" contributes zero variance — the simulation treats it as known with certainty, and the output's range is understated by exactly what that input could have swung. A point estimate hides the one thing a decision-maker needs: how wrong it could be.
*Enforced by VOSE-006 (hard-coded inputs), VOSE-010 (an input tagged but left constant).*

### 2. Carry parameter uncertainty

When you fit a distribution to data, fit *with* `uncertainty=TRUE`. A fit to a finite sample gives you a best *estimate* of the parameters — not their true values. Sampling only from the best-fit parameters pretends you know them exactly; you don't, and the smaller your sample, the larger that second layer of uncertainty. Omitting it makes the model over-confident and the tails too thin.
*Enforced by VOSE-003 (fit without uncertainty).*

### 3. Model discrete events as events

A "might happen" event — a breach, a lawsuit, a failure — is bimodal: most of the time it contributes nothing; occasionally it fires with its full impact. Model it as `VoseRiskEvent(p, impact)`, not `p × impact`. The multiplied version matches the *mean* but destroys the *shape*, smearing a little loss across every iteration instead of "usually nothing, rarely large." Every tail metric — the metric a risk event exists to inform — comes out wrong.
*Enforced by VOSE-007 (degenerate probability).* See also [modeling patterns: risk events](modeling-patterns.md#risk-events-fires-or-it-doesnt).

### 4. Aggregate frequency and severity properly

When a total is a random *number* of random *sizes* — annual loss as incidents × per-incident cost — use an aggregate function (`VoseAggregateMC`), not a fixed-length sum. The number of events is itself uncertain; a fixed loop freezes it and understates the spread of the total. This is principle 1 one level up.
*Apply by judgement — not yet auto-audited.* See [modeling patterns: frequency–severity](modeling-patterns.md#frequencyseverity-aggregation).

### 5. Correlate what's correlated

Inputs that move together in reality must be linked in the model, via a copula. Independent draws of correlated quantities cancel out and produce artificially tight outputs — the model claims things are safer than they are because it assumes the bad cases never coincide. Whether dependence matters varies by model; the only way to know is to include it and compare.
*Apply by judgement — confirm with `get_correlation_matrix` after the run.* See [modeling patterns: correlated inputs](modeling-patterns.md#correlated-inputs).

### 6. Mark your outputs

Only cells wrapped with `VoseOutput("name")` are recorded across iterations with history, percentiles, and sensitivity. An unmarked cell still recalculates every iteration, but its values are discarded — you run 10,000 iterations and then can't read the distribution of the very number you cared about. Give each a clear, unique name.
*Enforced by VOSE-008 (unnamed), VOSE-009 (duplicate names), VOSE-004 (output that depends on no random input).*

### 7. Mark your inputs

`VoseInput("name")` makes a cell trackable for correlation and tornado analysis. An unmarked distribution still varies, but the Results Viewer can't attribute output variance to it — your tornado is missing drivers and you can't see which assumption to tighten first.
*Enforced by VOSE-002 (bare distribution, no wrapper), VOSE-005 (arithmetic outside the wrapper).*

### 8. Don't simulate constants

If a value genuinely can't vary — days per year, a unit conversion — keep it a plain number. Wrapping a true constant as an input wastes samples and clutters the inputs list and tornado with a "driver" that contributes nothing, making the assumptions that *do* matter harder to see.
*Enforced by VOSE-010, read from the other direction: a wrapper around a constant.*

---

## Beyond the principles: what else the audit checks

Three classes of problem are about mechanical correctness or distribution choice rather than the principles above, but the audit catches them too:

- **Distribution selection** — a `VoseNormal` with a positive mean and large spread sends mass below zero; wrong for a price or quantity. *(VOSE-011. See the [distribution selection guide](distribution-selection.md).)*
- **Formula correctness** — typo'd function names, wrong argument counts. A model that won't calculate isn't a model. *(VOSE-001, VOSE-013.)*
- **Error propagation** — a cell resolving to `#DIV/0!` or `#REF!` inside a Vose call poisons every iteration. *(VOSE-012.)*

---

## How the methodology is enforced

| Surface | Role |
|---|---|
| `modelrisk://methodology` resource | loaded into Claude's context at `/build-risk-model` and `/audit-model` time — shapes what gets built |
| `audit_model` (13 rules) | checks a live workbook against the principles, cell by cell |
| `propose_distributions_for_inputs` + `modelrisk://distributions` | steers family selection toward sound choices |
| the `/build-risk-model` prompt | walks a new model through the principles in order |

A drift-guard test (`test_methodology_crossref.py`) keeps the principle ↔ rule cross-references honest: every rule must be referenced, and every reference must point to a real rule.

---

## Sources & further reading

The principles above are conservative statements of established quantitative-risk practice. To deepen this into a fully *sourced* knowledge base — with specific guidance, thresholds, and worked justifications attributed to authoritative material — the natural sources are:

- **Vose, D. — *Risk Analysis: A Quantitative Guide*** (the standard reference; the methodology behind ModelRisk).
- **The ModelRisk documentation / help** — function-level guidance and worked examples.
- **Vose Software's published technical notes and consulting canon.**

> *This section is a placeholder for sourced expansion. Content attributed to these works should be added with citations rather than paraphrased from memory — see the project maintainers before extending.*

---

## See also

- [Distribution selection](distribution-selection.md) — which family for which input
- [Modeling patterns](modeling-patterns.md) — techniques for composing sound structures
- [Scenarios](scenarios.md) — the principles applied end-to-end
- [Authoring audit rules](authoring-audit-rules.md) — add your own enforced principles
