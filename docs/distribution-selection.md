# Distribution Selection Guide

Which distribution for which uncertain quantity. This is the human-readable companion to the machine-readable selector in `src/modelrisk_mcp/data/distributions.yaml` — the same logic that `propose_distributions_for_inputs` uses when Claude suggests a family for one of your inputs.

You rarely need to consult this directly: describe the quantity to Claude ("lead time in days, usually about 10, occasionally 30") and it proposes a family with a rationale. This guide is for when you want to understand or override that choice.

> **One rule above all the others:** a distribution must respect the *bounds* of the real quantity. A price, a quantity, a duration cannot be negative — so don't model it with a distribution that puts mass below zero (the single most common mistake, caught by audit rule **VOSE-011**). Everything below is a refinement of that idea.

---

## Quick reference

| You're modelling… | Start with | Then consider |
|---|---|---|
| A subjective estimate (min / most-likely / max) | `VoseModPERT` | `VosePERT`, `VoseTriangle` |
| A count of events | `VosePoisson` | `VoseNegBin`, `VoseBinomial` |
| A yes/no event | `VoseBernoulli` | `VoseRiskEvent` (if it has an impact) |
| A duration / cost / positive quantity | `VoseLognormal` | `VoseGamma`, `VoseWeibull` |
| A price / rate over time | `VoseTimeGBM` | mean-reverting / AR time series |
| Anything you have data for | fit it — `Vose*Fit(…, uncertainty=TRUE)` | compare families on fit statistics |

---

## By scenario

### Subjective three-point estimate

You have expert judgement, not data: "best case 800, most likely 1,000, worst case 1,400."

- **`VoseModPERT(min, mode, max)`** — the default. A smooth, single-peaked shape that weights the most-likely value heavily and the extremes lightly. The "Mod" gamma parameter (default 4) sharpens the peak relative to classic PERT — usually what you want for an expert estimate.
- **`VosePERT(min, mode, max)`** — the textbook beta-PERT, slightly broader peak. Use if you specifically want the classic shape.
- **`VoseTriangle(min, mode, max)`** — simplest, linear sides. Use *only* when you genuinely have no prior beyond the three points; its straight edges overweight the extremes compared with PERT, which usually overstates tail risk.

**Why not just a normal?** A normal is symmetric and unbounded — it ignores your min/max and leaks probability past both ends. Three-point estimates are almost always asymmetric and bounded; PERT respects that.

### Count of events

"How many incidents next year?" "How many defects per batch?"

- **`VosePoisson(mean)`** — the workhorse. Independent events arriving at a constant mean rate. One parameter.
- **`VoseNegBin(...)`** — when counts are *over-dispersed* (variance noticeably exceeds the mean — clustering, contagion, heterogeneity). Poisson forces variance = mean; real count data often violates that.
- **`VoseBinomial(n, p)`** — when there's a known maximum number of trials `n`, each succeeding with probability `p` (e.g. "of 500 contracts, how many default?").

### Binary / yes-no event

"Does the customer renew?" "Does the plant trip?"

- **`VoseBernoulli(p)`** — a single yes/no trial. Returns 0 or 1.
- **`VoseRiskEvent(p, impact)`** — when the event *has a consequence*. This is the right structure for risk: with probability `p` it fires and you draw `impact`; otherwise it contributes nothing. Don't model this as `p × impact` — see [modeling patterns](modeling-patterns.md#risk-events-fires-or-it-doesnt) and methodology principle 3.

### Duration, cost, or any positive-only quantity

Lead times, project costs, repair times, demand — quantities that are naturally positive and usually right-skewed (a long upper tail).

- **`VoseLognormal(mean, stdev)`** — the standard. Bounded at zero, right-skewed, multiplicatively natural. The first choice for most positive quantities.
- **`VoseGamma(...)`** — a flexible right-skewed alternative when your priors aren't naturally on a log scale.
- **`VoseWeibull(...)`** — time-to-failure / lifetime modelling, when there's a known hazard pattern (wear-in, wear-out).

**The VOSE-011 trap:** it's tempting to reach for `VoseNormal(mean, stdev)` because you know the mean and stdev. But if `stdev > mean/2`, a normal sends ~16% of its draws below zero — nonsensical for a cost or quantity, and it quietly biases the result. Use lognormal/gamma, or `VoseNormal` truncated at zero, instead.

### Price or rate evolving over time

Stock prices, FX rates, commodity prices, interest rates — quantities with *memory* across periods.

- **`VoseTimeGBM(...)`** — Geometric Brownian Motion. The standard multiplicative random walk for prices and rates that drift and never go negative.
- **Mean-reverting / AR processes** — when the quantity pulls back toward a long-run level (commodity prices, interest rates, occupancy). See [modeling patterns: time-series choice](modeling-patterns.md#choosing-a-time-series-process). Built via `create_time_series`.

A single distribution per period would throw away the period-to-period correlation that defines these series — use a time-series function, not 24 independent draws.

### When you have data

Always prefer fitting over guessing when you have a real sample.

- **`Vose*Fit(data, uncertainty=TRUE)`** — fit the family to the data. **Always pass `uncertainty=TRUE`** so parameter uncertainty is carried through the simulation (methodology principle 2; audit rule VOSE-003). `fit_distribution_to_data` does this for you.
- **Compare families** when the data is ambiguous — `VoseLognormalFit` vs `VoseGammaFit` vs `VoseWeibullFit` — and choose on goodness-of-fit, not just visual shape.
- **Don't over-fit a tiny sample.** With few data points the parameter uncertainty dominates; that's a feature (`uncertainty=TRUE` shows it), not a bug to suppress.

---

## When nothing fits cleanly

If you can't characterise the quantity, the safe fallback is a `VoseModPERT` three-point estimate — ask the modeller for a min, most-likely, and max. It's better to capture rough bounds than to leave the input deterministic.

---

## See also

- [Modeling patterns](modeling-patterns.md) — how to compose these into frequency-severity, correlated, and time-series structures
- [Methodology](methodology.md) — the principles behind the choices
- `modelrisk://distributions` — the machine-readable selector (fetch from any MCP client)
- `modelrisk://functions` — the full 1,417-entry function reference
