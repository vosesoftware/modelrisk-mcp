# Modeling Patterns

Reusable *techniques* for composing a sound Monte Carlo model. Where the [scenarios](scenarios.md) are problem-shaped ("I'm modelling cost overrun"), these are technique-shaped ("how do I model a thing that happens a random number of times?"). Each pattern names the ModelRisk tool that implements it and the methodology principle behind it.

| Pattern | Use when | Tool |
|---|---|---|
| [Frequency–severity](#frequencyseverity-aggregation) | a random *number* of events, each a random *size* | `create_aggregate_mc` |
| [Risk events](#risk-events-fires-or-it-doesnt) | a discrete event with a consequence | `create_risk_event` |
| [Correlated inputs](#correlated-inputs) | inputs that move together in reality | `create_copula` |
| [Common random numbers](#common-random-numbers-paired-draws) | comparing two designs fairly | shared input cell |
| [Time-series choice](#choosing-a-time-series-process) | a quantity with memory across periods | `create_time_series` |
| [Iteration count](#how-many-iterations) | deciding how long to run | `run_simulation` |

---

## Frequency–severity aggregation

**Problem.** Total annual loss = (a random number of incidents) × (a random size per incident). Both the count and the size are uncertain.

**Pattern.** Model frequency and severity *separately*, then let an aggregate function convolve them:

- **Frequency** — usually `VosePoisson(mean)` (or `VoseNegBin` if over-dispersed).
- **Severity** — usually `VoseLognormal` / `VoseGamma` (positive, right-skewed).
- **Aggregate** — `create_aggregate_mc` builds `VoseAggregateMC(frequency, severity)`, which draws a fresh count each iteration and sums that many severity draws.

**Why not a fixed loop?** Multiplying "expected count × average severity", or summing a fixed number of severity cells, freezes the frequency and badly understates the spread of the annual total. The whole point is that *both* dimensions vary. (Methodology principle 4.)

**Where it shows up.** Operational risk, insurance claims, fraud, equipment failures, warranty costs — any "N things of size X per period." Walk-through: [scenario 3](scenarios.md#scenario-3--aggregate-operational-losses-frequency--severity).

---

## Risk events: "fires, or it doesn't"

**Problem.** A discrete event — a breach, a lawsuit, a permit rejection — that *might* happen, and *if* it does, costs an uncertain amount.

**Pattern.** `create_risk_event` → `VoseRiskEvent(p, impact)`. With probability `p` it returns a draw from the `impact` distribution; otherwise it returns 0.

**The anti-pattern to avoid.** `p × impact`. This reproduces the *mean* contribution but destroys the *shape*: it spreads a little bit of the loss across every iteration instead of "nothing most years, a large hit occasionally." Since risk events exist precisely to inform tail metrics (P95, P99, TVaR), smearing them makes the model wrong exactly where it matters. (Methodology principle 3; degenerate `p`=0/1 caught by audit rule VOSE-007.)

**Tip.** If the probability itself is uncertain, don't hard-code it — drive `p` from a `VoseBeta(α, β)` input so the model carries your uncertainty about the likelihood, not just the impact.

---

## Correlated inputs

**Problem.** Two or more inputs move together in the real world — price and demand, default rates across sectors, schedule and budget. Modelling them independently is not conservative; it's wrong.

**Pattern.** Set each input's marginal distribution first, then link them with a copula via `create_copula`:

- **Gaussian copula** — simple linear-ish association; the default.
- **Clayton / Gumbel** — when dependence is stronger in one tail (e.g. assets that crash together but rise independently — lower-tail dependence → Clayton).
- **t-copula** — symmetric tail dependence, heavier than Gaussian.

**Verify it landed.** After the run, `get_correlation_matrix` should show the dependence you intended. If it doesn't, the copula didn't wire correctly.

**Why it matters.** Independent draws of correlated quantities cancel out and produce artificially tight outputs — the model claims the portfolio is safer than it is, because it assumes the bad cases never coincide. Sometimes correlation barely moves the answer; sometimes it dominates. The only way to know is to model it and compare. (Methodology principle 5.) Walk-through: [scenario 4](scenarios.md#scenario-4--model-correlated-inputs).

---

## Common random numbers (paired draws)

**Problem.** You're comparing two designs — Plan A vs Plan B, vendor X vs vendor Y — and want the comparison to reflect the *difference between the plans*, not the noise of two separate simulations.

**Pattern.** Drive both alternatives from the **same** uncertain input cells. Sample demand, price, etc. once per iteration; feed that single draw into both Plan A's and Plan B's calculations. The difference `A − B` is then computed on identical scenarios each iteration, so its distribution isolates the real effect of the choice.

**Implementation.** Put the shared distribution in one input cell and reference it from both branches' formulas — don't create two independent `VoseNormal(...)` cells with the same parameters (those would draw *different* values each iteration and inject comparison noise). One Object / one input, many downstream references.

**Why it matters.** With independent draws, the variance of `A − B` is inflated by the variance of two separate sampling streams, and you may need vastly more iterations to detect a real difference. Pairing cancels the common noise — the same trick as a paired t-test. This is a variance-reduction technique, not just tidiness.

---

## Choosing a time-series process

**Problem.** A quantity with *memory*: its value next period depends on this period. A single distribution per period throws that structure away.

**Pattern.** `create_time_series`, picking the process that matches the quantity's behaviour:

| Behaviour | Process |
|---|---|
| Drifts, never negative, multiplicative (prices, FX) | Geometric Brownian Motion (`VoseTimeGBM`) |
| Pulls back toward a long-run level (commodities, rates, occupancy) | Mean-reverting (Ornstein–Uhlenbeck) |
| This period ≈ last period + shock (many economic series) | AR(1) |
| Occasional discrete jumps on top of a drift | Jump-diffusion |

**Key question to ask first:** *does this quantity revert to a level, or wander freely?* Reversion vs random-walk is the fork that picks the family. Getting it wrong produces price paths that either fan out far too wide (using GBM for something that actually reverts) or never explore the range (using reversion for something that genuinely drifts). Walk-through: [scenario 2, step 3](scenarios.md#scenario-2--fit-distributions-to-historical-data).

---

## How many iterations?

**Problem.** Too few iterations and your percentiles are noisy; too many and you wait longer than you need.

**Rules of thumb:**

- **Central statistics** (mean, P50) stabilise quickly — a few thousand iterations.
- **Tail percentiles** (P95, P99) need far more — a P99 is estimated from the worst 1% of iterations, so 10,000 iterations gives only ~100 samples in that tail. For decision-critical tail metrics, run 50,000+.
- **Convergence check** — run once, note the metric you care about; run again with more iterations; if it moved materially, you didn't have enough. ModelRisk also supports convergence monitoring.
- **Reproducibility** — set a fixed seed (`run_simulation` accepts one) when you need bit-for-bit repeatable results for an audit trail.

Practical default: **10,000** for exploration and central estimates; **50,000+** when a tail metric (VaR, TVaR, P99) drives the decision.

---

## See also

- [Distribution selection](distribution-selection.md) — which family for which input
- [Scenarios](scenarios.md) — these patterns composed into end-to-end recipes
- [Methodology](methodology.md) — the principles behind the patterns
