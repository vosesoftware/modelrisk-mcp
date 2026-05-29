# Walk-through Scenarios

Six end-to-end recipes, each for a different kind of risk problem. Find the one closest to yours and follow it — every step shows the exact prompt to type and what Claude does behind it.

New to the toolchain? Do the [15-minute quick-start](quick-start.md) first; it teaches the basic build → run → read loop on a simple NPV model. Unfamiliar with a term? See the [glossary](glossary.md).

| # | Scenario | For | Showcases |
|---|---|---|---|
| 1 | [Add uncertainty to a budget](#scenario-1--add-uncertainty-to-a-deterministic-budget) | FP&A / finance | convert deterministic → probabilistic |
| 2 | [Fit distributions to historical data](#scenario-2--fit-distributions-to-historical-data) | Analyst with data | fitting, parameter uncertainty, time-series |
| 3 | [Aggregate operational losses](#scenario-3--aggregate-operational-losses-frequency--severity) | Risk / insurance | aggregates, risk events, tail metrics |
| 4 | [Model correlated inputs](#scenario-4--model-correlated-inputs) | Finance / markets | copulas, correlation checks |
| 5 | [Stress-test with scenarios](#scenario-5--stress-test-with-scenarios) | Decision-makers | scenario sweeps, comparison |
| 6 | [Audit an inherited model](#scenario-6--audit-an-inherited-model) | Reviewers / auditors | audit rules, guided fixes |

Each scenario is self-contained. They share one convention: **prompts you type are shown as block quotes**, and the tools Claude fires are named so you can follow the mechanics.

A reminder that applies to every scenario: **every building tool previews first** (`dry_run=True`). Claude shows you what it would write; you confirm before it commits. `Ctrl+Z` in Excel undoes any write.

---

## Scenario 1 — Add uncertainty to a deterministic budget

**You have:** an Excel budget or forecast built from point estimates — "Q3 revenue = $4.2M", "unit cost = $25". One number per assumption, no sense of the range.

**You want:** a confidence range around the bottom line, and to know which assumption drives it.

### Steps

**1. Let Claude find the candidates.**

> Look at the active workbook. Find the hard-coded numbers that look like model inputs and tell me which ones are worth making uncertain.

Claude calls `find_hard_coded_inputs` (round numbers referenced by formulas) and `get_workbook_summary`, then lists candidates with its reasoning — a "Discount rate" cell gets treated differently from a "Units sold" cell.

**2. Give ranges for the ones that matter.**

> Make these three uncertain:
> - Units sold (B4): low 800, likely 1,000, high 1,400
> - Unit cost (B5): low $22, likely $25, high $32
> - Churn rate (B9): normal, mean 8%, stdev 2%

Claude calls `propose_distributions_for_inputs` (it'll suggest PERT for the three-point estimates, normal for churn), previews each formula, then `replace_constant_with_distribution` once per cell after you approve.

**3. Mark the output.**

> Wrap the bottom-line cell B20 as an output called "Net margin".

`wrap_with_output`.

**4. Run and read.**

> Run 10,000 iterations, then give me P10 / P50 / P90 of net margin and the top three drivers.

`run_simulation` → `get_simulation_results` + `get_sensitivity_ranking`.

**5. Report.**

> Build the executive report.

`build_executive_report` writes a styled sheet — histogram, percentiles, tornado, narrative.

### What you learn
Where your single-point budget actually sits as a range, and which assumption to refine first (the top of the tornado) if you want to tighten it.

**Shortcut:** the whole flow is wrapped in the `/add-uncertainty` slash command.

---

## Scenario 2 — Fit distributions to historical data

**You have:** actual history — weekly sales, monthly defect counts, daily prices — in a column somewhere.

**You want:** distributions grounded in that data rather than guessed three-point estimates, and a price that evolves over time.

### Steps

**1. Fit a distribution to the data.**

> Column D on the Data sheet has 3 years of weekly demand. Fit a distribution and use it for the demand input in cell B6.

Claude calls `fit_distribution_to_data`, which uses ModelRisk's own fitting routines. It returns the best-fit family **with `uncertainty=TRUE`** — so the simulation samples through parameter uncertainty, not just the best-fit point. (See the glossary on [parameter uncertainty](glossary.md#parameter-uncertainty) for why this matters.)

**2. Sanity-check the fit.**

> What family did it pick, and how good is the fit? Should I be worried about the tails?

Claude explains the fitted family and its parameters and flags tail behaviour.

**3. Model a price as a time series.**

> The wholesale price in row 12 should evolve over 24 months as a mean-reverting process — long-run level $80, current $72.

`create_time_series` builds the mean-reverting (Ornstein–Uhlenbeck) process across the 24 cells.

**4. Wire, run, interpret.**

> Wrap the gross-profit cell as an output, run 20,000 iterations, and interpret the results for me.

`wrap_with_output` → `run_simulation` → `get_simulation_results`, then `generate_executive_summary` for a narrative.

### What you learn
A model whose uncertainty is **earned from data**, not asserted — and a price path that behaves like a real commodity instead of an independent draw each period.

---

## Scenario 3 — Aggregate operational losses (frequency × severity)

**You have:** a low-frequency / high-severity exposure — cyber incidents, equipment failures, fraud, large claims. You know roughly how often and roughly how big, but not the annual total.

**You want:** the distribution of *total annual loss*, and the tail (how bad does a bad year get?).

### Steps

**1. Build the aggregate.**

> Build an annual cyber-loss model. Frequency: Poisson, mean 4 incidents/year. Severity per incident: lognormal, mean $250k, stdev $400k. Put the total in cell B3.

Claude calls `create_aggregate_mc` — this is a true compound distribution (`VoseAggregateMC`), not "4 × average severity". The number of incidents is itself random each iteration. (Glossary: [aggregate](glossary.md#aggregate--compound-distribution), [frequency × severity](glossary.md#frequency--severity).)

**2. Add a rare catastrophic event on top.**

> Add a separate tail risk: a major breach with 5% annual probability, impact PERT $5M / $15M / $40M. Add it to the total.

`create_risk_event` — `VoseRiskEvent(0.05, VosePERT(...))`. It fires or it doesn't; when it fires you draw the impact. (Glossary: [risk event](glossary.md#risk-event).)

**3. Run and pull the tail.**

> Wrap the total as an output, run 50,000 iterations. Give me the mean, the P95, the P99, and the average loss in the worst 1% of years.

`run_simulation` → `get_simulation_results` for percentiles, then `get_samples` so Claude can compute the **TVaR(99)** (average beyond P99) directly from the raw samples. (Glossary: [VaR / TVaR](glossary.md#var--tvar--value-at-risk--tail-value-at-risk).)

### What you learn
Not just the expected annual loss, but the shape of a bad year — the number that sets your capital reserve or insurance limit. The aggregate + risk-event combination is the foundation of operational-risk and insurance modelling.

---

## Scenario 4 — Model correlated inputs

**You have:** inputs that move together in real life — demand and price, default rates across sectors, exchange rate and import cost. Modelling them as independent produces artificially tight outputs and understates risk.

**You want:** a dependency structure that ties them together, and a check that it took.

### Steps

**1. Set up the marginal distributions first.**

> Make demand (B4) a lognormal mean 1,000 stdev 200, and wholesale price (B5) a PERT $18 / $22 / $30. Wrap both as inputs.

`replace_constant_with_distribution` ×2, `wrap_with_input` ×2.

**2. Link them with a copula.**

> Demand and price are negatively correlated — when price rises, demand falls. Link them with a copula at about −0.5.

Claude calls `create_copula` (it'll pick a family — Gaussian for a simple linear association, or a Clayton/Gumbel if you want tail dependence) and wires both inputs through it. (Glossary: [copula](glossary.md#copula).)

**3. Verify the correlation actually landed.**

> Run 10,000 iterations and show me the correlation matrix for the inputs and the output.

`run_simulation` → `get_correlation_matrix`. You should see the demand–price correlation come back near −0.5; if it doesn't, the copula didn't wire correctly and Claude can diagnose.

**4. Compare against the independent case.**

> Now temporarily remove the copula, re-run, and tell me how much wider the output range gets.

Shows the concrete cost of ignoring correlation — usually a materially wider (or narrower) P10–P90 on the output.

### What you learn
Whether correlation matters for *your* model (sometimes it barely moves the answer; sometimes it dominates) — and the confidence that the dependency you intended is actually in the simulation.

---

## Scenario 5 — Stress-test with scenarios

**You have:** a working probabilistic model and a decision to defend — "what if the discount rate is capped at 8%?", "what if the contract doesn't renew?".

**You want:** to see how the output distribution shifts as you fix one or more inputs at specific values, side by side.

### Steps

**1. Sweep a single driver.**

> Run scenarios sweeping the discount rate across 6%, 8%, 10%, 12%, and 14%. For each, give me the P50 and P90 of NPV.

Claude calls `run_scenarios` — it fixes the discount-rate input at each value, runs a simulation per setting, and returns the comparative percentiles in one table. (Glossary: [scenario](glossary.md#scenario).)

**2. Test a discrete what-if.**

> Now a binary scenario: Tier-1 customer renews vs doesn't. Compare the NPV distribution under each.

`run_scenarios` with the renewal flag fixed at 1 then 0.

**3. Combine and read.**

> Which scenario gives the best downside protection — highest P10 — and what does it cost in expected value?

Claude reads across the scenario results and frames the trade-off (downside floor vs mean).

**4. Capture it.**

> Summarise the scenario comparison as a markdown table I can paste into the board deck.

`generate_executive_summary` scoped to the scenario set.

### What you learn
How sensitive the decision is to the lever you control, expressed as a comparison a non-modeller can read — the difference between "NPV is ~$42M" and "capping the rate at 8% raises the P10 floor by $3M at a $0.4M expected-value cost".

---

## Scenario 6 — Audit an inherited model

**You have:** someone else's ModelRisk workbook (or your own, grown messy). Before you trust its numbers, you want to know what's wrong with it.

**You want:** a methodology review — the silent mistakes that make a model run fine but produce misleading results.

### Steps

**1. Run the audit.**

> Audit the active workbook and rank the findings by severity.

Claude calls `audit_model`, which runs all 13 rules (VOSE-001 … VOSE-013). Each finding names the cell, the rule, a one-line why, and a suggested fix. Severity is `error` / `warning` / `info`.

**2. Triage the errors first.**

> Walk me through the errors only. For each, explain the impact in plain terms.

Typical high-value catches:
- **VOSE-003** — a `Fit()` without `uncertainty=TRUE` (understates uncertainty silently)
- **VOSE-004** — a `VoseOutput` that doesn't depend on any random input (deterministic "output")
- **VOSE-008/009** — unnamed or duplicate output names (breaks result lookup)
- **VOSE-012** — a cell evaluating to `#DIV/0!` inside a Vose call (poisons every iteration)
- **VOSE-013** — a distribution called with the wrong number of arguments

**3. Fix them, guided.**

> Fix the VOSE-003 finding on the demand fit, and rename the duplicate "Output" cells to something meaningful.

Claude re-issues the fit with `uncertainty=TRUE` and uses `wrap_with_output` to rename — each as a previewed write you approve.

**4. Re-audit to confirm.**

> Re-run the audit and confirm the errors are cleared.

`audit_model` again — you should see the error count drop. Remaining `info`/`warning` items are judgement calls you can accept or address.

### What you learn
Whether the model is sound enough to base a decision on — and a clean bill of methodology health (or a punch-list) you can hand back to whoever built it.

**Shortcut:** `/audit-model` runs this as a guided workflow.

---

## Where to go next

- **[User manual](user-manual.md)** — the eight capabilities in depth, plus what the server does and doesn't do
- **[Quick-start](quick-start.md)** — the 15-minute first-simulation tutorial
- **[Glossary](glossary.md)** — Monte Carlo + MCP vocabulary
- **`modelrisk://methodology`** — fetch the eight Vose methodology principles from any MCP client

Have a problem that doesn't fit any of these six? Describe it to Claude directly — "I'm trying to model X, the uncertain parts are Y and Z" — and it'll propose a structure. The scenarios are starting points, not a closed set.
