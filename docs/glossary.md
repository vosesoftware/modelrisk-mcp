# ModelRisk MCP — Glossary

The vocabulary you'll bump into while using the server. Two camps: **Monte Carlo terms** (the domain language of probabilistic modelling) and **server terms** (the moving parts of ModelRisk MCP itself).

When in doubt, ask Claude — it understands all of these.

---

## Monte Carlo terms

### Aggregate / compound distribution

A combined distribution for "N events each of random size X". Used everywhere insurance and operational risk needs to express "we expect 5–20 incidents next year, each costing $10k–$2M". Built via `create_aggregate_mc` and the `VoseAggregateMC` family. Different from "sum of N samples from one distribution" — the *number of events* itself is random.

### CDF — cumulative distribution function

The probability that a random variable is less than or equal to a value. For a continuous distribution, the integral of the PDF from −∞ to that value. In a histogram report, the CDF line overlays the bars and grows from 0 to 1 across the range.

### Copula

A function that links the marginal distributions of two or more random variables into a joint distribution that preserves their *dependence structure*. Plain English: if your inputs are correlated (price and demand, default rates between sectors, project duration and budget), a copula models them as correlated *without forcing them to share a distribution family*. Families: Gaussian, Student t, Clayton, Frank, Gumbel. Built via `create_copula`.

### Correlation matrix

A grid of correlation coefficients between every pair of inputs (or every pair of outputs). Values run from −1 (perfectly inversely related) through 0 (no relationship) to +1 (perfectly correlated). Read via `get_correlation_matrix`. Not the same as causation — two inputs can correlate because both depend on a third.

### Distribution

A description of how likely each possible value of an uncertain quantity is. "Demand could be anywhere from 800 to 1200 units, most likely 1000" is a description of a distribution. Concretely: a function that says, for any value, how likely the variable is to take that value. Common families used in risk: Normal, Lognormal, PERT, Triangle, Beta, Gamma, Weibull, Exponential, Poisson, Discrete.

### Distribution fit

The process of finding which distribution family and which parameters best describe a sample of historical data. ModelRisk's `Fit` family (`VoseNormalFit`, `VoseLognormalFit`, etc.) takes a data range and returns a fitted distribution. **Always** pass `uncertainty=TRUE` — see Parameter uncertainty.

### Frequency × severity

The two-piece structure of an aggregate loss: how often something happens (frequency — often Poisson or Negative Binomial) times how much it costs each time (severity — often Lognormal, Gamma, or Weibull). Most insurance and operational-risk models reduce to this pattern. See Aggregate.

### Histogram

A bar chart showing how many simulation iterations fell into each value range for an output. The shape tells you whether the output is symmetric, skewed, bimodal, or fat-tailed. Generated automatically in `build_executive_report` and `build_drivers_report`.

### Iteration

One full recalculation of the workbook with a fresh random sample drawn from every input distribution. A 10,000-iteration simulation recalculates the workbook 10,000 times — each iteration produces one value per output.

### Methodology

Vose's set of opinionated rules for building defensible probabilistic models. Eight core principles, served at `modelrisk://methodology`. The audit ruleset (VOSE-001 through VOSE-013) is the operational enforcement of the methodology: it flags violations in your workbook.

### Mean / standard deviation / variance / skewness / kurtosis

The first four moments of an output's distribution:
- **Mean** — the average value
- **Standard deviation** — typical distance from the mean (square root of variance)
- **Variance** — the squared average distance from the mean
- **Skewness** — asymmetry; positive = long right tail, negative = long left tail
- **Kurtosis** — tail-heaviness; high = more extreme outliers than a normal would predict

All returned by `get_simulation_results`.

### Monte Carlo simulation

A method for analysing a model whose inputs are uncertain: instead of solving it once with point estimates, you solve it many times (10,000+) with random samples from the input distributions, then look at the *distribution of outputs*. Named after the casino — the technique was invented at Los Alamos in the 1940s for nuclear-weapon yield calculations.

### NPV / IRR / NPV at risk

Net Present Value and Internal Rate of Return. Standard finance outputs. "NPV at risk" — typically the P10 (10th percentile) of the NPV — is the value below which you have a 10% probability of landing. The whole point of risk-modelling NPV is replacing a single point estimate with this kind of range.

### Output

A cell whose value you want the simulation to track across iterations. Wrapped with `VoseOutput("Name")`. After the simulation, every result tool (`get_simulation_results`, percentiles, histograms, tornados) is keyed by output name.

### Parameter uncertainty

The recognition that even after fitting a distribution to data, the *parameters* of that distribution themselves are uncertain — you have a finite sample. A normal fit to 30 data points has a mean estimate, but the *true* mean could be a bit higher or lower. `uncertainty=TRUE` on a `Fit` call samples through both the natural variability AND the parameter uncertainty on each iteration. Without it, the simulation overstates how much you know about the input distribution.

### PERT distribution

A 3-parameter distribution defined by minimum, most-likely, and maximum. Skewed toward the most-likely value. Standard distribution for expert-elicited estimates ("if I had to give a range I'd say min $X, most-likely $Y, max $Z"). Built via `VosePERT`.

### Percentile (P10, P50, P90, …)

The value below which a given percentage of simulation iterations fell. P50 is the median (half above, half below). P10 is "10% of iterations were below this number" — a downside scenario. P90 is "90% were below" — an upside scenario. P10 / P50 / P90 together is the standard board-pack triplet.

### Risk event

A binary "fires-or-doesn't-fire" wrapper around an impact distribution. `VoseRiskEvent(probability, impact)`: on each iteration the event fires with probability `p`, and *if* it fires, samples the impact distribution. The bimodal result (often zero, sometimes a big value) is structurally different from "p × impact" — and risk events are how operational risk, project risk, and insurance models express low-frequency / high-severity threats. Built via `create_risk_event`.

### Scenario

A specific configuration of input values you want to run the simulation under. "Run the sim with discount rate fixed at 8%" or "Run the sim with no Tier-1 customer renewal". `run_scenarios` lets you sweep one or more inputs across a list of fixed values, getting back comparative percentiles for each. Useful for stress-tests and what-ifs.

### Sensitivity analysis / Tornado chart

Ranking inputs by how much each one contributes to the variance of an output. The result is usually visualised as a horizontal bar chart (a "tornado") with the biggest contributor at the top. Tells you where to focus model refinement — tighten your assumption about the input at the top first. Read via `get_sensitivity_ranking`; chart built via `create_tornado_chart`.

### Time series

A distribution whose value at one period depends on its value in the previous period — autoregressive (AR), geometric Brownian motion (GBM), mean-reverting, jump-diffusion. Used for prices, rates, populations, anything with momentum across periods. Built via `create_time_series`.

### Triangle / Triangular distribution

A 3-parameter distribution defined by minimum, most-likely, and maximum, like PERT but with sharper corners. Useful when you genuinely believe values near the extremes are roughly as likely as values near the most-likely point. Built via `VoseTriangle`.

### VaR / TVaR — Value at Risk / Tail Value at Risk

VaR(99) = the P99 of a loss distribution. TVaR(99) = the *average* loss conditional on the loss exceeding the P99. TVaR captures the tail; VaR only captures the cliff. Both terms come up in insurance / finance risk reporting; both are derivable from `get_samples` for any output.

### Variance / variance contribution

A measure of spread (the squared mean distance from the mean). "Variance contribution" or "share of variance" is the standard sensitivity output: the percentage of an output's variance attributable to one specific input. The numbers in a tornado chart.

### .vmrs file

ModelRisk's results file format. Contains every per-iteration sample for every output, plus computed statistics. Written automatically by `run_simulation` next to the workbook. Readable directly via MRService.dll — that's how ModelRisk MCP reads results without a COM round-trip per sample.

---

## Server / MCP terms

### Audit rule

A check the server can run against a workbook to flag a specific class of mistake. Thirteen rules ship in v0.3.0 (VOSE-001 through VOSE-013). The set is editable: see [docs/authoring-audit-rules.md](authoring-audit-rules.md) to add your own.

### Catalogue

The 1,417-entry function-by-function reference for every Vose worksheet function. Used internally to validate every formula before it lands in Excel — no hallucinated function names can ever reach your workbook. Browse via `modelrisk://functions` from any MCP client.

### Claude Desktop / Claude Code / Claude for Excel

Three Anthropic-shipped MCP clients ModelRisk MCP runs under:
- **Claude Desktop** — the standalone chat app. Uses stdio transport.
- **Claude Code** — the CLI / IDE coding assistant. Also stdio.
- **Claude for Excel** — runs inside Excel's Office.js sandbox. Uses HTTP transport (the sandbox can't spawn subprocesses).

### `dry_run=True`

The default setting on every building tool. When `dry_run=True`, the tool *returns* what it would change but doesn't actually write to Excel. Lets Claude preview a change and confirm before committing. To commit: pass `dry_run=False` explicitly.

### MCP — Model Context Protocol

The Anthropic-defined open standard for letting Claude (and other AI assistants) talk to external tools, resources, and prompts. ModelRisk MCP is an MCP *server*; Claude Desktop is an MCP *client*. The protocol is JSON-RPC over stdio or HTTP. See <https://modelcontextprotocol.io>.

### MCP Registry

The official discovery service at `registry.modelcontextprotocol.io`. ModelRisk MCP is registered under `io.github.vosesoftware/modelrisk-mcp`. The registry verifies ownership via GitHub OIDC and a `mcp-name:` marker in the package README.

### MRService.dll

Vose's official ModelRisk SDK. ModelRisk MCP calls it via ctypes to read `.vmrs` files directly — much faster than the COM Dispatch surface and immune to ATL CoClass quirks. The DLL ships with ModelRisk; activation is automatic in v0.3.0 via a bundled offline key.

### Prompt (MCP)

A reusable conversation template the server publishes to MCP clients. ModelRisk MCP ships 5: `/build-risk-model`, `/audit-model`, `/interpret-results`, `/add-uncertainty`, `/import-legacy-model`. In Claude Desktop / Code, they appear as slash commands.

### Resource (MCP)

A piece of data the server exposes for the client to fetch by URI. ModelRisk MCP ships 5: `modelrisk://functions`, `modelrisk://distributions`, `modelrisk://methodology`, `modelrisk://workbook/current`, `modelrisk://audit-rules`. Resources are read-only; tools are read/write.

### Safety mechanisms

The nine layered protections that make the server safe to use on a real workbook (spec §11):
1. `dry_run=True` default
2. Excel undo stack always works
3. Bulk-write guard (`confirm_bulk=True` required for >50 cells)
4. No automatic saves
5. No overwriting non-Vose formulas
6. JSONL audit log of every write
7. `--read-only` mode disables all building tools
8. Single-writer mutex (two MCP servers can't drive the same Excel)
9. `restore_cell` from the audit log (recovery even after Excel's undo is gone)

### Tool (MCP)

A function the server exposes for the client to call. ModelRisk MCP ships 40 in v0.3.0: 12 reading + 13 building + 5 simulation + 7 workflow/reporting + 3 VMRS. See the [user manual](user-manual.md) for what they do.

### Transport

How the MCP client talks to the server. ModelRisk MCP supports three:
- **stdio** — default; subprocess pipes. Used by Claude Desktop and Claude Code.
- **streamable-http** — modern remote MCP transport. Used by Claude for Excel.
- **sse** — legacy Server-Sent Events transport, still supported.

### VoseInput / VoseOutput

The two ModelRisk wrappers that hand-shake with the Results Viewer:
- **VoseInput("Name")** — marks a cell as a *named input* the Results Viewer should track. Required for the cell to appear in sensitivity rankings and correlation matrices.
- **VoseOutput("Name")** — marks a cell as a *named output* whose per-iteration value should be recorded to the `.vmrs` file.

Both wrappers can take either a string-literal name (`VoseInput("Demand")`), a cell-reference name (`VoseInput(A5)` where A5 holds the name text), or an expression-based name (`VoseOutput("Period " & B3)`).

### Writer mutex

A Windows named mutex the server holds while writing to Excel. Prevents two MCP server instances from racing the same workbook. The second instance to try a building tool call raises `ConcurrentWriterError`.
