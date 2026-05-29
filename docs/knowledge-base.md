# Risk-analysis Knowledge Base

A curated, **attributed distillation** of risk-modelling guidance from the **ModelRisk Help** (Vose Software) — written in our own words, each section citing the source article. It exists so the LLM consults authoritative methodology while building or critiquing a model, not just general intuition.

## Where it lives

To keep a single source of truth, the substance ships as a package **data file** and is served as an **MCP resource** rather than duplicated here:

- **Resource URI:** `modelrisk://knowledge` — fetchable from any MCP client; loaded into context by the `/build-risk-model` and `/audit-model` prompts.
- **Source file:** [`src/modelrisk_mcp/data/knowledge_base.md`](../src/modelrisk_mcp/data/knowledge_base.md) — read it directly on GitHub.

## What's in it

Six sections, each distilled from a ModelRisk Help article:

| Section | Source article |
|---|---|
| Why quantify uncertainty | *Introduction to risk analysis* |
| The two forms of uncertainty (variability vs risk events) | *Introduction to risk analysis* · *Vose Risk Event* |
| Selecting a distribution — the five properties | *Selecting the appropriate distributions for your model* |
| Eliciting expert opinion | *Modeling expert opinion* |
| Modeling correlation | *Modeling correlation* · *Copulas* |
| Modeling a quantity over time | *Time series introduction* |

## How it relates to the other knowledge docs

- **This knowledge base** — the authoritative *why*, distilled from the ModelRisk Help, served to the LLM at build time.
- [Methodology](methodology.md) — the 8 enforced principles, each tied to an audit rule.
- [Distribution selection](distribution-selection.md) — which exact distribution for which quantity.
- [Modeling patterns](modeling-patterns.md) — how to compose the structures (frequency-severity, correlation, time series).

> Provenance: the knowledge base is an attributed distillation in our own words, not verbatim ModelRisk Help text. For the complete, authoritative treatment of any topic, consult the corresponding article in the ModelRisk Help (Vose Software).
