# ModelRisk MCP — Demo Gallery

Eleven hands-on demo workbooks, each showing one feature area of ModelRisk MCP on a
simple, real-life use case. Every workbook is self-contained: open it, then follow the
numbered steps on its **README sheet** — each step gives the exact prompt to type to
Claude and what to expect back. A demo takes about 5–10 minutes.

Start with **[00_START_HERE.xlsx](00_START_HERE.xlsx)** — the index, prerequisites and
colour legend.

| Workbook | Real-life scenario | Features shown |
|---|---|---|
| [01_build_model_from_brief.xlsx](01_build_model_from_brief.xlsx) | Coffee-chain store expansion NPV | `build_model_from_brief`, `plan_risk_model`, `run_simulation`, undo |
| [02_fit_distributions.xlsx](02_fit_distributions.xlsx) | Bakery daily demand from 365 days of history | `fit_and_rank_distributions`, `fit_distribution_to_data`, `compute_distribution` |
| [03_fit_copula_and_wire.xlsx](03_fit_copula_and_wire.xlsx) | Construction input costs that spike together | `fit_copula_to_data`, `fit_all_data_and_wire` |
| [04_reverse_stress_test.xlsx](04_reverse_stress_test.xlsx) | Data-centre budget breach, post-mortem in advance | `build_model_from_brief`, `reverse_stress_test` |
| [05_simulation_and_charts.xlsx](05_simulation_and_charts.xlsx) | Consumer-gadget launch profit | `run_simulation`, `create_histogram_chart`, `create_cdf_chart`, `create_tornado_chart` |
| [06_audit_model.xlsx](06_audit_model.xlsx) | Inherited "messy" model health-check | `audit_model`, `diagnose_workbook`, `find_hard_coded_inputs`, `restore_cell` |
| [07_insurance_aggregate.xlsx](07_insurance_aggregate.xlsx) | Motor insurer annual loss + capital adequacy | `create_aggregate` (FFT), `create_risk_event`, `get_tail_risk` |
| [08_scenarios_pricing.xlsx](08_scenarios_pricing.xlsx) | SaaS pricing decision — $79 / $99 / $119 | `run_scenarios`, `replace_constant_with_distribution` |
| [09_tail_risk_op_losses.xlsx](09_tail_risk_op_losses.xlsx) | Bank operational-risk 99.9% capital | `fit_tail` (GPD), `get_tail_risk`, `create_aggregate_mc` |
| [10_time_series_and_construction.xlsx](10_time_series_and_construction.xlsx) | CFO 5-year stochastic revenue forecast | `create_time_series`, `create_copula`, `set_named_range`, `save_workbook_as` |
| [11_executive_report.xlsx](11_executive_report.xlsx) | Board pack for a quarry acquisition | `build_executive_report`, `build_drivers_report`, `generate_executive_summary` |

## Prerequisites

- Windows Excel with the **ModelRisk add-in** installed and licensed.
- **modelrisk-mcp** connected to Claude (Claude Desktop / Claude Code) — see the
  [setup guides](../../docs/).

## Design notes

- The workbooks contain **no live Vose formulas** — each is a deterministic starting
  point (data + plan). The Vose layer is built *by Claude, in front of you*, following
  the README steps. That's the demo: the agent does the modelling work, and every write
  is journaled and reversible.
- Shared formatting: navy title bands, **light-blue** cells = inputs you can change,
  **pale-yellow** cells = where Claude writes Vose formulas / results, **green** = the
  key decision metric, grey italic = a prompt to type to Claude.
- All synthetic data is generated with fixed seeds (noted on each Data sheet), so the
  demos are reproducible. All files recalculate with zero formula errors in Excel.
