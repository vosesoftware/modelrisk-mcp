"""/interpret-results prompt template."""

from __future__ import annotations

from modelrisk_mcp.server import mcp

description: str = (
    "Read the latest simulation results from the active workbook, "
    "produce a structured executive summary, then offer drill-downs "
    "into specific outputs (sensitivity, correlation, percentiles)."
)

template: str = """\
You are interpreting the user's ModelRisk simulation results.

Workflow:

1. **Confirm the workbook + run.** Call `get_active_workbook`. If
   `get_simulation_status` reports `idle` and `get_simulation_results`
   returns an empty list, the user hasn't run a simulation — offer
   to start one (with their parameters).

2. **Executive summary.** Call `generate_executive_summary` and
   present the markdown directly. Ask the user about deterministic
   values for the outputs first if they want the contingency
   comparison.

3. **Drill-down menu.** Offer:
   - "Show me the full percentile distribution for X." → call
     `get_simulation_results(output_names=[X])`.
   - "Which inputs drive X most?" → call
     `get_sensitivity_ranking(output_name=X)`.
   - "Are any of these outputs correlated?" → call
     `get_correlation_matrix(name_list=[...])`.

4. **Methodology checks.** Mention any caveats from the audit:
   if `audit_model` reports warnings on the workbook, surface
   them so the user knows what *not* to over-interpret.

Be quantitative but plain-spoken. Quote percentiles as ranges
("the 80% confidence range is X to Y") rather than statistics
jargon, unless the user has clearly opted in to that register.
"""


@mcp.prompt(name="interpret-results", description=description)
def interpret_results_prompt() -> str:
    return template
