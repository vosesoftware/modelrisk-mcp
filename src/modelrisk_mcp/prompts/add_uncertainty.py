"""/add-uncertainty prompt template."""

from __future__ import annotations

from modelrisk_mcp.server import mcp

description: str = (
    "Take a deterministic Excel model and add uncertainty to it. "
    "Walk the user through identifying which inputs are uncertain, "
    "choosing distributions for each, wrapping inputs/outputs, and "
    "running the simulation."
)

template: str = """\
You are converting a deterministic Excel model into a Monte Carlo
risk model. The user has built the model in Excel as plain numbers
and formulas; your job is to add the distributions and the Vose
wrappers without breaking the existing logic.

Workflow:

1. **Confirm the workbook + outputs.** Call `get_active_workbook`,
   then `list_modelrisk_outputs` to see if any cells are already
   wrapped. Ask the user: which output cell(s) do you want to
   put a confidence range around?

2. **Trace inputs.** Run `discover_inputs(workbook_name)` to get
   the ranked candidate list of hard-coded numeric inputs.

3. **For each candidate input**, ask the user:
   - Is this assumption genuinely uncertain?
   - If yes: what's your best estimate of the range (min / most
     likely / max), or do you have data to fit?

4. **Propose or fit distributions.**
   - *No data, just a range:* call
     `propose_distributions_for_inputs` with the user's
     description and walk through the top recommendation.
   - *Has data:* call `fit_and_rank_distributions` on the data
     range — it fits many families and ranks them by AIC/SIC/HQIC,
     so you recommend the best fit rather than guessing one. For a
     heavy tail (losses, claims), use `fit_tail` (GPD) on the
     exceedances.
   - Sanity-check a proposed distribution before committing with
     `compute_distribution(metric="summary", ...)` — confirm its
     mean/percentiles match the user's intent.
   - Commit with `replace_constant_with_distribution` (dry_run
     first, then commit).

5. **Correlate inputs that move together.** If two or more inputs
   are dependent, don't leave them independent (a classic
   methodology trap). With historical data, call
   `compute_correlation_matrix` on it, then `create_copula` with
   that matrix; otherwise ask the user for a correlation estimate.

6. **Mark outputs.** For each of the user's chosen output cells,
   `wrap_with_output(name=...)` so they appear in the Results
   Viewer.

7. **Audit, then run.** `audit_model` → fix any errors →
   `run_simulation(iterations=10000)`.

8. **Interpret.** Read the tail with `get_tail_risk` (VaR / CVaR);
   if you're comparing strategies, use `compare_distributions`
   (P(A>B), stochastic dominance). Then hand off to the same
   drill-down menu as `/interpret-results`.

Pace: this is a coaching conversation. Don't commit a distribution
until the user has explicitly confirmed both the family and the
parameters.
"""


@mcp.prompt(name="add-uncertainty", description=description)
def add_uncertainty_prompt() -> str:
    return template
