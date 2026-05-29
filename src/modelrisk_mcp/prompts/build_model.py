"""/build-risk-model prompt template."""

from __future__ import annotations

from modelrisk_mcp.server import mcp

description: str = (
    "Multi-turn workflow: walks the user from a deterministic question "
    "to a runnable Monte Carlo risk model. Asks about the decision, "
    "the output(s) of interest, the candidate inputs; proposes "
    "distributions; commits with confirmation; runs the simulation; "
    "interprets results."
)

template: str = """\
You are guiding the user through building a ModelRisk Monte Carlo risk
model from scratch. Lead the conversation; don't wait to be asked.

Workflow (do this in order; pause for the user's input at each step):

1. **Decision and outputs.** Ask the user: "What decision are we
   trying to inform, and what's the output of the model (e.g.
   total project cost, NPV, downtime hours)?" Confirm the active
   workbook with `get_active_workbook`; if the user wants to start
   in a new workbook, ask them to open one.

2. **Inputs.** Ask the user to walk you through every assumption
   that goes into the output. For each one, decide: is it
   (a) genuinely certain (constant), (b) uncertain (needs a
   distribution), or (c) a derived intermediate (formula).

3. **Discover hard-coded inputs.** Run `discover_inputs` against
   the workbook. Surface the top candidates to the user and ask
   which ones are genuinely uncertain.

4. **Propose distributions.** For each uncertain input, call
   `propose_distributions_for_inputs` with the user's description
   of the quantity. Walk through the top recommendation and its
   rationale. Confirm with the user before committing.

5. **Commit each input.** For each agreed input, call
   `replace_constant_with_distribution` with `dry_run=True` first
   to preview, then `dry_run=False` to commit. Use a descriptive
   `input_name`.

6. **Mark the output(s).** Use `wrap_with_output` (preview, then
   commit) on the cell(s) the user identified in step 1.

7. **Audit.** Run `audit_model` to catch any methodology mistakes
   before simulating. Address every `error` finding; surface
   `warning` and `info` for the user's judgement.

8. **Run.** Use `run_simulation` with `iterations=10000` (or what
   the user wants) and `seed` if reproducibility matters.

9. **Interpret.** Call `generate_executive_summary` and present
   the markdown to the user. Offer to drill into specific outputs
   via `get_simulation_results` and `get_sensitivity_ranking`.

Resources to load into context as you go:
- `modelrisk://methodology` — the 8 core principles.
- `modelrisk://knowledge` — the risk-analysis knowledge base distilled
  from the ModelRisk Help (why to quantify uncertainty, selecting
  distributions, expert opinion, correlation, time series).
- `modelrisk://distributions` — the distribution selection guide.
- `modelrisk://functions/{name}` — when explaining a specific
  function's parameters.

Defaults you should respect:
- Every building tool defaults to `dry_run=True`. Preview, confirm,
  then re-call with `dry_run=False`.
- Iterations: ask the user; default to 10000 if they don't specify.
"""


@mcp.prompt(
    name="build-risk-model",
    description=description,
)
def build_risk_model_prompt() -> str:
    return template
