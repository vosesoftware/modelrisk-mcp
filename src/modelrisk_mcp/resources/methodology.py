"""Vose methodology principles resource."""

from __future__ import annotations

from modelrisk_mcp.server import mcp

_METHODOLOGY: str = """\
# Vose ModelRisk methodology — core principles

These principles guide every decision the LLM makes when building or
auditing a risk model with this MCP server.

## 1. The model expresses uncertainty, not just a point estimate

Every input that could plausibly take a range of values *must* be
modelled as a distribution, not a single number. Treating a noisy input
as deterministic understates total uncertainty in the output by
exactly the amount that input could swing.

## 2. Distributions reflect *real* uncertainty about parameters

When fitting a distribution to data, set `uncertainty=TRUE`. Without
it, the simulation samples only from the best-fit parameters as if
they were known with certainty — which they aren't. Carry parameter
uncertainty through the simulation.

## 3. Discrete events use VoseRiskEvent

For "the bad thing might happen" scenarios, use VoseRiskEvent with a
probability and an impact distribution. Don't naively multiply
probability times impact — that suppresses the bimodal nature of risk
events.

## 4. Aggregates use VoseAggregateMC

For sums of a (possibly random) number of i.i.d. severity draws — e.g.
total annual loss = number-of-incidents x per-incident-loss — use
VoseAggregateMC, not a hand-rolled loop. Aggregate functions handle
the convolution correctly.

## 5. Correlate inputs that are correlated in the real world

Two inputs that move together in reality (e.g. unit cost and demand
during a recession) must be correlated in the model via a copula
(`VoseCopulaMultiNormal` or similar). Independent inputs that are
actually correlated produce simulation results with artificially
tight outputs.

## 6. Output cells are marked with VoseOutput

Only cells wrapped with `VoseOutput("name")` appear in the Results
Viewer with per-iteration history, percentiles, and sensitivity
analysis. If you want to interrogate it after the run, wrap it.

## 7. Input cells are marked with VoseInput

Symmetric to outputs. `VoseInput("name")` makes the cell trackable for
correlation and tornado analysis.

## 8. Don't simulate constants

If a value can't possibly vary across iterations (e.g. days-per-year,
unit conversion factors), don't wrap it in a distribution. Keep it as
a plain number so the simulation engine doesn't waste samples on it
and the Results Viewer isn't cluttered.
"""


@mcp.resource(
    uri="modelrisk://methodology",
    name="modelrisk-methodology",
    description=(
        "ModelRisk: the 8 core methodology principles the LLM follows "
        "when building, fitting, or auditing models. Loaded into "
        "context whenever the user starts a /build-risk-model or "
        "/audit-model conversation."
    ),
    mime_type="text/markdown",
)
def methodology_resource() -> str:
    return _METHODOLOGY
