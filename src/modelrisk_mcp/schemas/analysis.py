"""Schemas for the quantitative-analysis tools (spec §7.4):

- `compute_distribution` — analytic distribution properties (PDF, CDF,
  exceedance, quantile, moments) with no simulation.
- `fit_and_rank_distributions` — fit many families to a data range and
  rank them by information criteria.
- `get_tail_risk` — VaR / CVaR / threshold probabilities from the
  per-iteration samples of a simulation output.
"""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field


class DistributionProperty(BaseModel):
    """One analytic property of a distribution, plus the exact Vose
    expression evaluated (so the result is auditable)."""

    metric: str = Field(description="Property requested, e.g. 'cdf' or 'quantile'.")
    at: float | None = Field(
        default=None,
        description="The x (pdf/cdf/exceedance) or u (quantile) the metric was evaluated at.",
    )
    value: float = Field(description="The computed value.")
    expression: str = Field(description="The Vose worksheet expression evaluated.")


class DistributionSummary(BaseModel):
    """A one-call analytic summary of a distribution: central moments
    plus a percentile ladder. All values are exact (no sampling)."""

    distribution: str = Field(description="The distribution object expression summarised.")
    mean: float
    stdev: float
    variance: float
    skewness: float
    kurtosis: float
    cov: float = Field(description="Coefficient of variation (stdev / mean).")
    percentiles: dict[str, float] = Field(
        description="Percentile ladder keyed by percent label, e.g. {'P5': ..., 'P50': ...}."
    )


class FitCandidate(BaseModel):
    """Goodness-of-fit scores for one fitted family. Lower information
    criteria are better fits."""

    family: str
    aic: float = Field(description="Akaike information criterion (lower = better).")
    sic: float = Field(description="Schwarz/Bayesian information criterion (lower = better).")
    hqic: float = Field(description="Hannan-Quinn information criterion (lower = better).")
    rank: int = Field(description="1 = best fit by the chosen criterion.")


class FitRanking(BaseModel):
    """Result of fitting and ranking several distribution families to a
    data range."""

    data_range: str
    criterion: str = Field(description="Criterion the ranking is sorted by (AIC / SIC / HQIC).")
    sample_size: int
    best_family: str | None = Field(
        default=None, description="The top-ranked family, or null if every fit failed."
    )
    candidates: list[FitCandidate] = Field(
        description="Successfully-fitted families, best first."
    )
    skipped: list[dict[str, str]] = Field(
        default_factory=list,
        description="Families that could not be fitted, with a reason each.",
    )


class TailMetric(BaseModel):
    """VaR and CVaR at one tail probability."""

    alpha: float = Field(description="Confidence level, e.g. 0.95.")
    var: float = Field(description="Value-at-Risk: the alpha-quantile of the loss.")
    cvar: float = Field(
        description="Conditional VaR / expected shortfall: mean loss in the worst (1-alpha) tail."
    )


class ThresholdProbability(BaseModel):
    """Probability mass either side of a threshold."""

    threshold: float
    p_above: float = Field(description="P(X > threshold).")
    p_at_or_below: float = Field(description="P(X <= threshold).")


class TailRiskResult(BaseModel):
    """Tail-risk profile of a simulation output, computed from its
    per-iteration samples."""

    output_name: str
    sample_size: int
    tail: str = Field(description="'upper' (large = bad) or 'lower' (small = bad).")
    mean: float
    stdev: float
    minimum: float
    maximum: float
    tail_metrics: list[TailMetric] = Field(
        description="VaR / CVaR at each requested confidence level."
    )
    threshold_probabilities: list[ThresholdProbability] = Field(
        default_factory=list,
        description="P(X>t) / P(X<=t) for each requested threshold.",
    )


class CorrelationMatrixResult(BaseModel):
    """Rank-order correlation matrix of a data range, plus its nearest
    valid (positive-semidefinite) form."""

    data_range: str
    variable_count: int
    matrix: list[list[float]] = Field(
        description="Spearman rank-order correlation matrix (VoseCorrMatrix)."
    )
    is_valid: bool = Field(
        description="True if the matrix is already a valid (PSD) correlation matrix."
    )
    nearest_valid_matrix: list[list[float]] | None = Field(
        default=None,
        description="Nearest valid matrix (VoseValidCorrmat) — null when already valid.",
    )


class TailFit(BaseModel):
    """A fitted extreme-value / GPD tail and its analytic risk metrics."""

    family: str = Field(description="Tail family fitted, e.g. 'GPD' or 'GEV'.")
    data_range: str
    object_formula: str = Field(
        description="The Vose<Family>FitObject formula written (or previewed)."
    )
    written: bool
    mean: float
    percentiles: dict[str, float] = Field(
        description="Fitted-tail percentiles, e.g. {'P95': ..., 'P99': ..., 'P99.5': ...}."
    )


class PercentileDelta(BaseModel):
    label: str
    a: float
    b: float
    difference: float = Field(description="a - b at this percentile.")


class DistributionComparison(BaseModel):
    """Head-to-head comparison of two simulation outputs from their
    per-iteration samples. Dominance is reported under the convention
    that LARGER outcomes are preferred."""

    output_a: str
    output_b: str
    sample_size: int
    paired: bool = Field(
        description="True if equal-length samples were compared iteration-by-iteration."
    )
    mean_a: float
    mean_b: float
    mean_difference: float = Field(description="mean(A) - mean(B).")
    stdev_a: float
    stdev_b: float
    p_a_greater: float | None = Field(
        description="P(A > B). Paired if samples align, else null.",
    )
    first_order_dominance: str = Field(
        description="'A', 'B', or 'none' — first-order stochastic dominance (larger=better)."
    )
    second_order_dominance: str = Field(
        description="'A', 'B', or 'none' — second-order stochastic dominance (risk-averse)."
    )
    percentile_deltas: list[PercentileDelta] = Field(
        description="A vs B at a percentile ladder."
    )


class RiskModelPlan(BaseModel):
    """A blueprint for converting a deterministic workbook into a Monte
    Carlo risk model: where it stands and the ordered next actions."""

    workbook: str
    output_count: int
    outputs: list[str] = Field(description="Names of cells already wrapped with VoseOutput.")
    distribution_count: int = Field(description="Vose distribution cells already present.")
    input_candidate_count: int
    input_candidates: list[dict[str, Any]] = Field(
        description="Ranked hard-coded numeric cells that look like uncertain inputs."
    )
    readiness: str = Field(
        description="'ready', 'needs-outputs', 'needs-inputs', or 'empty'."
    )
    steps: list[str] = Field(description="Ordered, state-aware next actions.")


class IntervalCoverage(BaseModel):
    nominal: float = Field(description="Nominal central interval, e.g. 0.90.")
    lower: float
    upper: float
    empirical: float = Field(description="Fraction of actuals that fell inside.")


class BacktestResult(BaseModel):
    """Validation of a simulation output against realised actuals."""

    output_name: str
    sample_size: int
    n_actuals: int
    model_mean: float
    actuals_mean: float
    bias: float = Field(description="actuals_mean - model_mean.")
    mean_pit: float = Field(
        description="Mean Probability Integral Transform; ~0.5 if calibrated."
    )
    pit_uniformity_ks: float = Field(
        description="KS distance of the PIT values from Uniform(0,1); 0 = perfectly calibrated."
    )
    frac_below_median: float = Field(
        description="Fraction of actuals below the model median; ~0.5 if calibrated."
    )
    coverage: list[IntervalCoverage] = Field(
        description="Empirical vs nominal coverage of central prediction intervals."
    )
    verdict: str = Field(description="Short calibration verdict.")


class UncertaintyDecomposition(BaseModel):
    """Epistemic-vs-aleatory variance split of an output, via the law of
    total variance from a full run and an epistemic-frozen run."""

    total_output: str
    conditional_output: str
    total_variance: float
    aleatory_variance: float = Field(
        description="Variability remaining when epistemic (parameter) inputs are frozen."
    )
    epistemic_variance: float = Field(
        description="total - aleatory; the part driven by parameter uncertainty."
    )
    epistemic_share: float = Field(description="Epistemic fraction of total variance (0-1).")
    aleatory_share: float = Field(description="Aleatory fraction of total variance (0-1).")
    total_stdev: float
    aleatory_stdev: float
    epistemic_stdev: float = Field(description="sqrt(max(epistemic_variance, 0)).")
    interpretation: str = Field(
        description="Which uncertainty dominates and what reduces it."
    )


class CopulaFitCandidate(BaseModel):
    """Goodness-of-fit scores for one fitted copula family. Lower
    information criteria are better fits."""

    family: str = Field(description="Copula family, e.g. 'CopulaMultiClayton'.")
    aic: float = Field(description="Akaike information criterion (lower = better).")
    sic: float = Field(description="Schwarz/Bayesian information criterion (lower = better).")
    hqic: float = Field(description="Hannan-Quinn information criterion (lower = better).")
    rank: int = Field(description="1 = best fit by the chosen criterion.")
    tail_dependence: str = Field(
        description=(
            "Asymmetric tail-dependence character of the family: 'none' "
            "(Normal/Frank), 'lower' (Clayton — joint crashes), 'upper' "
            "(Gumbel — joint booms), or 'both' (T). This is what a single "
            "correlation coefficient throws away."
        )
    )


class CopulaFitRanking(BaseModel):
    """Result of fitting and ranking parametric copula families to a
    multi-column data range — the dependence structure *fitted from
    data*, not merely constructed."""

    data_range: str
    n_variables: int = Field(description="Number of variables (columns) the copula spans.")
    criterion: str = Field(description="Criterion the ranking is sorted by (AIC / SIC / HQIC).")
    sample_size: int = Field(description="Number of joint observations (rows).")
    best_family: str | None = Field(
        default=None, description="Top-ranked copula family, or null if every fit failed."
    )
    candidates: list[CopulaFitCandidate] = Field(
        description="Successfully-fitted copula families, best first."
    )
    skipped: list[dict[str, str]] = Field(
        default_factory=list,
        description="Families that could not be fitted, with a reason each.",
    )
    note: str = Field(description="Interpretation of the winning family's tail behaviour.")


class BreachDriver(BaseModel):
    """One input's behaviour conditional on the output breaching the
    stress threshold — the inverse of a forward tornado."""

    input_name: str
    marginal_mean: float = Field(description="The input's mean across all iterations.")
    breach_mean: float = Field(description="The input's mean across only the breach iterations.")
    mean_shift_sd: float = Field(
        description=(
            "Standardised departure = (breach_mean - marginal_mean) / marginal_stdev. "
            "Large magnitude ⇒ this input is systematically different when things go wrong."
        )
    )
    breach_share_of_own_tail: float = Field(
        description=(
            "Fraction of this input's own worst-decile iterations that fall in the "
            "breach set — how concentrated breaches are in this input's tail."
        )
    )
    rank: int = Field(description="1 = strongest breach driver by |mean_shift_sd|.")


class StressScenario(BaseModel):
    """A concrete named input state extracted from the breach iterations."""

    label: str
    input_values: dict[str, float] = Field(
        description="Per-input value (the mean input vector over the breach set)."
    )


class ReverseStressResult(BaseModel):
    """Reverse stress test: from a bad output outcome back to the joint
    input state that produced it (Solvency II / PRA style). Pure analysis
    over the recorded per-iteration sample matrix — impossible without
    the engine's joint samples."""

    output_name: str
    threshold: float
    direction: str = Field(description="'above' or 'below' — the breach side of the threshold.")
    iterations: int
    breach_count: int
    breach_probability: float
    drivers: list[BreachDriver] = Field(
        description="Inputs ranked by how far they shift in the breach set."
    )
    scenario: StressScenario | None = Field(
        default=None, description="The mean input vector over the breach iterations."
    )
    note: str


class WiredColumnResult(BaseModel):
    """Outcome of fitting + wiring one data column as a model input."""

    input_name: str
    target_cell: str
    best_family: str | None
    formula: str
    written: bool
    skipped_reason: str | None = None


class FitAndWireResult(BaseModel):
    """Outcome of fitting marginals + a copula from data and wiring them
    into the workbook as a correlated, simulation-ready model."""

    workbook: str
    sheet: str
    columns: list[WiredColumnResult]
    copula_family: str | None = Field(
        default=None, description="Best-fit copula family wired across the columns, if any."
    )
    copula_anchor: str | None = Field(
        default=None, description="Range where the correlated-U block was written."
    )
    dry_run: bool
    simulated: bool = Field(description="Whether a validating simulation was run.")
    achieved_correlation: list[list[float]] | None = Field(
        default=None, description="Rank correlation of the wired inputs from the validating run."
    )
    rolled_back: bool = Field(
        default=False, description="True if a mid-build failure triggered a full rollback."
    )
    steps: list[str] = Field(description="Ordered log of what happened.")
    note: str


class BuiltInput(BaseModel):
    """One input the brief-builder created."""

    cell: str
    input_name: str
    formula: str
    source: str = Field(description="'fitted-from-data', 'proposed', or 'existing'.")


class ModelFromBriefResult(BaseModel):
    """Outcome of turning a deterministic workbook into a simulation-ready
    Monte Carlo model in one orchestrated, reversible pass."""

    workbook: str
    dry_run: bool
    outputs_wrapped: list[str] = Field(default_factory=list)
    inputs_built: list[BuiltInput] = Field(default_factory=list)
    correlated: bool = Field(default=False)
    simulated: bool = Field(default=False)
    headline: dict[str, dict[str, float]] = Field(
        default_factory=dict,
        description="Per-output headline stats (mean/P10/P50/P90) from the validating run.",
    )
    rolled_back: bool = Field(default=False)
    change_set_size: int = Field(description="Number of cells written (0 if dry_run).")
    steps: list[str] = Field(description="Ordered, human-readable log of the build.")
    note: str
