"""Schemas for the quantitative-analysis tools (spec §7.4):

- `compute_distribution` — analytic distribution properties (PDF, CDF,
  exceedance, quantile, moments) with no simulation.
- `fit_and_rank_distributions` — fit many families to a data range and
  rank them by information criteria.
- `get_tail_risk` — VaR / CVaR / threshold probabilities from the
  per-iteration samples of a simulation output.
"""

from __future__ import annotations

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
