"""Quantitative-analysis tools (spec §7.4).

Three read-only analysis tools that answer questions ModelRisk can
answer but the server previously couldn't surface without a manual
formula + simulation round-trip:

- `compute_distribution` — analytic distribution properties (PDF, CDF,
  exceedance, quantile, moments, or a one-call summary). Uses
  `Application.Evaluate` on an inline distribution-object expression;
  deterministic, no cells written, no simulation.
- `fit_and_rank_distributions` — fit many families to a data range and
  rank them by AIC / SIC / HQIC (ModelRisk's own goodness-of-fit
  scores). Uses a transient scratch sheet that is always deleted.
- `get_tail_risk` — VaR / CVaR (expected shortfall) and threshold
  probabilities from the per-iteration samples of a simulation output.
  Pure Python over the samples the results reader already exposes.
"""

from __future__ import annotations

import math
from typing import Annotated, Any

from pydantic import Field

from modelrisk_mcp.bridge.formulas import build_distribution_formula
from modelrisk_mcp.errors import ModelRiskComputationError
from modelrisk_mcp.schemas.analysis import (
    DistributionProperty,
    DistributionSummary,
    FitCandidate,
    FitRanking,
    TailMetric,
    TailRiskResult,
    ThresholdProbability,
)
from modelrisk_mcp.server import mcp
from modelrisk_mcp.tools.reading import get_bridge

# Default candidate families for fit-and-rank: a broad continuous set.
# Families whose Vose<Family>FitObject is missing, or that can't fit the
# data, are reported as skipped — so an over-broad default is harmless.
_DEFAULT_FIT_FAMILIES = [
    "Normal",
    "Lognormal",
    "Gamma",
    "Weibull",
    "Expon",
    "Logistic",
    "LogLogistic",
    "Pareto",
    "InvGauss",
    "Beta",
]

_PERCENTILE_LADDER = [
    ("P1", 0.01),
    ("P5", 0.05),
    ("P10", 0.10),
    ("P25", 0.25),
    ("P50", 0.50),
    ("P75", 0.75),
    ("P90", 0.90),
    ("P95", 0.95),
    ("P99", 0.99),
]


def _object_expression(
    family: str | None,
    parameters: list[float] | None,
    object_sheet: str | None,
    object_cell: str | None,
) -> str:
    """Build the distribution-object expression to query — either an
    inline `Vose<Family>Object(p1, p2, ...)` or a reference to a cell
    that already holds an Object (e.g. a fitted distribution)."""
    if object_cell:
        ref = f"'{object_sheet}'!{object_cell}" if object_sheet else object_cell
        return ref
    if not family:
        raise ModelRiskComputationError(
            "Provide either `family` (+ `parameters`) or `object_cell`."
        )
    bridge = get_bridge()
    func = f"Vose{family}Object"
    formula = build_distribution_formula(func, list(parameters or []), bridge.catalogue)
    return formula.lstrip("=")


def _metric_expression(metric: str, obj: str, at: float | None) -> str:
    m = metric.lower()
    if m in ("pdf", "density"):
        return f"VoseProb({at},{obj},FALSE)"
    if m in ("cdf", "probability"):
        return f"VoseProb({at},{obj},TRUE)"
    if m in ("exceedance", "sf"):
        return f"1-VoseProb({at},{obj},TRUE)"
    if m in ("quantile", "percentile", "inverse"):
        return f"VosePercentile({obj},{at})"
    if m == "mean":
        return f"VoseMean({obj})"
    if m in ("stdev", "sd", "std"):
        return f"SQRT(VoseVariance({obj}))"
    if m == "variance":
        return f"VoseVariance({obj})"
    if m == "skewness":
        return f"VoseSkewness({obj})"
    if m == "kurtosis":
        return f"VoseKurtosis({obj})"
    if m in ("cov", "cofv"):
        return f"VoseCofV({obj})"
    raise ModelRiskComputationError(
        f"Unknown metric {metric!r}. Use one of: pdf, cdf, exceedance, "
        "quantile, mean, stdev, variance, skewness, kurtosis, cov, summary."
    )


@mcp.tool(
    description=(
        "ModelRisk: Analytic distribution calculator — compute a property of a "
        "distribution WITHOUT running a simulation. Give a `family` (e.g. "
        "'Normal', 'Lognormal', 'PERT') and its `parameters`, OR point at an "
        "`object_cell` that already holds a Vose distribution object (e.g. a "
        "fitted distribution). `metric` is one of: 'pdf' (density f(x)), 'cdf' "
        "(P(X<=x)), 'exceedance' (P(X>x)), 'quantile' (the x at cumulative "
        "probability u), 'mean', 'stdev', 'variance', 'skewness', 'kurtosis', "
        "'cov', or 'summary' (all moments + a percentile ladder in one call). "
        "`at` is x for pdf/cdf/exceedance and u (0-1) for quantile. Exact, "
        "read-only: nothing is written and no simulation is run."
    )
)
def compute_distribution(
    metric: Annotated[
        str,
        Field(
            description=(
                "pdf | cdf | exceedance | quantile | mean | stdev | variance | "
                "skewness | kurtosis | cov | summary"
            )
        ),
    ],
    family: Annotated[
        str | None,
        Field(description="Distribution family, e.g. 'Lognormal'. Omit if using object_cell."),
    ] = None,
    parameters: Annotated[
        list[float] | None,
        Field(description="Positional parameters for the family, in ModelRisk order."),
    ] = None,
    at: Annotated[
        float | None,
        Field(description="x (pdf/cdf/exceedance) or u in (0,1) for quantile."),
    ] = None,
    object_sheet: Annotated[
        str | None, Field(description="Sheet of object_cell, if used.")
    ] = None,
    object_cell: Annotated[
        str | None,
        Field(description="A1-style cell already holding a Vose distribution object."),
    ] = None,
) -> DistributionProperty | DistributionSummary:
    bridge = get_bridge()
    obj = _object_expression(family, parameters, object_sheet, object_cell)

    if metric.lower() == "summary":
        variance = bridge.evaluate_number(f"VoseVariance({obj})")
        percentiles = {
            label: bridge.evaluate_number(f"VosePercentile({obj},{u})")
            for label, u in _PERCENTILE_LADDER
        }
        return DistributionSummary(
            distribution=obj,
            mean=bridge.evaluate_number(f"VoseMean({obj})"),
            stdev=math.sqrt(variance) if variance >= 0 else float("nan"),
            variance=variance,
            skewness=bridge.evaluate_number(f"VoseSkewness({obj})"),
            kurtosis=bridge.evaluate_number(f"VoseKurtosis({obj})"),
            cov=bridge.evaluate_number(f"VoseCofV({obj})"),
            percentiles=percentiles,
        )

    needs_at = metric.lower() in (
        "pdf", "density", "cdf", "probability", "exceedance", "sf",
        "quantile", "percentile", "inverse",
    )
    if needs_at and at is None:
        raise ModelRiskComputationError(f"metric {metric!r} requires `at`.")
    expr = _metric_expression(metric, obj, at)
    return DistributionProperty(
        metric=metric.lower(),
        at=at if needs_at else None,
        value=bridge.evaluate_number(expr),
        expression=expr,
    )


@mcp.tool(
    description=(
        "ModelRisk: Fit several distribution families to a data range and rank "
        "them by goodness of fit. For each family it fits "
        "Vose<Family>FitObject and scores it with ModelRisk's information "
        "criteria — AIC, SIC (Schwarz/BIC) and HQIC — then ranks ascending "
        "(lower = better) by `criterion`. Families with no fit function, or "
        "that can't fit the data, are returned under `skipped` with a reason. "
        "Use this instead of guessing a single family for fit_distribution_to_data. "
        "Runs on a transient scratch sheet that is always deleted; the data is "
        "not modified."
    )
)
def fit_and_rank_distributions(
    workbook: Annotated[str, Field(description="Workbook file name, e.g. 'model.xlsx'.")],
    sheet: Annotated[str, Field(description="Sheet holding the data.")],
    data_range: Annotated[
        str, Field(description="A1-style range of the data, e.g. 'A1:A200'.")
    ],
    families: Annotated[
        list[str] | None,
        Field(description="Families to try. Omit for a broad continuous default set."),
    ] = None,
    criterion: Annotated[
        str, Field(description="Ranking criterion: 'SIC' (default), 'AIC', or 'HQIC'.")
    ] = "SIC",
    uncertainty: Annotated[
        bool,
        Field(description="Fit with parameter uncertainty (second-order). Default False."),
    ] = False,
) -> FitRanking:
    crit = criterion.upper()
    if crit not in ("AIC", "SIC", "HQIC"):
        raise ModelRiskComputationError(
            f"Unknown criterion {criterion!r}; use AIC, SIC, or HQIC."
        )
    fams = families or list(_DEFAULT_FIT_FAMILIES)
    qualified = f"'{sheet}'!{data_range}"
    scored, skipped, sample_size = get_bridge().fit_and_rank(
        qualified, fams, workbook=workbook, uncertainty=uncertainty
    )
    key = crit.lower()
    scored.sort(key=lambda d: d[key])
    candidates = [
        FitCandidate(
            family=d["family"],
            aic=d["aic"],
            sic=d["sic"],
            hqic=d["hqic"],
            rank=i + 1,
        )
        for i, d in enumerate(scored)
    ]
    return FitRanking(
        data_range=qualified,
        criterion=crit,
        sample_size=sample_size,
        best_family=candidates[0].family if candidates else None,
        candidates=candidates,
        skipped=skipped,
    )


def _quantile(sorted_xs: list[float], q: float) -> float:
    """Linear-interpolation quantile (Excel PERCENTILE.INC / numpy
    'linear' / type 7). `sorted_xs` must be ascending and non-empty."""
    n = len(sorted_xs)
    if n == 1:
        return sorted_xs[0]
    pos = q * (n - 1)
    lo = math.floor(pos)
    hi = math.ceil(pos)
    if lo == hi:
        return sorted_xs[lo]
    return sorted_xs[lo] + (sorted_xs[hi] - sorted_xs[lo]) * (pos - lo)


def compute_tail_risk(
    samples: list[float],
    *,
    alphas: list[float],
    thresholds: list[float],
    tail: str,
) -> dict[str, Any]:
    """Pure-Python VaR/CVaR + threshold probabilities from a sample
    array. `tail='upper'` treats large values as the loss (VaR is the
    alpha-quantile, CVaR the mean beyond it); `tail='lower'` treats
    small values as the loss. Separated from the tool for unit testing
    without Excel."""
    n = len(samples)
    if n == 0:
        raise ModelRiskComputationError("No samples to analyse.")
    xs = sorted(samples)
    mean = math.fsum(xs) / n
    var_pop = math.fsum((x - mean) ** 2 for x in xs) / n
    stdev = math.sqrt(var_pop)

    tail_metrics: list[dict[str, float]] = []
    for a in alphas:
        if tail == "lower":
            v = _quantile(xs, 1.0 - a)
            beyond = [x for x in xs if x <= v]
        else:
            v = _quantile(xs, a)
            beyond = [x for x in xs if x >= v]
        cvar = math.fsum(beyond) / len(beyond) if beyond else v
        tail_metrics.append({"alpha": a, "var": v, "cvar": cvar})

    thr: list[dict[str, float]] = []
    for t in thresholds:
        above = sum(1 for x in xs if x > t)
        thr.append(
            {"threshold": t, "p_above": above / n, "p_at_or_below": 1.0 - above / n}
        )

    return {
        "sample_size": n,
        "mean": mean,
        "stdev": stdev,
        "minimum": xs[0],
        "maximum": xs[-1],
        "tail_metrics": tail_metrics,
        "threshold_probabilities": thr,
    }


@mcp.tool(
    description=(
        "ModelRisk: Tail-risk profile of a simulation output from its "
        "per-iteration samples — Value-at-Risk (VaR) and Conditional VaR / "
        "expected shortfall (CVaR) at each confidence level, plus optional "
        "threshold probabilities P(X>t) / P(X<=t). `tail='upper'` (default) "
        "treats large values as the loss (e.g. cost, claims); `tail='lower'` "
        "treats small values as the loss (e.g. NPV, profit). VaR is the "
        "alpha-quantile; CVaR is the mean of the worst (1-alpha) tail beyond "
        "it. Reads the samples ModelRisk recorded in the .vmrs — run a "
        "simulation first."
    )
)
def get_tail_risk(
    output_name: Annotated[str, Field(description="VoseOutput (or VoseInput) name.")],
    alphas: Annotated[
        list[float] | None,
        Field(description="Confidence levels for VaR/CVaR. Default [0.95, 0.99]."),
    ] = None,
    thresholds: Annotated[
        list[float] | None,
        Field(description="Values to compute P(X>t) / P(X<=t) for. Optional."),
    ] = None,
    tail: Annotated[
        str, Field(description="'upper' (large = bad, default) or 'lower' (small = bad).")
    ] = "upper",
    workbook_name: Annotated[
        str | None, Field(description="Workbook name. Omit for the active workbook.")
    ] = None,
    max_n: Annotated[
        int,
        Field(ge=1, le=1_000_000, description="Max samples to read (default 100 000)."),
    ] = 100_000,
) -> TailRiskResult:
    t = tail.lower()
    if t not in ("upper", "lower"):
        raise ModelRiskComputationError("tail must be 'upper' or 'lower'.")
    a_list = alphas or [0.95, 0.99]
    if any(not 0.0 < a < 1.0 for a in a_list):
        raise ModelRiskComputationError("Each alpha must be strictly between 0 and 1.")
    samples = get_bridge().get_samples(output_name, workbook_name, max_n=max_n)
    stats = compute_tail_risk(
        list(samples), alphas=a_list, thresholds=thresholds or [], tail=t
    )
    return TailRiskResult(
        output_name=output_name,
        sample_size=stats["sample_size"],
        tail=t,
        mean=stats["mean"],
        stdev=stats["stdev"],
        minimum=stats["minimum"],
        maximum=stats["maximum"],
        tail_metrics=[TailMetric(**m) for m in stats["tail_metrics"]],
        threshold_probabilities=[
            ThresholdProbability(**p) for p in stats["threshold_probabilities"]
        ],
    )


__all__ = [
    "compute_distribution",
    "compute_tail_risk",
    "fit_and_rank_distributions",
    "get_tail_risk",
]
