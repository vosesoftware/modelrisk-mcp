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
    BacktestResult,
    CorrelationMatrixResult,
    DistributionComparison,
    DistributionProperty,
    DistributionSummary,
    FitCandidate,
    FitRanking,
    IntervalCoverage,
    PercentileDelta,
    TailFit,
    TailMetric,
    TailRiskResult,
    ThresholdProbability,
    UncertaintyDecomposition,
)
from modelrisk_mcp.schemas.workbook import CellRef
from modelrisk_mcp.server import mcp
from modelrisk_mcp.tools.reading import get_bridge

# Tail families that fit an extreme-value / Generalised-Pareto tail.
_TAIL_FAMILIES = {"GPD", "GEV", "ExtValueMax", "ExtValueMin"}

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


@mcp.tool(
    description=(
        "ModelRisk: Compute the rank-order (Spearman) correlation matrix of a "
        "data range via VoseCorrMatrix, and its nearest valid "
        "(positive-semidefinite) form via VoseValidCorrmat. Use this to turn "
        "historical data into the correlation matrix you feed to create_copula "
        "for correlated inputs. Variables are columns by default (set "
        "`data_in_rows=True` if each row is a variable). Read-only: runs on a "
        "transient scratch sheet that is always deleted."
    )
)
def compute_correlation_matrix(
    workbook: Annotated[str, Field(description="Workbook file name.")],
    sheet: Annotated[str, Field(description="Sheet holding the data.")],
    data_range: Annotated[
        str, Field(description="A1-style range of the data, e.g. 'A1:D200'.")
    ],
    data_in_rows: Annotated[
        bool, Field(description="True if each variable is a row. Default: columns.")
    ] = False,
) -> CorrelationMatrixResult:
    bridge = get_bridge()
    rng = bridge.excel.get_range_shape(workbook, sheet, data_range)
    rows, cols = rng
    n_vars = rows if data_in_rows else cols
    if n_vars < 2:
        raise ModelRiskComputationError(
            f"Need at least 2 variables to correlate; got {n_vars}. "
            "Check `data_range` orientation / `data_in_rows`."
        )
    qualified = f"'{sheet}'!{data_range}"
    matrix, nearest, is_valid = bridge.correlation_matrix_of_data(
        qualified, n_vars, data_in_rows=data_in_rows, workbook=workbook
    )
    return CorrelationMatrixResult(
        data_range=qualified,
        variable_count=n_vars,
        matrix=matrix,
        is_valid=is_valid,
        nearest_valid_matrix=None if is_valid else nearest,
    )


@mcp.tool(
    description=(
        "ModelRisk: Fit an extreme-value / Generalised-Pareto tail to data and "
        "read its risk. `family` is 'GPD' (peaks-over-threshold, the standard "
        "tail model), 'GEV' (block maxima), 'ExtValueMax', or 'ExtValueMin'. "
        "For GPD peaks-over-threshold, pass the range of exceedances above your "
        "threshold as `data_range`. Writes a Vose<Family>FitObject (dry_run "
        "previews) and returns the fitted tail's mean and high percentiles "
        "(P95 / P99 / P99.5 / P99.9) computed analytically — the tail risk "
        "without a simulation. Feed the written object cell to "
        "compute_distribution / get_tail_risk for more."
    )
)
def fit_tail(
    workbook: str,
    sheet: str,
    target_cell: str,
    data_range: Annotated[str, Field(description="A1-style range of the tail data.")],
    family: Annotated[
        str, Field(description="'GPD' (default), 'GEV', 'ExtValueMax', or 'ExtValueMin'.")
    ] = "GPD",
    uncertainty: Annotated[
        bool, Field(description="Fit with parameter uncertainty. Default True.")
    ] = True,
    dry_run: bool = True,
) -> TailFit:
    if family not in _TAIL_FAMILIES:
        raise ModelRiskComputationError(
            f"Unknown tail family {family!r}; use one of {sorted(_TAIL_FAMILIES)}."
        )
    bridge = get_bridge()
    func = f"Vose{family}FitObject"
    if func not in bridge.catalogue:
        raise ModelRiskComputationError(
            f"No fit function {func!r} in the ModelRisk catalogue."
        )
    qualified = f"'{sheet}'!{data_range}"
    unc = "TRUE" if uncertainty else "FALSE"
    object_formula = f"={func}({qualified},{unc})"

    ladder = [("P95", 0.95), ("P99", 0.99), ("P99.5", 0.995), ("P99.9", 0.999)]
    templates = ["VoseMean({obj})"] + [
        f"VosePercentile({{obj}},{u})" for _, u in ladder
    ]
    values = bridge.evaluate_object_metrics(object_formula, templates, workbook=workbook)
    if values[0] is None:
        raise ModelRiskComputationError(
            f"{family} fit did not produce a valid distribution for {qualified}."
        )
    percentiles = {
        label: v
        for (label, _), v in zip(ladder, values[1:], strict=True)
        if v is not None
    }

    written = False
    if not dry_run:
        ref = CellRef(workbook=workbook, sheet=sheet, cell=target_cell)
        get_bridge().safe_write_cell(ref, object_formula)
        written = True

    return TailFit(
        family=family,
        data_range=qualified,
        object_formula=object_formula,
        written=written,
        mean=values[0],
        percentiles=percentiles,
    )


def _empirical_cdf(sorted_xs: list[float], grid: list[float]) -> list[float]:
    """F(x) = fraction of samples <= x, evaluated at each grid point.
    `sorted_xs` must be ascending."""
    import bisect

    n = len(sorted_xs)
    return [bisect.bisect_right(sorted_xs, g) / n for g in grid]


def compare_samples(a: list[float], b: list[float]) -> dict[str, Any]:
    """Pure-Python head-to-head comparison of two sample sets. Dominance
    uses the convention that larger outcomes are preferred. Separated
    from the tool for unit testing without Excel."""
    if not a or not b:
        raise ModelRiskComputationError("Both outputs need samples to compare.")
    sa, sb = sorted(a), sorted(b)
    na, nb = len(a), len(b)
    mean_a, mean_b = math.fsum(a) / na, math.fsum(b) / nb
    sd_a = math.sqrt(math.fsum((x - mean_a) ** 2 for x in a) / na)
    sd_b = math.sqrt(math.fsum((x - mean_b) ** 2 for x in b) / nb)

    paired = na == nb
    p_a_greater = (
        sum(1 for x, y in zip(a, b, strict=True) if x > y) / na if paired else None
    )

    grid = sorted(set(sa) | set(sb))
    fa = _empirical_cdf(sa, grid)
    fb = _empirical_cdf(sb, grid)
    tol = 1e-9
    # First-order: A preferred (stochastically larger) if F_A(x) <= F_B(x) for all x.
    a_fo = all(x <= y + tol for x, y in zip(fa, fb, strict=True))
    b_fo = all(y <= x + tol for x, y in zip(fa, fb, strict=True))
    fo = "A" if a_fo and not b_fo else "B" if b_fo and not a_fo else "none"
    # Second-order: cumulative area of (F_B - F_A); A preferred if always >= 0.
    so = "none"
    if fo == "none":
        cum = 0.0
        a_dom = b_dom = True
        for i in range(1, len(grid)):
            width = grid[i] - grid[i - 1]
            cum += (fb[i - 1] - fa[i - 1]) * width
            if cum < -tol:
                a_dom = False
            if cum > tol:
                b_dom = False
        so = "A" if a_dom and not b_dom else "B" if b_dom and not a_dom else "none"

    ladder = [("P5", 0.05), ("P25", 0.25), ("P50", 0.50), ("P75", 0.75), ("P95", 0.95)]
    deltas = [
        {
            "label": label,
            "a": _quantile(sa, u),
            "b": _quantile(sb, u),
            "difference": _quantile(sa, u) - _quantile(sb, u),
        }
        for label, u in ladder
    ]

    return {
        "sample_size": min(na, nb),
        "paired": paired,
        "mean_a": mean_a,
        "mean_b": mean_b,
        "mean_difference": mean_a - mean_b,
        "stdev_a": sd_a,
        "stdev_b": sd_b,
        "p_a_greater": p_a_greater,
        "first_order_dominance": fo,
        "second_order_dominance": so,
        "percentile_deltas": deltas,
    }


@mcp.tool(
    description=(
        "ModelRisk: Compare two simulation outputs head-to-head from their "
        "per-iteration samples — mean/stdev/percentile differences, P(A > B), "
        "and first- and second-order stochastic dominance (under the "
        "convention that larger outcomes are preferred). First-order dominance "
        "means one option is better at every probability level; second-order "
        "adds risk-aversion. Use it to decide between strategies. Run a "
        "simulation that records both outputs first."
    )
)
def compare_distributions(
    output_a: Annotated[str, Field(description="First output (VoseOutput) name.")],
    output_b: Annotated[str, Field(description="Second output name.")],
    workbook_name: Annotated[
        str | None, Field(description="Workbook name. Omit for the active workbook.")
    ] = None,
    max_n: Annotated[
        int, Field(ge=1, le=1_000_000, description="Max samples per output (default 100 000).")
    ] = 100_000,
) -> DistributionComparison:
    bridge = get_bridge()
    a = bridge.get_samples(output_a, workbook_name, max_n=max_n)
    b = bridge.get_samples(output_b, workbook_name, max_n=max_n)
    c = compare_samples(list(a), list(b))
    return DistributionComparison(
        output_a=output_a,
        output_b=output_b,
        sample_size=c["sample_size"],
        paired=c["paired"],
        mean_a=c["mean_a"],
        mean_b=c["mean_b"],
        mean_difference=c["mean_difference"],
        stdev_a=c["stdev_a"],
        stdev_b=c["stdev_b"],
        p_a_greater=c["p_a_greater"],
        first_order_dominance=c["first_order_dominance"],
        second_order_dominance=c["second_order_dominance"],
        percentile_deltas=[PercentileDelta(**d) for d in c["percentile_deltas"]],
    )


def backtest_samples(
    samples: list[float], actuals: list[float], intervals: list[float]
) -> dict[str, Any]:
    """Pure-Python backtest of a predictive sample set against realised
    actuals: PIT calibration, central-interval coverage, and bias.
    Separated from the tool for unit testing without Excel."""
    if not samples:
        raise ModelRiskComputationError("No model samples to validate against.")
    if not actuals:
        raise ModelRiskComputationError("No actuals provided.")
    xs = sorted(samples)
    n = len(xs)
    import bisect

    model_mean = math.fsum(xs) / n
    actuals_mean = math.fsum(actuals) / len(actuals)
    median = _quantile(xs, 0.5)

    # PIT: F_model(actual) for each actual; ~Uniform(0,1) if calibrated.
    pit = [bisect.bisect_right(xs, a) / n for a in actuals]
    mean_pit = math.fsum(pit) / len(pit)
    frac_below_median = sum(1 for a in actuals if a < median) / len(actuals)

    # KS distance of the PIT values from Uniform(0,1).
    sp = sorted(pit)
    m = len(sp)
    ks = 0.0
    for i, p in enumerate(sp):
        ks = max(ks, abs((i + 1) / m - p), abs(p - i / m))

    coverage = []
    for nominal in intervals:
        lo_q = (1.0 - nominal) / 2.0
        lower = _quantile(xs, lo_q)
        upper = _quantile(xs, 1.0 - lo_q)
        inside = sum(1 for a in actuals if lower <= a <= upper) / len(actuals)
        coverage.append(
            {"nominal": nominal, "lower": lower, "upper": upper, "empirical": inside}
        )

    if ks < 0.1:
        verdict = "well calibrated"
    elif mean_pit > 0.6:
        verdict = "model runs low — actuals fall in the upper tail (under-forecasting)"
    elif mean_pit < 0.4:
        verdict = "model runs high — actuals fall in the lower tail (over-forecasting)"
    else:
        verdict = "miscalibrated spread — coverage deviates from nominal"

    return {
        "sample_size": n,
        "n_actuals": len(actuals),
        "model_mean": model_mean,
        "actuals_mean": actuals_mean,
        "bias": actuals_mean - model_mean,
        "mean_pit": mean_pit,
        "pit_uniformity_ks": ks,
        "frac_below_median": frac_below_median,
        "coverage": coverage,
        "verdict": verdict,
    }


@mcp.tool(
    description=(
        "ModelRisk: Backtest a simulation output against realised actuals — "
        "does the model's predicted distribution match what actually happened? "
        "Reports the Probability Integral Transform (PIT, ~0.5 mean and "
        "uniform if calibrated), the empirical coverage of central prediction "
        "intervals (e.g. ~90% of actuals should fall in the 90% interval), and "
        "bias. Pass the historical `actuals` you want to validate against. "
        "Reads the output's per-iteration samples — run the simulation first."
    )
)
def backtest_output(
    output_name: Annotated[str, Field(description="VoseOutput name to validate.")],
    actuals: Annotated[
        list[float], Field(min_length=1, description="Realised historical values.")
    ],
    intervals: Annotated[
        list[float] | None,
        Field(description="Central intervals to check coverage of. Default [0.5,0.8,0.9,0.95]."),
    ] = None,
    workbook_name: Annotated[
        str | None, Field(description="Workbook name. Omit for the active workbook.")
    ] = None,
    max_n: Annotated[
        int, Field(ge=1, le=1_000_000, description="Max samples to read (default 100 000).")
    ] = 100_000,
) -> BacktestResult:
    ivals = intervals or [0.5, 0.8, 0.9, 0.95]
    if any(not 0.0 < v < 1.0 for v in ivals):
        raise ModelRiskComputationError("Each interval must be strictly between 0 and 1.")
    samples = get_bridge().get_samples(output_name, workbook_name, max_n=max_n)
    r = backtest_samples(list(samples), list(actuals), ivals)
    return BacktestResult(
        output_name=output_name,
        sample_size=r["sample_size"],
        n_actuals=r["n_actuals"],
        model_mean=r["model_mean"],
        actuals_mean=r["actuals_mean"],
        bias=r["bias"],
        mean_pit=r["mean_pit"],
        pit_uniformity_ks=r["pit_uniformity_ks"],
        frac_below_median=r["frac_below_median"],
        coverage=[IntervalCoverage(**c) for c in r["coverage"]],
        verdict=r["verdict"],
    )


def _variance(xs: list[float]) -> tuple[float, float]:
    """Return (mean, population variance)."""
    n = len(xs)
    mean = math.fsum(xs) / n
    var = math.fsum((x - mean) ** 2 for x in xs) / n
    return mean, var


@mcp.tool(
    description=(
        "ModelRisk: Split an output's uncertainty into EPISTEMIC (parameter / "
        "knowledge uncertainty — reducible with more data) and ALEATORY "
        "(natural variability — irreducible), via the law of total variance. "
        "ModelRisk has no two-dimensional-simulation worksheet function, so "
        "this approximates it from two runs you provide as two outputs: "
        "`total_output` from a full run (everything varying), and "
        "`conditional_output` from a run with the epistemic/parameter inputs "
        "FROZEN at point estimates (only natural variability left). "
        "Epistemic variance = total - aleatory. Tells you whether collecting "
        "more data (cuts epistemic) or hedging variability (aleatory) is the "
        "lever."
    )
)
def decompose_uncertainty(
    total_output: Annotated[
        str, Field(description="Output name from the full run (all inputs varying).")
    ],
    conditional_output: Annotated[
        str,
        Field(description="Output name from the run with epistemic inputs frozen."),
    ],
    workbook_name: Annotated[
        str | None, Field(description="Workbook name. Omit for the active workbook.")
    ] = None,
    max_n: Annotated[
        int, Field(ge=1, le=1_000_000, description="Max samples to read (default 100 000).")
    ] = 100_000,
) -> UncertaintyDecomposition:
    bridge = get_bridge()
    total = bridge.get_samples(total_output, workbook_name, max_n=max_n)
    cond = bridge.get_samples(conditional_output, workbook_name, max_n=max_n)
    if not total or not cond:
        raise ModelRiskComputationError("Both outputs need samples.")
    _, total_var = _variance(list(total))
    _, aleatory_var = _variance(list(cond))
    epistemic_var = total_var - aleatory_var
    share_e = epistemic_var / total_var if total_var > 0 else 0.0
    share_a = aleatory_var / total_var if total_var > 0 else 0.0

    if share_e > 0.6:
        interp = (
            "Epistemic (parameter) uncertainty dominates — collecting more data "
            "to pin down the inputs is the highest-leverage way to narrow the output."
        )
    elif share_a > 0.6:
        interp = (
            "Aleatory (natural variability) dominates — this is largely "
            "irreducible; more data won't help much, so focus on hedging or capacity."
        )
    else:
        interp = "Epistemic and aleatory uncertainty are comparable; both levers matter."

    return UncertaintyDecomposition(
        total_output=total_output,
        conditional_output=conditional_output,
        total_variance=total_var,
        aleatory_variance=aleatory_var,
        epistemic_variance=epistemic_var,
        epistemic_share=share_e,
        aleatory_share=share_a,
        total_stdev=math.sqrt(total_var),
        aleatory_stdev=math.sqrt(aleatory_var),
        epistemic_stdev=math.sqrt(max(epistemic_var, 0.0)),
        interpretation=interp,
    )


__all__ = [
    "backtest_output",
    "backtest_samples",
    "compare_distributions",
    "compare_samples",
    "compute_correlation_matrix",
    "compute_distribution",
    "compute_tail_risk",
    "decompose_uncertainty",
    "fit_and_rank_distributions",
    "fit_tail",
    "get_tail_risk",
]
