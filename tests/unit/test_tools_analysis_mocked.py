"""Tests for the quantitative-analysis tools (`tools/analysis.py`).

`compute_tail_risk` is pure Python and gets exact numeric tests. The
three MCP tools are tested through the same `set_bridge_for_testing`
seam as the reading tools, with a MagicMock bridge — except
`compute_distribution`'s family path, which needs the real catalogue to
build `Vose<Family>Object(...)`, so the fixture wires in a real one.
"""

from __future__ import annotations

from collections.abc import Iterator
from unittest.mock import MagicMock

import pytest

from modelrisk_mcp.bridge.catalogue import load_catalogue
from modelrisk_mcp.errors import ModelRiskComputationError
from modelrisk_mcp.schemas.analysis import (
    DistributionProperty,
    DistributionSummary,
    FitRanking,
    TailRiskResult,
)
from modelrisk_mcp.tools import analysis, reading


@pytest.fixture
def bridge() -> Iterator[MagicMock]:
    b = MagicMock()
    b.catalogue = load_catalogue()  # real — build_distribution_formula needs it
    reading.set_bridge_for_testing(b)  # type: ignore[arg-type]
    yield b
    reading.set_bridge_for_testing(None)


# ----------------------------------------------------------------------
# compute_distribution
# ----------------------------------------------------------------------


class TestComputeDistribution:
    def test_cdf_builds_voseprob_true(self, bridge: MagicMock) -> None:
        bridge.evaluate_number.return_value = 0.9332
        r = analysis.compute_distribution(
            metric="cdf", family="Normal", parameters=[0, 1], at=1.5
        )
        assert isinstance(r, DistributionProperty)
        assert r.value == 0.9332
        assert r.metric == "cdf" and r.at == 1.5
        expr = bridge.evaluate_number.call_args.args[0]
        assert expr.startswith("VoseProb(1.5,") and expr.endswith(",TRUE)")
        assert "VoseNormalObject(" in expr

    def test_exceedance_is_one_minus_cdf(self, bridge: MagicMock) -> None:
        bridge.evaluate_number.return_value = 0.05
        r = analysis.compute_distribution(
            metric="exceedance", family="Normal", parameters=[0, 1], at=1.645
        )
        assert r.expression.startswith("1-VoseProb(")

    def test_quantile_via_object_cell_no_catalogue(self, bridge: MagicMock) -> None:
        bridge.evaluate_number.return_value = 1.96
        r = analysis.compute_distribution(
            metric="quantile", at=0.975, object_sheet="Sheet1", object_cell="B2"
        )
        assert r.value == 1.96
        assert bridge.evaluate_number.call_args.args[0] == "VosePercentile('Sheet1'!B2,0.975)"

    def test_stdev_is_sqrt_variance(self, bridge: MagicMock) -> None:
        bridge.evaluate_number.return_value = 9.0
        r = analysis.compute_distribution(metric="stdev", family="Normal", parameters=[0, 3])
        assert r.expression.startswith("SQRT(VoseVariance(")

    def test_summary_returns_moments_and_ladder(self, bridge: MagicMock) -> None:
        bridge.evaluate_number.return_value = 4.0  # variance=4 -> stdev=2
        r = analysis.compute_distribution(metric="summary", family="Normal", parameters=[0, 2])
        assert isinstance(r, DistributionSummary)
        assert r.variance == 4.0 and r.stdev == 2.0
        assert set(r.percentiles) == {"P1", "P5", "P10", "P25", "P50", "P75", "P90", "P95", "P99"}

    def test_missing_at_raises(self, bridge: MagicMock) -> None:
        with pytest.raises(ModelRiskComputationError):
            analysis.compute_distribution(metric="cdf", family="Normal", parameters=[0, 1])

    def test_no_family_no_cell_raises(self, bridge: MagicMock) -> None:
        with pytest.raises(ModelRiskComputationError):
            analysis.compute_distribution(metric="mean")

    def test_unknown_metric_raises(self, bridge: MagicMock) -> None:
        with pytest.raises(ModelRiskComputationError):
            analysis.compute_distribution(
                metric="bogus", object_cell="B2", at=1.0
            )


# ----------------------------------------------------------------------
# fit_and_rank_distributions
# ----------------------------------------------------------------------


class TestFitAndRank:
    def _scored(self) -> tuple:
        return (
            [
                {"family": "Normal", "aic": 194.6, "sic": 198.6, "hqic": 196.1},
                {"family": "Gamma", "aic": 188.5, "sic": 192.5, "hqic": 190.0},
            ],
            [{"family": "Beta", "reason": "fit failed"}],
            60,
        )

    def test_ranks_ascending_by_sic(self, bridge: MagicMock) -> None:
        bridge.fit_and_rank.return_value = self._scored()
        r = analysis.fit_and_rank_distributions(
            workbook="m.xlsx", sheet="Sheet1", data_range="A1:A60"
        )
        assert isinstance(r, FitRanking)
        assert r.best_family == "Gamma"
        assert [c.family for c in r.candidates] == ["Gamma", "Normal"]
        assert r.candidates[0].rank == 1 and r.candidates[1].rank == 2
        assert r.skipped == [{"family": "Beta", "reason": "fit failed"}]

    def test_qualifies_range_and_passes_families(self, bridge: MagicMock) -> None:
        bridge.fit_and_rank.return_value = ([], [], 0)
        analysis.fit_and_rank_distributions(
            workbook="m.xlsx", sheet="Data", data_range="B2:B99", families=["Normal"]
        )
        args, kwargs = bridge.fit_and_rank.call_args
        assert args[0] == "'Data'!B2:B99"
        assert args[1] == ["Normal"]
        assert kwargs["workbook"] == "m.xlsx"

    def test_criterion_aic_changes_sort(self, bridge: MagicMock) -> None:
        # Construct a case where AIC and SIC disagree on the winner.
        bridge.fit_and_rank.return_value = (
            [
                {"family": "A", "aic": 10.0, "sic": 5.0, "hqic": 7.0},
                {"family": "B", "aic": 1.0, "sic": 6.0, "hqic": 7.0},
            ],
            [],
            10,
        )
        by_sic = analysis.fit_and_rank_distributions(
            workbook="m.xlsx", sheet="S", data_range="A1:A10", criterion="SIC"
        )
        by_aic = analysis.fit_and_rank_distributions(
            workbook="m.xlsx", sheet="S", data_range="A1:A10", criterion="AIC"
        )
        assert by_sic.best_family == "A"
        assert by_aic.best_family == "B"

    def test_bad_criterion_raises(self, bridge: MagicMock) -> None:
        with pytest.raises(ModelRiskComputationError):
            analysis.fit_and_rank_distributions(
                workbook="m.xlsx", sheet="S", data_range="A1:A10", criterion="ZZZ"
            )


# ----------------------------------------------------------------------
# compute_tail_risk (pure)
# ----------------------------------------------------------------------


class TestComputeTailRisk:
    def test_upper_var_cvar_and_threshold(self) -> None:
        xs = [float(i) for i in range(1, 101)]  # 1..100
        out = compute = analysis.compute_tail_risk(
            xs, alphas=[0.95], thresholds=[90.0], tail="upper"
        )
        assert out["sample_size"] == 100
        assert out["mean"] == pytest.approx(50.5)
        m = out["tail_metrics"][0]
        assert m["var"] == pytest.approx(95.05)
        assert m["cvar"] == pytest.approx(98.0)  # mean of 96..100
        thr = out["threshold_probabilities"][0]
        assert thr["p_above"] == pytest.approx(0.10)  # 91..100
        assert thr["p_at_or_below"] == pytest.approx(0.90)
        assert compute is out

    def test_lower_tail_uses_opposite_end(self) -> None:
        xs = [float(i) for i in range(1, 101)]
        out = analysis.compute_tail_risk(xs, alphas=[0.95], thresholds=[], tail="lower")
        m = out["tail_metrics"][0]
        assert m["var"] == pytest.approx(5.95)
        assert m["cvar"] == pytest.approx(3.0)  # mean of 1..5

    def test_empty_samples_raise(self) -> None:
        with pytest.raises(ModelRiskComputationError):
            analysis.compute_tail_risk([], alphas=[0.95], thresholds=[], tail="upper")


# ----------------------------------------------------------------------
# get_tail_risk tool
# ----------------------------------------------------------------------


class TestGetTailRisk:
    def test_reads_samples_and_builds_result(self, bridge: MagicMock) -> None:
        bridge.get_samples.return_value = [float(i) for i in range(1, 101)]
        r = analysis.get_tail_risk(output_name="NPV", alphas=[0.95], thresholds=[90.0])
        assert isinstance(r, TailRiskResult)
        assert r.output_name == "NPV" and r.sample_size == 100
        assert r.tail == "upper"
        assert r.tail_metrics[0].cvar == pytest.approx(98.0)
        bridge.get_samples.assert_called_once()

    def test_bad_tail_raises(self, bridge: MagicMock) -> None:
        with pytest.raises(ModelRiskComputationError):
            analysis.get_tail_risk(output_name="NPV", tail="sideways")

    def test_bad_alpha_raises(self, bridge: MagicMock) -> None:
        with pytest.raises(ModelRiskComputationError):
            analysis.get_tail_risk(output_name="NPV", alphas=[1.5])
