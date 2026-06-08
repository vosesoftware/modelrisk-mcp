"""Integration tests for the §7.4 analysis / workflow tools added in
0.3.2 (compute_distribution, fit_and_rank_distributions, fit_tail,
compute_correlation_matrix, create_aggregate, plan_risk_model).

These exercise the real ModelRisk evaluation paths — inline
`Application.Evaluate`, the scratch-sheet array/fit-object dance, and
formula building — against a live add-in. The suite:

- skips entirely if Excel isn't running (conftest `excel_bridge`);
- soft-skips each add-in-dependent test if the ModelRisk add-in isn't
  *functional* (Vose functions don't resolve);
- writes all test data into a THROWAWAY workbook it creates and closes,
  so the user's open workbooks are never touched.

The pure sample-math tools (get_tail_risk / compare_distributions /
backtest_output / decompose_uncertainty) are covered by unit tests; the
only Excel-touching part they share — `get_samples` — is exercised by
the reading suite.
"""

from __future__ import annotations

import math
import random
from collections.abc import Iterator
from typing import Any

import pytest

from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.tools import analysis, reading, workflows


@pytest.fixture(autouse=True)
def _install_bridge(modelrisk_bridge: ModelRiskBridge) -> Iterator[None]:
    reading.set_bridge_for_testing(modelrisk_bridge)
    yield
    reading.set_bridge_for_testing(None)


@pytest.fixture
def addin(modelrisk_bridge: ModelRiskBridge) -> ModelRiskBridge:
    """Soft-skip when the in-Excel add-in isn't live (Vose functions
    return #NAME?). Distinct from MRService.dll being loaded."""
    if not modelrisk_bridge.probe_addin_functional():
        pytest.skip("ModelRisk add-in is not functional (Vose functions don't resolve).")
    return modelrisk_bridge


@pytest.fixture
def scratch_book(addin: ModelRiskBridge) -> Iterator[tuple[str, Any]]:
    """Create a throwaway workbook, yield (name, sheet), close it
    without saving. Keeps the user's workbooks untouched."""
    app = addin.excel._ensure()
    prev = app.api.DisplayAlerts
    app.api.DisplayAlerts = False
    book = app.books.add()
    try:
        yield book.name, book.sheets[0]
    finally:
        book.close()
        app.api.DisplayAlerts = prev


def test_compute_distribution_cdf_matches_known_value(addin: ModelRiskBridge) -> None:
    # Standard-normal CDF at 1.5 is ~0.93319 — a fixed analytic check.
    r = analysis.compute_distribution(
        metric="cdf", family="Normal", parameters=[0, 1], at=1.5
    )
    assert r.value == pytest.approx(0.9331928, abs=1e-4)


def test_compute_distribution_summary(addin: ModelRiskBridge) -> None:
    s = analysis.compute_distribution(
        metric="summary", family="Lognormal", parameters=[10, 3]
    )
    assert s.mean == pytest.approx(10.0, rel=1e-3)
    assert s.stdev == pytest.approx(3.0, rel=1e-3)
    assert "P50" in s.percentiles


def test_fit_and_rank_picks_a_winner(scratch_book: tuple[str, Any]) -> None:
    name, sheet = scratch_book
    random.seed(7)
    data = [round(math.exp(random.gauss(1.0, 0.5)), 4) for _ in range(60)]
    sheet.range("A1").options(transpose=True).value = data
    r = analysis.fit_and_rank_distributions(
        workbook=name, sheet=sheet.name, data_range="A1:A60",
        families=["Normal", "Lognormal", "Gamma", "Expon"],
    )
    assert r.sample_size == 60
    assert r.best_family is not None
    # Exponential should be the worst fit for lognormal-ish data.
    sics = {c.family: c.sic for c in r.candidates}
    if "Expon" in sics and "Lognormal" in sics:
        assert sics["Lognormal"] < sics["Expon"]


def test_fit_tail_returns_analytic_percentiles(
    scratch_book: tuple[str, Any],
) -> None:
    name, sheet = scratch_book
    random.seed(5)
    tail = [round(random.expovariate(1 / 50.0) + 10, 3) for _ in range(150)]
    sheet.range("A1").options(transpose=True).value = tail
    r = analysis.fit_tail(
        workbook=name, sheet=sheet.name, target_cell="C1", data_range="A1:A150",
        family="GPD",
    )
    assert r.family == "GPD" and r.written is False
    assert r.mean > 0
    assert r.percentiles["P99"] > r.percentiles["P95"]


def test_compute_correlation_matrix(scratch_book: tuple[str, Any]) -> None:
    name, sheet = scratch_book
    random.seed(3)
    rows = []
    for _ in range(80):
        z = random.gauss(0, 1)
        rows.append([round(z + random.gauss(0, 0.3), 4), round(2 * z + random.gauss(0, 0.6), 4)])
    sheet.range("A1").value = rows
    r = analysis.compute_correlation_matrix(
        workbook=name, sheet=sheet.name, data_range="A1:B80"
    )
    assert r.variable_count == 2
    assert r.matrix[0][0] == pytest.approx(1.0, abs=1e-6)
    # The two columns are strongly positively correlated.
    assert r.matrix[0][1] > 0.7


def test_create_aggregate_fft_object_is_readable(
    scratch_book: tuple[str, Any],
) -> None:
    name, sheet = scratch_book
    sheet.range("A1").formula = "=VosePoissonObject(5)"
    sheet.range("A2").formula = "=VoseLognormalObject(1000,400)"
    # Build the FFT aggregate object and read its analytic mean.
    from modelrisk_mcp.tools import building

    out = building.create_aggregate(
        name, sheet.name, "A3",
        frequency_object_cell="A1", severity_object_cell="A2",
        method="FFT", as_object=True, dry_run=False,
    )
    assert out.formula == "=VoseAggregateFFTObject(A1,A2)"
    # E[S] = E[N]*E[X] = 5*1000 = 5000 for a compound Poisson-lognormal.
    summary = analysis.compute_distribution(
        metric="mean", object_sheet=sheet.name, object_cell="A3"
    )
    assert summary.value == pytest.approx(5000.0, rel=1e-2)


def test_plan_risk_model_runs(addin: ModelRiskBridge) -> None:
    try:
        wb_name = addin.excel.get_active_workbook().name
    except Exception:
        pytest.skip("No active workbook in Excel.")
    plan = workflows.plan_risk_model(wb_name)
    assert plan.readiness in {"empty", "needs-outputs", "needs-inputs", "ready"}
    assert plan.output_count >= 0
    assert isinstance(plan.steps, list) and plan.steps
