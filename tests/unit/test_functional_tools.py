"""Unit tests for the four functional-capability tools:
`fit_copula_to_data`, `reverse_stress_test` (both in tools/analysis.py),
and `fit_all_data_and_wire`, `build_model_from_brief` (tools/workflows.py).

The reverse-stress partition logic is pure Python and gets an exact
numeric test via `reverse_stress_profile`. The tools themselves are
tested through the same `set_bridge_for_testing` seam as the rest of the
suite, with a MagicMock bridge whose `.catalogue` is real (the build
orchestrator needs it to render `Vose<Family>(...)`). Real-Excel wiring
is covered by the gated integration tests.
"""

from __future__ import annotations

from collections.abc import Iterator
from unittest.mock import MagicMock

import numpy as np
import pytest

from modelrisk_mcp.bridge.catalogue import load_catalogue
from modelrisk_mcp.errors import ModelRiskComputationError
from modelrisk_mcp.schemas.analysis import CopulaFitRanking, ReverseStressResult
from modelrisk_mcp.tools import analysis, reading, workflows


@pytest.fixture
def bridge() -> Iterator[MagicMock]:
    b = MagicMock()
    b.catalogue = load_catalogue()  # real — build_distribution_formula needs it
    reading.set_bridge_for_testing(b)  # type: ignore[arg-type]
    yield b
    reading.set_bridge_for_testing(None)


def _written(previous: str = "") -> MagicMock:
    r = MagicMock()
    r.previous_formula = previous
    return r


# ----------------------------------------------------------------------
# reverse_stress_profile — pure core
# ----------------------------------------------------------------------


class TestReverseStressProfile:
    def test_ranks_the_true_driver_first(self) -> None:
        rng = np.random.default_rng(0)
        n = 4000
        driver = rng.normal(0.0, 1.0, n)
        noise = rng.normal(0.0, 1.0, n)
        output = driver * 3.0 + noise * 0.3  # driver dominates the output
        mask = output >= np.quantile(output, 0.9)
        drivers, scenario = analysis.reverse_stress_profile(
            output, {"driver": driver, "noise": noise}, mask
        )
        assert drivers[0]["input_name"] == "driver"
        assert drivers[0]["rank"] == 1
        # The driver shifts far more (in SDs) than noise inside the breach.
        assert abs(drivers[0]["mean_shift_sd"]) > abs(drivers[1]["mean_shift_sd"])
        assert drivers[0]["mean_shift_sd"] > 1.0  # strongly positive
        assert scenario["driver"] > scenario["noise"]

    def test_no_breach_yields_no_drivers(self) -> None:
        arr = np.ones(100)
        mask = np.zeros(100, dtype=bool)
        drivers, scenario = analysis.reverse_stress_profile(arr, {"x": arr}, mask)
        assert drivers == []
        assert scenario == {}


# ----------------------------------------------------------------------
# fit_copula_to_data
# ----------------------------------------------------------------------


class TestFitCopulaToData:
    def _returns(self) -> tuple:
        return (
            [
                {"family": "CopulaMultiNormal", "aic": 100.0, "sic": 105.0, "hqic": 102.0},
                {"family": "CopulaMultiClayton", "aic": 90.0, "sic": 95.0, "hqic": 92.0},
            ],
            [{"family": "CopulaMultiT", "reason": "fit failed"}],
            999,
        )

    def test_ranks_by_criterion_and_reports_tail_dependence(
        self, bridge: MagicMock
    ) -> None:
        bridge.fit_and_rank_copulas.return_value = self._returns()
        r = analysis.fit_copula_to_data("m.xlsx", "Sheet1", "A2:C500")
        assert isinstance(r, CopulaFitRanking)
        # Clayton has the lower SIC → ranked first.
        assert r.best_family == "CopulaMultiClayton"
        assert r.candidates[0].rank == 1
        assert r.candidates[0].tail_dependence == "lower"
        assert r.n_variables == 3
        assert r.sample_size == 999 // 3
        assert "lower" in r.note.lower()
        assert r.skipped and r.skipped[0]["family"] == "CopulaMultiT"
        # Range was sheet-qualified for the scratch-sheet fit.
        assert bridge.fit_and_rank_copulas.call_args.args[0] == "'Sheet1'!A2:C500"

    def test_criterion_switch_changes_winner(self, bridge: MagicMock) -> None:
        # AIC also favours Clayton here; flip the numbers to prove sorting.
        bridge.fit_and_rank_copulas.return_value = (
            [
                {"family": "CopulaMultiNormal", "aic": 10.0, "sic": 999.0, "hqic": 1.0},
                {"family": "CopulaMultiGumbel", "aic": 999.0, "sic": 10.0, "hqic": 2.0},
            ],
            [],
            600,
        )
        by_aic = analysis.fit_copula_to_data("m.xlsx", "S", "A1:B300", criterion="AIC")
        by_sic = analysis.fit_copula_to_data("m.xlsx", "S", "A1:B300", criterion="SIC")
        assert by_aic.best_family == "CopulaMultiNormal"
        assert by_sic.best_family == "CopulaMultiGumbel"
        assert by_sic.candidates[0].tail_dependence == "upper"

    def test_bad_criterion_raises(self, bridge: MagicMock) -> None:
        with pytest.raises(ModelRiskComputationError, match="criterion"):
            analysis.fit_copula_to_data("m.xlsx", "S", "A1:B9", criterion="XYZ")


# ----------------------------------------------------------------------
# reverse_stress_test
# ----------------------------------------------------------------------


class TestReverseStressTest:
    def _wire(self, bridge: MagicMock, out: np.ndarray, drv: np.ndarray) -> None:
        data = {"NPV": out.tolist(), "Demand": drv.tolist()}

        def samples(name: str, wb: object = None, max_n: int = 100_000) -> list:
            return data[name]

        bridge.get_samples.side_effect = samples
        bridge.list_vmrs_variables.return_value = [
            {"name": "NPV", "kind": "output", "var_id": 1, "iterations": len(out)},
            {"name": "Demand", "kind": "input", "var_id": 2, "iterations": len(drv)},
        ]

    def test_threshold_above_partitions_and_ranks(self, bridge: MagicMock) -> None:
        out = np.concatenate([np.zeros(90), np.full(10, 100.0)])
        drv = np.concatenate([np.zeros(90), np.full(10, 5.0)])  # high in breach
        self._wire(bridge, out, drv)
        r = analysis.reverse_stress_test("NPV", threshold=50.0, direction="above")
        assert isinstance(r, ReverseStressResult)
        assert r.breach_count == 10
        assert r.iterations == 100
        assert r.breach_probability == pytest.approx(0.1)
        assert r.drivers[0].input_name == "Demand"
        assert r.drivers[0].mean_shift_sd > 0
        assert r.scenario is not None
        assert "Demand" in r.scenario.input_values

    def test_percentile_threshold(self, bridge: MagicMock) -> None:
        out = np.arange(1000, dtype=float)
        drv = np.arange(1000, dtype=float)
        self._wire(bridge, out, drv)
        r = analysis.reverse_stress_test(
            "NPV", threshold_percentile=0.95, direction="above"
        )
        # Top 5% of 0..999 → threshold ~949; ~50 breaches.
        assert 40 <= r.breach_count <= 60
        assert r.threshold == pytest.approx(np.quantile(out, 0.95))

    def test_requires_a_threshold(self, bridge: MagicMock) -> None:
        bridge.get_samples.return_value = [1.0, 2.0, 3.0]
        with pytest.raises(ModelRiskComputationError, match="threshold"):
            analysis.reverse_stress_test("NPV")


# ----------------------------------------------------------------------
# fit_all_data_and_wire
# ----------------------------------------------------------------------


class TestFitAllDataAndWire:
    def _wire_fits(self, bridge: MagicMock) -> None:
        bridge.fit_and_rank.return_value = (
            [{"family": "Normal", "aic": 1.0, "sic": 1.0, "hqic": 1.0}], [], 500
        )
        bridge.fit_and_rank_copulas.return_value = (
            [{"family": "CopulaMultiNormal", "aic": 1.0, "sic": 1.0, "hqic": 1.0}], [], 1000
        )

    def _cols(self) -> list[dict[str, str]]:
        return [
            {"input_name": "X", "target_cell": "E2"},
            {"input_name": "Y", "target_cell": "F2"},
        ]

    def test_dry_run_plans_but_does_not_write(self, bridge: MagicMock) -> None:
        self._wire_fits(bridge)
        r = workflows.fit_all_data_and_wire(
            "m.xlsx", "S", "A2:B500", self._cols(), copula_anchor="H2", dry_run=True
        )
        assert r.dry_run is True
        assert r.copula_family == "CopulaMultiNormal"
        assert not any(c.written for c in r.columns)
        bridge.safe_write_cell.assert_not_called()
        # Marginals are wired to the copula's U cells (H2, H3).
        assert "'S'!H2" in r.columns[0].formula
        assert "'S'!H3" in r.columns[1].formula
        assert "VoseNormalFit('S'!A2:A500" in r.columns[0].formula

    def test_writes_copula_block_plus_marginals(self, bridge: MagicMock) -> None:
        self._wire_fits(bridge)
        bridge.safe_write_cell.return_value = _written()
        r = workflows.fit_all_data_and_wire(
            "m.xlsx", "S", "A2:B500", self._cols(),
            copula_anchor="H2", dry_run=False, run=False,
        )
        assert all(c.written for c in r.columns)
        assert r.rolled_back is False
        # 1 copula block + 2 marginals.
        assert bridge.safe_write_cell.call_count == 3
        first_formula = bridge.safe_write_cell.call_args_list[0].args[1]
        assert first_formula.startswith("=VoseCopulaMultiNormalFit(")

    def test_mid_build_failure_rolls_back(self, bridge: MagicMock) -> None:
        self._wire_fits(bridge)
        seen: list = []

        def sw(ref: object, formula: str, *, allow_overwrite_non_vose: bool = False):
            seen.append(ref)
            if len(seen) == 2:  # fail on the first marginal
                raise RuntimeError("boom")
            return _written("old")

        bridge.safe_write_cell.side_effect = sw
        r = workflows.fit_all_data_and_wire(
            "m.xlsx", "S", "A2:B500", self._cols(), copula_anchor="H2", dry_run=False
        )
        assert r.rolled_back is True
        assert not any(c.written for c in r.columns)
        # The one successful write (the copula block) was rolled back.
        assert bridge.excel.write_cell.called


# ----------------------------------------------------------------------
# build_model_from_brief
# ----------------------------------------------------------------------


class TestBuildModelFromBrief:
    def _inputs(self) -> list[dict]:
        return [
            {
                "cell": "B4",
                "input_name": "Demand",
                "function_name": "VoseNormal",
                "parameters": [{"value": 100}, {"value": 10}],
            }
        ]

    def test_dry_run_previews_formulas(self, bridge: MagicMock) -> None:
        bridge.excel.get_cell.return_value = MagicMock(formula="=A1*B1", value=None)
        r = workflows.build_model_from_brief(
            "m.xlsx", "S", inputs=self._inputs(),
            outputs=[{"cell": "B12", "output_name": "NPV"}], dry_run=True,
        )
        assert r.dry_run is True
        assert r.change_set_size == 0
        assert r.inputs_built[0].input_name == "Demand"
        assert "VoseNormal(100,10)" in r.inputs_built[0].formula
        assert r.outputs_wrapped == ["NPV"]
        bridge.safe_write_cell.assert_not_called()

    def test_builds_and_simulates(self, bridge: MagicMock) -> None:
        bridge.excel.get_cell.return_value = MagicMock(formula="=A1*B1", value=None)
        bridge.safe_write_cell.return_value = _written()
        res = MagicMock(output_name="NPV", mean=42.0, percentiles={0.1: 10, 0.5: 40, 0.9: 80})
        bridge.get_simulation_results.return_value = [res]
        r = workflows.build_model_from_brief(
            "m.xlsx", "S", inputs=self._inputs(),
            outputs=[{"cell": "B12", "output_name": "NPV"}],
            dry_run=False, run=True,
        )
        assert r.simulated is True
        assert r.change_set_size == 2  # 1 output wrap + 1 input
        assert r.headline["NPV"]["p90"] == 80
        bridge.run_simulation.assert_called_once()

    def test_failure_rolls_back_whole_build(self, bridge: MagicMock) -> None:
        bridge.excel.get_cell.return_value = MagicMock(formula="=A1", value=None)

        def sw(ref: object, formula: str, *, allow_overwrite_non_vose: bool = False):
            if getattr(ref, "cell", None) == "B4":
                raise RuntimeError("bad input")
            return _written("prev")

        bridge.safe_write_cell.side_effect = sw
        r = workflows.build_model_from_brief(
            "m.xlsx", "S", inputs=self._inputs(),
            outputs=[{"cell": "B12", "output_name": "NPV"}],
            dry_run=False, run=False,
        )
        assert r.rolled_back is True
        assert r.change_set_size == 0
        # The output wrap that did succeed was rolled back.
        assert bridge.excel.write_cell.called


_ = pytest
