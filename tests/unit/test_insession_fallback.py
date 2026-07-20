"""Tests for the in-session (no-MRService.dll) results fallback.

The live semantics these tests encode were verified against real Excel
+ ModelRisk on 2026-07-20: `VoseSimValue(ref, k)` reads the live sample
store any time after a run (1-based k, "No simulation results" text
when out of range), and the extracted samples are identical to the
`.vmrs` content. See bridge/insession.py's module docstring.
"""

from __future__ import annotations

import math
from unittest.mock import MagicMock, patch

import pytest

from modelrisk_mcp.bridge import insession as ins_mod
from modelrisk_mcp.bridge.insession import InSessionSampleReader, _qualified_ref
from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.bridge.results import (
    correlation_matrix_from_samples,
    sensitivity_from_samples,
    simulation_result_from_samples,
)
from modelrisk_mcp.errors import (
    ModelRiskNotLoadedError,
    ReadOnlyModeError,
    SimulationFailedError,
)
from modelrisk_mcp.schemas.workbook import CellRef, ModelRiskInput, ModelRiskOutput

# ----------------------------------------------------------------------
# A fake reader that simulates a session store of a given size without
# any COM: _probe and _fill_and_read are the only Excel-touching sample
# paths, so overriding them exercises all the detection/orchestration
# logic for real.
# ----------------------------------------------------------------------


class _FakeStoreReader(InSessionSampleReader):
    def __init__(self, store: dict[str, list[float]]) -> None:
        super().__init__(excel=MagicMock())
        self.store = store  # ref -> samples
        self.probe_calls = 0

    # No real Excel: stub the COM plumbing.
    def _book(self, workbook):  # type: ignore[override]
        book = MagicMock()
        book.api.Names.side_effect = Exception("no names")
        return MagicMock(), book

    def _scratch_area(self, book, sheet_name):  # type: ignore[override]
        return MagicMock()

    def _clear_scratch(self, book, prev_saved):  # type: ignore[override]
        pass

    def _probe(self, scratch, ref, k):  # type: ignore[override]
        self.probe_calls += 1
        samples = self.store.get(ref, [])
        return 1 <= k <= len(samples)

    def _fill_and_read(self, scratch, ref, n):  # type: ignore[override]
        samples = self.store.get(ref, [])
        return tuple(samples[:n])


class TestQualifiedRef:
    def test_plain_and_dollar_forms(self) -> None:
        assert _qualified_ref("Model", "B2") == "'Model'!$B$2"
        assert _qualified_ref("Model", "$B$2") == "'Model'!$B$2"

    def test_sheet_quote_escaped(self) -> None:
        assert _qualified_ref("Tim's", "A1") == "'Tim''s'!$A$1"

    def test_range_rejected(self) -> None:
        with pytest.raises(SimulationFailedError):
            _qualified_ref("Model", "B2:B5")


class TestDetection:
    def test_hint_verified_and_used(self) -> None:
        r = _FakeStoreReader({"'M'!$B$2": [1.0] * 500})
        out = r.read_many("Book1", [("Total", "M", "B2")], n_hint=500)
        assert len(out["Total"]) == 500
        # hint path: probe(1), probe(500), probe(501) — no search.
        assert r.probe_calls == 3

    def test_stale_hint_falls_through_to_search(self) -> None:
        r = _FakeStoreReader({"'M'!$B$2": [2.0] * 300})
        out = r.read_many("Book1", [("Total", "M", "B2")], n_hint=500)
        assert len(out["Total"]) == 300  # true size, not the stale hint

    def test_no_hint_search_finds_exact_size(self) -> None:
        r = _FakeStoreReader({"'M'!$B$2": [3.0] * 1234})
        out = r.read_many("Book1", [("Total", "M", "B2")])
        assert len(out["Total"]) == 1234

    def test_max_n_truncates(self) -> None:
        r = _FakeStoreReader({"'M'!$B$2": list(map(float, range(1000)))})
        out = r.read_many(
            "Book1", [("Total", "M", "B2")], max_n=100, n_hint=1000
        )
        assert len(out["Total"]) == 100
        assert out["Total"][0] == 0.0

    def test_empty_store_raises_with_guidance(self) -> None:
        r = _FakeStoreReader({})
        with pytest.raises(SimulationFailedError, match="in-session"):
            r.read_many("Book1", [("Total", "M", "B2")])

    def test_second_target_reuses_detected_n(self) -> None:
        r = _FakeStoreReader({
            "'M'!$B$2": [1.0] * 400,
            "'M'!$C$2": [2.0] * 400,
        })
        out = r.read_many(
            "Book1",
            [("Total", "M", "B2"), ("Cost", "M", "C2")],
            n_hint=400,
        )
        assert set(out) == {"Total", "Cost"}
        # detection ran once (3 probes) — the second target reused N.
        assert r.probe_calls == 3


class TestReadOnlyRefusal:
    def test_read_only_blocks_scratch_writes(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setenv("MODELRISK_MCP_READ_ONLY", "1")
        r = _FakeStoreReader({"'M'!$B$2": [1.0]})
        with pytest.raises(ReadOnlyModeError, match="scratch"):
            r.read_many("Book1", [("Total", "M", "B2")])


class TestSimOptParsing:
    def test_literal_parsed(self) -> None:
        book = MagicMock()
        book.api.Names.return_value.RefersTo = "=1000"
        assert InSessionSampleReader._simopt_samples(book) == 1000

    def test_cell_ref_ignored(self) -> None:
        book = MagicMock()
        book.api.Names.return_value.RefersTo = "=Sheet1!$A$1"
        assert InSessionSampleReader._simopt_samples(book) is None

    def test_missing_name_ignored(self) -> None:
        book = MagicMock()
        book.api.Names.side_effect = Exception("not found")
        assert InSessionSampleReader._simopt_samples(book) is None


class TestSampleAssembly:
    def test_simulation_result_stats(self) -> None:
        samples = [1.0, 2.0, 3.0, 4.0]
        res = simulation_result_from_samples("Total", samples)
        assert res.iterations == 4
        assert res.mean == pytest.approx(2.5)
        assert res.variance == pytest.approx(1.25)  # population
        assert res.stdev == pytest.approx(math.sqrt(1.25))
        assert res.min == 1.0 and res.max == 4.0
        assert res.percentiles[0.50] == pytest.approx(2.5)
        assert "in-session" in res.source

    def test_empty_samples_raise(self) -> None:
        with pytest.raises(SimulationFailedError):
            simulation_result_from_samples("Total", [])

    def test_correlation_matrix(self) -> None:
        x = [1.0, 2.0, 3.0, 4.0]
        y = [2.0, 4.0, 6.0, 8.0]
        m = correlation_matrix_from_samples([("a", x), ("b", y)])
        assert m.names == ["a", "b"]
        assert m.pearson[0][1] == pytest.approx(1.0)
        assert m.spearman[0][1] == pytest.approx(1.0)
        assert m.iterations == 4
        assert "in-session" in m.source

    def test_sensitivity_sorted_by_abs_correlation(self) -> None:
        out = [1.0, 2.0, 3.0, 4.0, 5.0]
        strong = [1.1, 2.0, 2.9, 4.2, 5.0]
        weak = [3.0, 1.0, 4.0, 2.0, 3.5]
        rank = sensitivity_from_samples(
            "Total", out, [("weak", weak), ("strong", strong)]
        )
        assert next(e.input_name for e in rank.entries) == "strong"
        assert rank.iterations == 5
        assert "in-session" in rank.source


# ----------------------------------------------------------------------
# Bridge-level dispatch: DLL error → in-session fallback
# ----------------------------------------------------------------------


def _bridge_with_failing_dll() -> ModelRiskBridge:
    excel = MagicMock()
    excel.get_active_workbook.return_value.name = "Book1"
    excel.list_workbooks.return_value = []
    results = MagicMock()
    dll_err = ModelRiskNotLoadedError("MRService.dll not found (test)")
    results.get_simulation_results.side_effect = dll_err
    results.get_samples.side_effect = dll_err
    results.get_correlation_matrix.side_effect = dll_err
    results.get_sensitivity_ranking.side_effect = dll_err
    results.list_variables.side_effect = dll_err
    bridge = ModelRiskBridge(
        excel=excel, results=results, mrservice=MagicMock(),
        simulation=MagicMock(),
    )
    return bridge


def _out(name: str, cell: str) -> ModelRiskOutput:
    return ModelRiskOutput(
        ref=CellRef(workbook="Book1", sheet="Model", cell=cell),
        name=name,
        formula=f'=VoseOutput("{name}")+1',
    )


def _inp(name: str, cell: str) -> ModelRiskInput:
    return ModelRiskInput(
        ref=CellRef(workbook="Book1", sheet="Model", cell=cell),
        name=name,
        formula=f'=VoseInput("{name}")',
    )


class TestBridgeFallbackDispatch:
    def test_get_simulation_results_falls_back(self) -> None:
        bridge = _bridge_with_failing_dll()
        with (
            patch.object(
                bridge, "list_outputs", return_value=[_out("Total", "B2")]
            ),
            patch.object(bridge, "list_inputs", return_value=[]),
            patch.object(
                bridge._insession, "read_many",
                return_value={"Total": (1.0, 2.0, 3.0)},
            ) as read_many,
        ):
            results = bridge.get_simulation_results("Book1")
        assert len(results) == 1
        assert results[0].output_name == "Total"
        assert results[0].iterations == 3
        assert "in-session" in results[0].source
        targets = read_many.call_args.args[1]
        assert targets == [("Total", "Model", "B2")]

    def test_get_samples_falls_back_and_matches(self) -> None:
        bridge = _bridge_with_failing_dll()
        with (
            patch.object(
                bridge, "list_outputs", return_value=[_out("Total", "B2")]
            ),
            patch.object(bridge, "list_inputs", return_value=[]),
            patch.object(
                bridge._insession, "read_many",
                return_value={"Total": (9.0, 8.0)},
            ),
        ):
            assert bridge.get_samples("Total", "Book1") == [9.0, 8.0]

    def test_get_sensitivity_falls_back(self) -> None:
        bridge = _bridge_with_failing_dll()
        samples = {
            "Total": tuple(float(i) for i in range(10)),
            "Cost": tuple(float(i) * 2 for i in range(10)),
        }
        with (
            patch.object(
                bridge, "list_outputs", return_value=[_out("Total", "B2")]
            ),
            patch.object(
                bridge, "list_inputs", return_value=[_inp("Cost", "B1")]
            ),
            patch.object(
                bridge._insession, "read_many", return_value=dict(samples)
            ),
        ):
            rank = bridge.get_sensitivity_ranking("Total", "Book1")
        assert rank.output_name == "Total"
        assert [e.input_name for e in rank.entries] == ["Cost"]
        assert "in-session" in rank.source

    def test_no_matching_cells_raises_actionable_error(self) -> None:
        bridge = _bridge_with_failing_dll()
        with (
            patch.object(bridge, "list_outputs", return_value=[]),
            patch.object(bridge, "list_inputs", return_value=[]),
            pytest.raises(SimulationFailedError, match=r"vosesoftware\.com"),
        ):
            bridge.get_simulation_results("Book1", ["Total"])

    def test_list_vmrs_variables_falls_back(self) -> None:
        bridge = _bridge_with_failing_dll()
        with (
            patch.object(
                bridge, "list_outputs", return_value=[_out("Total", "B2")]
            ),
            patch.object(
                bridge, "list_inputs", return_value=[_inp("Cost", "B1")]
            ),
            patch.object(
                bridge._insession, "read_many",
                return_value={"Total": (1.0,) * 5, "Cost": (2.0,) * 5},
            ),
        ):
            entries = bridge.list_vmrs_variables("Book1")
        by_name = {e["name"]: e for e in entries}
        assert by_name["Total"]["kind"] == "output"
        assert by_name["Cost"]["kind"] == "input"
        assert all(e["iterations"] == 5 for e in entries)
        assert all(e["source"] == "in-session" for e in entries)

    def test_vmrs_path_untouched_when_dll_present(self) -> None:
        """When the DLL path works, the fallback must never run."""
        excel = MagicMock()
        excel.get_active_workbook.return_value.name = "Book1"
        excel.list_workbooks.return_value = []
        results = MagicMock()
        results.get_simulation_results.return_value = []
        bridge = ModelRiskBridge(
            excel=excel, results=results, mrservice=MagicMock(),
            simulation=MagicMock(),
        )
        with (
            patch.object(bridge, "list_outputs", return_value=[]),
            patch.object(bridge, "list_inputs", return_value=[]),
            patch.object(bridge._insession, "read_many") as read_many,
        ):
            bridge.get_simulation_results("Book1", ["Total"])
        read_many.assert_not_called()


class TestRunRecordsIterationHint:
    def test_hint_recorded_after_run(self) -> None:
        assert (
            _bridge_with_failing_dll()._last_sim_iterations is None
        )
        # The attribute is set inside run_simulation right after the
        # vmrs pin; exercising the full run path needs Excel, so this
        # is covered by the gated integration test — here we only pin
        # the initial state.


# ----------------------------------------------------------------------
# read_many uses ExcelBridge lazily — no COM at import time
# ----------------------------------------------------------------------


def test_module_importable_without_excel() -> None:
    assert ins_mod.InSessionSampleReader is not None
