"""End-to-end integration test for the v0.3 run_simulation pipeline.

What this exercises:

1. Build a minimal Monte Carlo model in Excel programmatically (two
   cells: one input `X = N(0,1)`, one output `Y = 2X`).
2. Save the workbook to a tmp_path so the `.vmrs` lands there and not
   on the user's Desktop.
3. Call `bridge.run_simulation()` — the XLL command path
   (`VoseStartSimulCustom12` + `VoseGetDataSZ12`). Blocks until the
   simulation completes.
4. Verify the `.vmrs` appears on disk.
5. Call `bridge.get_simulation_results()` and confirm the stats match
   what `Y = 2 * N(0,1)` should produce.
6. Exercise `list_vmrs_variables` — should find both X and Y.
7. Exercise `get_samples(Y)` — should return ~1000 floats with the
   right mean / stdev.
8. Exercise `diagnose_workbook` — should report a clean session.

Every test in this file is gated on Excel + ModelRisk XLL + MRService
.dll all being installed and reachable. If any piece is missing, the
test skips with a clear reason rather than failing.

Run with: `uv run pytest tests/integration/test_e2e_run_simulation.py -v`
"""

from __future__ import annotations

import math
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.errors import SimulationFailedError

# Skipping is handled by the conftest fixtures (`excel_bridge` /
# `modelrisk_bridge` call pytest.skip() when their preconditions
# aren't met) — no module-level marker needed.


# ---------------------------------------------------------------------------
# Test workbook fixture
# ---------------------------------------------------------------------------


@pytest.fixture(scope="function")
def test_workbook(
    modelrisk_bridge: ModelRiskBridge, tmp_path: Path
) -> Iterator[tuple[str, Path]]:
    """Create a tiny Monte Carlo model in a fresh workbook saved to
    tmp_path. Yields (workbook_name, workbook_path). Closes the
    workbook without saving on teardown — the .xlsx file remains in
    tmp_path until pytest cleans the directory."""

    wb_path = tmp_path / "e2e_test.xlsx"

    # `xw.Book()` creates an unsaved book; `SaveAs` persists it.
    # Doing this through the bridge's connected app keeps everything
    # in one Excel instance.
    excel = modelrisk_bridge.excel
    app = excel._app  # bridge-internal; the connected xlwings App
    assert app is not None, "ExcelBridge must be connected"

    book = app.books.add()
    try:
        sheet = book.sheets[0]
        sheet.name = "Model"
        # A1: input "X" sampled from N(0, 1)
        sheet.range("A1").formula = '=VoseInput("X")+VoseNormal(0,1)'
        # B1: output "Y" = 2 * A1 ; mean 0, stdev 2
        sheet.range("B1").formula = '=VoseOutput("Y")+A1*2'
        # SaveAs to tmp_path. xlwings' .save() takes a path on first save.
        book.save(str(wb_path))
        # `book.name` after save is the basename of the path.
        yield book.name, wb_path
    finally:
        try:
            book.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


class TestRunSimulationEndToEnd:
    def test_run_simulation_produces_vmrs(
        self,
        modelrisk_bridge: ModelRiskBridge,
        test_workbook: tuple[str, Path],
    ) -> None:
        wb_name, wb_path = test_workbook

        result = modelrisk_bridge.run_simulation(
            workbook=wb_name, samples=1000, seed=1,
        )
        assert result.workbook_name == wb_name
        assert result.iterations == 1000
        # .vmrs landed next to the workbook.
        vmrs = Path(result.vmrs_path)
        assert vmrs.is_file(), f"Expected .vmrs at {vmrs}"
        assert vmrs.parent == wb_path.parent
        assert vmrs.stat().st_size > 0

    def test_simulation_results_match_analytic(
        self,
        modelrisk_bridge: ModelRiskBridge,
        test_workbook: tuple[str, Path],
    ) -> None:
        """Y = 2 * N(0, 1): analytic mean 0, stdev 2. With 1000
        iterations and a fixed seed the empirical stats should land
        within ~3 standard errors (~0.2 for the mean, ~0.07 for the
        stdev). Use slightly wider bands to absorb seed quirks."""
        wb_name, _ = test_workbook
        modelrisk_bridge.run_simulation(
            workbook=wb_name, samples=1000, seed=1,
        )
        results = modelrisk_bridge.get_simulation_results(
            workbook=wb_name, output_names=["Y"]
        )
        assert len(results) == 1
        r = results[0]
        assert r.output_name == "Y"
        assert r.iterations == 1000
        assert abs(r.mean) < 0.3, f"mean too far from 0: {r.mean}"
        assert 1.8 < r.stdev < 2.2, f"stdev too far from 2: {r.stdev}"
        # P50 should be near 0; P5/P95 near ±3.29 (1.645 * 2).
        p50 = r.percentiles.get(0.50)
        p5 = r.percentiles.get(0.05)
        p95 = r.percentiles.get(0.95)
        assert p50 is not None and abs(p50) < 0.3
        assert p5 is not None and -4.5 < p5 < -2.5
        assert p95 is not None and 2.5 < p95 < 4.5

    def test_list_vmrs_variables_finds_both(
        self,
        modelrisk_bridge: ModelRiskBridge,
        test_workbook: tuple[str, Path],
    ) -> None:
        wb_name, _ = test_workbook
        modelrisk_bridge.run_simulation(
            workbook=wb_name, samples=500, seed=1,
        )
        entries = modelrisk_bridge.list_vmrs_variables(workbook=wb_name)
        names = {e["name"] for e in entries}
        assert "X" in names, f"expected input X in {names}"
        assert "Y" in names, f"expected output Y in {names}"
        # Iteration count carries through.
        for e in entries:
            assert e["iterations"] == 500

    def test_get_samples_returns_raw_array(
        self,
        modelrisk_bridge: ModelRiskBridge,
        test_workbook: tuple[str, Path],
    ) -> None:
        """Raw per-iteration samples for Y should be ~1000 floats
        whose empirical mean and stdev match the analytic moments."""
        wb_name, _ = test_workbook
        modelrisk_bridge.run_simulation(
            workbook=wb_name, samples=1000, seed=1,
        )
        samples = modelrisk_bridge.get_samples(
            "Y", workbook=wb_name, max_n=10_000,
        )
        assert isinstance(samples, list)
        assert 950 <= len(samples) <= 1000  # ~1000 minus any filter drops
        mean = sum(samples) / len(samples)
        var = sum((s - mean) ** 2 for s in samples) / (len(samples) - 1)
        stdev = math.sqrt(var)
        assert abs(mean) < 0.3
        assert 1.8 < stdev < 2.2

    def test_diagnose_workbook_reports_clean(
        self,
        modelrisk_bridge: ModelRiskBridge,
        test_workbook: tuple[str, Path],
    ) -> None:
        """After a successful run_simulation, diagnose_workbook should
        find a healthy model: Excel connected, MRService activated,
        outputs present, .vmrs on disk, no issues raised."""
        wb_name, _ = test_workbook
        modelrisk_bridge.run_simulation(
            workbook=wb_name, samples=500, seed=1,
        )
        # diagnose_workbook lives on the bridge as the workflow tool's
        # internals — call via the tool wrapper for a true end-to-end.
        from modelrisk_mcp.tools import reading, workflows

        reading.set_bridge_for_testing(modelrisk_bridge)
        try:
            out: Any = workflows.diagnose_workbook(workbook_name=wb_name)
        finally:
            reading.set_bridge_for_testing(None)

        assert out["excel_connected"] is True
        assert out["modelrisk_loaded"] is True
        assert out["active_workbook"] == wb_name
        assert out["output_count"] >= 1
        assert out["input_count"] >= 1
        assert out["vmrs_exists"] is True
        # An empty issues list is the bar for a "healthy" model.
        assert out["issues"] == [], f"unexpected issues: {out['issues']}"


class TestRunSimulationFailureModes:
    """Failure-shape tests against real Excel — these confirm our
    error translation is accurate, not just plausible."""

    def test_unknown_workbook_raises(
        self, modelrisk_bridge: ModelRiskBridge
    ) -> None:
        from modelrisk_mcp.errors import WorkbookNotFoundError

        with pytest.raises(WorkbookNotFoundError):
            modelrisk_bridge.run_simulation(
                workbook="not_open.xlsx", samples=100, seed=1,
            )

    def test_get_samples_unknown_name_raises(
        self,
        modelrisk_bridge: ModelRiskBridge,
        test_workbook: tuple[str, Path],
    ) -> None:
        wb_name, _ = test_workbook
        modelrisk_bridge.run_simulation(
            workbook=wb_name, samples=100, seed=1,
        )
        with pytest.raises(SimulationFailedError, match="not found"):
            modelrisk_bridge.get_samples(
                "NonExistent", workbook=wb_name,
            )
