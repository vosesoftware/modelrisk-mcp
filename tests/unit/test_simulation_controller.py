"""Unit tests for the SimulationController XLL-command driver.

The controller's job is to:
1. Pack SimulationOptions into the exact `[Key]:Value` line format
   `CSimulationOptions::PackToStringList` (C++) emits.
2. Call `Application.Run("VoseSetSimulOptions12", options_2d)` to persist
   the options (esp. the seed) to the workbook's SimOpt_* defined-names —
   required so the per-cell-twister Manual-Seed engine honors the seed
   (AB#2742).
3. Call `Application.Run("VoseStartSimulCustom12", options_2d)` with the
   same 1-row 2D string array.
4. Call `Application.Run("VoseGetDataSZ12", session_name, target_path)`
   with the session name in the form `h<hwnd>_SaveResultsToFile_<book>`.
5. Prove the saved .vmrs is this run (header: workbook, start time,
   iterations), rebuilding it from ModelRisk's run file when the save
   handed back another run, and raise SimulationFailedError otherwise.

Tests use a fake ExcelBridge + a fake Application that plays ModelRisk:
it records every Run call (for exact-string conformance with the C++
side), writes a run file per run, and saves like the real XLL does.
"""

from __future__ import annotations

from collections.abc import Callable
from datetime import datetime, timedelta
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from modelrisk_mcp.bridge import simulation
from modelrisk_mcp.bridge.simulation import (
    SimulationController,
    SimulationOptions,
)
from modelrisk_mcp.bridge.vmrs_file import read_vmrs_header
from modelrisk_mcp.errors import SimulationFailedError, WorkbookNotFoundError
from modelrisk_mcp.schemas.workbook import WorkbookInfo
from tests.unit._results_files import vmrs_bytes, write_dmr, write_vmrs


@pytest.fixture(autouse=True)
def run_files(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """ModelRisk's simulation folder for these tests: the only place the
    controller looks for run files, never the machine's real one."""
    folder = tmp_path / "SimulationStorage"
    folder.mkdir()
    monkeypatch.setattr(simulation, "_simulation_folders", lambda: [folder])
    return folder


class _RunRecorder:
    """Fake Application.api: records every Run() invocation and plays
    ModelRisk.

    VoseStartSimulCustom12 runs the ACTIVE workbook and writes its run file,
    `f<hwnd hex><book>.dmr`, into `run_dir` (none when `run_dir` is None).
    VoseGetDataSZ12 saves as the real XLL does: the run file the Results
    Viewer holds (`viewer_holds`) if set, else the named workbook's own.
    `on_save` replaces that save entirely."""

    def __init__(
        self,
        *,
        hwnd: int = 12345,
        on_save: Callable[..., Any] | None = None,
        run_dir: Path | None = None,
        viewer_holds: Path | None = None,
        engine_hwnd: int | None = None,
    ) -> None:
        self.Hwnd = hwnd
        # The window handle the engine fixed at start-up and names run files
        # with; Application.Hwnd follows the active workbook's window.
        self.engine_hwnd = hwnd if engine_hwnd is None else engine_hwnd
        self.calls: list[tuple[str, tuple[Any, ...]]] = []
        self._on_save = on_save
        self.run_dir = run_dir
        self.viewer_holds = viewer_holds
        self.registered: list[str] = []
        # One installed ModelRisk add-in, so code that walks AddIns to
        # re-register it would find something to register.
        self.AddIns = _FakeAddIns()
        self.bridge: _FakeBridge | None = None

    def Run(self, name: str, *args: Any) -> Any:  # noqa: N802 (COM API name)
        self.calls.append((name, args))
        if name == "VoseStartSimulCustom12" and self.run_dir is not None:
            assert self.bridge is not None
            book = self.bridge.get_active_workbook().name
            samples = next(
                int(s.split(":", 1)[1]) for s in args[0][0] if s.startswith("[Samples]:")
            )
            write_dmr(
                self.run_file(book), book, start=datetime.now(), iterations=samples
            )
        if name == "VoseGetDataSZ12":
            if self._on_save is not None:
                self._on_save(*args)
                return 1
            session, path = args
            book = session.split("_SaveResultsToFile_", 1)[1]
            source = self.viewer_holds or (
                self.run_file(book) if self.run_dir is not None else None
            )
            if source is not None and source.is_file():
                Path(path).write_bytes(vmrs_bytes(source.read_bytes()))
        return 1

    def run_file(self, book: str) -> Path:
        assert self.run_dir is not None
        return self.run_dir / f"f{self.engine_hwnd:X}{book}.dmr"

    def RegisterXLL(self, path: str) -> bool:  # noqa: N802 (COM API name)
        self.registered.append(path)
        return True


class _FakeAddIns:
    Count = 1

    def __call__(self, index: int) -> Any:
        return SimpleNamespace(
            Name="ModelRisk.xll",
            Installed=True,
            FullName=r"C:\Program Files\Vose Software\ModelRisk\ModelRisk.xll",
        )


class _FakeApp:
    def __init__(self, recorder: _RunRecorder) -> None:
        self.api = recorder


class _FakeBridge:
    """Stand-in for ExcelBridge. Only implements what
    SimulationController touches."""

    def __init__(
        self,
        active: WorkbookInfo,
        workbooks: list[WorkbookInfo] | None = None,
        recorder: _RunRecorder | None = None,
    ) -> None:
        self._active = active
        self._books = workbooks or [active]
        self._recorder = recorder or _RunRecorder()
        self._recorder.bridge = self
        self._app = _FakeApp(self._recorder)
        self.activated: list[str] = []

    def list_workbooks(self) -> list[WorkbookInfo]:
        return list(self._books)

    def get_active_workbook(self) -> WorkbookInfo:
        return self._active

    def activate_workbook(self, workbook: str) -> None:
        self.activated.append(workbook)
        self._active = next(b for b in self._books if b.name == workbook)

    def is_connected(self) -> bool:
        return True

    def connect(self) -> None:
        pass

    def evaluate(self, expr: str) -> Any:
        # bug #38: report the add-in as live so run_simulation's
        # ensure_modelrisk_functional() short-circuits.
        return 0.0


# ---------------------------------------------------------------------------
# SimulationOptions packing
# ---------------------------------------------------------------------------


class TestOptionsPacking:
    def test_defaults_match_c_plus_plus_field_order(self) -> None:
        """The CSimulationOptions::PackToStringList macros emit keys in
        a strict order — N, Samples, CntNames, name<i>*, SeedFixed,
        SeedMultiplyType, CntSeeds, seed<i>*, RefreshExcel, RefreshRate,
        StopOnOutputError, ShowResultsAtEnd, HideProgressWindow,
        MinSimBufferSize, MacrosUsage, Macros0..3."""
        opts = SimulationOptions()
        keys = [line.split(":", 1)[0] for line in opts.to_string_list()]
        expected_prefix = [
            "[N]", "[Samples]", "[CntNames]",
            "[SeedFixed]", "[SeedMultiplyType]", "[CntSeeds]",
            "[seed0]", "[RefreshExcel]", "[RefreshRate]",
            "[StopOnOutputError]", "[ShowResultsAtEnd]",
            "[HideProgressWindow]", "[MinSimBufferSize]", "[MacrosUsage]",
            "[Macros0]", "[Macros1]", "[Macros2]", "[Macros3]",
        ]
        assert keys == expected_prefix

    def test_named_outputs_emit_indexed_keys(self) -> None:
        opts = SimulationOptions(output_names=("Revenue", "Profit"))
        out = opts.to_string_list()
        assert "[CntNames]:2" in out
        assert "[name0]:Revenue" in out
        assert "[name1]:Profit" in out

    def test_seed_serialised_as_int(self) -> None:
        opts = SimulationOptions(seeds=(42,))
        out = opts.to_string_list()
        assert "[seed0]:42" in out
        assert "[CntSeeds]:1" in out

    def test_booleans_serialise_as_01(self) -> None:
        opts = SimulationOptions(
            seed_fixed=True, refresh_excel=False, hide_progress_window=True
        )
        out = opts.to_string_list()
        assert "[SeedFixed]:1" in out
        assert "[RefreshExcel]:0" in out
        assert "[HideProgressWindow]:1" in out


# ---------------------------------------------------------------------------
# Run flow
# ---------------------------------------------------------------------------


def _make_wb(name: str, folder: Path) -> WorkbookInfo:
    return WorkbookInfo(
        name=name,
        path=str(folder / name),
        sheets=["Sheet1"],
        active_sheet="Sheet1",
    )


class TestRunSimulation:
    def test_invokes_start_then_save_in_order(
        self, tmp_path: Path, run_files: Path
    ) -> None:
        target = tmp_path / "model.vmrs"

        recorder = _RunRecorder(hwnd=9999, run_dir=run_files)
        wb = _make_wb("model.xlsx", tmp_path)
        bridge = _FakeBridge(active=wb, recorder=recorder)
        controller = SimulationController(bridge)  # type: ignore[arg-type]

        result = controller.run_simulation()

        assert [c[0] for c in recorder.calls] == [
            "VoseSetSimulOptions12",
            "VoseStartSimulCustom12",
            "VoseGetDataSZ12",
        ]
        # Persist + start calls both carry the same 1-row 2D options array.
        set_args = recorder.calls[0][1]
        start_args = recorder.calls[1][1]
        assert len(start_args) == 1
        options_2d = start_args[0]
        assert len(options_2d) == 1, "options must be a 1-row 2D array"
        assert any(s.startswith("[Samples]:") for s in options_2d[0])
        assert set_args == start_args, "persist must get the same options as start"
        # Third call: session name + path.
        save_args = recorder.calls[2][1]
        assert save_args[0] == "h9999_SaveResultsToFile_model.xlsx"
        assert save_args[1] == str(target)
        # Result reflects the discovered file, which holds this run.
        assert result.vmrs_path == str(target)
        assert result.iterations == 1000
        assert result.note is None
        assert read_vmrs_header(target).spreadsheet_name == "model.xlsx"

    def test_custom_samples_and_seed_propagate(
        self, tmp_path: Path, run_files: Path
    ) -> None:
        recorder = _RunRecorder(run_dir=run_files)
        wb = _make_wb("m.xlsx", tmp_path)
        bridge = _FakeBridge(active=wb, recorder=recorder)
        controller = SimulationController(bridge)  # type: ignore[arg-type]

        controller.run_simulation(samples=5000, seed=99)

        options_row = recorder.calls[0][1][0][0]
        assert "[Samples]:5000" in options_row
        assert "[seed0]:99" in options_row

    def test_custom_save_path_honoured(
        self, tmp_path: Path, run_files: Path
    ) -> None:
        target = tmp_path / "custom" / "out.vmrs"
        target.parent.mkdir()

        recorder = _RunRecorder(run_dir=run_files)
        bridge = _FakeBridge(
            active=_make_wb("m.xlsx", tmp_path), recorder=recorder
        )
        controller = SimulationController(bridge)  # type: ignore[arg-type]

        result = controller.run_simulation(save_to=str(target))
        assert result.vmrs_path == str(target)
        # calls: [0]=SetSimulOptions, [1]=StartSimul, [2]=GetDataSZ(session, path)
        assert recorder.calls[2][1][1] == str(target)

    def test_raises_when_file_not_produced(self, tmp_path: Path) -> None:
        # No run_dir: the fake run writes no run file, so the save has
        # nothing to write either.
        recorder = _RunRecorder()
        bridge = _FakeBridge(
            active=_make_wb("m.xlsx", tmp_path), recorder=recorder
        )
        controller = SimulationController(bridge)  # type: ignore[arg-type]

        with pytest.raises(SimulationFailedError) as exc:
            controller.run_simulation()
        msg = str(exc.value)
        assert "no .vmrs was produced" in msg
        assert "VoseOutput" in msg

    def test_named_workbook_lookup(self, tmp_path: Path, run_files: Path) -> None:
        a = _make_wb("a.xlsx", tmp_path)
        b = _make_wb("b.xlsx", tmp_path)

        recorder = _RunRecorder(run_dir=run_files)
        bridge = _FakeBridge(active=a, workbooks=[a, b], recorder=recorder)
        controller = SimulationController(bridge)  # type: ignore[arg-type]

        result = controller.run_simulation(workbook_name="b.xlsx")
        assert result.workbook_name == "b.xlsx"
        # Session name uses b.xlsx, not the active a.xlsx (save = 3rd call now)
        assert "b.xlsx" in recorder.calls[2][1][0]

    def test_unknown_workbook_raises(self, tmp_path: Path) -> None:
        bridge = _FakeBridge(active=_make_wb("a.xlsx", tmp_path))
        controller = SimulationController(bridge)  # type: ignore[arg-type]
        with pytest.raises(WorkbookNotFoundError):
            controller.run_simulation(workbook_name="missing.xlsx")

    def test_onedrive_path_falls_back_to_desktop(
        self, tmp_path: Path, run_files: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """When `path` is empty (the OneDrive fallback case) and the
        caller didn't supply an explicit save_to, save next to the
        Desktop instead of the unresolvable workbook folder."""
        fake_home = tmp_path / "home"
        fake_desktop = fake_home / "Desktop"
        fake_desktop.mkdir(parents=True)
        monkeypatch.setattr(Path, "home", lambda: fake_home)

        recorder = _RunRecorder(run_dir=run_files)
        wb = WorkbookInfo(
            name="onedrive.xlsx", path="", sheets=[], active_sheet=None
        )
        bridge = _FakeBridge(active=wb, recorder=recorder)
        controller = SimulationController(bridge)  # type: ignore[arg-type]

        result = controller.run_simulation()
        assert result.vmrs_path == str(fake_desktop / "onedrive.vmrs")


class TestSessionNameFormat:
    def test_matches_c_plus_plus_packsessionname_layout(
        self, tmp_path: Path
    ) -> None:
        """`PackSessionName` (ModelRiskSimulationResults.cpp:54) emits
        `h%d_%s_%s` with hWndExcel, Operation, book_name. We must
        produce byte-identical strings — the XLL handler dispatches on
        a strict prefix match."""

        def fake_save(session: str, path: str) -> None:
            assert session == "h42_SaveResultsToFile_book.xlsx"
            write_vmrs(Path(path), "book.xlsx", start=datetime.now())

        recorder = _RunRecorder(hwnd=42, on_save=fake_save)
        bridge = _FakeBridge(
            active=_make_wb("book.xlsx", tmp_path), recorder=recorder
        )
        controller = SimulationController(bridge)  # type: ignore[arg-type]
        controller.run_simulation()


class TestStartCallShape:
    def test_application_run_failure_wraps_to_typed_error(
        self, tmp_path: Path
    ) -> None:
        recorder = _RunRecorder()

        def boom(name: str, *args: Any) -> Any:
            raise RuntimeError("XLL not loaded")

        recorder.Run = boom  # type: ignore[method-assign]
        bridge = _FakeBridge(
            active=_make_wb("m.xlsx", tmp_path), recorder=recorder
        )
        controller = SimulationController(bridge)  # type: ignore[arg-type]
        with pytest.raises(SimulationFailedError) as exc:
            controller.run_simulation()
        assert "VoseStartSimulCustom12" in str(exc.value)
        assert "ModelRisk add-in" in str(exc.value)


class TestBridgeIntegration:
    def test_run_simulation_auto_pins_vmrs(
        self, tmp_path: Path, run_files: Path
    ) -> None:
        """ModelRiskBridge.run_simulation must call
        ResultsReader.set_active_vmrs(path) so the next
        get_simulation_results call doesn't need a sibling-search."""
        from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge

        # Build a controller backed by a fake bridge that writes the
        # file synchronously.
        recorder = _RunRecorder(run_dir=run_files)
        wb = _make_wb("m.xlsx", tmp_path)
        excel_fake = _FakeBridge(active=wb, recorder=recorder)
        controller = SimulationController(excel_fake)  # type: ignore[arg-type]

        # Stub the ResultsReader to capture set_active_vmrs.
        captured: dict[str, Any] = {}

        class _StubReader:
            def set_active_vmrs(self, path: str | None) -> None:
                captured["path"] = path

        bridge = ModelRiskBridge(
            excel=excel_fake,  # type: ignore[arg-type]
            simulation=controller,
            results=_StubReader(),  # type: ignore[arg-type]
        )
        result = bridge.run_simulation()
        assert captured["path"] == result.vmrs_path


def _earlier_run(folder: Path, book: str) -> Path:
    """A run file of `book` from last night, like the one ModelRisk's
    Results Viewer held on 2026-10-03."""
    return write_dmr(
        folder / f"f3039{book}.dmr", book, start=datetime.now() - timedelta(hours=13)
    )


class TestSavedRunIsThisRun:
    """Field report, 2026-10-03: ModelRisk's save hands the job to its
    Results Viewer, which saves the run IT holds, whatever workbook is
    named. 22 saves wrote OG-05's run under other workbooks' names, and
    the tool answered each with a bare error."""

    def test_another_workbooks_run_is_replaced_by_this_run(
        self, tmp_path: Path, run_files: Path
    ) -> None:
        a, b = _make_wb("a.xlsx", tmp_path), _make_wb("b.xlsx", tmp_path)
        recorder = _RunRecorder(
            run_dir=run_files, viewer_holds=_earlier_run(run_files, "a.xlsx")
        )
        bridge = _FakeBridge(active=b, workbooks=[a, b], recorder=recorder)
        started = datetime.now().replace(microsecond=0)

        result = SimulationController(bridge).run_simulation(  # type: ignore[arg-type]
            workbook_name="b.xlsx"
        )

        header = read_vmrs_header(result.vmrs_path)
        assert header.spreadsheet_name == "b.xlsx"
        assert header.start_date is not None and header.start_date >= started
        assert header.iterations == 1000
        assert result.note is not None
        assert "the results of 'a.xlsx'" in result.note

    def test_run_file_is_found_whatever_window_named_it(
        self, tmp_path: Path, run_files: Path
    ) -> None:
        """Live repro, 2026-10-04: with one window per workbook the engine
        named run files after the first window (f50108…) while
        Application.Hwnd read the active one (0xB0822). An older run file of
        the same workbook from another Excel session lies beside it."""
        a, b = _make_wb("a.xlsx", tmp_path), _make_wb("b.xlsx", tmp_path)
        _earlier_run(run_files, "b.xlsx").rename(run_files / "f10025Ab.xlsx.dmr")
        recorder = _RunRecorder(
            hwnd=0xB0822,
            engine_hwnd=0x50108,
            run_dir=run_files,
            viewer_holds=_earlier_run(run_files, "a.xlsx"),
        )
        bridge = _FakeBridge(active=b, workbooks=[a, b], recorder=recorder)

        result = SimulationController(bridge).run_simulation(  # type: ignore[arg-type]
            workbook_name="b.xlsx", samples=500
        )

        header = read_vmrs_header(result.vmrs_path)
        assert header.spreadsheet_name == "b.xlsx"
        assert header.iterations == 500
        assert (run_files / "f50108b.xlsx.dmr").is_file()

    def test_wrong_run_without_a_run_file_fails_and_leaves_no_wrong_file(
        self, tmp_path: Path
    ) -> None:
        a, b = _make_wb("a.xlsx", tmp_path), _make_wb("b.xlsx", tmp_path)
        # No run_dir: nothing to rebuild from.
        recorder = _RunRecorder(viewer_holds=_earlier_run(tmp_path / "rv", "a.xlsx"))
        bridge = _FakeBridge(active=b, workbooks=[a, b], recorder=recorder)

        with pytest.raises(SimulationFailedError) as exc:
            SimulationController(bridge).run_simulation(  # type: ignore[arg-type]
                workbook_name="b.xlsx"
            )

        msg = str(exc.value)
        assert "the results of 'a.xlsx'" in msg
        assert "Results Viewer" in msg
        assert not (tmp_path / "b.vmrs").exists()

    def test_earlier_run_of_the_same_workbook_is_not_this_run(
        self, tmp_path: Path
    ) -> None:
        recorder = _RunRecorder(viewer_holds=_earlier_run(tmp_path / "rv", "m.xlsx"))
        bridge = _FakeBridge(active=_make_wb("m.xlsx", tmp_path), recorder=recorder)

        with pytest.raises(SimulationFailedError) as exc:
            SimulationController(bridge).run_simulation()  # type: ignore[arg-type]

        assert "an earlier run of 'm.xlsx'" in str(exc.value)
        assert not (tmp_path / "m.vmrs").exists()

    def test_iteration_count_must_match(self, tmp_path: Path) -> None:
        def short_save(session: str, path: str) -> None:
            write_vmrs(Path(path), "m.xlsx", start=datetime.now(), iterations=500)

        recorder = _RunRecorder(on_save=short_save)
        bridge = _FakeBridge(active=_make_wb("m.xlsx", tmp_path), recorder=recorder)

        with pytest.raises(SimulationFailedError) as exc:
            SimulationController(bridge).run_simulation()  # type: ignore[arg-type]

        assert "a run of 500 iterations, not 1000" in str(exc.value)

    def test_failed_save_is_recovered_from_the_run_file(
        self, tmp_path: Path, run_files: Path
    ) -> None:
        def broken_save(session: str, path: str) -> None:
            raise RuntimeError("The RPC server is unavailable.")

        recorder = _RunRecorder(run_dir=run_files, on_save=broken_save)
        bridge = _FakeBridge(active=_make_wb("m.xlsx", tmp_path), recorder=recorder)

        result = SimulationController(bridge).run_simulation()  # type: ignore[arg-type]

        assert read_vmrs_header(result.vmrs_path).spreadsheet_name == "m.xlsx"
        assert result.note is not None and "RPC server" in result.note

    def test_named_workbook_is_made_active_before_the_run(
        self, tmp_path: Path, run_files: Path
    ) -> None:
        """ModelRisk simulates the ACTIVE workbook. Asked for b.xlsx while
        a.xlsx is active, the run must be b.xlsx's."""
        a, b = _make_wb("a.xlsx", tmp_path), _make_wb("b.xlsx", tmp_path)
        recorder = _RunRecorder(run_dir=run_files)
        bridge = _FakeBridge(active=a, workbooks=[a, b], recorder=recorder)

        result = SimulationController(bridge).run_simulation(  # type: ignore[arg-type]
            workbook_name="b.xlsx"
        )

        assert bridge.activated == ["b.xlsx"]
        assert recorder.run_file("b.xlsx").is_file()
        assert not recorder.run_file("a.xlsx").exists()
        assert read_vmrs_header(result.vmrs_path).spreadsheet_name == "b.xlsx"
        assert result.note is None

    def test_no_run_when_the_workbook_cannot_be_made_active(
        self, tmp_path: Path
    ) -> None:
        a, b = _make_wb("a.xlsx", tmp_path), _make_wb("b.xlsx", tmp_path)
        recorder = _RunRecorder()
        bridge = _FakeBridge(active=a, workbooks=[a, b], recorder=recorder)

        def refuse(workbook: str) -> None:
            raise RuntimeError("a dialog is open")

        bridge.activate_workbook = refuse  # type: ignore[method-assign]

        with pytest.raises(SimulationFailedError, match="active workbook"):
            SimulationController(bridge).run_simulation(  # type: ignore[arg-type]
                workbook_name="b.xlsx"
            )
        assert recorder.calls == []

    def test_never_re_registers_the_add_in(
        self, tmp_path: Path, run_files: Path
    ) -> None:
        """RegisterXLL on a live add-in re-runs its xlAutoOpen, the call that
        once wiped a licence activation; a run must not make it."""
        recorder = _RunRecorder(run_dir=run_files)
        bridge = _FakeBridge(active=_make_wb("m.xlsx", tmp_path), recorder=recorder)

        SimulationController(bridge).run_simulation()  # type: ignore[arg-type]

        assert recorder.registered == []


__all__ = [
    "TestBridgeIntegration",
    "TestOptionsPacking",
    "TestRunSimulation",
    "TestSavedRunIsThisRun",
    "TestSessionNameFormat",
    "TestStartCallShape",
]
