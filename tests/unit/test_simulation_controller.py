"""Unit tests for the SimulationController XLL-command driver.

The controller's job is to:
1. Pack SimulationOptions into the exact `[Key]:Value` line format
   `CSimulationOptions::PackToStringList` (C++) emits.
2. Call `Application.Run("VoseStartSimulCustom12", options_2d)` with a
   1-row 2D string array.
3. Call `Application.Run("VoseGetDataSZ12", session_name, target_path)`
   with the session name in the form `h<hwnd>_SaveResultsToFile_<book>`.
4. Verify the .vmrs file appeared, raising SimulationFailedError if not.

Tests use a fake ExcelBridge + a fake Application to record every Run
call and let us assert exact-string conformance with the C++ side.
"""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from modelrisk_mcp.bridge.simulation import (
    SimulationController,
    SimulationOptions,
)
from modelrisk_mcp.errors import SimulationFailedError, WorkbookNotFoundError
from modelrisk_mcp.schemas.workbook import WorkbookInfo


class _RunRecorder:
    """Fake Application.api that records every Run() invocation."""

    def __init__(self, *, hwnd: int = 12345, on_save: Any = None) -> None:
        self.Hwnd = hwnd
        self.calls: list[tuple[str, tuple[Any, ...]]] = []
        self._on_save = on_save

    def Run(self, name: str, *args: Any) -> Any:  # noqa: N802 (COM API name)
        self.calls.append((name, args))
        if name == "VoseGetDataSZ12" and self._on_save is not None:
            self._on_save(*args)
        return 1


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
        self._app = _FakeApp(self._recorder)

    def list_workbooks(self) -> list[WorkbookInfo]:
        return list(self._books)

    def get_active_workbook(self) -> WorkbookInfo:
        return self._active

    def is_connected(self) -> bool:
        return True

    def connect(self) -> None:
        pass


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
    def test_invokes_start_then_save_in_order(self, tmp_path: Path) -> None:
        target = tmp_path / "model.vmrs"

        def fake_save(session: str, path: str) -> None:
            Path(path).write_bytes(b"")  # pretend XLL wrote the file

        recorder = _RunRecorder(hwnd=9999, on_save=fake_save)
        wb = _make_wb("model.xlsx", tmp_path)
        bridge = _FakeBridge(active=wb, recorder=recorder)
        controller = SimulationController(bridge)  # type: ignore[arg-type]

        result = controller.run_simulation()

        assert [c[0] for c in recorder.calls] == [
            "VoseStartSimulCustom12",
            "VoseGetDataSZ12",
        ]
        # First call: 1-row 2D options array.
        start_args = recorder.calls[0][1]
        assert len(start_args) == 1
        options_2d = start_args[0]
        assert len(options_2d) == 1, "options must be a 1-row 2D array"
        assert any(s.startswith("[Samples]:") for s in options_2d[0])
        # Second call: session name + path.
        save_args = recorder.calls[1][1]
        assert save_args[0] == "h9999_SaveResultsToFile_model.xlsx"
        assert save_args[1] == str(target)
        # Result reflects the discovered file.
        assert result.vmrs_path == str(target)
        assert result.iterations == 1000

    def test_custom_samples_and_seed_propagate(self, tmp_path: Path) -> None:
        def fake_save(session: str, path: str) -> None:
            Path(path).write_bytes(b"")

        recorder = _RunRecorder(on_save=fake_save)
        wb = _make_wb("m.xlsx", tmp_path)
        bridge = _FakeBridge(active=wb, recorder=recorder)
        controller = SimulationController(bridge)  # type: ignore[arg-type]

        controller.run_simulation(samples=5000, seed=99)

        options_row = recorder.calls[0][1][0][0]
        assert "[Samples]:5000" in options_row
        assert "[seed0]:99" in options_row

    def test_custom_save_path_honoured(self, tmp_path: Path) -> None:
        target = tmp_path / "custom" / "out.vmrs"
        target.parent.mkdir()

        def fake_save(session: str, path: str) -> None:
            Path(path).write_bytes(b"")

        recorder = _RunRecorder(on_save=fake_save)
        bridge = _FakeBridge(
            active=_make_wb("m.xlsx", tmp_path), recorder=recorder
        )
        controller = SimulationController(bridge)  # type: ignore[arg-type]

        result = controller.run_simulation(save_to=str(target))
        assert result.vmrs_path == str(target)
        assert recorder.calls[1][1][1] == str(target)

    def test_raises_when_file_not_produced(self, tmp_path: Path) -> None:
        # on_save = None means the fake never writes the file.
        recorder = _RunRecorder(on_save=None)
        bridge = _FakeBridge(
            active=_make_wb("m.xlsx", tmp_path), recorder=recorder
        )
        controller = SimulationController(bridge)  # type: ignore[arg-type]

        with pytest.raises(SimulationFailedError) as exc:
            controller.run_simulation()
        msg = str(exc.value)
        assert "no .vmrs was produced" in msg
        assert "VoseOutput" in msg

    def test_named_workbook_lookup(self, tmp_path: Path) -> None:
        a = _make_wb("a.xlsx", tmp_path)
        b = _make_wb("b.xlsx", tmp_path)

        def fake_save(session: str, path: str) -> None:
            Path(path).write_bytes(b"")

        recorder = _RunRecorder(on_save=fake_save)
        bridge = _FakeBridge(active=a, workbooks=[a, b], recorder=recorder)
        controller = SimulationController(bridge)  # type: ignore[arg-type]

        result = controller.run_simulation(workbook_name="b.xlsx")
        assert result.workbook_name == "b.xlsx"
        # Session name uses b.xlsx, not the active a.xlsx
        assert "b.xlsx" in recorder.calls[1][1][0]

    def test_unknown_workbook_raises(self, tmp_path: Path) -> None:
        bridge = _FakeBridge(active=_make_wb("a.xlsx", tmp_path))
        controller = SimulationController(bridge)  # type: ignore[arg-type]
        with pytest.raises(WorkbookNotFoundError):
            controller.run_simulation(workbook_name="missing.xlsx")

    def test_onedrive_path_falls_back_to_desktop(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """When `path` is empty (the OneDrive fallback case) and the
        caller didn't supply an explicit save_to, save next to the
        Desktop instead of the unresolvable workbook folder."""
        fake_home = tmp_path / "home"
        fake_desktop = fake_home / "Desktop"
        fake_desktop.mkdir(parents=True)
        monkeypatch.setattr(Path, "home", lambda: fake_home)

        def fake_save(session: str, path: str) -> None:
            Path(path).write_bytes(b"")

        recorder = _RunRecorder(on_save=fake_save)
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
            Path(path).write_bytes(b"")

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
    def test_run_simulation_auto_pins_vmrs(self, tmp_path: Path) -> None:
        """ModelRiskBridge.run_simulation must call
        ResultsReader.set_active_vmrs(path) so the next
        get_simulation_results call doesn't need a sibling-search."""
        from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge

        # Build a controller backed by a fake bridge that writes the
        # file synchronously.
        def fake_save(session: str, path: str) -> None:
            Path(path).write_bytes(b"")

        recorder = _RunRecorder(on_save=fake_save)
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


__all__ = [
    "TestBridgeIntegration",
    "TestOptionsPacking",
    "TestRunSimulation",
    "TestSessionNameFormat",
    "TestStartCallShape",
]


def _unused_imports_anchor() -> None:  # pragma: no cover
    _ = SimpleNamespace
