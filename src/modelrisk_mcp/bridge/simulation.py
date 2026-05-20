"""SimulationController — drive a ModelRisk Monte Carlo run via XLL commands.

Architecture (v0.3.0-alpha.2, post-pivot):

The previous v0.2 attempts dispatched ATL CoClasses directly from Python.
That approach failed because `ModelRiskAtl.dll`'s coclasses don't expose
IDispatch at runtime, so cross-process automation cannot reach them.

Instead, this controller replicates exactly what the ATL itself does
when its `IModelRiskSimulation::StartSimulation` and
`IModelRiskSimulationResults::SaveResultsToFile` methods are invoked:

1. Start the sim by calling the XLL command **VoseStartSimulCustom12**
   via `Application.Run`. It takes a 1xN VARIANT array of `[Key]:Value`
   string options packed by `CSimulationOptions::PackToStringList`
   (ModelRiskCloude/SimulationObj.cpp:94). The call is synchronous —
   `Application.Run` returns when the simulation has finished.

2. Save the resulting `.vmrs` by calling the XLL command
   **VoseGetDataSZ12** with the session name
   `h<hWndExcel>_SaveResultsToFile_<book.xlsx>` and the target path as
   `xlParam1`. The handler in
   ModelRiskCloude/SimulationObj_VBA.cpp:805 dispatches on the
   operation prefix and routes to
   `CSimulationsManager::SaveWorkbookResults(sc, path)` directly when
   the path argument is non-empty. The handler internally writes a
   success/failure code to a memory-mapped file for the ATL's benefit,
   but `IPC_helpers.cpp:Send_sz_to_ATL` self-initialises that MMF — we
   don't need to set up anything on the Python side. We confirm
   success by checking that the file actually appeared on disk.

References:
- VoseStartSimulCustom12 export: ModelRiskCloude/XllAddIn.cpp:210
- VoseGetDataSZ12 export:        ModelRiskCloude/XllAddIn.cpp:207
- Session-name format:           ModelRiskAtl/ModelRiskSimulationResults.cpp:54
- Save handler:                  ModelRiskCloude/SimulationObj_VBA.cpp:805
- Options packing format:        ModelRiskAtl/SimulationObj.cpp:94
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Any

from modelrisk_mcp.errors import (
    ExcelNotRunningError,
    SimulationFailedError,
    WorkbookNotFoundError,
)

if TYPE_CHECKING:
    from modelrisk_mcp.bridge.excel import ExcelBridge


# ---------------------------------------------------------------------------
# Options
# ---------------------------------------------------------------------------


@dataclass
class SimulationOptions:
    """Mirrors `CSimulationOptions` (SimulationObj.h). Defaults are tuned
    for headless MCP usage — no progress dialog, no auto-open results
    window, no during-sim Excel refresh (faster + no flicker)."""

    samples: int = 1000
    sim_count: int = 1
    seeds: tuple[int, ...] = (1,)
    seed_fixed: bool = True
    seed_multiply_type: int = 0
    refresh_excel: bool = False
    refresh_rate: int = 1
    stop_on_output_error: bool = False
    show_results_at_end: bool = False
    hide_progress_window: bool = True
    min_sim_buffer_size: int = 0
    output_names: tuple[str, ...] = ()  # empty → all outputs
    macros_usage: int = 0
    macro_names: tuple[str, str, str, str] = ("", "", "", "")

    def to_string_list(self) -> list[str]:
        """Reproduce `CSimulationOptions::PackToStringList`. Returns a
        flat list of `[Key]:Value` strings in the exact order the C++
        emits them. The XLL parses by key, but we match the order for
        defence in depth."""
        out: list[str] = []
        out.append(f"[N]:{self.sim_count}")
        out.append(f"[Samples]:{self.samples}")
        out.append(f"[CntNames]:{len(self.output_names)}")
        for i, name in enumerate(self.output_names):
            out.append(f"[name{i}]:{name}")
        out.append(f"[SeedFixed]:{1 if self.seed_fixed else 0}")
        out.append(f"[SeedMultiplyType]:{self.seed_multiply_type}")
        out.append(f"[CntSeeds]:{len(self.seeds)}")
        for i, seed in enumerate(self.seeds):
            out.append(f"[seed{i}]:{seed}")
        out.append(f"[RefreshExcel]:{1 if self.refresh_excel else 0}")
        out.append(f"[RefreshRate]:{self.refresh_rate}")
        out.append(f"[StopOnOutputError]:{1 if self.stop_on_output_error else 0}")
        out.append(f"[ShowResultsAtEnd]:{1 if self.show_results_at_end else 0}")
        out.append(f"[HideProgressWindow]:{1 if self.hide_progress_window else 0}")
        out.append(f"[MinSimBufferSize]:{self.min_sim_buffer_size}")
        out.append(f"[MacrosUsage]:{self.macros_usage}")
        for i, macro in enumerate(self.macro_names):
            out.append(f"[Macros{i}]:{macro}")
        return out


@dataclass
class SimulationRunResult:
    workbook_name: str
    vmrs_path: str
    iterations: int
    options: SimulationOptions = field(default_factory=SimulationOptions)


# ---------------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------------


# XLL command names registered by ModelRisk.xll. CMDSTR_VISIBLE(X) →
# "Vose" + X (ModelRiskCloude/CLAUDE.md:63).
_CMD_START_SIM = "VoseStartSimulCustom12"
_CMD_GET_DATA_SZ = "VoseGetDataSZ12"

# Operation prefix the SimulationObj_VBA dispatcher matches on for the
# save path. PackSessionName format: "h<hwnd>_<Operation>_<book_name>"
# (ModelRiskAtl/ModelRiskSimulationResults.cpp:54).
_OP_SAVE_RESULTS = "SaveResultsToFile"


class SimulationController:
    """Drives ModelRisk simulations through the XLL command surface
    exposed to `Application.Run`. No direct ATL COM dispatch.

    All methods raise typed `modelrisk_mcp.errors` exceptions; raw COM
    HRESULTs and xlwings stack traces never leak.
    """

    def __init__(self, excel: ExcelBridge) -> None:
        self._excel = excel

    # ----- public API ----------------------------------------------------

    def run_simulation(
        self,
        workbook_name: str | None = None,
        *,
        samples: int = 1000,
        seed: int = 1,
        seed_fixed: bool = True,
        hide_dialogs: bool = True,
        save_to: str | None = None,
    ) -> SimulationRunResult:
        """Run a simulation on `workbook_name` (defaults to the active
        workbook) and save the resulting `.vmrs` to `save_to` (defaults
        to a sibling of the workbook).

        The call blocks until the simulation completes — that's how
        `VoseStartSimulCustom12` is implemented (synchronous Application.Run).

        Returns a SimulationRunResult with the resolved vmrs path.
        Raises SimulationFailedError if the file doesn't appear after
        the save call returns.
        """
        wb_info = self._resolve_workbook(workbook_name)
        opts = SimulationOptions(
            samples=samples,
            seeds=(seed,),
            seed_fixed=seed_fixed,
            hide_progress_window=hide_dialogs,
            show_results_at_end=False,
            refresh_excel=False,
        )
        target = self._resolve_save_path(wb_info, save_to)

        self._invoke_start_simulation(opts)
        self._invoke_save_results(wb_info.name, target)

        if not Path(target).is_file():
            raise SimulationFailedError(
                f"Simulation ran but no .vmrs was produced at {target!r}. "
                "Common causes: the workbook has no VoseOutput cells, "
                "ModelRisk failed to acquire a writer lock on the target "
                "path, or the simulation was cancelled by the user."
            )
        return SimulationRunResult(
            workbook_name=wb_info.name,
            vmrs_path=target,
            iterations=samples,
            options=opts,
        )

    # ----- internal ------------------------------------------------------

    def _resolve_workbook(self, workbook_name: str | None) -> _WorkbookCoords:
        """Resolve to (book name, folder path) tolerating OneDrive
        path-resolution failure. The save needs a folder; we fall back
        to the user's Desktop when xlwings can't tell us the path."""
        if workbook_name:
            books = self._excel.list_workbooks()
            info = next((b for b in books if b.name == workbook_name), None)
            if info is None:
                raise WorkbookNotFoundError(
                    f"Workbook {workbook_name!r} is not open."
                )
        else:
            info = self._excel.get_active_workbook()

        folder = self._folder_for(info.path)
        return _WorkbookCoords(name=info.name, folder=folder)

    @staticmethod
    def _folder_for(path: str) -> Path:
        if path:
            try:
                p = Path(path)
                if p.parent.is_dir():
                    return p.parent
            except (OSError, ValueError):
                pass
        # OneDrive workbooks or untrackable paths: fall back to Desktop,
        # which is where the user is most likely to find the file.
        desktop = Path.home() / "Desktop"
        if desktop.is_dir():
            return desktop
        return Path.home()

    @staticmethod
    def _resolve_save_path(wb: _WorkbookCoords, override: str | None) -> str:
        if override:
            return str(Path(override).expanduser())
        # ModelRisk's default file dialog suggests "<book>.vmrs"
        # (SimulationObj_VBA.cpp:813). Mirror that.
        stem = Path(wb.name).stem
        return str(wb.folder / f"{stem}.vmrs")

    def _invoke_start_simulation(self, opts: SimulationOptions) -> None:
        """Application.Run("VoseStartSimulCustom12", options_array).

        `options_array` must be a 1-row 2D SAFEARRAY of BSTRs. pywin32
        converts a list-of-lists into a SAFEARRAY automatically when the
        target argument is a VARIANT."""
        app = self._app()
        try:
            options_2d = [opts.to_string_list()]  # 1 row x N cols
            app.api.Run(_CMD_START_SIM, options_2d)
        except Exception as exc:
            raise SimulationFailedError(
                f"Application.Run({_CMD_START_SIM!r}, ...) failed: {exc}. "
                "ModelRisk add-in must be loaded and the workbook must "
                "contain at least one VoseOutput cell."
            ) from exc

    def _invoke_save_results(self, book_name: str, target_path: str) -> None:
        """Application.Run("VoseGetDataSZ12", session_name, target_path).

        Mirrors the ATL's IModelRiskSimulationResults::SaveResultsToFile
        path (ModelRiskAtl/ModelRiskSimulationResults.cpp:1196). The
        XLL handler in SimulationObj_VBA.cpp:805 reads xlParam1 as the
        target file and skips the file dialog when non-empty."""
        app = self._app()
        try:
            hwnd = int(app.api.Hwnd)
        except Exception as exc:
            raise SimulationFailedError(
                "Could not read Application.Hwnd to compose the save "
                "session name. Excel may have closed."
            ) from exc
        session_name = f"h{hwnd}_{_OP_SAVE_RESULTS}_{book_name}"
        try:
            app.api.Run(_CMD_GET_DATA_SZ, session_name, target_path)
        except Exception as exc:
            raise SimulationFailedError(
                f"Application.Run({_CMD_GET_DATA_SZ!r}, ...) failed during "
                f"SaveResultsToFile: {exc}. The simulation may have run "
                "but the .vmrs was not persisted."
            ) from exc

    def _app(self) -> Any:
        if not self._excel.is_connected():
            self._excel.connect()
        app = self._excel._app  # bridge-internal access; same module family
        if app is None:
            raise ExcelNotRunningError(
                "Excel is not connected. Open Excel and load the workbook."
            )
        return app


# ---------------------------------------------------------------------------
# Internal types
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class _WorkbookCoords:
    name: str
    folder: Path


__all__ = [
    "SimulationController",
    "SimulationOptions",
    "SimulationRunResult",
]
