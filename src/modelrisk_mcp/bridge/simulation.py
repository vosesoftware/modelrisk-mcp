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
   `Application.Run` returns when the simulation has finished. The run
   belongs to the ACTIVE workbook, so a named workbook is activated first.

2. Save the resulting `.vmrs` by calling the XLL command
   **VoseGetDataSZ12** with the session name
   `h<hWndExcel>_SaveResultsToFile_<book.xlsx>` and the target path as
   `xlParam1`. The handler in
   ModelRiskCloude/SimulationObj_VBA.cpp:805 dispatches on the
   operation prefix and hands the save to the Results Viewer process
   first. The viewer saves the run IT has loaded, whatever workbook was
   named; only when it has none does the XLL save the named workbook's
   run with `CSimulationsManager::SaveWorkbookResults(sc, path)`. The
   handler internally writes a
   success/failure code to a memory-mapped file for the ATL's benefit,
   but `IPC_helpers.cpp:Send_sz_to_ATL` self-initialises that MMF — we
   don't need to set up anything on the Python side.

3. Prove the saved file is this run: its header must name the workbook,
   and its start time and iteration count must match the run. If the
   viewer handed back another run, rebuild the `.vmrs` from the
   workbook's own ModelRisk run file (`bridge/vmrs_file.py`).

References:
- VoseStartSimulCustom12 export: ModelRiskCloude/XllAddIn.cpp:210
- VoseGetDataSZ12 export:        ModelRiskCloude/XllAddIn.cpp:207
- Session-name format:           ModelRiskAtl/ModelRiskSimulationResults.cpp:54
- Save handler:                  ModelRiskCloude/SimulationObj_VBA.cpp:805
- Viewer's save handler:         ModelRiskResultsViewer/simul_funcs_custom.cpp:788
- Options packing format:        ModelRiskAtl/SimulationObj.cpp:94
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import TYPE_CHECKING, Any

from modelrisk_mcp.bridge.vmrs_file import (
    ResultsHeader,
    read_dmr_header,
    read_vmrs_header,
    write_vmrs_from_dmr,
)
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
    # Set when ModelRisk's save handed back another run and the .vmrs was
    # rebuilt from this run's own file (SimulationController._save_results).
    note: str | None = None


# ---------------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------------


# XLL command names registered by ModelRisk.xll. CMDSTR_VISIBLE(X) →
# "Vose" + X (ModelRiskCloude/CLAUDE.md:63).
_CMD_START_SIM = "VoseStartSimulCustom12"
_CMD_GET_DATA_SZ = "VoseGetDataSZ12"
# Persists the simulation options to the workbook as SimOpt_* defined-names via
# SaveOptionsToWorkbook (SimulationCommonOptionsReadWrite.cpp). REQUIRED for the
# seed: ModelRisk's per-cell-twister Manual-Seed mode reads its base seed from the
# workbook-persisted SimOpt_SeedFixed / SimOpt_Seed0 names, NOT from the per-call
# options handed to VoseStartSimulCustom12. Without this call the per-call
# [SeedFixed]/[seed0] are parsed but ignored by the twister -> Random mode ->
# non-reproducible streams (AB#2742; verified end-to-end against the engine oracle).
_CMD_SET_SIM_OPTS = "VoseSetSimulOptions12"

# Operation prefix the SimulationObj_VBA dispatcher matches on for the
# save path. PackSessionName format: "h<hwnd>_<Operation>_<book_name>"
# (ModelRiskAtl/ModelRiskSimulationResults.cpp:54).
_OP_SAVE_RESULTS = "SaveResultsToFile"

# ModelRisk stamps a run's start to the second; allow that much when
# comparing it with the moment this process started the run.
_CLOCK_SLACK = timedelta(seconds=1)


class SimulationController:
    """Drives ModelRisk simulations through the XLL command surface
    exposed to `Application.Run`. No direct ATL COM dispatch.

    All methods raise typed `modelrisk_mcp.errors` exceptions; raw COM
    HRESULTs and xlwings stack traces never leak.

    It never calls `Application.RegisterXLL`: re-registering a live
    add-in re-runs its `xlAutoOpen`, which once destroyed a licence
    activation (ModelChoice, 2026-09-21). `ModelRiskBridge.run_simulation`
    proves the add-in live before any run (`ensure_modelrisk_functional`)
    and registers it only when a Vose function does not resolve.
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
        output_names: tuple[str, ...] = (),
    ) -> SimulationRunResult:
        """Run a simulation on `workbook_name` (defaults to the active
        workbook) and save the resulting `.vmrs` to `save_to` (defaults
        to a sibling of the workbook).

        The call blocks until the simulation completes — that's how
        `VoseStartSimulCustom12` is implemented (synchronous Application.Run).

        Bug #31 (alpha.30): defensive sanity check on `samples`.
        When a non-MCP caller (a direct script, an integration test,
        future Python clients) passes `samples <= 0`, the value used
        to flow straight through to the XLL which would throw a
        C++ exception (OLE error 0xe06d7363). The user saw
        "Application.Run failed" — opaque. The MCP tool layer's
        Pydantic field validates `ge=1` but the bridge had no such
        guard. Adding it here ensures every caller path produces a
        clear message before hitting the XLL.

        `output_names`, when supplied, is threaded into the XLL command's
        options payload as `[CntNames]:N` + `[name0]:...` etc. The
        original C++ header comment claims "empty → all outputs" but
        end-user testing against alpha.17 showed sims completing
        without registering ANY outputs in the .vmrs unless the
        ribbon path was used. The working hypothesis: the ribbon
        populates this list from the discovered VoseOutput cells, and
        the XLL command actually requires it. Defaulted to `()` here
        so the controller stays general-purpose; the bridge layer
        populates it from `list_outputs(workbook)`.

        Returns a SimulationRunResult with the resolved vmrs path.
        Raises SimulationFailedError unless the file at that path holds
        this run (see `_save_results`).
        """
        if samples < 1:
            raise SimulationFailedError(
                f"samples must be >= 1; got {samples}. ModelRisk's "
                f"XLL command would throw an opaque C++ exception "
                f"otherwise."
            )
        if samples > 10_000_000:
            # Soft cap: nothing in the bridge enforces this, but
            # production users rarely want more than a few million
            # iterations and the .vmrs file gets huge. Reject with
            # a clear message rather than letting MRService thrash.
            raise SimulationFailedError(
                f"samples={samples} exceeds the 10M soft cap. "
                f"If you genuinely need more, run multiple sims and "
                f"aggregate."
            )
        wb_info = self._resolve_workbook(workbook_name)
        opts = SimulationOptions(
            samples=samples,
            seeds=(seed,),
            seed_fixed=seed_fixed,
            hide_progress_window=hide_dialogs,
            show_results_at_end=False,
            refresh_excel=False,
            output_names=output_names,
        )
        target = self._resolve_save_path(wb_info, save_to)

        self._make_active(wb_info.name)
        started = datetime.now().replace(microsecond=0)
        self._invoke_start_simulation(opts)
        note = self._save_results(
            wb_info.name, target, started=started, samples=samples
        )
        return SimulationRunResult(
            workbook_name=wb_info.name,
            vmrs_path=target,
            iterations=samples,
            options=opts,
            note=note,
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

    def _make_active(self, book_name: str) -> None:
        """Make the workbook to simulate the active one. A run records the
        outputs of every open workbook but belongs to the ACTIVE one: its
        results and run file are filed under that workbook's name (seen
        live on 2026-10-04: asked for A while B was active, the run was
        B's, and the save for A was labelled B)."""
        if self._active_name() == book_name.casefold():
            return
        try:
            self._excel.activate_workbook(book_name)
        except Exception as exc:
            raise SimulationFailedError(
                f"Could not make {book_name!r} the active workbook: {exc}. "
                "ModelRisk simulates the active workbook, so the run was "
                "not started."
            ) from exc
        if self._active_name() != book_name.casefold():
            raise SimulationFailedError(
                f"{book_name!r} did not become the active workbook, so the "
                "run was not started: ModelRisk simulates the active "
                "workbook. Close any dialog open in Excel and retry."
            )

    def _active_name(self) -> str | None:
        try:
            return self._excel.get_active_workbook().name.casefold()
        except Exception:
            return None

    def _invoke_start_simulation(self, opts: SimulationOptions) -> None:
        """Application.Run("VoseStartSimulCustom12", options_array).

        `options_array` must be a 1-row 2D SAFEARRAY of BSTRs. pywin32
        converts a list-of-lists into a SAFEARRAY automatically when the
        target argument is a VARIANT."""
        app = self._app()
        try:
            options_2d = [opts.to_string_list()]  # 1 row x N cols
            # Persist options (esp. SeedFixed/seed0) to the workbook FIRST so the
            # per-cell-twister Manual-Seed engine actually honors the seed. See
            # _CMD_SET_SIM_OPTS above for the full rationale (AB#2742).
            app.api.Run(_CMD_SET_SIM_OPTS, options_2d)
            app.api.Run(_CMD_START_SIM, options_2d)
        except Exception as exc:
            raise SimulationFailedError(
                f"Application.Run({_CMD_START_SIM!r}, ...) failed: {exc}. "
                "ModelRisk add-in must be loaded and the workbook must "
                "contain at least one VoseOutput cell."
            ) from exc

    def _save_results(
        self, book_name: str, target: str, *, started: datetime, samples: int
    ) -> str | None:
        """Save the run just made of `book_name` to `target`, and prove it.

        ModelRisk's save hands the job to its Results Viewer process, and
        the viewer saves the run IT has loaded, whatever workbook was named
        (ModelRiskResultsViewer/simul_funcs_custom.cpp:788). Only when the
        viewer holds no run does the XLL save the named workbook's own run
        (SimulationObj_VBA.cpp:832). Our runs never load the viewer
        (ShowResultsAtEnd=0), so once it has shown one workbook's results
        (a run from the ribbon, the Results button), every later save
        writes that workbook's run file under the new name. Seen in the
        field on 2026-10-03: 22 saves held OG-05's run while each workbook
        showed its own, correct, results.

        So the saved file's header must name `book_name`, start no earlier
        than this run and hold `samples` iterations. If it does not, the
        `.vmrs` is rebuilt from the workbook's own ModelRisk run file,
        which is the file a correct save compresses. Returns a note when
        that was needed; raises when neither gives this run, after
        removing a wrong file the save wrote."""
        hwnd = self._hwnd()
        problem: str | None
        try:
            self._invoke_save_results(book_name, target, hwnd)
        except SimulationFailedError as exc:
            problem = f"an error ({exc})"
        else:
            problem = _file_problem(target, book_name, started, samples)
        if problem is None:
            return None

        run_file, misses = _find_run_file(book_name, started, samples)
        if run_file is not None:
            try:
                write_vmrs_from_dmr(run_file, target)
            except OSError as exc:
                misses.append(f"rebuilding it from {run_file} failed ({exc})")
            else:
                rebuilt = _file_problem(target, book_name, started, samples)
                if rebuilt is None:
                    return (
                        f"ModelRisk's save returned {problem}; the .vmrs was "
                        "rebuilt from ModelRisk's own run file for "
                        f"{book_name!r}."
                    )
                misses.append(f"the file rebuilt from {run_file} held {rebuilt}")

        _discard_if_written_since(target, started)
        raise SimulationFailedError(
            f"No results of this run of {book_name!r} could be saved to "
            f"{target!r}: ModelRisk's save returned {problem}, and "
            f"{'; '.join(misses)}. Another workbook's run comes back when "
            "ModelRisk's Results Viewer has that run loaded: close the "
            f"Results Viewer, or show {book_name!r}'s results in it, and run "
            "again. Nothing or an earlier run comes back when the simulation "
            "did not run: check that the workbook has VoseOutput cells and "
            "that the run was not cancelled."
        )

    def _hwnd(self) -> int:
        try:
            return int(self._app().api.Hwnd)
        except Exception as exc:
            raise SimulationFailedError(
                "Could not read Application.Hwnd to compose the save "
                "session name. Excel may have closed."
            ) from exc

    def _invoke_save_results(
        self, book_name: str, target_path: str, hwnd: int
    ) -> None:
        """Application.Run("VoseGetDataSZ12", session_name, target_path).

        Mirrors the ATL's IModelRiskSimulationResults::SaveResultsToFile
        path (ModelRiskAtl/ModelRiskSimulationResults.cpp:1196). The
        XLL handler in SimulationObj_VBA.cpp:805 reads xlParam1 as the
        target file and skips the file dialog when non-empty."""
        session_name = f"h{hwnd}_{_OP_SAVE_RESULTS}_{book_name}"
        try:
            self._app().api.Run(_CMD_GET_DATA_SZ, session_name, target_path)
        except Exception as exc:
            raise SimulationFailedError(
                f"Application.Run({_CMD_GET_DATA_SZ!r}, ...) failed during "
                f"SaveResultsToFile: {exc}"
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


# ---------------------------------------------------------------------------
# Checking a saved run
# ---------------------------------------------------------------------------


def _run_problem(
    header: ResultsHeader, book_name: str, started: datetime, samples: int
) -> str | None:
    """What `header` holds instead of this run, or None if it is this run."""
    when = (
        f", started {header.start_date:%Y-%m-%d %H:%M:%S}"
        if header.start_date
        else ""
    )
    if header.spreadsheet_name.casefold() != book_name.casefold():
        return f"the results of {header.spreadsheet_name!r}{when}"
    if header.start_date is None or header.start_date < started - _CLOCK_SLACK:
        return (
            f"an earlier run of {book_name!r}{when}, not the run started "
            f"{started:%Y-%m-%d %H:%M:%S}"
        )
    if header.iterations != samples:
        return f"a run of {header.iterations} iterations, not {samples}"
    return None


def _file_problem(
    path: str, book_name: str, started: datetime, samples: int
) -> str | None:
    """What the `.vmrs` at `path` holds instead of this run, or None."""
    if not Path(path).is_file():
        return "nothing (no .vmrs was produced)"
    try:
        header = read_vmrs_header(path)
    except (OSError, ValueError) as exc:
        return f"a file that is not a ModelRisk results file ({exc})"
    return _run_problem(header, book_name, started, samples)


def _find_run_file(
    book_name: str, started: datetime, samples: int
) -> tuple[Path | None, list[str]]:
    """ModelRisk's own file for this run of `book_name`: `f<hWndExcel as
    %X><book>.dmr` in its simulation folder (SimulationObj.cpp:6163).

    The prefix is matched as any hex number: the engine fixes hWndExcel at
    start-up, and in Excel's one-window-per-workbook UI that is the first
    workbook's window, while Application.Hwnd follows the active one (seen
    live: run files `f50108…` while Hwnd read 0xB0822). The header check
    picks this run among the candidates, newest first. Returns that file,
    and what was found instead in every place looked at."""
    pattern = re.compile(r"f[0-9A-F]+" + re.escape(book_name) + r"\.dmr", re.I)
    misses: list[str] = []
    for folder in _simulation_folders():
        try:
            candidates = [p for p in folder.iterdir() if pattern.fullmatch(p.name)]
        except OSError:
            continue
        candidates.sort(key=_mtime, reverse=True)
        for path in candidates:
            try:
                problem = _run_problem(
                    read_dmr_header(path), book_name, started, samples
                )
            except (OSError, ValueError) as exc:
                problem = f"nothing readable ({exc})"
            if problem is None:
                return path, misses
            misses.append(f"its run file {path} holds {problem}")
    if not misses:
        looked = ", ".join(str(f) for f in _simulation_folders())
        misses.append(f"no run file f…{book_name}.dmr was found in {looked}")
    return None, misses


def _mtime(path: Path) -> float:
    try:
        return path.stat().st_mtime
    except OSError:
        return 0.0


def _simulation_folders() -> list[Path]:
    """Where ModelRisk keeps run files, in the order its GetAppDataFolder
    picks them (ModelRiskCloude/GetAppData.cpp:390). A custom folder set
    in ModelRisk's folders file is not read; the header check rejects any
    file that is not this run, so a wrong guess cannot cause harm."""
    folders = [Path("C:/Vose Software/ModelRisk/SimulationStorage")]
    temp = os.environ.get("TEMP")
    if temp:
        folders.append(Path(temp) / "SimulationStorage")
    folders.append(Path.home() / "Documents" / "SimulationStorage")
    return folders


def _discard_if_written_since(path: str, started: datetime) -> None:
    """Remove a file written during this run, so a wrong `.vmrs` is not
    left for a later reader. An older file at `path` is left alone."""
    try:
        target = Path(path)
        if target.stat().st_mtime >= started.timestamp():
            target.unlink()
    except OSError:
        pass


__all__ = [
    "SimulationController",
    "SimulationOptions",
    "SimulationRunResult",
]
