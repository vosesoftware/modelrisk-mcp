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

`engine="turbo"` replaces step 1 with ModelRisk's Turbo command,
**VoseStartFastSimulation**, after the same VoseSetSimulOptions12 call:
Turbo reads its samples and seed from those options. Turbo cannot
evaluate every function, so the engine's own check runs first on a saved
copy of the workbook, and the classic engine runs instead when Turbo
cannot run it. The command's message boxes are answered while it runs
(`bridge/turbo.py`). Its run file is `f<hWndExcel>vsmre_<stem>.dmr`.

References:
- VoseStartSimulCustom12 export: ModelRiskCloude/XllAddIn.cpp:210
- VoseGetDataSZ12 export:        ModelRiskCloude/XllAddIn.cpp:207
- Session-name format:           ModelRiskAtl/ModelRiskSimulationResults.cpp:54
- Save handler:                  ModelRiskCloude/SimulationObj_VBA.cpp:805
- Viewer's save handler:         ModelRiskResultsViewer/simul_funcs_custom.cpp:788
- Options packing format:        ModelRiskAtl/SimulationObj.cpp:94
- Turbo command and run:         ModelRiskCloude/simul_funcs.cpp:605, MREngine.cpp:1024
"""

from __future__ import annotations

import os
import re
import shutil
import tempfile
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import TYPE_CHECKING, Any

from modelrisk_mcp.bridge import turbo
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
    # rebuilt from this run's own file (SimulationController._save_results),
    # or when Turbo was asked for and the classic engine ran instead.
    note: str | None = None
    # The engine whose run was saved: "classic" or "turbo".
    engine: str = "classic"


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
# The Turbo engine's run of the active workbook (simul_funcs.cpp:605). It
# takes no options and returns 1 whatever happens; a failure shows a
# message box. Samples and seeds come from VoseSetSimulOptions12.
_CMD_START_TURBO = "VoseStartFastSimulation"
ENGINES = ("classic", "turbo")

# Operation prefix the SimulationObj_VBA dispatcher matches on for the
# save path. PackSessionName format: "h<hwnd>_<Operation>_<book_name>"
# (ModelRiskAtl/ModelRiskSimulationResults.cpp:54).
_OP_SAVE_RESULTS = "SaveResultsToFile"

# Excel's xlCalculationManual.
_XL_CALC_MANUAL = -4135

# A workbook activated through COM can be switched back within half a
# second when another process's window is in front: the Results Viewer,
# which a Turbo run brings forward, hands the focus back to the Excel
# window it is attached to (seen live on 2026-10-04). So an activation is
# checked again after this long, and repeated; the second one holds,
# because Excel is then in front.
_ACTIVATE_SETTLE = 0.5
_ACTIVATE_ATTEMPTS = 3


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
        engine: str = "classic",
    ) -> SimulationRunResult:
        """Run a simulation on `workbook_name` (defaults to the active
        workbook) and save the resulting `.vmrs` to `save_to` (defaults
        to a sibling of the workbook).

        The call blocks until the simulation completes — that's how
        `VoseStartSimulCustom12` is implemented (synchronous Application.Run).

        `engine="turbo"` runs ModelRisk's Turbo engine. When Turbo cannot
        run the workbook (`_turbo_obstacle`), the classic engine runs and
        `note` says why; `engine` on the result names the engine whose run
        was saved.

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
        if engine not in ENGINES:
            raise SimulationFailedError(
                f"engine must be one of {', '.join(map(repr, ENGINES))}; "
                f"got {engine!r}."
            )
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

        notes: list[str] = []
        if engine == "turbo":
            obstacle = self._turbo_obstacle(wb_info.name)
            if obstacle is not None:
                notes.append(f"{obstacle}, so the classic engine ran instead.")
                engine = "classic"
        self._make_active(wb_info.name)
        started = _fresh_second()
        messages: list[str] = []
        if engine == "turbo":
            messages = self._invoke_turbo_simulation(opts)
        else:
            self._invoke_start_simulation(opts)
        try:
            saved = self._save_results(
                wb_info.name, target, started=started, samples=samples,
                engine=engine,
            )
        except SimulationFailedError as exc:
            if messages:
                raise SimulationFailedError(
                    f"ModelRisk's Turbo run of {wb_info.name!r} stopped with "
                    f"this message: {_quoted(messages)} No results were "
                    "saved. engine='classic' runs the workbook with "
                    "ModelRisk's own engine."
                ) from exc
            raise
        if messages:
            notes.append(f"ModelRisk said: {_quoted(messages)}")
        if saved:
            notes.append(saved)
        return SimulationRunResult(
            workbook_name=wb_info.name,
            vmrs_path=target,
            iterations=samples,
            options=opts,
            note=" ".join(notes) or None,
            engine=engine,
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
        B's, and the save for A was labelled B). An activation must still
        hold after `_ACTIVATE_SETTLE`; one that is switched back is
        repeated."""
        wanted = book_name.casefold()
        if self._active_name() == wanted:
            return
        for _ in range(_ACTIVATE_ATTEMPTS):
            try:
                self._excel.activate_workbook(book_name)
            except Exception as exc:
                raise SimulationFailedError(
                    f"Could not make {book_name!r} the active workbook: {exc}. "
                    "ModelRisk simulates the active workbook, so the run was "
                    "not started."
                ) from exc
            if self._active_name() != wanted:
                continue
            time.sleep(_ACTIVATE_SETTLE)
            if self._active_name() == wanted:
                return
        raise SimulationFailedError(
            f"{book_name!r} did not stay the active workbook, so the run was "
            "not started: ModelRisk simulates the active workbook. Close any "
            "dialog open in Excel and retry."
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
            self._persist_options(app, options_2d)
            app.api.Run(_CMD_START_SIM, options_2d)
        except Exception as exc:
            raise SimulationFailedError(
                f"Application.Run({_CMD_START_SIM!r}, ...) failed: {exc}. "
                "ModelRisk add-in must be loaded and the workbook must "
                "contain at least one VoseOutput cell."
            ) from exc

    @staticmethod
    def _persist_options(app: Any, options_2d: list[list[str]]) -> None:
        """Application.Run("VoseSetSimulOptions12", options_array), under
        manual calculation. The command writes each option as a workbook
        name (29 of them), and under automatic calculation every write
        recalculates the workbook: 52.7 s, against 1.9 s for the single
        recalculation that restoring the mode makes, on CF-01 with a
        10,000-iteration run attached (191 VoseSim cells, 2026-10-04).
        ModelRisk holds manual calculation for its own option writes the
        same way (MREngine.cpp:1066)."""
        try:
            mode = app.api.Calculation
        except Exception:
            mode = _XL_CALC_MANUAL  # unreadable: leave it alone
        if mode != _XL_CALC_MANUAL:
            app.api.Calculation = _XL_CALC_MANUAL
        try:
            app.api.Run(_CMD_SET_SIM_OPTS, options_2d)
        finally:
            if mode != _XL_CALC_MANUAL:
                app.api.Calculation = mode

    def _turbo_obstacle(self, book_name: str) -> str | None:
        """Why Turbo cannot run `book_name`, or None when it can. Turbo
        reads .xlsx and .xlsm files only (MREngine.cpp:210), and an output
        that depends on a function it cannot evaluate comes back NaN with
        no warning. The engine's own check of those functions runs on a
        saved copy of the workbook, as it is in Excel now."""
        suffix = Path(book_name).suffix.lower()
        if suffix and suffix not in turbo.TURBO_FORMATS:
            return "Turbo runs only .xlsx and .xlsm workbooks"
        if self._excel_pid() is None:
            return (
                "Excel's process could not be identified, and Turbo's "
                "message boxes could not be answered without it"
            )
        folder = Path(tempfile.mkdtemp(prefix="modelrisk-mcp-turbo-"))
        try:
            copy = folder / f"turbo-check{suffix or '.xlsx'}"
            try:
                self._excel.save_workbook_as(book_name, str(copy), overwrite=True)
            except Exception as exc:
                return (
                    "a copy of the workbook for Turbo's compatibility check "
                    f"could not be saved ({exc})"
                )
            check = turbo.check_workbook(str(copy))
        finally:
            shutil.rmtree(folder, ignore_errors=True)
        if check.problem is not None:
            return f"Turbo's compatibility check could not run ({check.problem})"
        if check.unsupported:
            return (
                f"Turbo cannot evaluate {', '.join(check.unsupported)}, "
                f"which {book_name!r} uses"
            )
        return None

    def _invoke_turbo_simulation(self, opts: SimulationOptions) -> list[str]:
        """Application.Run("VoseStartFastSimulation") on the active
        workbook, after VoseSetSimulOptions12 has set the samples and seed
        Turbo reads (MREngine.cpp:1130, MREngine_GenerateSeeds). Its
        message boxes are answered while it runs: No to the offer to check
        Turbo against the classic engine (which would run the classic
        engine too, and wait for a click on its result), OK to a message.
        Returns the text of each message, which is how Turbo reports a
        failed run."""
        app = self._app()
        pid = self._excel_pid()
        if pid is None:
            raise SimulationFailedError(
                "Excel's process could not be identified; the Turbo run was "
                "not started."
            )
        watch = turbo.DialogWatch(pid)
        try:
            self._persist_options(app, [opts.to_string_list()])
            with watch:
                app.api.Run(_CMD_START_TURBO)
        except Exception as exc:
            raise SimulationFailedError(
                f"Application.Run({_CMD_START_TURBO!r}) failed: {exc}. The "
                "ModelRisk add-in must be loaded, in a version with the Turbo "
                "engine."
            ) from exc
        return [
            " ".join(box.text.split())
            for box in watch.answered
            if not box.is_question
        ]

    def _excel_pid(self) -> int | None:
        try:
            return int(self._app().pid)
        except Exception:
            return None

    def _save_results(
        self,
        book_name: str,
        target: str,
        *,
        started: datetime,
        samples: int,
        engine: str = "classic",
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
        removing a wrong file the save wrote.

        A Turbo run ends by loading itself into the viewer, so the save
        normally holds it; its run file is named after the engine's copy
        of the workbook, `vsmre_<stem>`, and names the workbook inside."""
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

        stems: tuple[str, ...] = (book_name,)
        if engine == "turbo":
            stems = (f"vsmre_{Path(book_name).stem}", book_name)
        run_file, misses = _find_run_file(book_name, started, samples, stems)
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
        turbo_hint = (
            " A Turbo run also ends without results when its engine stops, "
            "or when it is stopped or sent to the background in its progress "
            "window; engine='classic' runs the workbook with ModelRisk's own "
            "engine."
            if engine == "turbo"
            else ""
        )
        raise SimulationFailedError(
            f"No results of this run of {book_name!r} could be saved to "
            f"{target!r}: ModelRisk's save returned {problem}, and "
            f"{'; '.join(misses)}. Another workbook's run comes back when "
            "ModelRisk's Results Viewer has that run loaded: close the "
            f"Results Viewer, or show {book_name!r}'s results in it, and run "
            "again. Nothing or an earlier run comes back when the simulation "
            "did not run: check that the workbook has VoseOutput cells and "
            f"that the run was not cancelled.{turbo_hint}"
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
    if header.start_date is None or header.start_date < started:
        return (
            f"an earlier run of {book_name!r}{when}, not the run started "
            f"{started:%Y-%m-%d %H:%M:%S}"
        )
    if header.iterations != samples:
        return f"a run of {header.iterations} iterations, not {samples}"
    return None


def _fresh_second() -> datetime:
    """The start of the next whole second, returned once it has come.
    ModelRisk stamps a run's start to the second, so a run started earlier
    in the current second would carry the same stamp as this one; from the
    next second on, only runs started after this call can. Seen live on
    2026-10-04: a Turbo run stamped 18:49:14 and repeated at once with the
    classic engine was taken for the classic run, started 18:49:15, under
    the one-second allowance this replaces."""
    target = datetime.now().replace(microsecond=0) + timedelta(seconds=1)
    while (wait := (target - datetime.now()).total_seconds()) > 0:
        time.sleep(wait)
    return target


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
    book_name: str,
    started: datetime,
    samples: int,
    stems: tuple[str, ...] | None = None,
) -> tuple[Path | None, list[str]]:
    """ModelRisk's own file for this run of `book_name`: `f<hWndExcel as
    %X><stem>.dmr` in its simulation folder, where the stem is the
    workbook's name (SimulationObj.cpp:6163) or, for a Turbo run, the name
    of the engine's copy of it, `vsmre_<stem>` (SimulationObj.cpp:14690).

    The prefix is matched as any hex number: the engine fixes hWndExcel at
    start-up, and in Excel's one-window-per-workbook UI that is the first
    workbook's window, while Application.Hwnd follows the active one (seen
    live: run files `f50108…` while Hwnd read 0xB0822). The header check
    picks this run among the candidates, newest first. Returns that file,
    and what was found instead in every place looked at."""
    names = "|".join(re.escape(stem) for stem in (stems or (book_name,)))
    pattern = re.compile(r"f[0-9A-F]+(?:" + names + r")\.dmr", re.I)
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
        wanted = " or ".join(f"f…{stem}.dmr" for stem in (stems or (book_name,)))
        misses.append(f"no run file {wanted} was found in {looked}")
    return None, misses


def _quoted(messages: list[str]) -> str:
    return " | ".join(f'"{m}"' for m in messages)


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
    "ENGINES",
    "SimulationController",
    "SimulationOptions",
    "SimulationRunResult",
]
