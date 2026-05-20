"""ModelRiskBridge — domain layer on top of ExcelBridge.

Phase 1 skeleton: implements the read-only inspection methods needed by
the §7.1 reading tools (`list_modelrisk_inputs`, `list_modelrisk_outputs`,
`list_distributions`, `find_hard_coded_inputs`). Building / simulation /
results methods land in later phases.

Nothing here touches COM directly — all Excel access goes through
`ExcelBridge`. Catalogue access goes through `FunctionCatalogue`.
"""

from __future__ import annotations

import json
import re
import struct
import sys
import winreg
from datetime import datetime
from pathlib import Path
from typing import Any

from modelrisk_mcp.bridge.catalogue import FunctionCatalogue, load_catalogue
from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.bridge.progids import (
    CLSID_DISTRIBUTIONS,
    CLSID_SIMULATION,
    PROGID_DISTRIBUTIONS,
)
from modelrisk_mcp.bridge.results import ResultsReader
from modelrisk_mcp.bridge.simulation import SimulationController
from modelrisk_mcp.config import Settings
from modelrisk_mcp.errors import CellReferenceError
from modelrisk_mcp.safety import (
    WriterMutex,
    append_write_log,
    extract_call_heads,
    is_vose_formula,
)
from modelrisk_mcp.schemas.distributions import InsertResult
from modelrisk_mcp.schemas.results import (
    CorrelationMatrix,
    SensitivityRanking,
    SimulationResult,
)
from modelrisk_mcp.schemas.workbook import (
    CellInfo,
    CellRef,
    DistributionCell,
    ModelRiskInput,
    ModelRiskOutput,
    WorkbookSummary,
)

# Pattern for `VoseInput("foo")` or `VoseOutput("foo")` at any position
# in a formula. The captured group is the name.
_VOSE_INPUT_RE = re.compile(r'VoseInput\(\s*"((?:[^"\\]|\\.|"")*?)"\s*\)')
_VOSE_OUTPUT_RE = re.compile(r'VoseOutput\(\s*"((?:[^"\\]|\\.|"")*?)"\s*\)')

# Categories whose presence in a cell makes it a "distribution cell"
# (per spec §7.1 detection rules).
_DISTRIBUTION_CATEGORIES: frozenset[str] = frozenset({
    "continuous",
    "discrete",
    "aggregate",
    "time-series",
    "copula",
    "fitting",
    "object",
})


class ModelRiskBridge:
    def __init__(
        self,
        excel: ExcelBridge,
        catalogue: FunctionCatalogue | None = None,
        results: ResultsReader | None = None,
        simulation: SimulationController | None = None,
        settings: Settings | None = None,
        writer_mutex: WriterMutex | None = None,
    ) -> None:
        self._excel = excel
        self._catalogue = catalogue or load_catalogue()
        self._results = results or ResultsReader()
        self._simulation = simulation or SimulationController()
        self._settings = settings or Settings()
        self._writer_mutex = writer_mutex or WriterMutex()

    @property
    def catalogue(self) -> FunctionCatalogue:
        return self._catalogue

    @property
    def excel(self) -> ExcelBridge:
        return self._excel

    @property
    def results(self) -> ResultsReader:
        return self._results

    @property
    def simulation(self) -> SimulationController:
        return self._simulation

    # ------------------------------------------------------------------
    # Environment checks
    # ------------------------------------------------------------------

    def is_modelrisk_loaded(self) -> bool:
        """Returns True if the ModelRisk COM surface is reachable.

        First attempts a direct Dispatch. If that fails, asks Excel to
        enable any ModelRisk-named add-in (Vose's COM-add-in or XLL),
        then retries. Most callers should use this instead of probing
        Dispatch themselves — it handles the common 'ModelRisk
        installed but not loaded into the active Excel session' case
        transparently."""
        if self._try_dispatch():
            return True
        # Dispatch failed. Try to enable ModelRisk inside Excel.
        try:
            self.ensure_modelrisk_active()
        except Exception:
            pass
        return self._try_dispatch()

    def _try_dispatch(self) -> bool:
        ok, _err = self._try_dispatch_with_error()
        return ok

    def _try_dispatch_with_error(self) -> tuple[bool, str | None]:
        """Returns (success, error_message). Captures the COM exception
        so the diagnostic can surface the actual HRESULT to the LLM —
        invaluable for distinguishing 'class not registered' from a
        bitness mismatch from E_NOINTERFACE (the usual gen_py-cache
        or wrong-coclass culprit)."""
        ok, _strategy, err = self._dispatch_via_first_working_strategy()
        return ok, err

    def _dispatch_via_first_working_strategy(
        self,
    ) -> tuple[bool, str | None, str | None]:
        """Try every plausible Dispatch path and return the first one
        that returns a usable object. Returns (success, strategy_name,
        error_message). On success, error_message is None; on failure,
        strategy_name is None and error_message summarises all the
        attempts."""
        results = self.diagnose_dispatch_strategies()
        for strategy in (
            "dispatch_ex",     # always-late-bound; bypasses gen_py cache
            "co_create",       # bypasses pywin32 entirely
            "dispatch",        # may hit gen_py cache
            "via_comaddin",    # ask Excel for its in-process instance
        ):
            outcome = results.get(strategy, {})
            if outcome.get("ok"):
                return True, strategy, None
        # Compose a one-line summary of every attempt for the diagnostic.
        summary = "; ".join(
            f"{name}={outcome.get('error') or 'ok'}"
            for name, outcome in results.items()
        )
        return False, None, summary

    def diagnose_dispatch_strategies(self) -> dict[str, dict[str, Any]]:
        """Run every dispatch strategy and report each outcome.

        Strategies:
          - dispatch:    win32com.client.Dispatch(progid) — the standard
                         path; may pick up a stale gen_py cache.
          - dispatch_ex: win32com.client.DispatchEx(progid) — always
                         late-bound; bypasses gen_py.
          - co_create:   pythoncom.CoCreateInstance(clsid, ...,
                         IID_IDispatch) — pywin32's lowest-level path.
          - via_comaddin: walk Excel.COMAddIns for a ModelRisk entry and
                         return its `.Object`. Useful when the COM coclass
                         requires Office to have initialised it.
        """
        out: dict[str, dict[str, Any]] = {}
        try:
            import win32com.client as com
        except ImportError as exc:
            for name in ("dispatch", "dispatch_ex", "co_create", "via_comaddin"):
                out[name] = {"ok": False, "error": f"pywin32 missing: {exc}"}
            return out

        # 1. plain Dispatch
        try:
            obj = com.Dispatch(PROGID_DISTRIBUTIONS)
            out["dispatch"] = {"ok": obj is not None, "error": None}
        except Exception as exc:
            out["dispatch"] = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}

        # 2. DispatchEx — late-bound, ignores gen_py cache
        try:
            obj = com.DispatchEx(PROGID_DISTRIBUTIONS)
            out["dispatch_ex"] = {"ok": obj is not None, "error": None}
        except Exception as exc:
            out["dispatch_ex"] = {
                "ok": False,
                "error": f"{type(exc).__name__}: {exc}",
            }

        # 3. CoCreateInstance via pythoncom
        try:
            import pythoncom
            clsid = pythoncom.MakeIID(CLSID_DISTRIBUTIONS)
            obj = pythoncom.CoCreateInstance(
                clsid,
                None,
                pythoncom.CLSCTX_INPROC_SERVER | pythoncom.CLSCTX_LOCAL_SERVER,
                pythoncom.IID_IDispatch,
            )
            out["co_create"] = {"ok": obj is not None, "error": None}
        except Exception as exc:
            out["co_create"] = {
                "ok": False,
                "error": f"{type(exc).__name__}: {exc}",
            }

        # 4. via Excel.COMAddIns.Object — the add-in's own IDispatch
        out["via_comaddin"] = self._try_via_comaddin()
        return out

    def _try_via_comaddin(self) -> dict[str, Any]:
        """Walk Excel.COMAddIns for a ModelRisk entry and grab its
        `.Object` property. Returns the same {ok, error} shape as the
        other strategies; on success also records the add-in name we
        found."""
        try:
            app = getattr(self._excel, "_app", None)
            if app is None:
                # Try to attach; tolerate failure.
                self._excel.connect()
                app = getattr(self._excel, "_app", None)
            if app is None:
                return {"ok": False, "error": "Excel not reachable"}
            for addin in app.api.COMAddIns:
                desc = str(getattr(addin, "Description", "") or "")
                progid = str(getattr(addin, "ProgID", "") or "")
                if "modelrisk" not in (desc + progid).lower() and "vose" not in (
                    desc + progid
                ).lower():
                    continue
                try:
                    obj = addin.Object
                except Exception as exc:
                    return {
                        "ok": False,
                        "error": f"COMAddIn({desc!r}).Object: {exc}",
                    }
                if obj is None:
                    continue
                return {"ok": True, "error": None, "addin_name": desc or progid}
        except Exception as exc:
            return {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
        return {"ok": False, "error": "no ModelRisk COMAddIn found"}

    def ensure_modelrisk_active(self) -> dict[str, Any]:
        """Make sure the ModelRisk add-in is loaded inside the running
        Excel session.

        Scans both `Excel.COMAddIns` (the ModelRisk extended UI is a
        COM add-in) and `Excel.AddIns` (the worksheet-function XLL
        registers here), flips `.Connect = True` / `.Installed = True`
        on any entry whose name mentions ModelRisk or Vose, and reports
        which entries it touched. Idempotent: skips entries already on.

        Returns a diagnostic dict the LLM can surface to the user:
            {
              "com_addins_enabled": [...],
              "excel_addins_enabled": [...],
              "com_addins_seen": [...],   # for diagnostics
              "excel_addins_seen": [...],
              "modelrisk_dispatchable": bool,   # final Dispatch outcome
            }
        """
        def _is_modelrisk(info: dict[str, Any]) -> bool:
            blob = " ".join(str(v) for v in info.values()).lower()
            return "modelrisk" in blob or "vose" in blob

        com_seen = self._excel.list_com_addins()
        xll_seen = self._excel.list_excel_addins()
        com_enabled = self._excel.enable_com_addin(_is_modelrisk)
        xll_enabled = self._excel.enable_excel_addin(_is_modelrisk)
        dispatchable, dispatch_error = self._try_dispatch_with_error()
        diag: dict[str, Any] = {
            "com_addins_enabled": com_enabled,
            "excel_addins_enabled": xll_enabled,
            "com_addins_seen": [
                a["description"] or a["progid"] for a in com_seen
            ],
            "excel_addins_seen": [a["name"] for a in xll_seen],
            "modelrisk_dispatchable": dispatchable,
            "dispatch_error": dispatch_error,
            "com_addins_already_connected": [
                a["description"] or a["progid"]
                for a in com_seen
                if a["connected"] and _is_modelrisk(a)
            ],
            "excel_addins_already_installed": [
                a["name"] for a in xll_seen if a["installed"] and _is_modelrisk(a)
            ],
        }
        if not dispatchable:
            diag["bitness"] = self._bitness_report()
            diag["dispatch_strategies"] = self.diagnose_dispatch_strategies()
            diag["root_cause_hypothesis"] = self._classify_root_cause(diag)
        return diag

    def _bitness_report(self) -> dict[str, Any]:
        """Gather a Python / Excel / ModelRiskAtl.dll bitness snapshot.

        Each field is best-effort — failures degrade to None rather
        than raising, because this is diagnostic code that runs *after*
        something else is already broken."""
        report: dict[str, Any] = {
            "python_bits": 64 if sys.maxsize > 2**32 else 32,
            "excel_path": None,
            "excel_bits_guess": None,
            "modelriskatl_clsid": None,
            "modelriskatl_path": None,
            "modelriskatl_bits": None,
        }
        # Excel path → bitness guess via "Program Files (x86)" marker.
        try:
            app = getattr(self._excel, "_app", None)
            if app is not None:
                excel_path = str(app.api.Path)
                report["excel_path"] = excel_path
                report["excel_bits_guess"] = (
                    32 if "(x86)" in excel_path else 64
                )
        except Exception:
            pass
        # ModelRiskAtl.dll registered path via HKCR.
        try:
            clsid, dll_path = _lookup_modelrisk_inproc_server()
            report["modelriskatl_clsid"] = clsid
            report["modelriskatl_path"] = dll_path
            if dll_path:
                report["modelriskatl_bits"] = _pe_bits(dll_path)
        except Exception:
            pass
        return report

    def _classify_root_cause(self, diag: dict[str, Any]) -> str:
        bits = diag.get("bitness") or {}
        py_bits = bits.get("python_bits")
        dll_bits = bits.get("modelriskatl_bits")
        strategies = diag.get("dispatch_strategies") or {}
        # If one strategy worked, point at that.
        winners = [n for n, r in strategies.items() if r.get("ok")]
        if winners:
            return (
                f"Dispatch succeeded via {winners[0]!r}. The primary "
                f"path is broken (gen_py cache pollution, wrong coclass, "
                f"or integrity-level mismatch) but the fallback works. "
                f"This run will use that path."
            )
        # Heuristic on the dominant error.
        first_err = str(
            (strategies.get("dispatch") or {}).get("error") or ""
        ).lower()
        if "no such interface" in first_err or "e_nointerface" in first_err:
            return (
                "E_NOINTERFACE on every strategy. Likely causes: "
                "(1) stale pywin32 gen_py cache — delete "
                "%LOCALAPPDATA%\\Temp\\gen_py\\ and retry. "
                "(2) Excel running elevated while Python isn't, or "
                "vice versa — COM Dispatch across integrity levels "
                "returns this exact HRESULT. (3) ModelRiskAtl.dll's "
                "TypeLib registration is broken — try "
                "`regsvr32 \"" + (bits.get("modelriskatl_path") or "ModelRiskAtl.dll") + "\"` "
                "from an elevated prompt."
            )
        if dll_bits and py_bits and dll_bits != py_bits:
            return (
                f"BITNESS MISMATCH: ModelRiskAtl.dll is {dll_bits}-bit; "
                f"this Python is {py_bits}-bit. COM can only load "
                f"matching bitness in-process. Install a "
                f"{dll_bits}-bit Python (or run modelrisk-mcp under one) "
                f"and rerun `uv sync` against that interpreter."
            )
        if dll_bits is None and bits.get("modelriskatl_clsid"):
            return (
                "ModelRiskAtl.dll path was found in the registry but "
                "the file isn't readable or has no PE header — the "
                "registered path may be broken. Reinstall ModelRisk "
                "or run `regsvr32 ModelRiskAtl.dll` from the install dir."
            )
        if not bits.get("modelriskatl_clsid"):
            return (
                "ModelRisk's CLSID isn't registered in HKCR — the COM "
                "self-registration step didn't run, or it ran under a "
                "different bitness's registry hive (HKCR is reflected). "
                "Try `regsvr32 ModelRiskAtl.dll` from the ModelRisk "
                "install folder; on x64 systems with a 32-bit DLL, use "
                "`%SystemRoot%\\SysWOW64\\regsvr32.exe`."
            )
        if diag.get("com_addins_already_connected") or diag.get(
            "excel_addins_already_installed"
        ):
            return (
                "Add-ins are loaded in Excel but Dispatch still fails. "
                "Most common cause: bitness mismatch couldn't be confirmed "
                "automatically. Compare your Excel install path "
                f"({bits.get('excel_path')!r}) and the Python "
                f"interpreter bitness ({py_bits}-bit)."
            )
        return "Unrecognised failure mode. Inspect `dispatch_error`."

    # ------------------------------------------------------------------
    # Reading tools (§7.1)
    # ------------------------------------------------------------------

    def list_inputs(self, workbook: str) -> list[ModelRiskInput]:
        result: list[ModelRiskInput] = []
        for cell in self._excel.iterate_cells(workbook):
            if not cell.formula:
                continue
            m = _VOSE_INPUT_RE.search(cell.formula)
            if not m:
                continue
            result.append(
                ModelRiskInput(
                    ref=cell.ref,
                    name=_unescape_excel_string(m.group(1)),
                    formula=cell.formula,
                    current_value=_coerce_displayable(cell.value),
                )
            )
        return result

    def list_outputs(self, workbook: str) -> list[ModelRiskOutput]:
        result: list[ModelRiskOutput] = []
        for cell in self._excel.iterate_cells(workbook):
            if not cell.formula:
                continue
            m = _VOSE_OUTPUT_RE.search(cell.formula)
            if not m:
                continue
            result.append(
                ModelRiskOutput(
                    ref=cell.ref,
                    name=_unescape_excel_string(m.group(1)),
                    formula=cell.formula,
                    current_value=_coerce_displayable(cell.value),
                )
            )
        return result

    def list_distributions(
        self, workbook: str, *, sheet: str | None = None
    ) -> list[DistributionCell]:
        result: list[DistributionCell] = []
        for cell in self._excel.iterate_cells(workbook, sheet=sheet):
            if not cell.formula:
                continue
            dist_head = self._first_distribution_head(cell.formula)
            if dist_head is None:
                continue
            result.append(
                DistributionCell(
                    ref=cell.ref,
                    function_name=dist_head,
                    parameters=self._extract_top_level_args(cell.formula, dist_head),
                    has_input_wrapper=bool(_VOSE_INPUT_RE.search(cell.formula)),
                    has_output_wrapper=bool(_VOSE_OUTPUT_RE.search(cell.formula)),
                    formula=cell.formula,
                )
            )
        return result

    def get_workbook_summary(self, workbook: str) -> WorkbookSummary:
        """Aggregate counts over every used cell in the workbook in a
        single pass, so the LLM gets the high-level picture without
        running four separate tools."""
        wb_info = next(
            (w for w in self._excel.list_workbooks() if w.name == workbook),
            None,
        )
        sheets = list(wb_info.sheets) if wb_info else []
        input_count = 0
        output_count = 0
        distribution_count = 0
        formula_cell_count = 0
        numeric_cell_count = 0
        for cell in self._excel.iterate_cells(workbook):
            if cell.formula:
                formula_cell_count += 1
                if _VOSE_INPUT_RE.search(cell.formula):
                    input_count += 1
                if _VOSE_OUTPUT_RE.search(cell.formula):
                    output_count += 1
                if self._first_distribution_head(cell.formula) is not None:
                    distribution_count += 1
            elif isinstance(cell.value, (int, float)) and not isinstance(
                cell.value, bool
            ):
                numeric_cell_count += 1
        return WorkbookSummary(
            workbook=workbook,
            sheets=sheets,
            input_count=input_count,
            output_count=output_count,
            distribution_count=distribution_count,
            formula_cell_count=formula_cell_count,
            numeric_cell_count=numeric_cell_count,
            modelrisk_loaded=self.is_modelrisk_loaded(),
        )

    def get_simulation_results(
        self, output_names: list[str] | None = None
    ) -> list[SimulationResult]:
        return self._results.get_simulation_results(output_names)

    def get_correlation_matrix(
        self, names: list[str] | None = None
    ) -> CorrelationMatrix:
        return self._results.get_correlation_matrix(names)

    def get_sensitivity_ranking(self, output_name: str) -> SensitivityRanking:
        return self._results.get_sensitivity_ranking(output_name)

    # ------------------------------------------------------------------
    # Building-tool support (Phase 3) — every write goes through here
    # ------------------------------------------------------------------

    def safe_write_cell(
        self,
        ref: CellRef,
        new_formula: str,
        *,
        allow_overwrite_non_vose: bool = False,
    ) -> InsertResult:
        """Wraps every cell write with the §11 safety mechanisms:
        writer-mutex acquisition, non-Vose-formula refusal,
        before/after audit log append. Returns an InsertResult whose
        `previous_formula` field carries what we replaced."""
        with self._writer_mutex.held(timeout_ms=0):
            existing = self._excel.get_cell(ref.workbook, ref.sheet, ref.cell)
            previous_formula = existing.formula or ""
            previous_value = existing.value
            if (
                previous_formula
                and not allow_overwrite_non_vose
                and not is_vose_formula(previous_formula, self._catalogue)
            ):
                raise CellReferenceError(
                    f"Refusing to overwrite cell {ref.a1!r}: it contains a "
                    f"non-Vose formula ({previous_formula!r}). Use a tool "
                    f"that explicitly replaces (e.g. "
                    f"replace_constant_with_distribution), or pass a "
                    f"cell that's empty or already contains a Vose formula."
                )
            self._excel.write_cell(ref.workbook, ref.sheet, ref.cell, new_formula)
            append_write_log(
                cell=f"{ref.workbook}!{ref.sheet}!{ref.cell}",
                before_formula=previous_formula,
                before_value=previous_value,
                after_formula=new_formula,
                log_path=self._settings.writes_log_path,
            )
        return InsertResult(
            cell=ref,
            formula=new_formula,
            written=True,
            previous_formula=previous_formula or None,
        )

    def restore_cell(
        self,
        ref: CellRef,
        *,
        since: datetime | None = None,
    ) -> InsertResult:
        """Find the oldest matching entry in writes.log (optionally
        filtered to entries newer than `since`) and restore that cell's
        pre-write formula. Returns an InsertResult describing what was
        written back."""
        records = _read_audit_log(self._settings.writes_log_path)
        target = f"{ref.workbook}!{ref.sheet}!{ref.cell}"
        matches = [
            r for r in records
            if r.get("cell") == target
            and (since is None or _parse_ts(r.get("ts", "")) >= since)
        ]
        if not matches:
            raise CellReferenceError(
                f"No audit-log entry found for {ref.a1!r}"
                + (f" since {since.isoformat()}" if since else "")
                + "."
            )
        # Oldest first — we want to roll back to the state before the
        # earliest write in the window.
        matches.sort(key=lambda r: r.get("ts", ""))
        pre_formula = matches[0].get("before_formula") or ""
        with self._writer_mutex.held(timeout_ms=0):
            current = self._excel.get_cell(ref.workbook, ref.sheet, ref.cell)
            self._excel.write_cell(
                ref.workbook, ref.sheet, ref.cell, pre_formula
            )
            append_write_log(
                cell=f"{ref.workbook}!{ref.sheet}!{ref.cell}",
                before_formula=current.formula or "",
                before_value=current.value,
                after_formula=pre_formula,
                log_path=self._settings.writes_log_path,
            )
        return InsertResult(
            cell=ref,
            formula=pre_formula,
            written=True,
            previous_formula=current.formula or None,
        )

    def find_hard_coded_inputs(self, workbook: str) -> list[CellRef]:
        """Heuristic: numeric cells that are referenced by at least one
        formula cell elsewhere in the workbook. v0.1 version returns the
        list of candidate cells; ranking by usage count is left for a
        future tool refinement."""
        numeric_cells: dict[str, CellInfo] = {}
        formula_cells: list[CellInfo] = []
        for cell in self._excel.iterate_cells(workbook):
            if cell.formula:
                formula_cells.append(cell)
            elif isinstance(cell.value, (int, float)) and not isinstance(
                cell.value, bool
            ):
                key = f"{cell.ref.sheet}!{cell.ref.cell}"
                numeric_cells[key] = cell
        if not numeric_cells:
            return []
        # Build a set of cell tokens referenced by any formula. Excel cell
        # references inside formulas look like A1, $A$1, Sheet!A1.
        referenced: set[str] = set()
        ref_token_re = re.compile(
            r"(?:(?P<sheet>[A-Za-z_][\w]*)!)?\$?(?P<col>[A-Z]+)\$?(?P<row>\d+)"
        )
        for fc in formula_cells:
            for m in ref_token_re.finditer(fc.formula):
                sheet = m.group("sheet") or fc.ref.sheet
                token = f"{sheet}!{m.group('col')}{m.group('row')}"
                referenced.add(token)
        return [
            cell.ref for key, cell in numeric_cells.items() if key in referenced
        ]

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _first_distribution_head(self, formula: str) -> str | None:
        """Return the first call-expression head that is a Vose
        distribution-shaped function (one of the categories in
        `_DISTRIBUTION_CATEGORIES`). Wrapper calls (`VoseInput`,
        `VoseOutput`) are intentionally skipped."""
        for head in extract_call_heads(formula):
            spec = self._catalogue.get(head)
            if spec is None:
                continue
            if spec.category in _DISTRIBUTION_CATEGORIES:
                return head
        return None

    @staticmethod
    def _extract_top_level_args(formula: str, function_name: str) -> list[str]:
        """Best-effort: return the comma-separated argument strings of
        the first call to `function_name` in `formula`. Respects nested
        parens and string literals."""
        idx = formula.find(function_name + "(")
        if idx < 0:
            return []
        i = idx + len(function_name) + 1
        depth = 1
        in_string = False
        args: list[str] = []
        buf: list[str] = []
        while i < len(formula) and depth > 0:
            ch = formula[i]
            if ch == '"':
                # Excel escapes embedded quotes by doubling — "" inside a string.
                if in_string and i + 1 < len(formula) and formula[i + 1] == '"':
                    buf.append('""')
                    i += 2
                    continue
                in_string = not in_string
                buf.append(ch)
                i += 1
                continue
            if not in_string:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0:
                        break
                elif ch == "," and depth == 1:
                    args.append("".join(buf).strip())
                    buf = []
                    i += 1
                    continue
            buf.append(ch)
            i += 1
        if buf:
            args.append("".join(buf).strip())
        return args


def _unescape_excel_string(raw: str) -> str:
    return raw.replace('""', '"')


def _coerce_displayable(value: Any) -> float | str | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        return float(value)
    return str(value)


def _read_audit_log(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    records: list[dict[str, Any]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            records.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return records


def _parse_ts(ts: str) -> datetime:
    try:
        # fromisoformat handles "2026-05-20T10:30:00+00:00" since 3.11.
        return datetime.fromisoformat(ts)
    except (TypeError, ValueError):
        return datetime.min


def _lookup_modelrisk_inproc_server() -> tuple[str | None, str | None]:
    """Read HKCR\\CLSID\\{clsid}\\InprocServer32 for ModelRisk's main
    coclass. Returns (clsid, dll_path) — either may be None if the
    registry entry isn't there or can't be read."""
    for clsid in (CLSID_DISTRIBUTIONS, CLSID_SIMULATION):
        for hive in (winreg.HKEY_CLASSES_ROOT, winreg.HKEY_CURRENT_USER):
            sub = (
                f"CLSID\\{clsid}\\InprocServer32"
                if hive == winreg.HKEY_CLASSES_ROOT
                else f"Software\\Classes\\CLSID\\{clsid}\\InprocServer32"
            )
            try:
                with winreg.OpenKey(hive, sub) as key:
                    dll, _ = winreg.QueryValueEx(key, "")
                    return clsid, str(dll)
            except OSError:
                continue
    return None, None


def _pe_bits(dll_path: str) -> int | None:
    """Read a Windows PE header and return 32 or 64. Best-effort —
    returns None if the file isn't a recognisable PE."""
    try:
        with open(dll_path, "rb") as f:
            # DOS header → e_lfanew at offset 0x3C points to the PE header.
            f.seek(0x3C)
            pe_offset_bytes = f.read(4)
            if len(pe_offset_bytes) < 4:
                return None
            pe_offset = struct.unpack("<I", pe_offset_bytes)[0]
            f.seek(pe_offset)
            sig = f.read(4)
            if sig != b"PE\x00\x00":
                return None
            # IMAGE_FILE_HEADER.Machine at PE+4 (16-bit little-endian).
            machine_bytes = f.read(2)
            if len(machine_bytes) < 2:
                return None
            machine = struct.unpack("<H", machine_bytes)[0]
            if machine == 0x014C:  # IMAGE_FILE_MACHINE_I386
                return 32
            if machine in (0x8664, 0xAA64):  # AMD64 or ARM64
                return 64
    except OSError:
        return None
    return None


__all__ = ["ModelRiskBridge"]
