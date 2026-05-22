"""ModelRiskBridge — domain layer on top of `ExcelBridge` + `MrServiceBridge`.

Architecture (v0.3, post-pivot):
- Reading + building tools use xlwings (via `ExcelBridge`).
- Simulation results come from `.vmrs` files, opened by `MrServiceBridge`
  (via the SDK's `MRService.dll`). No COM, no ATL, no VBA helper.
- Excel itself runs the simulation (the user presses Start or triggers
  via a cell formula); we just read the `.vmrs` output.

Everything in this file is pure Python — the actual ctypes / xlwings
calls live in the dedicated bridge modules.
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from modelrisk_mcp.bridge.catalogue import FunctionCatalogue, load_catalogue
from modelrisk_mcp.bridge.charts import TornadoChartResult, TornadoChartWriter
from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.bridge.mrservice import MrServiceBridge
from modelrisk_mcp.bridge.name_parser import (
    ExpressionName,
    LiteralName,
    extract_vose_first_arg,
)
from modelrisk_mcp.bridge.reports import (
    DriversReportBuilder,
    DriversReportResult,
    ExecutiveReportBuilder,
    ExecutiveReportResult,
    default_subtitle,
)
from modelrisk_mcp.bridge.results import ResultsReader
from modelrisk_mcp.bridge.simulation import (
    SimulationController,
    SimulationRunResult,
)
from modelrisk_mcp.config import Settings
from modelrisk_mcp.errors import CellReferenceError, SimulationFailedError
from modelrisk_mcp.safety import (
    WriterMutex,
    append_write_log,
    extract_call_heads,
    is_vose_formula,
)
from modelrisk_mcp.schemas.distributions import InsertResult
from modelrisk_mcp.schemas.results import (
    CorrelationMatrix,
    ScenarioOutcome,
    ScenarioRun,
    ScenarioSweepResult,
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
        mrservice: MrServiceBridge | None = None,
        settings: Settings | None = None,
        writer_mutex: WriterMutex | None = None,
        simulation: SimulationController | None = None,
    ) -> None:
        self._excel = excel
        self._catalogue = catalogue or load_catalogue()
        self._mrservice = mrservice or MrServiceBridge()
        self._results = results or ResultsReader(self._mrservice)
        self._settings = settings or Settings()
        self._writer_mutex = writer_mutex or WriterMutex()
        self._simulation = simulation or SimulationController(excel)

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
    def mrservice(self) -> MrServiceBridge:
        return self._mrservice

    @property
    def simulation(self) -> SimulationController:
        return self._simulation

    # ------------------------------------------------------------------
    # Run sim — wraps SimulationController + auto-pins the resulting
    # .vmrs as the active source so subsequent get_simulation_results
    # finds it without further setup.
    # ------------------------------------------------------------------

    def run_simulation(
        self,
        workbook: str | None = None,
        *,
        samples: int = 1000,
        seed: int = 1,
        save_to: str | None = None,
    ) -> SimulationRunResult:
        # Bug #20 fix: capture the list of VoseOutput names BEFORE the
        # simulation so we can verify at least one ended up registered
        # in the .vmrs afterwards. Without this, run_simulation
        # reports success on simulations where ModelRisk's
        # post-simulation phase quietly failed — leaving the .vmrs
        # missing output metadata and every downstream reader unable
        # to find anything.
        resolved_workbook = (
            workbook if workbook else self._excel.get_active_workbook().name
        )
        expected_output_names: list[str] = []
        try:
            expected_output_names = [
                o.name for o in self.list_outputs(resolved_workbook)
            ]
        except Exception:
            # Scanner failure shouldn't abort the sim — but we lose
            # the ability to verify outputs. The .vmrs-existence
            # check downstream is still in force.
            pass

        # alpha.18 experiment: pre-populate output_names into the XLL
        # command's options payload. The C++ header comment says
        # "empty → all outputs" but real-world testing showed sims
        # completing without registering any outputs in the .vmrs
        # unless the ribbon path was used. Working hypothesis: the
        # ribbon populates this list and the XLL command actually
        # requires it. Re-using `expected_output_names` (already
        # gathered for post-condition verification above) is free.
        result = self._simulation.run_simulation(
            workbook_name=workbook,
            samples=samples,
            seed=seed,
            save_to=save_to,
            output_names=tuple(expected_output_names),
        )
        # Pin the produced file so the existing reader tools find it.
        self._results.set_active_vmrs(result.vmrs_path)

        # Post-condition verification. Only check names we can
        # statically resolve — bug #32 (alpha.31): ExpressionName
        # outputs like `VoseOutput("prefix "&B8&" suffix")` have
        # their actual runtime name computed by Excel at simulation
        # time, so the scanner's partial prefix can't be looked up
        # in the .vmrs. Including those in the check would
        # false-positively claim every such workbook failed.
        verifiable_names = [
            n for n in expected_output_names
            if n and not n.endswith("…")  # "…" marker for dynamic
            and n != "<dynamic name>"
        ]
        if verifiable_names:
            try:
                self._verify_simulation_post_conditions(
                    result.vmrs_path, verifiable_names,
                )
            except SimulationFailedError:
                try:
                    self._excel.recalculate_workbook(resolved_workbook)
                except Exception:
                    pass
                raise
        return result

    def _verify_simulation_post_conditions(
        self, vmrs_path: str, expected_output_names: list[str],
    ) -> None:
        """Open the produced .vmrs and confirm at least one of the
        workbook's VoseOutput names resolves to a var_id. That's the
        cheapest available proof that ModelRisk's post-simulation
        phase actually wrote output metadata (vs. the bug #20
        symptom: sim completes, .vmrs is created, but no outputs are
        registered)."""
        try:
            with self._mrservice.open_vmrs(vmrs_path) as handle:
                for name in expected_output_names:
                    try:
                        if handle.lookup_var_id(name) is not None:
                            return  # post-condition satisfied
                    except SimulationFailedError:
                        # lookup_var_id can raise on pathological names;
                        # one bad name doesn't condemn the whole sim.
                        continue
        except SimulationFailedError:
            raise
        except Exception as exc:
            raise SimulationFailedError(
                f"Simulation completed but the produced .vmrs at "
                f"{vmrs_path!r} could not be opened for verification: "
                f"{exc!s}. The file may be incomplete or corrupt."
            ) from exc
        raise SimulationFailedError(
            f"Simulation completed but the produced .vmrs at "
            f"{vmrs_path!r} does not register any of the expected "
            f"VoseOutput names: {expected_output_names!r}. This is the "
            f"signature of ModelRisk's post-simulation phase crashing "
            f"silently (the sim ran but the output metadata didn't "
            f"persist). The workbook may also be left with VoseOutput "
            f"cells stuck on sample values rather than deterministic "
            f"ones — run `restore_deterministic_state` to recover, then "
            f"retry the simulation."
        )

    def restore_deterministic_state(
        self, workbook: str | None = None,
    ) -> dict[str, Any]:
        """Recalculate the workbook to clear any VoseOutput cells that
        are stuck on sample values from a previous run.

        Recovery path for bug #21: when `run_simulation` fails post-
        condition (#20), the workbook's VoseOutput cells often retain
        the final per-iteration sample instead of the deterministic
        value. Subsequent reads (`list_modelrisk_outputs`,
        `find_hard_coded_inputs`) would then return misleading
        `current_value`s. A full recalculation restores the
        deterministic state by re-evaluating every formula."""
        resolved_workbook = (
            workbook if workbook else self._excel.get_active_workbook().name
        )
        self._excel.recalculate_workbook(resolved_workbook)
        return {
            "workbook_name": resolved_workbook,
            "recalculated": True,
        }

    # ------------------------------------------------------------------
    # Environment checks
    # ------------------------------------------------------------------

    def is_modelrisk_loaded(self) -> bool:
        """True if MRService.dll can be loaded and activated. The DLL is
        what we use to read `.vmrs` files; if this returns False, the
        results-reading tools won't work."""
        try:
            self._mrservice.ensure_ready()
        except Exception:
            return False
        return True

    # ------------------------------------------------------------------
    # Reading tools (§7.1)
    # ------------------------------------------------------------------

    def list_inputs(self, workbook: str) -> list[ModelRiskInput]:
        result: list[ModelRiskInput] = []
        for cell in self._excel.iterate_cells(workbook):
            if not cell.formula:
                continue
            name = self._resolve_vose_name(
                cell.formula, "VoseInput", workbook, cell.ref.sheet,
            )
            if name is None:
                continue
            result.append(
                ModelRiskInput(
                    ref=cell.ref,
                    name=name,
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
            name = self._resolve_vose_name(
                cell.formula, "VoseOutput", workbook, cell.ref.sheet,
            )
            if name is None:
                continue
            result.append(
                ModelRiskOutput(
                    ref=cell.ref,
                    name=name,
                    formula=cell.formula,
                    current_value=_coerce_displayable(cell.value),
                )
            )
        return result

    def _resolve_vose_name(
        self,
        formula: str,
        wrapper: str,
        workbook: str,
        same_sheet: str,
    ) -> str | None:
        """Extract and resolve a Vose wrapper's name argument.

        Handles all three forms:
          - `VoseInput("WidgetCost")`         → "WidgetCost"
          - `VoseInput(A5)`                   → reads A5
          - `VoseInput("prefix "&B8)`         → returns the static
            prefix (post-alpha.31 ExpressionName). The runtime-
            computed name is unknown but at least the user sees
            "Total net revenue from " in `list_outputs` instead of
            the cell disappearing entirely.

        Returns None if the wrapper isn't present, the form is
        unrecognized, or the referenced cell is empty / unreadable.
        """
        arg = extract_vose_first_arg(formula, wrapper)
        if arg is None:
            return None
        if isinstance(arg, LiteralName):
            return _unescape_excel_string(arg.name)
        if isinstance(arg, ExpressionName):
            # Show the partial prefix so the user knows the cell exists
            # and roughly what it's called. Marker syntax `prefix…`
            # makes it obvious this is a dynamic name, not a literal.
            prefix = _unescape_excel_string(arg.static_prefix).rstrip()
            return f"{prefix}…" if prefix else "<dynamic name>"
        # CellRefName — read the target cell to get the actual name.
        target_sheet = arg.sheet or same_sheet
        try:
            target = self._excel.get_cell(workbook, target_sheet, arg.cell)
        except Exception:
            return None
        value = target.value
        if value is None:
            return None
        text = str(value).strip()
        return text or None

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
            # Wrapper-presence flags now cover both literal AND cell-ref
            # name forms — we don't need to resolve the name here, just
            # know whether the wrapper exists in the formula.
            has_input = (
                extract_vose_first_arg(cell.formula, "VoseInput") is not None
            )
            has_output = (
                extract_vose_first_arg(cell.formula, "VoseOutput") is not None
            )
            result.append(
                DistributionCell(
                    ref=cell.ref,
                    function_name=dist_head,
                    parameters=self._extract_top_level_args(cell.formula, dist_head),
                    has_input_wrapper=has_input,
                    has_output_wrapper=has_output,
                    formula=cell.formula,
                )
            )
        return result

    def get_workbook_summary(self, workbook: str) -> WorkbookSummary:
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
            # Bug #27: only count cells whose .Formula starts with `=`
            # as actual formulas. xlwings returns the text content in
            # .Formula for text cells, which would otherwise count as
            # formulas (and shadow the numeric_cell_count for the
            # cells that hold numbers).
            is_real_formula = (
                bool(cell.formula) and cell.formula.lstrip().startswith("=")
            )
            if is_real_formula:
                formula_cell_count += 1
                # Use the parser so cell-ref name forms count too.
                if extract_vose_first_arg(cell.formula, "VoseInput") is not None:
                    input_count += 1
                if extract_vose_first_arg(cell.formula, "VoseOutput") is not None:
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

    def find_hard_coded_inputs(self, workbook: str) -> list[CellRef]:
        """Locate numeric cells referenced by at least one formula.

        Bug #27 (alpha.25): xlwings returns the cell's text content
        in `.Formula` even for non-formula cells — so a cell holding
        the label "Total Revenue" comes back as formula="Total
        Revenue", which the previous `if cell.formula:` check
        misclassified as a formula. On models with text labels every
        cell got bucketed as formula, no numeric inputs were
        candidates, and the function returned [] — silently broken
        on exactly the workbooks where it's most useful (the
        "convert this Excel model" workflow). Fix: a cell counts as
        a formula only when its `.Formula` starts with `=`.
        """
        numeric_cells: dict[str, CellInfo] = {}
        formula_cells: list[CellInfo] = []
        for cell in self._excel.iterate_cells(workbook):
            if cell.formula and cell.formula.lstrip().startswith("="):
                formula_cells.append(cell)
            elif isinstance(cell.value, (int, float)) and not isinstance(
                cell.value, bool
            ):
                key = f"{cell.ref.sheet}!{cell.ref.cell}"
                numeric_cells[key] = cell
        if not numeric_cells:
            return []
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
    # Results (read .vmrs via MRService.dll)
    # ------------------------------------------------------------------

    def get_simulation_results(
        self,
        workbook: str | None = None,
        output_names: list[str] | None = None,
    ) -> list[SimulationResult]:
        wb_path, names = self._resolve_workbook_and_outputs(
            workbook, output_names
        )
        return self._results.get_simulation_results(wb_path, names)

    def get_correlation_matrix(
        self,
        workbook: str | None = None,
        names: list[str] | None = None,
    ) -> CorrelationMatrix:
        wb_path, resolved_names = self._resolve_workbook_and_outputs(
            workbook, names, include_inputs=True
        )
        return self._results.get_correlation_matrix(wb_path, resolved_names)

    def build_drivers_report(
        self,
        output_name: str,
        *,
        workbook: str | None = None,
        title: str | None = None,
        subtitle: str | None = None,
        sheet_name: str = "Drivers_Report",
    ) -> DriversReportResult:
        """Render a single-sheet sensitivity report that explains the
        tornado in plain English: key findings, driver ranking table,
        a 'how to read this chart' panel, and tiered recommendations.
        """
        wb_name = workbook or self._excel.get_active_workbook().name
        wb_path, _ = self._resolve_workbook_and_outputs(wb_name, [output_name])

        # Fetch sensitivity + iteration count.
        input_names = [i.name for i in self.list_inputs(wb_name)]
        sensitivity = self._results.get_sensitivity_ranking(
            output_name, input_names, wb_path,
        )

        if not self._excel.is_connected():
            self._excel.connect()
        app = self._excel._app
        assert app is not None
        book = app.books[wb_name]
        return DriversReportBuilder.build(
            book,
            sheet_name=sheet_name,
            output_name=output_name,
            sensitivity=sensitivity,
            iterations=sensitivity.iterations,
            title=title,
            subtitle=subtitle,
        )

    def build_executive_report(
        self,
        primary_output: str,
        *,
        workbook: str | None = None,
        title: str | None = None,
        subtitle: str | None = None,
        secondary_outputs: list[str] | None = None,
        contingency_percentile: float = 0.90,
        top_drivers: int = 5,
        sheet_name: str = "Executive_Report",
    ) -> ExecutiveReportResult:
        """Render a single-sheet decision-maker dashboard for one
        primary output (with optional secondary outputs in the stats
        table). Gathers all the data via existing read methods, then
        hands off to `ExecutiveReportBuilder.build`."""
        wb_name = workbook or self._excel.get_active_workbook().name

        # Gather: stats for primary + secondary, raw samples for the
        # histogram, sensitivity ranking for the tornado.
        all_outputs = [primary_output, *list(secondary_outputs or [])]
        wb_path, _ = self._resolve_workbook_and_outputs(wb_name, all_outputs)
        stats = self._results.get_simulation_results(wb_path, all_outputs)
        results_by_name = {r.output_name: r for r in stats}
        if primary_output not in results_by_name:
            raise SimulationFailedError(
                f"Primary output {primary_output!r} not found in the "
                "active simulation results. Run a simulation first or "
                "check the output name."
            )
        primary_result = results_by_name[primary_output]
        secondary_results = [
            results_by_name[name] for name in (secondary_outputs or [])
            if name in results_by_name
        ]

        # Raw samples for the histogram.
        primary_samples = list(
            self._results.get_samples(primary_output, wb_path, max_n=10_000)
        )

        # Sensitivity — input names from the workbook.
        input_names = [i.name for i in self.list_inputs(wb_name)]
        sensitivity = self._results.get_sensitivity_ranking(
            primary_output, input_names, wb_path,
        )

        # Default title / subtitle if not provided.
        effective_title = title or f"Simulation Report — {primary_output}"
        effective_subtitle = subtitle or default_subtitle(
            samples=primary_result.iterations,
        )

        # Render.
        if not self._excel.is_connected():
            self._excel.connect()
        app = self._excel._app
        assert app is not None
        book = app.books[wb_name]
        return ExecutiveReportBuilder.build(
            book,
            sheet_name=sheet_name,
            title=effective_title,
            subtitle=effective_subtitle,
            primary_output=primary_output,
            primary_result=primary_result,
            primary_samples=primary_samples,
            sensitivity=sensitivity,
            secondary_results=secondary_results,
            contingency_percentile=contingency_percentile,
            top_drivers=top_drivers,
        )

    def create_tornado_chart(
        self,
        output_name: str,
        workbook: str | None = None,
        *,
        sheet_name: str | None = None,
    ) -> TornadoChartResult:
        """Render a tornado chart of input sensitivity for `output_name`.

        Reuses the existing `get_sensitivity_ranking` machinery for the
        data; writes the sorted table + a native Excel BarClustered
        chart to a new sheet via `TornadoChartWriter`. Idempotent — if
        the target sheet already exists, it's replaced."""
        wb_name = workbook or self._excel.get_active_workbook().name
        ranking = self.get_sensitivity_ranking(output_name, wb_name)
        # Force a connect if we haven't yet (the chart writer needs the
        # raw xlwings book handle, not just bridge methods).
        if not self._excel.is_connected():
            self._excel.connect()
        app = self._excel._app  # bridge-internal access
        assert app is not None  # connect() guarantees this
        book = app.books[wb_name]
        return TornadoChartWriter.write(
            book, output_name, ranking.entries, sheet_name=sheet_name
        )

    def run_scenarios(
        self,
        sheet: str,
        cell: str,
        values: list[float],
        *,
        workbook: str | None = None,
        samples: int = 1000,
        seed: int = 1,
    ) -> ScenarioSweepResult:
        """Sweep a fixed input across `values`, running a full simulation
        per value. Returns aggregate stats per output per scenario.

        Workflow:
        1. Capture the cell's current formula (so we can restore it).
        2. For each value: write the literal, run a sim, pull P5/P50/P95
           + mean for every output.
        3. Restore the original formula — even if a scenario raises.

        The cell ends up holding its original content after the sweep
        regardless of outcome, so this is safe to interrupt."""
        wb_name = workbook or self._excel.get_active_workbook().name
        original = self._excel.get_cell(wb_name, sheet, cell)
        original_formula = original.formula or ""
        # If the cell holds a raw number (no formula), restoring means
        # writing the literal back; serialize via Excel's expected form.
        if not original_formula and original.value is not None:
            original_formula = f"={original.value}"

        result = ScenarioSweepResult(
            workbook_name=wb_name,
            sheet=sheet,
            cell=cell,
            original_formula=original_formula,
            samples_per_scenario=samples,
        )
        try:
            for value in values:
                # Write the override (literal value, no formula prefix).
                self._excel.write_cell(wb_name, sheet, cell, str(value))
                # Run the sim — produces .vmrs and auto-pins it.
                self.run_simulation(
                    workbook=wb_name, samples=samples, seed=seed,
                )
                # Read every output's stats from the just-produced .vmrs.
                output_names = [o.name for o in self.list_outputs(wb_name)]
                stats_list = self._results.get_simulation_results(
                    None, output_names, percentiles=(0.05, 0.50, 0.95),
                )
                outcomes = [
                    ScenarioOutcome(
                        output_name=s.output_name,
                        mean=s.mean,
                        p5=s.percentiles.get(0.05, s.min),
                        p50=s.percentiles.get(0.50, s.mean),
                        p95=s.percentiles.get(0.95, s.max),
                    )
                    for s in stats_list
                ]
                result.scenarios.append(
                    ScenarioRun(scenario_value=value, outputs=outcomes)
                )
        finally:
            # Restore — even on exception. This is non-negotiable; an
            # interrupted sweep that left the cell at a random scenario
            # value would be a nightmare for the user.
            if original_formula:
                try:
                    self._excel.write_cell(
                        wb_name, sheet, cell, original_formula
                    )
                except Exception:
                    pass
        return result

    def list_vmrs_variables(
        self,
        workbook: str | None = None,
    ) -> list[dict[str, str | int]]:
        """Enumerate VoseInput / VoseOutput names from the workbook
        that also have data in the active `.vmrs`. Each entry includes
        name, kind ('input' / 'output'), variable ID, and iteration count."""
        wb_path, _ = self._resolve_workbook_and_outputs(workbook, None)
        wb = workbook or self._excel.get_active_workbook().name
        candidates: list[tuple[str, str]] = []
        for o in self.list_outputs(wb):
            candidates.append((o.name, "output"))
        for i in self.list_inputs(wb):
            candidates.append((i.name, "input"))
        entries = self._results.list_variables(wb_path, candidates)
        return [e.to_dict() for e in entries]

    def get_samples(
        self,
        output_name: str,
        workbook: str | None = None,
        *,
        max_n: int = 10_000,
    ) -> list[float]:
        """Raw per-iteration sample array for one variable."""
        wb_path, _ = self._resolve_workbook_and_outputs(workbook, None)
        samples = self._results.get_samples(
            output_name, wb_path, max_n=max_n
        )
        return list(samples)

    def get_sensitivity_ranking(
        self,
        output_name: str,
        workbook: str | None = None,
    ) -> SensitivityRanking:
        wb_path, _ = self._resolve_workbook_and_outputs(workbook, None)
        wb = workbook or self._excel.get_active_workbook().name
        input_names = [i.name for i in self.list_inputs(wb)]
        return self._results.get_sensitivity_ranking(
            output_name, input_names, wb_path
        )

    def _resolve_workbook_and_outputs(
        self,
        workbook: str | None,
        names: list[str] | None,
        *,
        include_inputs: bool = False,
    ) -> tuple[str | None, list[str]]:
        wb_name = workbook or self._excel.get_active_workbook().name
        try:
            books = self._excel.list_workbooks()
            wb_info = next((b for b in books if b.name == wb_name), None)
            wb_path = wb_info.path if wb_info else None
        except Exception:
            wb_path = None
        resolved: list[str] = list(names) if names else []
        if not resolved:
            resolved = [o.name for o in self.list_outputs(wb_name)]
            if include_inputs:
                resolved += [i.name for i in self.list_inputs(wb_name)]
        return wb_path, resolved

    # ------------------------------------------------------------------
    # Building-tool support — every write goes through here (spec §11)
    # ------------------------------------------------------------------

    def safe_write_cell(
        self,
        ref: CellRef,
        new_formula: str,
        *,
        allow_overwrite_non_vose: bool = False,
    ) -> InsertResult:
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

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _first_distribution_head(self, formula: str) -> str | None:
        for head in extract_call_heads(formula):
            spec = self._catalogue.get(head)
            if spec is None:
                continue
            if spec.category in _DISTRIBUTION_CATEGORIES:
                return head
        return None

    @staticmethod
    def _extract_top_level_args(formula: str, function_name: str) -> list[str]:
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
        return datetime.fromisoformat(ts)
    except (TypeError, ValueError):
        return datetime.min


__all__ = ["ModelRiskBridge"]
