"""ModelRiskBridge — domain layer on top of ExcelBridge.

Phase 1 skeleton: implements the read-only inspection methods needed by
the §7.1 reading tools (`list_modelrisk_inputs`, `list_modelrisk_outputs`,
`list_distributions`, `find_hard_coded_inputs`). Building / simulation /
results methods land in later phases.

Nothing here touches COM directly — all Excel access goes through
`ExcelBridge`. Catalogue access goes through `FunctionCatalogue`.
"""

from __future__ import annotations

import re
from typing import Any

from modelrisk_mcp.bridge.catalogue import FunctionCatalogue, load_catalogue
from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.bridge.progids import PROGID_DISTRIBUTIONS
from modelrisk_mcp.bridge.results import ResultsReader
from modelrisk_mcp.safety import extract_call_heads
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
    ) -> None:
        self._excel = excel
        self._catalogue = catalogue or load_catalogue()
        self._results = results or ResultsReader()

    @property
    def catalogue(self) -> FunctionCatalogue:
        return self._catalogue

    @property
    def excel(self) -> ExcelBridge:
        return self._excel

    @property
    def results(self) -> ResultsReader:
        return self._results

    # ------------------------------------------------------------------
    # Environment checks
    # ------------------------------------------------------------------

    def is_modelrisk_loaded(self) -> bool:
        """Best-effort: returns True if the ModelRisk COM object can be
        Dispatched on this machine. Doesn't guarantee the Excel add-in is
        currently loaded — but if it's installed at all, Dispatch works."""
        try:
            import win32com.client as com
        except ImportError:
            return False
        try:
            obj = com.Dispatch(PROGID_DISTRIBUTIONS)
        except Exception:
            return False
        return obj is not None

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


__all__ = ["ModelRiskBridge"]
