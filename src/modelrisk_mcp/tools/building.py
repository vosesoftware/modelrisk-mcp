"""Building tools (spec §7.2) — every tool defaults `dry_run=True`.

Each tool builds a Vose formula via `FormulaBuilder`, validates against
the catalogue (UnknownFunctionError + close-match suggestion on bad
function names), and either:

- returns the formula it would write (dry_run=True, the default)
- or calls `ModelRiskBridge.safe_write_cell` (dry_run=False), which
  applies the §11 safety mechanisms: writer-mutex acquisition,
  non-Vose-formula refusal, audit-log append.

Time-series and copula tools that write contiguous ranges set
`exempt=True` on the bulk-write guard (spec §11.3).
"""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import Field

from modelrisk_mcp.bridge.formulas import (
    build_aggregate_mc,
    build_copula,
    build_distribution_formula,
    build_input_wrapper,
    build_output_wrapper,
    build_risk_event,
    build_time_series,
)
from modelrisk_mcp.errors import ParameterMismatchError, UnknownFunctionError
from modelrisk_mcp.safety import check_bulk_write
from modelrisk_mcp.schemas.distributions import InsertResult, WrapResult
from modelrisk_mcp.schemas.workbook import CellRef
from modelrisk_mcp.server import mcp
from modelrisk_mcp.tools.reading import get_bridge


def _dry_or_write(
    ref: CellRef,
    formula: str,
    dry_run: bool,
    *,
    allow_overwrite_non_vose: bool = False,
) -> InsertResult:
    if dry_run:
        return InsertResult(cell=ref, formula=formula, written=False)
    return get_bridge().safe_write_cell(
        ref, formula, allow_overwrite_non_vose=allow_overwrite_non_vose
    )


def _params_dict(parameters: list[dict[str, Any]] | dict[str, Any]) -> dict[str, Any]:
    """Accept either a dict {name: value} or a list of {name, value}
    dicts (Pydantic models often serialise to the latter)."""
    if isinstance(parameters, dict):
        return parameters
    out: dict[str, Any] = {}
    for p in parameters:
        if "name" not in p:
            raise ParameterMismatchError(
                f"Parameter entry missing 'name': {p!r}."
            )
        out[p["name"]] = p.get("value")
    return out


# ----------------------------------------------------------------------
# §7.2 tools
# ----------------------------------------------------------------------


@mcp.tool(
    description=(
        "ModelRisk: Write a Vose distribution formula (e.g. =VoseModPERT(...)) "
        "into a cell. Validates the function name against the ModelRisk "
        "catalogue. Defaults to dry_run=True — Claude must explicitly pass "
        "dry_run=False to commit."
    )
)
def insert_distribution(
    workbook: str,
    sheet: str,
    cell: Annotated[str, Field(description="A1 cell reference like 'B12'.")],
    function_name: str,
    parameters: list[dict[str, Any]],
    dry_run: bool = True,
) -> InsertResult:
    ref = CellRef(workbook=workbook, sheet=sheet, cell=cell)
    bridge = get_bridge()
    formula = build_distribution_formula(
        function_name, _params_dict(parameters), bridge.catalogue
    )
    return _dry_or_write(ref, formula, dry_run)


@mcp.tool(
    description=(
        "ModelRisk: Wrap an existing distribution cell with "
        "VoseInput(\"name\")+ so it appears in the input list and "
        "the Results Viewer."
    )
)
def wrap_with_input(
    workbook: str,
    sheet: str,
    cell: str,
    name: str,
    dry_run: bool = True,
) -> WrapResult:
    ref = CellRef(workbook=workbook, sheet=sheet, cell=cell)
    bridge = get_bridge()
    current = bridge.excel.get_cell(workbook, sheet, cell)
    inner = _strip_existing_input_wrapper(current.formula)
    formula = build_input_wrapper(name, inner)
    if dry_run:
        return WrapResult(cell=ref, formula=formula, written=False)
    result = bridge.safe_write_cell(ref, formula, allow_overwrite_non_vose=False)
    return WrapResult(
        cell=ref,
        formula=result.formula,
        written=True,
        previous_formula=result.previous_formula,
    )


@mcp.tool(
    description=(
        "ModelRisk: Wrap an existing output cell with VoseOutput(\"name\")+ "
        "so it appears in the output list and Results Viewer."
    )
)
def wrap_with_output(
    workbook: str,
    sheet: str,
    cell: str,
    name: str,
    dry_run: bool = True,
) -> WrapResult:
    ref = CellRef(workbook=workbook, sheet=sheet, cell=cell)
    bridge = get_bridge()
    current = bridge.excel.get_cell(workbook, sheet, cell)
    inner = _strip_existing_output_wrapper(current.formula)
    formula = build_output_wrapper(name, inner)
    if dry_run:
        return WrapResult(cell=ref, formula=formula, written=False)
    result = bridge.safe_write_cell(ref, formula, allow_overwrite_non_vose=False)
    return WrapResult(
        cell=ref,
        formula=result.formula,
        written=True,
        previous_formula=result.previous_formula,
    )


@mcp.tool(
    description=(
        "ModelRisk: Replace a hard-coded number in a cell with a Vose "
        "distribution wrapped by VoseInput. Use after find_hard_coded_inputs "
        "identifies candidates. This is the only tool that overwrites a "
        "non-Vose cell — it does so by design."
    )
)
def replace_constant_with_distribution(
    workbook: str,
    sheet: str,
    cell: str,
    function_name: str,
    parameters: list[dict[str, Any]],
    input_name: str,
    dry_run: bool = True,
) -> InsertResult:
    ref = CellRef(workbook=workbook, sheet=sheet, cell=cell)
    bridge = get_bridge()
    inner = build_distribution_formula(
        function_name, _params_dict(parameters), bridge.catalogue
    )
    formula = build_input_wrapper(input_name, inner)
    return _dry_or_write(ref, formula, dry_run, allow_overwrite_non_vose=True)


@mcp.tool(
    description=(
        "ModelRisk: Fit a distribution family to a data range and write the "
        "result. 'family' is e.g. 'Normal', 'Lognormal', 'Beta', 'Gamma', "
        "'Weibull' — the tool maps it to the catalogue function (Vose<Family>"
        "Fit). Set uncertainty=True (the default) to include parameter "
        "uncertainty in the fitted distribution."
    )
)
def fit_distribution_to_data(
    workbook: str,
    sheet: str,
    target_cell: str,
    data_range: str,
    family: Annotated[str, Field(description="Distribution family, e.g. 'Normal'.")],
    uncertainty: bool = True,
    dry_run: bool = True,
) -> InsertResult:
    ref = CellRef(workbook=workbook, sheet=sheet, cell=target_cell)
    bridge = get_bridge()
    fit_function = f"Vose{family}Fit"
    if fit_function not in bridge.catalogue:
        suggestions = bridge.catalogue.suggest(fit_function)
        hint = (
            f" Did you mean: {', '.join(suggestions)}?" if suggestions else ""
        )
        raise UnknownFunctionError(
            f"No fitting function {fit_function!r} in the ModelRisk catalogue."
            + hint
        )
    formula = build_distribution_formula(
        fit_function,
        {"data": data_range, "uncertainty": uncertainty},
        bridge.catalogue,
    )
    return _dry_or_write(ref, formula, dry_run)


@mcp.tool(
    description=(
        "ModelRisk: Build a VoseAggregateMC(n, distribution, ...) formula "
        "that simulates the aggregate of a (possibly random) number n of "
        "i.i.d. severity draws. n and distribution are cell references — "
        "n points at a frequency cell, distribution at a severity object "
        "cell (built with a Vose<Family>Object function)."
    )
)
def create_aggregate_mc(
    workbook: str,
    sheet: str,
    target_cell: str,
    frequency_object_cell: str,
    severity_object_cell: str,
    min_limit: float | None = None,
    max_limit: float | None = None,
    distribution_shift: float | None = None,
    dry_run: bool = True,
) -> InsertResult:
    ref = CellRef(workbook=workbook, sheet=sheet, cell=target_cell)
    bridge = get_bridge()
    formula = build_aggregate_mc(
        frequency_object_cell,
        severity_object_cell,
        bridge.catalogue,
        min_limit=min_limit,
        max_limit=max_limit,
        distribution_shift=distribution_shift,
    )
    return _dry_or_write(ref, formula, dry_run)


@mcp.tool(
    description=(
        "ModelRisk: Build a VoseRiskEvent(probability, impact_object) "
        "formula. The impact_function_name must be a distribution-object "
        "function (category 'object') — typically a Vose<Family>Object "
        "variant."
    )
)
def create_risk_event(
    workbook: str,
    sheet: str,
    target_cell: str,
    probability: float | str,
    impact_function_name: str,
    impact_parameters: list[dict[str, Any]],
    dry_run: bool = True,
) -> InsertResult:
    ref = CellRef(workbook=workbook, sheet=sheet, cell=target_cell)
    bridge = get_bridge()
    formula = build_risk_event(
        probability,
        impact_function_name,
        _params_dict(impact_parameters),
        bridge.catalogue,
    )
    return _dry_or_write(ref, formula, dry_run)


@mcp.tool(
    description=(
        "ModelRisk: Build a time-series formula (VoseTimeGBM, VoseTimeAR1, "
        "etc.) and write it into the target_range. Time-series tools are "
        "exempt from the >50-cell bulk-write guard because the dimension "
        "of the time series is the whole point."
    )
)
def create_time_series(
    workbook: str,
    sheet: str,
    target_range: Annotated[str, Field(description="A1 range to spill into, e.g. 'D2:D101'.")],
    function_name: str,
    parameters: list[dict[str, Any]],
    dry_run: bool = True,
) -> InsertResult:
    bridge = get_bridge()
    formula = build_time_series(
        function_name, _params_dict(parameters), bridge.catalogue
    )
    cell_count = _approximate_range_size(target_range)
    check_bulk_write(cell_count, exempt=True)
    first_cell = target_range.split(":", 1)[0]
    ref = CellRef(workbook=workbook, sheet=sheet, cell=first_cell)
    if dry_run:
        return InsertResult(cell=ref, formula=formula, written=False)
    # Time-series formulas spill via Excel 365 dynamic arrays; we write
    # the formula into the first cell and let Excel handle the spill.
    return bridge.safe_write_cell(ref, formula)


@mcp.tool(
    description=(
        "ModelRisk: Build a copula formula (e.g. VoseCopulaMultiNormal) "
        "that produces a correlated u-array for downstream distribution "
        "calls. Copula tools are exempt from the bulk-write guard."
    )
)
def create_copula(
    workbook: str,
    sheet: str,
    u_array_target_range: str,
    function_name: str,
    parameters: list[dict[str, Any]],
    dry_run: bool = True,
) -> InsertResult:
    bridge = get_bridge()
    formula = build_copula(
        function_name, _params_dict(parameters), bridge.catalogue
    )
    cell_count = _approximate_range_size(u_array_target_range)
    check_bulk_write(cell_count, exempt=True)
    first_cell = u_array_target_range.split(":", 1)[0]
    ref = CellRef(workbook=workbook, sheet=sheet, cell=first_cell)
    if dry_run:
        return InsertResult(cell=ref, formula=formula, written=False)
    return bridge.safe_write_cell(ref, formula)


@mcp.tool(
    description=(
        "ModelRisk: Create or overwrite a workbook-level named range. "
        "Useful for giving cells clear identities the LLM can reference "
        "by name later. The reference must be A1-style (e.g. 'Sheet1!"
        "$A$1:$A$10')."
    )
)
def set_named_range(
    workbook: str,
    name: str,
    range_ref: str,
    dry_run: bool = True,
) -> dict[str, str]:
    if dry_run:
        return {
            "workbook": workbook,
            "name": name,
            "range_ref": range_ref,
            "written": "false",
        }
    bridge = get_bridge()
    bridge.excel.set_named_range(workbook, name, range_ref)
    return {
        "workbook": workbook,
        "name": name,
        "range_ref": range_ref,
        "written": "true",
    }


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------


def _strip_existing_input_wrapper(formula: str) -> str:
    return _strip_wrapper(formula, "VoseInput")


def _strip_existing_output_wrapper(formula: str) -> str:
    return _strip_wrapper(formula, "VoseOutput")


def _strip_wrapper(formula: str, wrapper_name: str) -> str:
    """Remove a `VoseInput("name")+` or `VoseOutput("name")+` prefix from
    a formula, returning the inner formula. If no wrapper is present,
    returns the original (sans leading '=')."""
    body = formula.lstrip()
    if body.startswith("="):
        body = body[1:].lstrip()
    if not body.startswith(wrapper_name + "("):
        return body
    depth = 0
    i = len(wrapper_name) + 1
    in_string = False
    while i < len(body):
        ch = body[i]
        if ch == '"':
            if in_string and i + 1 < len(body) and body[i + 1] == '"':
                i += 2
                continue
            in_string = not in_string
        elif not in_string:
            if ch == "(":
                depth += 1
            elif ch == ")":
                if depth == 0:
                    i += 1
                    break
                depth -= 1
        i += 1
    rest = body[i:].lstrip()
    if rest.startswith("+"):
        rest = rest[1:].lstrip()
    return rest


def _approximate_range_size(range_ref: str) -> int:
    """Cheap estimate of how many cells a range covers, used by the
    bulk-write guard. Doesn't require Excel — parses A1 notation."""
    if ":" not in range_ref:
        return 1
    start, end = range_ref.split(":", 1)
    sc, sr = _split_a1(start)
    ec, er = _split_a1(end)
    return max(1, (ec - sc + 1) * (er - sr + 1))


def _split_a1(ref: str) -> tuple[int, int]:
    col_letters = "".join(c for c in ref if c.isalpha()).upper()
    row_digits = "".join(c for c in ref if c.isdigit())
    col = 0
    for c in col_letters:
        col = col * 26 + (ord(c) - ord("A") + 1)
    row = int(row_digits) if row_digits else 1
    return col, row
