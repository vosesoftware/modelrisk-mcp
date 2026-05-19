"""Pure functions for building Vose formula strings.

Every formula written to Excel must pass through this module. The
contract is simple: it validates `function_name` against the catalogue
(via `FunctionCatalogue.require`, which raises `UnknownFunctionError`
with a close-match suggestion if the name is bogus) and validates
parameter shape (no unknown names, all required names present), then
renders an Excel formula string.

These functions never touch Excel — they're pure string construction
and easily unit-tested.
"""

from __future__ import annotations

import re
from typing import Any

from modelrisk_mcp.bridge.catalogue import FunctionCatalogue, FunctionSpec
from modelrisk_mcp.errors import ParameterMismatchError

# A token that looks like an A1 reference (possibly with a sheet prefix).
# Used to distinguish "B12"/"Sheet1!A1:A10" from a literal string.
_REF_RE = re.compile(
    r"^[A-Za-z_][\w\.]*!\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?$"
    r"|^\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?$"
)

# Tokens that look like Excel/Vose function calls or arithmetic — treated
# as formula fragments and emitted verbatim (no quoting). The token must
# start with a letter and contain a `(` somewhere or look like an arithmetic
# expression (contains an operator).
_FORMULA_FRAGMENT_RE = re.compile(
    r"^[A-Za-z_][\w]*\("       # FunctionName(...
    r"|^[\(\-\+]"              # leading paren / unary operator
    r"|.*[+\-*/^&].*"          # contains an Excel arithmetic operator
)


def render_value(value: Any) -> str:
    """Render a single argument value as an Excel formula fragment.

    Numbers and booleans become Excel literals. A list becomes an array
    literal like `{1,2,3}` (or `{1,2;3,4}` for 2D). A string is used
    verbatim if it looks like an A1 reference or a formula fragment;
    otherwise it's emitted as a quoted Excel string with embedded
    double-quotes escaped.
    """
    if value is True:
        return "TRUE"
    if value is False:
        return "FALSE"
    if isinstance(value, (int, float)):
        return _render_number(value)
    if isinstance(value, list):
        return _render_array(value)
    if isinstance(value, str):
        return _render_string_arg(value)
    raise ParameterMismatchError(
        f"Cannot render value of type {type(value).__name__}: {value!r}."
    )


def _render_number(value: int | float) -> str:
    # Avoid Python's scientific notation for plausible Excel inputs and
    # avoid trailing ".0" for whole numbers.
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return repr(value) if isinstance(value, float) else str(value)


def _render_array(values: list[Any]) -> str:
    if not values:
        raise ParameterMismatchError("Cannot render an empty array literal.")
    if all(isinstance(row, list) for row in values):
        # 2D — Excel uses ';' as row separator, ',' as column separator
        rows = [",".join(render_value(c) for c in row) for row in values]
        return "{" + ";".join(rows) + "}"
    if any(isinstance(c, list) for c in values):
        raise ParameterMismatchError(
            "Array values must be either fully 1D or fully 2D (a list of lists)."
        )
    return "{" + ",".join(render_value(c) for c in values) + "}"


def _render_string_arg(s: str) -> str:
    stripped = s.strip()
    if not stripped:
        return '""'
    if stripped.startswith("="):
        # Strip the leading "=" (Excel doesn't allow nested "=" inside formulas).
        stripped = stripped[1:].lstrip()
    if _REF_RE.match(stripped):
        return stripped
    if _FORMULA_FRAGMENT_RE.match(stripped):
        return stripped
    # Plain string literal — quote and double up embedded quotes.
    return '"' + stripped.replace('"', '""') + '"'


def _validate_param_names(
    spec: FunctionSpec, supplied: dict[str, Any]
) -> None:
    known = {p.name: p for p in spec.parameters}
    unknown = sorted(set(supplied) - set(known))
    if unknown:
        raise ParameterMismatchError(
            f"Unknown parameter(s) {unknown} for {spec.name}. "
            f"Known parameters: {list(known)}."
        )
    missing = [
        name for name, p in known.items() if p.required and name not in supplied
    ]
    if missing:
        raise ParameterMismatchError(
            f"Missing required parameter(s) {missing} for {spec.name}."
        )


def build_distribution_formula(
    function_name: str,
    params: dict[str, Any] | list[Any],
    catalogue: FunctionCatalogue,
) -> str:
    """Build a `=VoseXXX(...)` formula string.

    `params` is either a dict keyed by parameter name (preferred) or a
    positional list (only accepted if every required parameter is
    supplied in catalogue order — trailing optional parameters may be
    omitted). The function name is validated against the catalogue and
    bad names raise `UnknownFunctionError` with a close-match suggestion.
    """
    spec = catalogue.require(function_name)
    if isinstance(params, list):
        params = _positional_to_named(spec, params)
    _validate_param_names(spec, params)
    ordered_args: list[str] = []
    last_supplied_index = -1
    for i, p in enumerate(spec.parameters):
        if p.name in params:
            last_supplied_index = i
    for i, p in enumerate(spec.parameters):
        if i > last_supplied_index:
            break
        if p.name in params:
            ordered_args.append(render_value(params[p.name]))
        else:
            # An optional parameter skipped in the middle of the supplied
            # range — represent as empty (Excel uses "," with nothing between).
            ordered_args.append("")
    return f"={spec.name}({','.join(ordered_args)})"


def _positional_to_named(
    spec: FunctionSpec, values: list[Any]
) -> dict[str, Any]:
    if len(values) > len(spec.parameters):
        raise ParameterMismatchError(
            f"Too many positional arguments for {spec.name}: "
            f"got {len(values)}, takes at most {len(spec.parameters)}."
        )
    return {p.name: v for p, v in zip(spec.parameters, values, strict=False)}


# ----------------------------------------------------------------------
# Wrapper formulas
# ----------------------------------------------------------------------


def build_input_wrapper(name: str, inner_formula: str) -> str:
    """Wrap an inner formula with `VoseInput("name") +`.

    The inner formula may be passed with or without a leading "=". The
    returned formula always starts with "=".
    """
    inner = _strip_leading_equals(inner_formula)
    safe_name = _quote_vose_string(name)
    return f"=VoseInput({safe_name})+{inner}"


def build_output_wrapper(name: str, inner_formula: str) -> str:
    """Wrap with `VoseOutput("name") +`. Same rules as the input wrapper."""
    inner = _strip_leading_equals(inner_formula)
    safe_name = _quote_vose_string(name)
    return f"=VoseOutput({safe_name})+{inner}"


def _strip_leading_equals(formula: str) -> str:
    f = formula.strip()
    if f.startswith("="):
        return f[1:].lstrip()
    return f


def _quote_vose_string(s: str) -> str:
    if not s:
        raise ParameterMismatchError("Wrapper name must be non-empty.")
    return '"' + s.replace('"', '""') + '"'


# ----------------------------------------------------------------------
# Composite builders for tools listed in spec §7.2
# ----------------------------------------------------------------------


def build_aggregate_mc(
    frequency_obj_cell: str,
    severity_obj_cell: str,
    catalogue: FunctionCatalogue,
    min_limit: float | str | None = None,
    max_limit: float | str | None = None,
    distribution_shift: float | str | None = None,
) -> str:
    params: dict[str, Any] = {
        "n": frequency_obj_cell,
        "distribution": severity_obj_cell,
    }
    if min_limit is not None:
        params["MinLimit"] = min_limit
    if max_limit is not None:
        params["MaxLimit"] = max_limit
    if distribution_shift is not None:
        params["DistributionShift"] = distribution_shift
    return build_distribution_formula("VoseAggregateMC", params, catalogue)


def build_risk_event(
    probability: float | str,
    impact_function_name: str,
    impact_params: dict[str, Any] | list[Any],
    catalogue: FunctionCatalogue,
) -> str:
    """Build a `VoseRiskEvent(prob, impact_obj)` formula.

    The impact must be a Vose distribution object — i.e. a function from
    the `object` category. This is enforced against the catalogue.
    """
    impact_spec = catalogue.require(impact_function_name)
    if impact_spec.category != "object":
        raise ParameterMismatchError(
            f"VoseRiskEvent impact must be a distribution-object function "
            f"(category 'object'), but {impact_function_name} is "
            f"{impact_spec.category!r}. Look for a Vose*Object variant."
        )
    impact_formula = build_distribution_formula(
        impact_function_name, impact_params, catalogue
    )
    return build_distribution_formula(
        "VoseRiskEvent",
        {
            "probability": probability,
            "impact": impact_formula,  # treated as a formula fragment
        },
        catalogue,
    )


def build_time_series(
    function_name: str,
    params: dict[str, Any] | list[Any],
    catalogue: FunctionCatalogue,
) -> str:
    """Build a `=VoseTimeXXX(...)` formula. Validates that the function is
    in the time-series category."""
    spec = catalogue.require(function_name)
    if spec.category != "time-series":
        raise ParameterMismatchError(
            f"{function_name} is in category {spec.category!r}, not "
            f"'time-series'. Use `build_distribution_formula` for it instead, "
            f"or pick a Vose*Time* function."
        )
    return build_distribution_formula(function_name, params, catalogue)


def build_copula(
    function_name: str,
    params: dict[str, Any] | list[Any],
    catalogue: FunctionCatalogue,
) -> str:
    """Build a copula formula (e.g. VoseCopulaMultiNormal). Validates that
    the function is in the copula category."""
    spec = catalogue.require(function_name)
    if spec.category != "copula":
        raise ParameterMismatchError(
            f"{function_name} is in category {spec.category!r}, not "
            f"'copula'. Pick a Vose*Copula* function."
        )
    return build_distribution_formula(function_name, params, catalogue)
