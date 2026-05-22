"""Parse VoseInput / VoseOutput name arguments.

ModelRisk allows the name argument to be either a string literal or a
cell reference:

    =VoseInput("WidgetCost")         <- string literal
    =VoseInput(A5)                   <- same-sheet cell reference
    =VoseInput(Sheet1!A5)            <- qualified cell reference
    =VoseInput('Sheet with spaces'!A5)  <- quoted-sheet ref

The cell-reference form is documented and common in real workbooks (it
lets users keep input labels in their headers and reference them
without manually retyping). Previous versions of the scanner only
matched string literals — workbooks built around the cell-reference
pattern returned empty input/output lists, which cascaded into empty
sensitivity rankings, empty diagnose_workbook counts, and empty
results from get_sensitivity_ranking / build_drivers_report.

This module exposes `extract_vose_first_arg(formula, wrapper)` which
returns:

- `LiteralName("the name")` for a string-literal arg
- `CellRefName(sheet, cell)` for a cell-reference arg (sheet may be
  empty string to mean "same sheet as the wrapper cell")
- `None` if the wrapper isn't in the formula or the first arg form is
  unrecognised (functions, ranges, etc. — we don't try to evaluate
  Excel expressions)

The caller is responsible for resolving CellRefName to an actual string
by reading the target cell (typically via `ExcelBridge.get_cell`).
"""

from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass(frozen=True)
class LiteralName:
    """The wrapper's first arg was a string literal — `name` is the
    actual name with Excel doubled-quote escapes already collapsed."""

    name: str


@dataclass(frozen=True)
class CellRefName:
    """The wrapper's first arg was a cell reference. `sheet` is the
    sheet name from the ref (empty string means "same sheet as the
    wrapper cell" — caller must supply the context). `cell` is the
    A1 ref with any leading $ stripped."""

    sheet: str  # "" means same-sheet (context-dependent)
    cell: str   # A1-style, e.g. "A5" or "AB12"


@dataclass(frozen=True)
class ExpressionName:
    """The wrapper's first arg was an Excel expression (e.g.
    `"prefix"&B8&" suffix"` or `CONCAT(B8, B23)` or any arithmetic).
    The runtime-evaluated name is only knowable after Excel computes
    the formula — typically only ModelRisk's XLL sees the final
    string.

    `static_prefix` is whatever literal-string portion we could
    extract from the start of the expression. Useful for display
    ("we see the name starts with X"), but MUST NOT be used as the
    canonical name when comparing against the simulation's .vmrs
    output registry — the actual name will be longer or different.

    Bug #32 (alpha.31): without this type the parser silently
    returned a partial-LiteralName for expression-based wrappers
    like Vose's own sample `VoseOutput("Total net revenue from "
    &B8&" to "&B23,"$k")`. The bridge then asked MRService to look
    up the partial name, which never matched the runtime-evaluated
    name, and `run_simulation`'s post-condition verification
    false-positively claimed the sim failed."""

    static_prefix: str  # informational only; not a usable lookup key


# Matches a cell reference at the start of a string. Captures:
#   group "sheet"  — sheet name (without quotes), optional
#   group "col"    — column letters
#   group "row"    — row digits
_CELL_REF_RE = re.compile(
    r"^"
    # Optional sheet prefix. If present it MUST be followed by `!` —
    # without that requirement the regex was greedily matching the
    # first letter of multi-letter columns (e.g. "AB123" parsed as
    # sheet="A", col="B", row="123").
    r"(?:"
        r"(?:"
            r"'(?P<sheet_quoted>[^']+)'"        # 'Sheet with spaces'
            r"|"
            r"(?P<sheet_bare>[A-Za-z_][\w\.]*)"  # Sheet1 or My_Sheet
        r")"
        r"!"                                     # required separator
    r")?"
    r"\$?(?P<col>[A-Z]+)"
    r"\$?(?P<row>\d+)"
    r"$"
)


def extract_vose_first_arg(
    formula: str, wrapper: str,
) -> LiteralName | CellRefName | ExpressionName | None:
    """Extract and classify the first argument of a Vose wrapper call.

    `wrapper` is the function name without the opening paren, e.g.
    "VoseInput" or "VoseOutput". Match is case-sensitive — ModelRisk
    function names are always capitalised consistently.

    Returns the typed argument, or `None` if the wrapper isn't present
    or the first arg form isn't supported.
    """
    if not formula:
        return None
    idx = formula.find(wrapper + "(")
    if idx < 0:
        return None
    start = idx + len(wrapper) + 1
    # Skip leading whitespace
    while start < len(formula) and formula[start] in " \t":
        start += 1
    if start >= len(formula):
        return None

    # Case 1: string literal (possibly the start of an expression).
    if formula[start] == '"':
        i = start + 1
        buf: list[str] = []
        while i < len(formula):
            ch = formula[i]
            if ch == '"':
                # Doubled "" inside a string literal escapes to a single "
                if i + 1 < len(formula) and formula[i + 1] == '"':
                    buf.append('"')
                    i += 2
                    continue
                # End of string literal at position i. Now check what
                # follows: a `,` or `)` (possibly after whitespace)
                # means the literal IS the whole first argument. A
                # `&`, `+`, or any other operator means the literal
                # was just the start of an expression — bug #32
                # surfaced by Vose's own `Inputs Outputs.xlsx` sample
                # which uses `VoseOutput("prefix "&B8&" suffix",...)`.
                j = i + 1
                while j < len(formula) and formula[j] in " \t":
                    j += 1
                if j < len(formula) and formula[j] in ",)":
                    return LiteralName(name="".join(buf))
                # Expression — return the partial prefix as informational.
                return ExpressionName(static_prefix="".join(buf))
            buf.append(ch)
            i += 1
        return None  # unterminated string literal

    # Case 2: not a literal — capture until first comma or close-paren
    # at depth 0 and classify the captured token.
    end = start
    depth = 0
    while end < len(formula):
        ch = formula[end]
        if ch == "(":
            depth += 1
        elif ch == ")":
            if depth == 0:
                break
            depth -= 1
        elif ch == "," and depth == 0:
            break
        end += 1
    arg = formula[start:end].strip()
    if not arg:
        return None

    # Try to parse as a cell reference.
    m = _CELL_REF_RE.match(arg)
    if m:
        sheet = m.group("sheet_quoted") or m.group("sheet_bare") or ""
        cell = f"{m.group('col')}{m.group('row')}"
        return CellRefName(sheet=sheet, cell=cell)

    # Anything else (a function call, a range, arithmetic, etc.) — we
    # don't try to evaluate. Returning None means the scanner skips
    # this cell. Conservative on purpose: false negatives are better
    # than confidently mis-attributing a name.
    return None


__all__ = [
    "CellRefName",
    "ExpressionName",
    "LiteralName",
    "extract_vose_first_arg",
]
