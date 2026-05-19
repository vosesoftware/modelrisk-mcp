"""Extract the ModelRisk function catalogue into functions.json.

Parses two source files in the ModelRisk repository and cross-references
them by base function name:

- `ModelRiskAtl.idl` — COM IDL. Authoritative for: function name, parameter
  names, [optional] flags. Pattern matched: `FNC_VISIBLE_VBA(Name)(...)`.

- `XllAddIn_English.h` — XLL UDF registration table. Authoritative for:
  category ID (constants from `XllAddIn.h`), one-line description, and
  per-parameter help text.

Outputs `src/modelrisk_mcp/data/functions.json` with one entry per
Vose-prefixed function name (spec §6 schema).

Run from a checkout of the modelrisk-mcp repo with the ModelRisk repo
cloned alongside:

    uv run python scripts/extract_catalogue.py \
        --modelrisk-repo C:/Users/timou/source/repos/ModelRisk
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

# Source-of-truth constants from XllAddIn.h:72-87
CATEGORY_VALUES: dict[str, int] = {
    "VOSE_GENERAL_ID": 0,
    "VOSE_DISTR_ID": 1,
    "VOSE_DISTR_OBJECT_ID": 2,
    "VOSE_AGGREGATE_ID": 3,
    "VOSE_COPULA_ID": 4,
    "VOSE_FIT_ID": 5,
    "VOSE_PROB_ID": 6,
    "VOSE_PROB10_ID": 7,
    "VOSE_TIMESERIES_ID": 8,
    "VOSE_OPT_ID": 9,
    "VOSE_SIXSIGMA_ID": 10,
    "VOSE_DATA_ID": 11,
    "VOSE_INTERNAL_ID": 0xFFF,
    "LOCKIN_STANDARD": 0x100,
    "VOSE_FUNCS_MASK": 0xFF,
}

CATEGORY_OF_ID: dict[int, str] = {
    0: "utility",
    1: "continuous",  # may be reclassified to "discrete" by function name
    2: "object",
    3: "aggregate",
    4: "copula",
    5: "fitting",
    6: "property",
    7: "property",
    8: "time-series",
    9: "utility",
    10: "utility",
    11: "utility",
}

# Names of discrete distributions inside VOSE_DISTR_ID (so VOSE_DISTR_ID
# alone is not enough — these get reclassified). Source: domain knowledge
# of Monte Carlo distribution families.
DISCRETE_BASE_NAMES: frozenset[str] = frozenset({
    "Bernoulli",
    "BetaBinomial",
    "BetaGeometric",
    "BetaNegBin",
    "Binomial",
    "Delaporte",
    "Discrete",
    "DUniform",
    "Geometric",
    "HyperGeo",
    "HyperGeoD",
    "HypGeoD",
    "HyperGeometric",
    "IntUniform",
    "NegBin",
    "NegativeBinomial",
    "Poisson",
    "PoissonUnif",
})

# Parameter names that take arrays/ranges rather than scalars. Heuristic
# applied because the IDL types every parameter as VARIANT.
ARRAY_PARAM_HINTS: frozenset[str] = frozenset({
    "values",
    "probabilities",
    "probs",
    "weights",
    "data",
    "samples",
    "range",
    "ranges",
    "matrix",
    "correlations",
    "alpha",  # Dirichlet takes vector alpha
})

BOOLEAN_PARAM_NAMES: frozenset[str] = frozenset({
    "cumulative",
    "uncertainty",
})

# IDs whose entries are excluded from the catalogue (internal-only).
EXCLUDED_CATEGORY_IDS: frozenset[int] = frozenset({0xFFF})

ENGLISH_PREFIX = "Vose"


# ----------------------------------------------------------------------
# IDL parsing
# ----------------------------------------------------------------------

# Matches: FNC_VISIBLE_VBA(Name)(VARIANT a, VARIANT b, [optional]VARIANT c, [out,retval] VARIANT *res)
_IDL_FUNC_RE = re.compile(
    r"FNC_VISIBLE_VBA\(\s*(?P<name>\w+)\s*\)\s*\((?P<params>[^)]*)\)\s*;",
    re.DOTALL,
)


@dataclass
class IdlParam:
    name: str
    required: bool


@dataclass
class IdlFunction:
    name: str
    params: list[IdlParam] = field(default_factory=list)


def parse_idl(idl_path: Path) -> dict[str, IdlFunction]:
    text = idl_path.read_text(encoding="utf-8", errors="replace")
    funcs: dict[str, IdlFunction] = {}
    for m in _IDL_FUNC_RE.finditer(text):
        name = m.group("name")
        params_str = m.group("params")
        params = _parse_idl_params(params_str)
        # If a name appears more than once in the IDL, keep the first
        # occurrence (avoids #ifdef ARCHER/XLSTAT branches double-counting).
        if name not in funcs:
            funcs[name] = IdlFunction(name=name, params=params)
    return funcs


def _parse_idl_params(params_str: str) -> list[IdlParam]:
    # Bracketed attributes like `[out,retval]` and `[in, optional]` contain
    # commas. Replace them with placeholders so the comma split doesn't
    # tear them apart, then recover the attribute content for inspection.
    tokenised = re.sub(
        r"\[([^\]]*)\]",
        lambda m: f"<<ATTR:{m.group(1).replace(',', ';')}>>",
        params_str,
    )

    result: list[IdlParam] = []
    for raw in tokenised.split(","):
        raw = raw.strip()
        if not raw:
            continue
        attrs = re.findall(r"<<ATTR:([^>]*)>>", raw)
        # Each attr is a single block; multiple comma-separated attrs are
        # rejoined with ';' by the substitution above.
        attr_text = " ".join(attrs).lower()
        if "out" in attr_text and "retval" in attr_text:
            continue  # return value, not an input parameter
        optional = "optional" in attr_text
        # Strip attribute markers, leaving just the C type and name.
        type_and_name = re.sub(r"<<ATTR:[^>]*>>", " ", raw).strip()
        tokens = type_and_name.split()
        if not tokens:
            continue
        name = tokens[-1].lstrip("*").strip()
        if not name:
            continue
        result.append(IdlParam(name=name, required=not optional))
    return result


# ----------------------------------------------------------------------
# XLL header parsing
# ----------------------------------------------------------------------

# Matches one record body. Records start with "{" optionally followed by a
# "// ..." comment (the comment is sometimes a number like "//198",
# sometimes a category label like "//Fit Function", sometimes empty "//",
# and sometimes absent entirely). We require the next non-whitespace token
# to be PREFIX_IN_CODE_T — that distinguishes a real function record from
# the outer "{" that wraps the whole table. The closing "},"  must be at
# the start of a (possibly whitespace-prefixed) line — otherwise a literal
# "},u" inside a TSTR string like "min,max,{values},u" would terminate
# the record prematurely.
_XLL_RECORD_RE = re.compile(
    r"\{(?:\s*//[^\n]*)?\s*\n(?=\s*PREFIX_IN_CODE_T)(.*?)^\s*\}\s*,",
    re.DOTALL | re.MULTILINE,
)

_TSTR_RE = re.compile(r'TSTR\(\s*"((?:[^"\\]|\\.)*)"\s*\)')
_CATEGORY_EXPR_RE = re.compile(r"\(\s*Tchar\s*\*\s*\)\s*\(([^)]+)\)")

# XLL type-code strings: 1-20 uppercase letters optionally followed by
# zero or more of #, !, $, ^, &, % (the SDK's "uncalced", "volatile",
# "thread-safe" etc. modifiers). Examples seen in source:
#   "RPPPPP#", "RPP", "RB", "RP!$", "RRRPPPP!"
_XLL_TYPECODE_RE = re.compile(r"^[A-Z]{1,30}[#!$^&%]*$")


@dataclass
class XllEntry:
    name: str
    category_id: int
    description: str
    param_names_csv: str = ""  # raw "n,distribution,MinLimit,..." from the .h
    param_help: list[str] = field(default_factory=list)


def parse_xll_header(xll_path: Path) -> dict[str, XllEntry]:
    text = xll_path.read_text(encoding="utf-8", errors="replace")
    entries: dict[str, XllEntry] = {}
    for m in _XLL_RECORD_RE.finditer(text):
        body = m.group(1)
        entry = _parse_xll_record(body)
        if entry is None:
            continue
        # First-occurrence wins (handles any duplicate or localised variants).
        entries.setdefault(entry.name, entry)
    return entries


def _parse_xll_record(body: str) -> XllEntry | None:
    strings = _TSTR_RE.findall(body)
    if len(strings) < 4:
        return None
    name = strings[0]
    cat_match = _CATEGORY_EXPR_RE.search(body)
    if not cat_match:
        return None
    cat_value = _eval_category_expr(cat_match.group(1))

    # Each .h record has this fixed TSTR ordering (the FNCSTR_VISIBLE_T
    # macro and the (Tchar*) category expression aren't TSTR calls, so
    # they're invisible to our regex):
    #   [0] function name
    #   [1] XLL type code
    #   [2] parameter-name CSV (single name if 1 param, e.g. "Shift")
    #   [3] small numeric ("1")
    #   [4] empty stub ("")
    #   [5] help reference ("Ma.chm!N") — sometimes missing
    #   [6] description
    #   [7..N-2] per-parameter help strings
    #   [N-1] trailing ""
    #
    # The Ma.chm ref is the most reliable anchor — description always sits
    # immediately after it. If it isn't present, fall back to position 6.
    param_names_csv = strings[2] if len(strings) > 2 else ""
    description = ""
    param_help: list[str] = []
    ref_idx = next(
        (i for i, s in enumerate(strings) if s.startswith("Ma.chm!")),
        -1,
    )
    desc_start = ref_idx + 1 if ref_idx >= 0 else 6
    if desc_start < len(strings):
        description = strings[desc_start].strip()
        tail = strings[desc_start + 1:]
        # Drop trailing empty markers.
        while tail and tail[-1].strip() == "":
            tail = tail[:-1]
        param_help = [s.strip() for s in tail if s.strip()]
    return XllEntry(
        name=name,
        category_id=cat_value,
        description=description,
        param_names_csv=param_names_csv,
        param_help=param_help,
    )


def _eval_category_expr(expr: str) -> int:
    """Evaluate something like 'VOSE_DISTR_ID | LOCKIN_STANDARD' to an int."""
    expr = expr.strip()
    total = 0
    for part in re.split(r"\|", expr):
        ident = part.strip()
        if ident in CATEGORY_VALUES:
            total |= CATEGORY_VALUES[ident]
    return total


# ----------------------------------------------------------------------
# Cross-reference + classification
# ----------------------------------------------------------------------


def classify(name: str, category_id: int) -> str | None:
    """Return the spec-level category string, or None if the function
    should be excluded."""
    if category_id in EXCLUDED_CATEGORY_IDS:
        return None
    masked = category_id & CATEGORY_VALUES["VOSE_FUNCS_MASK"]
    base = CATEGORY_OF_ID.get(masked)
    if base is None:
        return None
    if base == "continuous" and name in DISCRETE_BASE_NAMES:
        return "discrete"
    return base


def param_type(param_name: str) -> str:
    n = param_name.lower()
    if n in BOOLEAN_PARAM_NAMES:
        return "boolean"
    if n in ARRAY_PARAM_HINTS:
        return "array"
    return "number"


def return_type(spec_category: str) -> str:
    if spec_category in ("object", "fitting"):
        return "object"
    if spec_category in ("time-series", "copula"):
        return "array"
    return "number"


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------


_OPTIONAL_HINT_RE = re.compile(
    r"\b(?:optional|if omitted|not applied if omitted|defaults?\s+to|"
    r"\(\s*optional\s*\))",
    re.IGNORECASE,
)


def _parse_xll_param_csv(csv: str) -> list[str]:
    """Parse a param-name CSV like 'n,distribution,{MinLimit},{MaxLimit},u'.

    The .h file uses braces to mark array-like / range parameters. We strip
    them here; the array-vs-scalar decision happens later in `param_type`.
    """
    if not csv:
        return []
    names: list[str] = []
    for raw in csv.split(","):
        token = raw.strip().strip("{}").strip()
        if token:
            names.append(token)
    return names


def _params_from_xll(entry: XllEntry) -> list[IdlParam]:
    """Reconstruct a param list for functions absent from the IDL.

    Uses the parameter-name CSV for names, and the per-parameter help-text
    strings for required/optional inference. Heuristic but reliable for
    the small set of XLL-only functions we've seen (e.g. `AggregateMC`,
    `RiskEvent`).
    """
    names = _parse_xll_param_csv(entry.param_names_csv)
    helps = entry.param_help
    out: list[IdlParam] = []
    for i, name in enumerate(names):
        help_text = helps[i] if i < len(helps) else ""
        optional = bool(_OPTIONAL_HINT_RE.search(help_text)) if help_text else False
        out.append(IdlParam(name=name, required=not optional))
    return out


def apply_optional_overrides(
    catalogue: dict[str, dict[str, Any]], overrides_path: Path
) -> int:
    """Flip selected parameters from required=True to required=False, and
    record a 'default' value for documentation. Returns the number of
    overrides applied. Silently skips entries whose function or
    parameter aren't in the catalogue (so the file can be edited
    without breaking the build).
    """
    if not overrides_path.is_file():
        return 0
    raw = yaml.safe_load(overrides_path.read_text(encoding="utf-8"))
    if not raw:
        return 0
    applied = 0
    for func_base_name, override_list in raw.items():
        full_name = (
            func_base_name if func_base_name.startswith("Vose")
            else f"Vose{func_base_name}"
        )
        entry = catalogue.get(full_name)
        if entry is None:
            continue
        for override in override_list or []:
            target_name = override.get("param")
            default = override.get("default")
            if not target_name:
                continue
            for param in entry["parameters"]:
                if param["name"] == target_name:
                    param["required"] = False
                    if default is not None:
                        param["default"] = default
                    applied += 1
                    break
    return applied


def build_catalogue(idl_path: Path, xll_path: Path) -> dict[str, dict]:
    idl = parse_idl(idl_path)
    xll = parse_xll_header(xll_path)

    catalogue: dict[str, dict] = {}
    skipped_no_xll: list[str] = []
    skipped_no_idl_no_csv: list[str] = []
    skipped_internal: list[str] = []
    skipped_unknown_category: list[str] = []
    xll_only_included: list[str] = []

    # Pass 1: functions in IDL — IDL is the source of truth for signatures,
    # XLL adds category and description.
    for base_name, idl_func in idl.items():
        xll_entry = xll.get(base_name)
        if xll_entry is None:
            skipped_no_xll.append(base_name)
            continue
        spec_category = classify(base_name, xll_entry.category_id)
        if spec_category is None:
            if xll_entry.category_id in EXCLUDED_CATEGORY_IDS:
                skipped_internal.append(base_name)
            else:
                skipped_unknown_category.append(
                    f"{base_name} (category_id=0x{xll_entry.category_id:x})"
                )
            continue

        params_out = [
            {"name": p.name, "type": param_type(p.name), "required": p.required}
            for p in idl_func.params
        ]
        catalogue[f"{ENGLISH_PREFIX}{base_name}"] = {
            "category": spec_category,
            "parameters": params_out,
            "returns": return_type(spec_category),
            "description": xll_entry.description,
        }

    # Pass 2: XLL-only functions (registered as Excel UDFs but not exposed
    # via COM/VBA). Reconstruct signatures from the .h param-name CSV.
    for base_name, xll_entry in xll.items():
        if base_name in idl:
            continue  # already handled in pass 1
        spec_category = classify(base_name, xll_entry.category_id)
        if spec_category is None:
            if xll_entry.category_id in EXCLUDED_CATEGORY_IDS:
                skipped_internal.append(base_name)
            else:
                skipped_unknown_category.append(
                    f"{base_name} (category_id=0x{xll_entry.category_id:x})"
                )
            continue
        params = _params_from_xll(xll_entry)
        if not params and not xll_entry.param_names_csv:
            skipped_no_idl_no_csv.append(base_name)
            continue
        params_out = [
            {"name": p.name, "type": param_type(p.name), "required": p.required}
            for p in params
        ]
        catalogue[f"{ENGLISH_PREFIX}{base_name}"] = {
            "category": spec_category,
            "parameters": params_out,
            "returns": return_type(spec_category),
            "description": xll_entry.description,
        }
        xll_only_included.append(base_name)

    # Report counts so a human reviewer can spot regressions.
    if skipped_no_xll:
        print(
            f"Note: {len(skipped_no_xll)} IDL function(s) had no matching XLL "
            f"entry (likely COM-only Array variants); skipped. "
            f"E.g.: {', '.join(skipped_no_xll[:5])}.",
            file=sys.stderr,
        )
    if xll_only_included:
        print(
            f"Note: {len(xll_only_included)} XLL-only function(s) reconstructed "
            f"from .h (e.g. {', '.join(xll_only_included[:5])}).",
            file=sys.stderr,
        )
    if skipped_no_idl_no_csv:
        print(
            f"Note: {len(skipped_no_idl_no_csv)} XLL-only function(s) had no "
            f"parsable parameter CSV; skipped.",
            file=sys.stderr,
        )
    if skipped_internal:
        print(
            f"Note: {len(skipped_internal)} function(s) excluded as internal-only.",
            file=sys.stderr,
        )
    if skipped_unknown_category:
        print(
            f"Warning: {len(skipped_unknown_category)} function(s) had an "
            f"unrecognised category ID:",
            file=sys.stderr,
        )
        for s in skipped_unknown_category[:10]:
            print(f"  - {s}", file=sys.stderr)

    return catalogue


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument(
        "--modelrisk-repo",
        required=True,
        type=Path,
        help="Path to the ModelRisk repository checkout.",
    )
    ap.add_argument(
        "--output",
        default=Path(__file__).resolve().parent.parent
        / "src"
        / "modelrisk_mcp"
        / "data"
        / "functions.json",
        type=Path,
        help="Where to write functions.json (default: src/modelrisk_mcp/data/functions.json).",
    )
    args = ap.parse_args()

    idl_path = args.modelrisk_repo / "ModelRisk_Project" / "VBAProject" / "ModelRiskAtl" / "ModelRiskAtl.idl"
    xll_path = args.modelrisk_repo / "ModelRisk_Project" / "ModelRisk" / "XllAddIn_English.h"

    if not idl_path.is_file():
        print(f"IDL not found: {idl_path}", file=sys.stderr)
        return 2
    if not xll_path.is_file():
        print(f"XLL header not found: {xll_path}", file=sys.stderr)
        return 2

    catalogue = build_catalogue(idl_path, xll_path)

    overrides_path = (
        Path(__file__).resolve().parent.parent
        / "src"
        / "modelrisk_mcp"
        / "data"
        / "optional_overrides.yaml"
    )
    overrides_applied = apply_optional_overrides(catalogue, overrides_path)
    if overrides_applied:
        print(
            f"Applied {overrides_applied} optional-parameter override(s) from "
            f"{overrides_path.name}.",
            file=sys.stderr,
        )

    out_path: Path = args.output
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8", newline="\n") as f:
        json.dump(catalogue, f, indent=2, ensure_ascii=False, sort_keys=True)
        f.write("\n")

    print(f"Wrote {len(catalogue)} function entries to {out_path}")

    # Quick category histogram for sanity.
    histogram: dict[str, int] = {}
    for entry in catalogue.values():
        histogram[entry["category"]] = histogram.get(entry["category"], 0) + 1
    print("Category histogram:")
    for cat, count in sorted(histogram.items(), key=lambda kv: -kv[1]):
        print(f"  {cat:<14} {count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
