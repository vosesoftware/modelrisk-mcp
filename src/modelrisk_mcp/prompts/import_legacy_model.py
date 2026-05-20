"""/import-legacy-model prompt template."""

from __future__ import annotations

from modelrisk_mcp.server import mcp

description: str = (
    "Open a workbook built with another Monte Carlo add-in (notably "
    "the legacy RiskXXX(...) functions) and propose ModelRisk "
    "equivalents cell by cell."
)

template: str = """\
You are migrating a workbook from a legacy Monte Carlo add-in to
ModelRisk. The most common case is workbooks built with `RiskXXX`
functions from another vendor.

Workflow:

1. **Confirm the workbook.** Call `get_active_workbook` and
   `list_distributions`. If the listed function names don't start
   with `Vose`, they may be from another add-in still resolving in
   Excel (and the cells show #NAME? or stale values).

2. **Catalogue the legacy calls.** Iterate cells via `get_cell`
   or `read_range` and identify every `Risk*(`, `@Risk*(`,
   `Crystal*(`, or other vendor prefix.

3. **Propose mappings.** For each legacy function, find the
   closest ModelRisk equivalent in the catalogue:
   - `RiskNormal(mu, sigma)` → `VoseNormal(mu, sigma)`
   - `RiskTriang(a, b, c)` → `VoseTriangle(a, b, c)`
   - `RiskPert(min, ml, max)` → `VosePERT(min, ml, max)`
     (or `VoseModPERT` for the gamma-sharpened variant)
   - `RiskOutput()+...` → `VoseOutput("...")+...`
   - `RiskMakeInput(...)+...` → `VoseInput("...")+...`
   For anything you don't recognise, look it up against
   `modelrisk://functions` and confirm with the user.

4. **Commit one cell at a time.** Use `insert_distribution`,
   `wrap_with_input`, or `wrap_with_output` with `dry_run=True`
   first, then commit. Restore via `restore_cell` if needed.

5. **Re-audit + run.** Run `audit_model` to catch any leftover
   non-Vose distribution calls, then `run_simulation` to verify
   the converted workbook produces sensible results.

Don't translate logic the user didn't ask about. If a cell uses a
non-distribution function from the legacy add-in (e.g. their
correlation matrix function), flag it but don't auto-convert.
"""


@mcp.prompt(name="import-legacy-model", description=description)
def import_legacy_model_prompt() -> str:
    return template
