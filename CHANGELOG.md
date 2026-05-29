# Changelog

All notable changes to ModelRisk MCP. Follows [Keep a Changelog](https://keepachangelog.com/) and [Semantic Versioning](https://semver.org/).

## [Unreleased]

## [0.3.1-alpha.7] — 2026-05-29

### Added

- **SS-004 inconsistent_formula_in_block** (warning) — the single most valuable spreadsheet-integrity check: it catches the classic and most dangerous spreadsheet error (per the EuSpRIG literature) where a row or column of formulas was filled correctly, then one interior cell was overtyped, silently breaking the pattern.

  Detection is tuned for **near-zero false positives**: formulas are normalised to a position-relative form (so a correctly-filled run collapses to one identical string regardless of which cells it references), and a cell is flagged only when it is an **interior cell whose pattern differs from both neighbours while those neighbours agree with each other** — the unambiguous "odd one out in the middle" signature. Edge cells (legitimate first/last-period differences) are never flagged; heterogeneous rows with no agreeing neighbours are never flagged; relative fills that differ literally but share a pattern are never flagged. Works on both horizontal and vertical runs; skips Vose and errored cells.

  `audit_model` now runs **17 rules** (13 VOSE methodology + 4 SS spreadsheet-integrity).

### Tests

`TestInconsistentFormulaInBlock` — 6 cases: horizontal overtype, vertical overtype, clean run, heterogeneous row, edge cell, relative-fill-same-pattern. 519 unit tests pass.

## [0.3.1-alpha.6] — 2026-05-29

### Added

- **A spreadsheet-integrity audit family (SS-001 … SS-003)** — a new class of audit rule, distinct from the VOSE-* Monte-Carlo-methodology rules, that checks whether the *deterministic scaffolding* of the workbook is sound. A model can be methodologically perfect and still wrong if the spreadsheet underneath it is broken. Drawn from the established spreadsheet-error / model-control discipline (O'Beirne, *Spreadsheet Check and Control*; Rees, *Principles of Financial Modelling*; the EuSpRIG literature).
  - **SS-001 magic_number_in_formula** (info) — a parameter-like decimal (1.21, 0.85) buried in a formula instead of a labelled input cell. Skips Vose cells (their literals are distribution parameters).
  - **SS-002 number_stored_as_text** (warning) — a numeric value held as text in a cell a formula references; silently dropped by SUM/arithmetic. Only fires when the cell is actually referenced, to keep false positives near zero.
  - **SS-003 overly_complex_formula** (info) — a single non-Vose formula doing too much (many operators / very long); best practice is one calculation step per cell.

  `audit_model` now runs **16 rules**. The new family is tuned for a low false-positive rate (decimal-only magic numbers, referenced-only text numbers, generous complexity thresholds, Vose cells excluded where their literals are legitimate).

- **An annotated bibliography** in `docs/methodology.md` — the standard references for risk modelling (Vose; Savage's *Flaw of Averages*; Hubbard's *How to Measure Anything*; O'Beirne; Rees; Winston; Charnes; Hulett; Grinstead & Snell; Grzelak & Oosterlee), each with a one-line relevance note. External references for further study; no book content reproduced.

### Changed

- The `modelrisk://methodology` resource now documents the SS-* family in its appendix. The drift-guard test was generalised to cover any rule-id prefix (`PREFIX-###`), so every one of the 16 rules must still be cross-referenced.

### Tests

`TestMagicNumberInFormula`, `TestNumberStoredAsText`, `TestOverlyComplexFormula` (positive + false-positive-avoidance cases each). 513 unit tests pass.

## [0.3.1-alpha.5] — 2026-05-29

### Changed

- **Added a distribution-family quick-reference (§10) to the `modelrisk://knowledge` resource** — a "what it models / when to reach for it / watch-out" table for ~18 of the most-used families (Normal, Lognormal, PERT/ModPERT, Triangle, Uniform, Beta, Gamma, Weibull, Exponential, Pareto, Student-t; Bernoulli, Binomial, Poisson, Negative Binomial, Geometric, Hypergeometric, Discrete). Distilled from the per-family "Uses" sections in the ModelRisk Help's Continuous and Discrete distribution articles. Gives Claude family-level selection guidance at build time to pair with the five-properties framework (§3). Knowledge base is now ten sections; still all from the corpus already on hand, attributed, our own words.

### Tests

The knowledge-base resource test asserts the family quick-reference is present. 502 unit tests pass.

## [0.3.1-alpha.4] — 2026-05-29

### Changed

- **Expanded the `modelrisk://knowledge` resource from six sections to nine**, distilling three more ModelRisk Help articles (still attributed, still our own words):
  - **Fitting distributions to data** — ModelRisk fits by MLE and ranks competing families by information criteria (SIC / HQIC / AIC); the `uncertainty` flag defaults to FALSE only to match common practice but should be TRUE, generating parameter uncertainty by parametric bootstrapping (which captures parameter correlation and non-normal marginals). This is the concrete mechanism behind methodology principle 2 / audit rule VOSE-003.
  - **Choosing an aggregation method** — Monte Carlo (general), Panjer / De Pril (recursive analytic), FFT (fast convolution), and the multivariate variants — which to reach for and when.
  - **Interpreting sensitivity (tornado types)** — conditional mean (default, decision-meaningful), conditional cumulative percentile (tail sensitivity), and rank correlation (crude screening); which to use for which question.

  All three drawn from material already in the ModelRisk Help corpus — no new external sources. The build/audit prompts already point Claude at this resource, so the added depth flows into model construction and critique automatically.

### Tests

The knowledge-base resource test now asserts all nine topics are present. 502 unit tests pass.

## [0.3.1-alpha.3] — 2026-05-29

### Added

- **A curated risk-analysis knowledge base, served as the `modelrisk://knowledge` resource.** An attributed distillation — in our own words, not verbatim — of foundational guidance from the **ModelRisk Help** (Vose Software): why to quantify uncertainty (the vaccine/investment insight), the two forms of uncertainty (variability vs risk events), selecting a distribution via the five properties (discrete/continuous, bounded/unbounded, parametric/non-parametric, univariate/multivariate, first/second order), eliciting expert opinion, modeling correlation (rank-order → envelope → lookup → conditional → copulas), and modeling a quantity over time. Each section cites its source article.

  The new resource is wired into the `/build-risk-model` and `/audit-model` prompts so it's loaded into the LLM's context at build/critique time — grounding model construction in the authoritative source, not just general intuition. This is the sixth MCP resource.

  Single source of truth: the substance ships as the package data file `data/knowledge_base.md` (force-included in the wheel), is served by the resource, and is pointed to from `docs/knowledge-base.md` for GitHub readers — no duplicated copies to drift.

### Tests

`test_knowledge_base_resource_returns_cited_markdown` (provenance marker + all six topics present) and the registration test now assert `modelrisk://knowledge`. 502 unit tests pass.

## [0.3.1-alpha.2] — 2026-05-29

### Changed

- **The `modelrisk://methodology` resource is now a methodology *knowledge base*, not just a list.** Each of the 8 core principles gains a **Why**, a **Failure mode**, and an **Enforced by** line naming the audit rule(s) that police it — so a principle is never just advice, it's tied to a check `audit_model` runs against the live workbook. The resource is loaded into the LLM's context at `/build-risk-model` and `/audit-model` time, so this directly sharpens how Claude builds and critiques models.

  A closing section maps the remaining rules that are about *correctness* or *distribution selection* rather than the core principles (VOSE-001, VOSE-011, VOSE-012, VOSE-013) — so all 13 audit rules are now cross-referenced from the methodology. Nothing in the knowledge base is invented: it's grounded in the existing principles and the existing rule set.

### Added

- `test_methodology_crossref.py` — a **drift guard**: every `VOSE-0NN` cited in the methodology must exist in `audit_rules.yaml`, and every rule must be referenced back. Renaming, adding, or removing a rule without updating the methodology now fails CI. Knowledge that drifts from code is worse than none; this makes drift unmergeable.

### Tests

501 unit tests pass (+4 cross-reference / structure guards).

## [0.3.1-alpha.1] — 2026-05-23

### Changed

- **Report charts now follow a complete styling ruleset** (`docs/chart-style-guide.md`), taking the native Excel histogram and tornado from "generic" to "designer-perfect" — while staying **native, editable chart objects** (no embedded images). The histogram in `build_executive_report` gains:
  - **Round-number bins** (`_nice_bins`): bin edges floor/ceil to a `1/2/2.5/5 × 10ⁿ` width, so the X axis reads `2M 3M 4M…` instead of irregular raw bin centres (`2,182,219 …`). This was the single worst aesthetic problem in the old charts.
  - **Magnitude-aware tick formats** (`_axis_scale_format`): `4M` / `850K` / `420` depending on scale; thinned labels via `TickLabelSpacing`.
  - **Central-80% bar shading**: bars inside [P10, P90] solid brand-blue, tails muted — the confidence interval shown directly on the bars, no extra series.
  - **Decluttered**: count-axis labels removed (absolute frequency isn't decision-relevant), gridlines moved to the cumulative-% axis, no tick marks, no chart/plot borders, tight `GapWidth=16`.
  - **Secondary axis hard-capped at 100%** (`Max=1.0, MajorUnit=0.2`).
  - **Brand typography**: one font family chart-wide, left-aligned navy semibold title.

  The tornado and any chart using the shared `_style_chart_frame` / `_style_chart_axes` helpers inherit the font, border-removal, tick-mark removal, and left-aligned title automatically.

### Added

- `docs/chart-style-guide.md` — the 11-rule chart styling standard, documented (not just coded), with the colour palette table and the function map.
- `_HistogramBins`, `_nice_bins`, `_percentile`, `_axis_scale_format` in `reports.py` — the pure, testable core of the binning/scaling logic.

### Tests

`test_report_binning.py` — 18 cases covering round-boundary binning, label spacing, monotone cumulative, degenerate/empty inputs, percentile interpolation, and magnitude-format selection. 490 unit tests pass.

## [0.3.0] — 2026-05-23

First stable release of the 0.3 line. Promotion from `0.3.0-alpha.38` with no functional changes — version bump only.

### What ships in 0.3.0

A read/build/run MCP server over Vose Software's ModelRisk Excel add-in. The full surface area is callable from Claude Desktop, Claude Code, Claude for Excel, Cursor, Zed, and any MCP-compliant client.

**Tools (40 total):**
- **Reading (12):** `list_open_workbooks`, `get_active_workbook`, `get_workbook_summary`, `list_modelrisk_inputs`, `list_modelrisk_outputs`, `list_distributions`, `get_cell`, `read_range`, `get_simulation_results`, `get_correlation_matrix`, `get_sensitivity_ranking`, `find_hard_coded_inputs`
- **Building (13):** `insert_distribution`, `wrap_with_input`, `wrap_with_output`, `replace_constant_with_distribution`, `fit_distribution_to_data`, `create_aggregate_mc`, `create_copula`, `create_time_series`, `create_risk_event`, `set_named_range`, `write_formula`, `propose_distributions_for_inputs`, `discover_inputs`
- **Simulation (5):** `run_simulation`, `run_scenarios`, `get_samples`, `restore_deterministic_state`, `restore_cell`
- **Workflows / reporting (7):** `audit_model`, `diagnose_workbook`, `build_drivers_report`, `build_executive_report`, `create_tornado_chart`, `generate_executive_summary`, `save_workbook_as`
- **VMRS (3):** `read_vmrs`, `set_active_vmrs`, `list_vmrs_variables`

**Resources (5 URIs):** `modelrisk://functions`, `modelrisk://distributions`, `modelrisk://methodology`, `modelrisk://workbook/current`, `modelrisk://audit-rules`.

**Prompts (5):** `build_model`, `audit_model`, `interpret_results`, `add_uncertainty`, `import_legacy_model`.

**Audit rules (13):** VOSE-001 through VOSE-013. See the `0.3.0-alpha.34/.35` entries for the most recent additions.

**Transports:** `stdio` (default for Claude Desktop) and `streamable-http` / `sse` with bearer-token auth (for Claude for Excel + remote deployments).

**Distribution:**
- PyPI: `pip install modelrisk-mcp` (via OIDC trusted publishing)
- Windows single-file: `modelrisk-mcp.exe` (PyInstaller, ~39 MB, attached to the GitHub release)
- MCP Registry: `io.github.vosesoftware/modelrisk-mcp`
- CLI installer: `modelrisk-mcp install` configures Claude Desktop + Claude Code automatically

### Alpha-cycle highlights since 0.2.0-alpha

- **MRService.dll integration** (α.1–α.7): direct ctypes bridge to ModelRisk's simulation engine — no longer drives Excel's UI for sim execution, ~10× faster and more reliable than COM-driven sims.
- **Bridge layer rewrite** (α.8–α.20): xlwings + pywin32 with stale-reconnect, OneDrive path-resolution fallback, COM CVErr handling, multi-strategy dispatch.
- **Audit rule set** (α.4, α.25, α.33–α.35): grew from 8 → 13 rules, with VOSE-012 (errored cells) and VOSE-013 (arg-count mismatch) catching common LLM hallucination classes.
- **Reporting** (α.13, α.19–α.23): drivers report (tornado + scatter + narrative), executive report (KPI + histogram + chart polish), corporate styling palette.
- **Robustness** (α.21–α.32): named-range scanner sees cell-reference forms, expression-based VoseInput/VoseOutput names, `SaveCopyAs` instead of `SaveAs`, `RegisterXLL` before simulation triggers, post-condition verification on sim completion.
- **Bug surfacing** (α.33–α.36): Excel error cells now distinguished from empty cells in `CellInfo.error`, with bulk `Range.Value2` detection that's robust across Excel versions.
- **Performance** (α.37): 10× speedup on `iterate_cells` / `audit_model` / `get_workbook_summary` by caching the sheet name once instead of once-per-cell.
- **Release pipeline** (α.38): hardened GitHub Actions retries on `actions/checkout` and `mcp-publisher login` / `publish` against transient GitHub auth outages.

### Quality bar at cut

- **472 unit tests** pass; ruff + mypy clean
- **13 audit rules** verified firing end-to-end on a live Excel workbook
- **20 k-cell workbook** audits in **~1 second**; 100 k-cell extrapolation: ~5 s
- **PyPI + MCP Registry**: every alpha tag from α.33 through α.38 round-tripped successfully through the publish pipeline (4 of 5 had clean MCP-registry publishes; one hit a transient GitHub auth outage which alpha.38's retry hardening now covers)
- **Activation-key obfuscation:** `scripts/scan_exe_for_key.py` runs in CI; no plain key in shipped artifacts

## [0.3.0-alpha.38] — 2026-05-23

### Fixed

- **Bug #37 — `publish-mcp-registry` job has zero retry tolerance.** The alpha.37 release ran straight into a transient GitHub Actions auth outage; `actions/checkout@v4` failed three times in a row inside the action with `fatal: could not read Username for 'https://github.com'`, and the registry update for that tag was lost (PyPI publish succeeded — registry was 1 version behind until manually resynced). The same outage window also fired "Failed to save: Our services aren't available right now" warnings against `actions/upload-artifact` and the cache service. Class of failure we'll hit again.

  Fix to `release.yml`'s `publish-mcp-registry` job:
  - Try `actions/checkout@v4` up to **3 times** across separate step invocations (so each attempt re-issues a fresh GITHUB_TOKEN), with a 30-second sleep before the final attempt. Only the third attempt is fatal.
  - Wrap `mcp-publisher login github-oidc` in a shell retry loop (3 attempts, 20-second sleep between).
  - Wrap `mcp-publisher publish` in a shell retry loop (3 attempts, 30-second sleep between).

  Each independent failure mode now needs to lose three coin flips in a row before the job actually fails. Same workflow that produced 4/5 successful registry updates in the past week becomes much more resilient.

## [0.3.0-alpha.37] — 2026-05-22

### Fixed

- **Bug #36 — `iterate_cells` called `sh.name` once per cell.** Each `sh.name` access on an xlwings sheet wrapper triggers a COM round-trip (`ISheet::Get_Name`) costing ~150μs. The inner loop's `CellRef(workbook=workbook, sheet=sh.name, cell=ref)` made that round-trip for every cell it yielded — so a 10k-cell scan paid 1.5 seconds of pure attribute overhead. Caught by `cProfile`: `sh.name` accounted for **81%** of `iterate_cells`' total runtime.

  Fix: cache `sh.name` once per sheet before the inner loop. One line. Verified speedups on the 20k-cell live perf probe:

  | Op | α.36 | α.37 | Speedup |
  |---|---|---|---|
  | `iterate_cells` | 3.37 s | **0.23 s** | 14.7× |
  | `get_workbook_summary` | 3.15 s | **0.36 s** | 8.8× |
  | `find_hard_coded_inputs` | 3.16 s | **0.34 s** | 9.3× |
  | `run_audit` (13 rules) | 7.35 s | **1.04 s** | 7.1× |

### Verified

Round-10 perf probe on a 20k-cell synthetic workbook (2 sheets × 200 rows × 50 cols) — every read/audit path now runs in well under 5 seconds. Headroom: a 100k-cell enterprise model would audit in ~5s rather than ~35s.

## [0.3.0-alpha.36] — 2026-05-22

### Fixed

- **Bug #35 — bulk `Range.Text` returns `None` on some Excel versions, regressing VOSE-012 on real workbooks.** The alpha.33 `iterate_cells` / `read_range` bulk error-detection path relied on `used_range.api.Text` returning a 2D tuple of cell text. On the dev Excel (Office 365) this property returns `None` for any multi-cell range — single-cell `.Text` still works fine. Result: audit scans saw `error=None` on every cell of every sheet, and VOSE-012 silently couldn't fire on errored cells even though `get_cell` (which only reads single cells) worked correctly. Round-10 live-workbook probe caught this — the audit found 0 VOSE-012 findings on a workbook with a deliberate `=1/0` cell.

  Fix: prefer `Range.Value2` for bulk error detection. On a multi-cell range Value2 reliably returns a tuple-of-tuples with the COM CVErr **integer code** in each errored cell's slot (e.g. `-2146826281` for `#DIV/0!`, `-2146826259` for `#NAME?`). The mapping is stable across Excel versions because the lower 16 bits are the well-known `xlCVError` constants. Text remains as a secondary fallback for any cell Value2 didn't classify. Empirically verified all seven canonical errors round-trip correctly: `#DIV/0!`, `#N/A`, `#NAME?`, `#NULL!`, `#NUM!`, `#REF!`, `#VALUE!`.

  `_detect_excel_error` (used by `get_cell`) now also has the Value2 fallback for defence in depth.

### Verified

Round-10 live-workbook audit probe: built one Excel sheet with cells engineered to trigger each of the 13 audit rules, ran `audit_model` against it, asserted every rule fires at least once. Before alpha.36: 12/13 (VOSE-012 missing). After alpha.36: **13/13 — all rules fire on a real workbook.**

### Tests

- 12 new cases in `TestCoerceErrorValue` (every canonical CVErr code → string, plus plain numbers / floats / strings / None / booleans must NOT be misinterpreted)
- 3 new cases in `TestDetectExcelErrorValue2Fallback` (Text=None + Value2 code = detected; Text wins when both present; neither = None)

472 unit tests pass.

## [0.3.0-alpha.35] — 2026-05-22

### Added

- **VOSE-013 audit rule — `arg_count_mismatch`** (severity: `error`). Catches the classic LLM hallucination class that `VOSE-001` (unknown function) misses: a *real* Vose function called with the wrong number of arguments. Examples flagged: `VosePERT(min, max)` (missing mode), `VoseLognormal(mean)` (missing stdev), `VoseTriangle(1, 2)` (missing mode), `VosePERT(1,2,3,4,5,6,7,8)` (too many — beyond the catalogue's max). The cell would `#VALUE!` or `#NUM!` at Excel calc time, but the formula is well-formed enough that VOSE-001 stays silent. With VOSE-013 we flag it statically before the sim runs.

  Rule compares actual arg count against the catalogue's `required` (min) and `len(parameters)` (max). Skips:
  - Functions not in the catalogue (VOSE-001's job)
  - `VoseInput` / `VoseOutput` wrappers (VOSE-008's job)
  - `VoseChoose`, `VoseDiscrete`, `VoseDiscreteUniform` — variadic shapes the catalogue can't fully describe.

  13 audit rules now ship; all 13 wired up in `RULES_BY_NAME`.

- **New helper `safety.count_call_args(formula, function_name)`**: returns one count per occurrence of `function_name(...)` in `formula`. Walks the raw formula so a single string-literal arg correctly counts as 1 (not 0 — the prototype bug caught by the alpha.35 dev pass). Skips strings, nested calls, and array literals atomically. Underpins VOSE-013 but generally useful for static analysis.

### Tests

11 new cases in `TestArgCountMismatch` (too-few PERT, too-few Normal, too-many, correct arity, optional-trailing-args allowed, wrapper exemption, unknown-function silence, non-Vose silence, nested-call detection, suggested-fix content, one-finding-per-cell) + 16 new cases in `TestCountCallArgs` (covering string literals with embedded commas + doubled quotes, nested calls, array literals, multi-call formulas, malformed inputs).

## [0.3.0-alpha.34] — 2026-05-22

### Added

- **VOSE-012 audit rule — `cell_evaluates_to_error`** (severity: `error`). The natural pairing with bug #34: now that `iterate_cells` surfaces Excel error literals via `CellInfo.error`, the audit can flag them. The message is sharper when the errored cell's formula contains a Vose call (e.g. `VosePERT(10, #DIV/0!, 30)` → "the distribution call is broken — the simulation will produce error samples from this cell on every iteration") versus a vanilla broken formula ("trace the formula back to find the root cause"). 12 audit rules now ship; all 12 wired up in `RULES_BY_NAME`.

### Tests

5 new cases in `TestCellEvaluatesToError`: Vose-call diagnostic, generic-formula diagnostic, silent on clean cells, one-finding-per-errored-cell, severity inherits from rule spec.

## [0.3.0-alpha.33] — 2026-05-22

### Fixed

- **Bug #34 — error cells (`#DIV/0!`, `#REF!`, `#NAME?`, ...) were indistinguishable from empty cells.** When a cell evaluated to an Excel error, xlwings' `Range.value` returned `None` — the same value an empty cell returns. `get_cell` therefore reported `value=null, cell_type="formula"` for a broken cell and `value=null, cell_type="empty"` for an empty one, with no way for the LLM to tell them apart. Worse: a Vose call with an errored argument (e.g. `VosePERT(10, #DIV/0!, 30)`) showed up as a normal formula in `read_range`, and audit scans missed broken distributions entirely.

  Fix: detect errors via `Range.Text`, which always renders error cells as their literal (`"#DIV/0!"` etc.). `CellInfo` gains a new optional `error: str | None` field and `cell_type="error"` is a recognised classification. `RangeInfo` gains a parallel `errors: list[list[str | None]]` 2D array (empty list when no cells in the range errored, so the common case stays compact). `iterate_cells` (used by `audit_model` and `find_hard_coded_inputs`) does a single bulk Text read per sheet so the per-cell error info costs no extra COM round-trips. All detection paths fail open: if Text isn't available, the read still returns values + formulas as before.

### Tests

`test_excel_bridge.py` gains `TestDetectExcelError` (12 cases covering Excel error literals, normal cells, edge cases like `#hashtag` text, COM failure, non-string Text) and `TestClassifyCellWithError` (error wins over formula classification).

## [0.3.0-alpha.32] — 2026-05-22

### Fixed

- **Bug #33 — drop the alpha.18 `output_names` pre-populate.** Round-7 testing revealed the alpha.18 hypothesis was wrong. The `output_names` payload to `VoseStartSimulCustom12` is a **filter**, not an enable list — pass empty and the XLL auto-scans the workbook and registers every VoseOutput; pass a list and only matching outputs get registered. So alpha.18, which pre-populated names from the workbook scanner, was strictly worse for any workbook with expression-named outputs (like Vose's own `Inputs Outputs.xlsx` sample): the scanner-extracted prefix never matched the runtime-evaluated name, and NOTHING got registered.

  The alpha.17-era symptom that motivated alpha.18 ("sims completing without registering outputs") was almost certainly bug #29 — XLL commands not callable when Excel was started programmatically — which we fixed properly in alpha.27 via `Application.RegisterXLL`. With #29 fixed, the XLL's auto-scan works correctly. Removing alpha.18's pre-populate makes expression-named outputs register too.

### Verified

Live test on the Vose `Inputs Outputs.xlsx` sample: with empty `output_names`, the .vmrs registered `Period 1` (var_id=1) — variables now register where they didn't before. On the NPV workbook: `NPV (10%)`, `Market growth`, `Sales Price`, and `Conservatives get in? (1=yes)` all resolved as before. Both styles of workbook keep working; the expression-named case is now fixed.

### Tests

408 unit tests pass. Renamed `test_run_simulation_passes_voseoutput_names_to_xll` → `test_run_simulation_does_not_filter_xll_outputs` (sentinel: the bridge MUST NOT pre-filter the XLL).

## [0.3.0-alpha.31] — 2026-05-22

### Fixed

- **Bug #32 — expression-based VoseInput/VoseOutput names false-positive-failed post-condition verification.** Vose's own `Inputs Outputs.xlsx` sample declares its output as `VoseOutput("Total net revenue from "&B8&" to "&B23,"$k")` — the name is an Excel expression, not a static literal. The runtime-evaluated name (e.g. `"Total net revenue from 2020 to 2027"`) is only knowable after Excel computes the formula at simulation time. Our `name_parser` was returning the literal prefix as a `LiteralName` (since it stopped at the closing quote), so:
  - The bridge's `expected_output_names` contained the partial prefix.
  - `run_simulation` passed it to the XLL, which couldn't match it against the actual VoseOutput cell.
  - Post-condition verification looked it up in the produced .vmrs, didn't find it, and raised `SimulationFailedError` — claiming the sim's post-phase had crashed when in fact the only issue was the name-resolution mismatch.

  Fix:
  - New `ExpressionName` type in `name_parser.py` for first-args that turn out to be expressions (literal followed by `&`, `+`, etc., rather than a closing `,` or `)`).
  - The parser detects this by checking what follows the closing quote.
  - `_resolve_vose_name` returns the partial prefix marked with a `…` ellipsis so `list_modelrisk_outputs` still surfaces the cell with an informational name (`"Total net revenue from …"`) instead of dropping it.
  - Post-condition verification filters out `…`-marked names — we can't statically verify them, so we don't try, rather than failing loudly.

### Why this matters

Workbooks that build output names from cell content are a real pattern (year-range labels, scenario-specific outputs, anything dynamic). Before alpha.31 every one of those workbooks looked broken to the bridge. The deeper fix — actually evaluating the Excel expression to get the runtime name and registering THAT with the XLL — is a separate larger investigation; this release is the honesty improvement: don't claim failure when the sim ran fine.

### Tests

408 unit tests pass (+4 in `test_name_parser.py::TestExpressionForm` covering the Vose-sample literal-concat-cellref case, simple `"prefix"&A1`, two-arg literal-with-units, and whitespace tolerance around the closing quote).

## [0.3.0-alpha.30] — 2026-05-22

### Fixed

- **Bug #31 — `samples<=0` passed to `bridge.run_simulation` reached the XLL and triggered an opaque C++ exception** (`OLE error 0xe06d7363`). The MCP tool layer's Pydantic validation enforces `ge=1`, but direct callers (integration tests, automation scripts, future Python clients) bypassed that. Surfaced by the round-3 input-validation probe. Fix: defensive sanity check at the bridge boundary (`samples >= 1` and a 10M soft cap) so every code path produces a clear actionable error before invoking ModelRisk's XLL. The message names the offending value and explains why we're rejecting it.

### Verified end-to-end (no fixes needed)

Round 3 confirmed:
- **MCP tool envelope shapes**: all `list_*` and `find_*` tools return the alpha.17 `{noun: [...], count: N}` envelope correctly. Brand prefix on all 40 tool descriptions.
- **Distribution catalogue breadth**: 9 of 11 sampled families (Lognormal, Uniform, Triangle, Beta, Gamma, Weibull, Poisson, Binomial, Bernoulli) round-trip through `insert_distribution`. The other two failures were test-script errors (`VoseExpon`'s param is `beta` not `mean`; the discrete uniform is `VoseDiscreteU`, not `VoseDiscreteUniform`).
- **MCP resources**: 5 resources registered (`modelrisk://audit-rules`, `/distributions`, `/functions`, `/methodology`, `/workbook/current`) and readable.
- **50K-iteration stress**: simulation completes in 18.4s wall-clock; all 50K samples readable in 0.13s.
- **End-to-end convert workflow** (separate run): non-MR workbook → `discover_inputs` → `propose_distributions_for_inputs` → `replace_constant_with_distribution` → `wrap_with_output` → `run_simulation` → `get_simulation_results` → `build_executive_report`. Produced real randomness (mean $34.9M, stdev $537) on a fully-converted noMR model.

### Tests

404 unit tests pass.

## [0.3.0-alpha.29] — 2026-05-22

### Polished

- **`generate_executive_summary` markdown formatting.** Previously the per-output stats table used `.3g` format which switches to scientific notation past `1e4` — so a mean of $63,300 came out as `6.33e+04`, unreadable in a corporate context. New `_fmt_num` helper uses thousands-separated decimals with two decimal places for normal-range values (`63,300.00`) and only falls back to `.3g` for extreme magnitudes (≥1e9 or <1e−2 in absolute value) where decimal form would be unreadable. Same treatment for the contingency-vs-deterministic delta columns via `_fmt_signed`, which keeps the explicit +/- sign.

### Tests

404 unit tests pass.

## [0.3.0-alpha.28] — 2026-05-22

### Fixed

- **Bug #30 — `diagnose_workbook` mixed data sources when called with an explicit `workbook_name`.** Prior versions always assigned `active_workbook = <Excel-active book's name>` and `workbook_path = <active book's path>`, regardless of which workbook the caller asked to diagnose. Result: calling `diagnose_workbook("foo.xlsx")` while `bar.xlsx` was active in Excel reported `active_workbook="bar.xlsx"` and `workbook_path=<bar's path>` alongside foo's input/output counts — misleading. Worse, the downstream `.vmrs` lookup used `workbook_path` (bar's) and would silently find bar's sibling vmrs instead of foo's. Fix: when an explicit `workbook_name` is supplied, look up that book's path from `list_workbooks` and report it in `workbook_path`. The `active_workbook` field still reflects Excel's active book (useful informationally), but `workbook_path` now consistently describes the workbook being diagnosed.

### Tests

404 unit tests pass. Live verification via the round-2 test pass.

## [0.3.0-alpha.27] — 2026-05-22

### Fixed

- **Bug #29 — `run_simulation` failed when Excel was started programmatically** (e.g. via xlwings' `xw.App()` from an automation context, CI, or any service-driven setup). The ModelRisk XLL would show up as `Installed=True` in the AddIns collection, but `Application.Run('VoseStartSimulCustom12', ...)` failed with `Cannot run the macro 'VoseStartSimulCustom12'`. Root cause: Excel's normal startup flow runs the XLL's `xlAutoOpen` which registers each command via `xlfRegister`; the programmatic-launch path skips that step. Fix: `SimulationController` now calls `Application.RegisterXLL(path)` for every loaded ModelRisk*.xll before its first sim run. `RegisterXLL` is idempotent (re-runs `xlAutoOpen`) so safe to call whether or not the XLL is already fully registered. Cached per controller instance so we only run it once per session, not before every sim.

### Why this matters

The user-driven Claude Desktop session worked fine all along because Excel's normal startup loaded the XLL properly. But anyone running modelrisk-mcp in a non-interactive context — an automation script, a CI test, an MCP server spawned by a daemon, or the autonomous E2E test harness I just ran — would hit `Cannot run the macro` and not know why. This unblocks every non-Claude-Desktop usage pattern.

### Verified

Live probe before fix: `app.api.Run('VoseStartSimulCustom12')` → "Cannot run the macro". After `app.api.RegisterXLL(path)`: the same call resolves. Now part of `_invoke_start_simulation`'s preamble.

## [0.3.0-alpha.26] — 2026-05-22

### Fixed

- **Bug #28 — `get_correlation_matrix` crashed when only one variable resolved.** For a 1×N input matrix (one variable), `numpy.corrcoef` returns a 0-d scalar of value 1.0 (the variable's self-correlation) instead of a 2-d (1, 1) matrix. The downstream `_matrix_to_optional_list` then died with `TypeError: iteration over a 0-d array`. Surfaces in real use when `get_correlation_matrix` is called with a single name (or where several names are requested but only one resolves — the failure mode in the autonomous test pass). Fix: `_corrcoef` now promotes a 0-d numpy result to a (1, 1) array before returning. The downstream JSON envelope correctly serialises the trivial `[[1.0]]` matrix.

### Tests

404 unit tests pass (+2 in `test_mrservice.py::TestCorrcoefHelper` covering single-row and multi-row cases — sentinel against the 0-d regression).

## [0.3.0-alpha.25] — 2026-05-22

Two bugs surfaced by the autonomous end-to-end test pass against a model with extensive text labels.

### Fixed

- **Bug #26 — audit's VOSE-002 / VOSE-005 / VOSE-004 / VOSE-010 detectors false-positive on cell-ref-form VoseInput / VoseOutput wrappers.** The detectors used a regex (`_VOSE_INPUT_RE`) that only matched the string-literal form `VoseInput("Name")`. Workbooks using the cell-reference form `VoseInput(B20)` — which is what most real ModelRisk models use — triggered "not wrapped" warnings on every distribution cell. Same root cause as bug #13 (alpha.14), but the audit didn't get the parser-based check at the time. Fix: detectors now use `extract_vose_first_arg()` (the same function the scanner uses), so cell-ref-form wrappers are recognised consistently. Live test against the NPV workbook: audit findings dropped from 34 (all false-positive VOSE-002/005) to 0 on the same model.
- **Bug #27 — `_classify_cell` and downstream tools misclassified text cells as formulas.** xlwings' `Range.Formula` accessor returns the cell's text content even for non-formula cells — so a cell holding the label `"Total Revenue"` came back with `formula="Total Revenue"`, which the prior check `if formula:` flagged as a formula. Two consequences: `formula_cell_count` was inflated, and `find_hard_coded_inputs` returned `[]` on any model with text labels (every cell got bucketed as "formula", no numeric inputs were candidates). The "convert this Excel model to ModelRisk" workflow was silently broken on exactly the workbooks where it's most useful. Fix: a cell counts as a formula only when its `.Formula` starts with `=`. Applied at three sites: `_classify_cell`, `ModelRiskBridge.get_workbook_summary`, `ModelRiskBridge.find_hard_coded_inputs`.

### Tests

402 unit tests pass. The bug-#26 regression is covered by the live E2E pass (the broken audit was the symptom — same workbook now reports 0 false positives). Bug-#27 likewise: `find_hard_coded_inputs` against the non-MR workbook now returns candidate numeric cells instead of empty.

## [0.3.0-alpha.24] — 2026-05-22

Bug surfaced by the autonomous end-to-end test pass: `save_workbook_as` was renaming the open workbook in place instead of saving a copy.

### Fixed

- **Bug #25 — `save_workbook_as` renamed the live workbook.** Prior versions used `book.save(path)`, which xlwings translates to `Workbook.SaveAs(path)`. `SaveAs` doesn't save a copy — it renames the open workbook to the new path and rebinds it in Excel's books collection. Subsequent tool calls referencing the original workbook name then failed with "Workbook 'X.xlsx' is not open" because Excel only knew the new name. Not the contract callers expect from a "save as" operation in an MCP context where downstream tools chain after a save.
- Fix: use `book.api.SaveCopyAs(path)` directly via COM. SaveCopyAs writes the file without touching the open workbook's identity — the original stays open under its original name; the saved copy is an independent file on disk.
- `overwrite=True` now also `unlink()`s any pre-existing target first, since `SaveCopyAs` refuses to overwrite (whereas the old `SaveAs` happily clobbered).

### Tests

402 unit tests pass (+3 in `test_excel_bridge.py::TestSaveWorkbookAsUsesSaveCopyAs`): the API actually called is `SaveCopyAs` (regression sentinel asserting `book.save` was NOT called), overwrite=True clears the target first, overwrite=False refuses with `CellReferenceError`.

## [0.3.0-alpha.23] — 2026-05-22

Corporate-grade polish pass on the report charts. Live screenshot review showed the data was correct but the visual default-Excel-blue, unformatted axis numbers ("-156508.3276"), single-series legends floating off to the side, and stats-table overlap looked unfinished. This release moves the reports closer to "screenshot-and-paste-into-a-deck" quality.

### Polished

- **Centralised chart palette.** New module-level constants `_COLOR_CHART_PRIMARY` (steel blue, matched to the title band), `_COLOR_CHART_LINE` (burnt orange), `_COLOR_BAR_POSITIVE` (forest green), `_COLOR_BAR_NEGATIVE` (brick red), plus axis text + gridline tones. Both report builders now share one identity.
- **Histogram chart**:
  - Bars: steel-blue, no outlines (`Format.Line.Visible = False`).
  - Cumulative line: burnt-orange, 2.25pt weight, no markers (clean monotonic curve).
  - X-axis: tick labels now thousands-separated via `'#,##0;(#,##0);-'` — bin centres render as `-156,508` instead of `-156508.3276`. Smaller font, gray colour.
  - Primary Y-axis: integer counts; subtle gridlines.
  - Secondary Y-axis: 0% format on the cumulative line.
  - No legend (two-series chart with self-evident roles via title + colour).
  - Soft gray border around the chart area.
- **Tornado chart**:
  - Bars colour-coded by sign: positive correlations green, negative red — at-a-glance signal of "this driver helps" vs "this driver hurts".
  - No legend (single series, colour-coded directly per point).
  - X-axis: `0.00` format on the correlation values.
  - No major gridlines on the category axis; subtle ones on the value axis.
  - Soft gray border.
- **Layout fix**: `STATS_TABLE_TOP` bumped from 26 → 32. Taller charts (240pt) in alpha.20 had pushed the chart band to end around row 30, which overlapped the stats table. Row 32 leaves a full row of margin between chart bottom and table header.

### Tests

399 unit tests pass. 1 stats-table position test updated (B26 → B32).

### Why corporate styling, not just default

Decision-maker reports get screenshotted into decks, pasted into emails, printed for board meetings. Default Excel colours and unformatted axis labels read as "generated, not designed" — which lowers trust in the analysis sitting next to them. The cost of styling here is a one-time write of helper functions (`_style_chart_axes`, `_style_chart_frame`); the value carries across every report the LLM ever builds.

## [0.3.0-alpha.22] — 2026-05-22

Layout consistency fix: alpha.20 polished `build_executive_report`'s layout (narrow gutters at A/M, content in B–L) but I missed applying the same change to `build_drivers_report`. Live screenshot review caught the asymmetry — exec sheet had the polished gutter pattern, drivers sheet still had labels and tornado flush against column A.

### Fixed

- **DriversReportBuilder now uses the same B-shifted layout as ExecutiveReportBuilder.** Title band B:L (was A:J), KEY FINDINGS at B4 (was A4), bullets in B (was A), tornado chart shifted right by ~16pt to align with column B, driver-ranking table moved from G:J to H:K (sits after the mid-gutter at column G), HOW TO READ THIS CHART and RECOMMENDED ACTIONS narrative sections shifted to B with merge ranges expanded to L. Recommendations rows put labels in B and values in C (was A/B); the value cell merges B-merged-to-L. Now both reports look like siblings.

### Tests

399 unit tests pass. 6 drivers-report tests updated to assert against the new column positions (B-shifted findings, H-K driver table, C-shifted recommendations).

## [0.3.0-alpha.21] — 2026-05-22

Hotfix for a regression introduced by alpha.16's `_ModelRiskReports` helper sheet: the second run of `build_executive_report` (and `build_drivers_report`) failed with `Move method of Worksheet class failed` on real Excel.

### Fixed

- **Bug #24 — adding a sheet after a very-hidden sheet fails.** Both report builders anchored the new sheet via `book.sheets.add(name, after=book.sheets[-1])`. The trailing sheet became `_ModelRiskReports` (xlSheetVeryHidden) after the first report build, and Excel refuses to position a new sheet "after" a very-hidden anchor — COM raises `Move method of Worksheet class failed`. Manifested only on the second `build_*_report` call within a session.
- Fix: new `_last_visible_sheet(book)` helper walks `book.sheets` and picks the last sheet whose `Visible = -1` (xlSheetVisible). Both report builders + the helper-sheet creator now anchor against that.

### Tests

399 unit tests pass. The fake-Excel sheet class doesn't model the `Visible` attribute precisely, but the production fix is small and the call-site change is purely about which sheet object gets passed to `after=`. Integration test against real Excel is the regression sentinel.

## [0.3.0-alpha.20] — 2026-05-22

Polish pass on `build_executive_report` after a live screenshot review against the `NPV_of_a_capital_investment` workbook. Two real issues found: the histogram chart was rendering with completely wrong semantics, and the layout was visually cramped with column A pulling double duty as label-holder and edge.

### Fixed

- **Histogram chart was inverted (#18b).** Prior versions called `SetSourceData` on the full 3-column block `[Bin, Count, Cumulative %]`, which made Excel turn all three columns into data series — visible result was "Bin" values plotted as random-height blue bars (the user's screenshot showed bars at 50k, 100k, 150k etc.), "Count" plotted as a red line, and Cumulative invisible. Fix: bind to ONLY the Count + Cumulative columns, then explicitly assign `SeriesCollection(i).XValues` on both series to the Bin column. Result: proper frequency-histogram bars with bin centres on the X axis, cumulative-% line overlay on a secondary Y axis. The chart now looks like what the report description says.

### Polished

- **Column A is now a narrow gutter (width 2).** Previously content started flush against the left edge with column A holding both labels AND being the page edge. Now the layout has narrow gutters at A and M, with content in B–L. Same change applied to the title band merge range (B:L instead of A:J), headline numbers (MEAN now at B6 instead of A6), stats table (Output at B26 instead of A26), and callouts (`•  ...` at B instead of A).
- **Stats-table CV column no longer overflows.** Bumped width to 16 (was 14 implicit) so values like `1.296` render in full instead of `####`.
- **Alternating row tint** on the stats table for readability when there are multiple outputs.
- **High-CV cells now bold** in addition to coloured, so they survive print-to-PDF where colour fidelity drops.
- **Chart sizes bumped** to 400×240 (histogram) and 360×240 (tornado) — the original 380×220 / 340×220 felt small relative to the title band. Charts shifted right by ~16pt to align with the new column-B content start.

### Why this matters

The report is the primary user-facing deliverable. A broken chart isn't "a bug to fix later" — it's the LLM lying to the user about what got built (`chart_count: 2` while the chart was visually wrong). Same goes for the cramped layout: a stakeholder sees the report and forms an opinion about modelrisk-mcp from that single screenshot. Worth getting right.

### Tests

399 unit tests pass. 4 existing tests updated to assert against the new column positions (B-shifted).

## [0.3.0-alpha.19] — 2026-05-22

Fixes the bug-#23 lookup-after-samples regression discovered while end-to-end testing alpha.18 against a real workbook: `get_sensitivity_ranking` returned empty on the first call after `run_simulation`, then worked on the second identical call. The diagnostic trace was unambiguous — the output looked up fine, its samples loaded, then every input lookup against the same handle returned None.

### Fixed

- **MRLIB_GetModelData poisons subsequent MRLIB_GetModelVarID calls on the same handle.** Resolution: every reader that interleaves name lookups with sample fetches now resolves ALL var_ids first, THEN pulls samples. Applies to:
  - `ResultsReader.get_sensitivity_ranking` — was failing on the first call after a fresh simulation (output looked up, output samples loaded, all inputs then refused to resolve). Now: output lookup → all input lookups → all sample fetches → ranking.
  - `ResultsReader.get_simulation_results` — same risk on multi-output calls. Same fix.
  - `ResultsReader.get_correlation_matrix` — same risk on multi-name correlation requests. Same fix.
  - `ResultsReader.get_samples` (single name, single fetch — no change needed).

This is a *contract* finding about MRService.dll: the call sequence within one open handle must be all `GetModelVarID` calls first, then all `GetModelData` calls. Inverting them or interleaving is unsafe. Worth flagging upstream to the ModelRisk SDK team — and worth knowing for any future readers that touch the same surface.

### Why this matters end-to-end

Without alpha.19, the user's first sensitivity-ranking call after a sim returned silently empty. The LLM would tell them "no drivers detected" — completely wrong on a model that clearly has Spearman correlations up to +0.72. After alpha.19 the first call works correctly. Verified live against the `NPV_of_a_capital_investment complete.xlsx` workbook: 6 driver entries returned, top driver Market growth (r = +0.72), bottom three in noise territory.

### Tests

399 unit tests still pass — the bug only manifests against the real DLL, so the regression test is the integration smoke run.

## [0.3.0-alpha.18] — 2026-05-22

Targeted experiment for the empty-`.vmrs` blocker surfaced by alpha.17's post-condition verification. The bridge correctly detected that `VoseStartSimulCustom12 + VoseGetDataSZ12` was producing `.vmrs` files with zero registered outputs — sim ran, file existed, but no variable metadata. Ribbon-driven simulations on the same workbook worked fine, suggesting the ribbon path threads an option the headless XLL path skipped.

### Changed

- **Pre-populate `output_names` into the XLL command payload.** Previously `SimulationOptions.output_names` defaulted to `()` based on a C++ header comment that said "empty → all outputs". Real-world testing showed that interpretation was wrong — sims completed but the `.vmrs` didn't register any outputs unless the names were enumerated explicitly. alpha.18 changes the bridge to populate this list from `list_outputs(workbook)` before invoking the XLL command, threading the result through `SimulationController.run_simulation(output_names=...)` into the payload's `[CntNames]:N` + `[name0]:Profit` + … entries. The ribbon path presumably does this implicitly during its setup phase; we now mimic that explicitly.

### If this works

The empty-`.vmrs` symptom goes away and downstream readers find the outputs. Post-condition verification (added in alpha.17) becomes the test: if it stops firing on workbooks where it fired before, the hypothesis is confirmed.

### If it doesn't

The asymmetry is elsewhere (variable-registration timing, session handle threading, save-finalisation phase). Next step would be to compare what the ribbon does on the C++ side that this codepath skips — likely needs a diff against ModelRiskAtl's `IModelRiskSimulation::StartSimulation` entry vs. the XLL command handler.

### Tests

399 unit tests pass (+1: `test_run_simulation_passes_voseoutput_names_to_xll` confirms the bridge populates the names from the workbook scan).

## [0.3.0-alpha.17] — 2026-05-22

Full sweep against the running bug list — the biggest correctness release since the v0.3 pivot. Tackles every still-broken item: the response-envelope cross-cutting fix (#1, #2, validates #15), `run_simulation` false-positive reliability (#20), and the workbook-recovery tool (#21), plus a CI guard so the envelope category can't regress.

### Fixed

- **Envelope sweep across every list-returning MCP tool (#1, #2, validates #15).** FastMCP serialises a bare `list[T]` return as one MCP content block per element — which makes the LLM see N concatenated objects instead of a single array. Symptoms ranged from "list_modelrisk_outputs returns a single record" (#1) to "list_vmrs_variables returns concatenated JSON objects" (#2) to "get_samples wraps each float in a text-block dict" (#15, fixed in alpha.14). Fixed all of them in one pass by wrapping every list-typed response in a dict envelope with a semantic noun key: `list_open_workbooks` → `{workbooks, count}`, `list_modelrisk_inputs` → `{inputs, count}`, `list_modelrisk_outputs` → `{outputs, count}`, `list_distributions` → `{distributions, count}`, `get_simulation_results` → `{results, count}`, `find_hard_coded_inputs` → `{candidates, count}`, `list_vmrs_variables` → `{variables, count}`, `read_vmrs` → `{results, count}`, `propose_distributions_for_inputs` → `{proposals, count}`, `discover_inputs` → `{candidates, count}`. New CI guard test (`test_no_tool_returns_bare_list`) scans every tool module and fails if any uses `-> list[`, so the next instance of this category gets caught before merge.
- **`run_simulation` no longer reports false-positive success (#20).** Previously the tool returned `samples: 10000` and a valid `.vmrs` path even when ModelRisk's post-simulation phase crashed silently — leaving the `.vmrs` without registered output metadata and every downstream reader unable to find anything. The "samples" number was just echoing the input parameter, not measuring actual completion. Fix: post-condition verification. Before running, the bridge captures the list of expected VoseOutput names. After the simulation returns, it opens the produced `.vmrs` and confirms at least one expected output resolves to a `var_id`. If none do, raises `SimulationFailedError` with an actionable message ("the simulation's post-sim phase failed to register outputs; run `restore_deterministic_state` to recover").
- **`restore_deterministic_state` recovery tool (#21).** New MCP tool that recalculates the workbook to clear any VoseOutput cells stuck on per-iteration sample values from a previous run. Triggers `Application.CalculateFull` which re-evaluates every formula. Wired into the auto-recovery path on `run_simulation` post-condition failure — so the workbook is restored even if the user doesn't call the tool explicitly.
- **MRService.dll activation error message (#8).** The "no key supplied" error now lists both activation flavours (single-int64 via `MRSERVICE_ACTIVATION_KEY`, split-int64 via `MRSERVICE_ACTIVATION_KEY1/2`), explains what `MRSERVICE_DISABLE_BUNDLED_KEY` does, and points at the activation docs URL.

### Already fixed (acknowledged from the running bug list)

- **#4 (`wrap_with_output` refuses non-Vose formulas)** — current code passes `allow_overwrite_non_vose=True` and an existing test covers the Workflow-1-Step-6 pattern. The bug was real in an earlier alpha; the current implementation is correct.
- **#5 (`save_workbook_as` tool)** — registered as an MCP tool in `tools/building.py` since alpha.2.
- **#6 (`set_cell_formula` / guarded write)** — exposed as `write_formula` in `tools/building.py` since alpha.2.
- **#7 (`get_active_workbook` OneDrive)** — fallback path in `excel.py::get_active_workbook` already returns an empty path when xlwings' OneDrive resolution fails.
- **#12 (unsaved-workbook path)** — `_workbook_info` detects path strings missing any separator (the unsaved-workbook signature) and returns empty path.

### Obsoleted

- **#10 (`use_vba_helper_for_simulation` hangs)** and **#11 (`ensure_modelrisk_active` overfit)** — both refer to code that was removed in the v0.3 MRService.dll pivot. The new architecture doesn't have a VBA helper or an add-in-activation gate; simulations run directly via `Application.Run` on the XLL command surface and `.vmrs` reads go through MRService.

### New tools

- `restore_deterministic_state(workbook_name?)` — workbook recovery from the frozen-sample state.

### Tests

398 unit tests pass (+5: post-condition happy path, post-condition fails when no output registered, auto-restore fires on post-condition failure, restore tool with explicit workbook, restore tool defaults to active). Plus the new envelope CI guard.

## [0.3.0-alpha.16] — 2026-05-22

Two paired bugs in the report builders, both surfaced by the same end-user testing session — charts came out blank, and the staging data leaked onto the user-visible report sheet. Both `build_executive_report` and `build_drivers_report` are affected because they share the same chart-construction helpers.

### Fixed

- **Charts no longer render blank (bug #18).** Previously the flow was: create chart, call `chart.set_source_data(range)` (xlwings wrapper), then configure type / title / colours. On real Excel this looked correct in the COM trace but `SeriesCollection(1).Formula` came back empty — the bind silently dropped, Excel auto-filled a placeholder series during chart creation, and the chart went blank once that placeholder cleared. Fix: bind via the COM `chart_api.SetSourceData(Source=range.api, PlotBy=2)` call directly (skipping the xlwings wrapper that was where the regression lived), then probe `SeriesCollection(1).Formula` to verify the binding actually stuck. If the probe comes back empty the chart counts as failed (`chart_count` decrements) so the LLM doesn't mislead the user about how much of the report rendered.
- **Staging data no longer leaks onto the visible report sheet (bug #19).** Previously the histogram / tornado source ranges were written to columns M:Q of the report sheet itself and then hidden via `EntireColumn.Hidden`. Cosmetic until the user scrolled or printed, then it became visible noise. Fix: all staging data now goes on a workbook-scoped helper sheet `_ModelRiskReports` marked `xlSheetVeryHidden` (unreachable from the right-click "Unhide" menu). Block ownership: executive report uses columns A:C (histogram) and E:F (tornado); drivers report uses I:J (tornado). The two reports can coexist in one workbook without stomping. Each builder clears its own block before re-writing so re-running a report doesn't blend new + stale data.

### Tests

5 new tests in `test_reports.py`: no staging-data leak on the visible sheet, helper sheet created with the right headers, helper sheet has `Visible = xlSheetVeryHidden`, drivers + executive use distinct helper blocks, chart binding produces a non-empty `SeriesCollection(1).Formula` (with a negative test that proves an empty formula causes `chart_count` to decrement).

### Why this matters

These are the two bugs that made `build_executive_report` look broken to end users — "you said you built 2 charts but I see blank squares, and there's some weird data in column M". After alpha.16 the charts render, the report sheet has nothing on it but the intended content, and a binding regression on real Excel won't be silent — the chart count drops and the LLM can flag it.

## [0.3.0-alpha.15] — 2026-05-22

Fixes the second bug from yesterday's Claude Desktop testing session: `get_samples` (and every other read tool that resolves a variable name) no longer hangs forever when the workbook contains a VoseInput / VoseOutput name with characters that confuse MRService.dll's name lookup.

### Fixed

- **`get_samples` hang on names containing `?`, `(`, or `)` (bug #16).** `MRLIB_GetModelVarID` has been observed to spin indefinitely on names with those characters — looks like a wildcard/glob matcher that misinterprets them. Without a timeout, Claude Desktop's 4-minute hard limit was the only thing that stopped the request, and the user never got a useful error.
  - New `_call_with_timeout` helper in `bridge/mrservice.py` runs an individual ctypes call in a daemon thread with a wall-clock deadline. On expiry it raises `SimulationFailedError` with an actionable message that names the likely cause (`?` / `(` / `)` in the variable name) and the workaround (rename the input/output in the workbook).
  - New `VmrsHandle.lookup_var_id(name, *, timeout=None)` method moves the name-resolution logic from `ResultsReader._lookup_var_id` onto the handle where it belongs, and applies the timeout wrapper. Default budget is **8 seconds**; overridable via the `MRSERVICE_VARID_TIMEOUT_S` environment variable for environments where the SDK is unusually slow.
  - `ResultsReader._lookup_var_id` now delegates to `handle.lookup_var_id`. Every read tool that touches a variable by name benefits: `get_samples`, `get_simulation_results`, `get_correlation_matrix`, `get_sensitivity_ranking`, `list_vmrs_variables`, `read_vmrs`, `build_drivers_report`, `build_executive_report`.

### Why this matters

A workbook with a name like `"Conservatives get in? (1=yes)"` used to lock up every read tool against it. Post-fix, the user gets a clear error in ~8 seconds telling them which variable is the problem and what to do about it.

5 new tests cover the timeout wrapper itself plus the lookup-via-handle success / not-found / timeout / env-override paths.

## [0.3.0-alpha.14] — 2026-05-22

Two real bridge bugs found in a Claude Desktop end-user testing session. The first one is the critical fix — it unblocks roughly half of the read-side tool surface for workbooks that use the cell-reference name form (the most common pattern in production ModelRisk models).

### Fixed

- **CRITICAL: workbook scanner now recognises `VoseInput(Cell!Ref)` and `VoseOutput(Cell!Ref)` name forms.** Previous versions only matched the string-literal form `VoseInput("WidgetCost")`. But ModelRisk documents and most real workbooks use the cell-reference form — labels live in column headers and `VoseInput(A5)` / `VoseInput(Sheet1!A5)` pulls the name from there. The scanner missed every one of these, so `list_modelrisk_inputs`, `list_modelrisk_outputs`, `list_distributions`, and `get_workbook_summary` returned empty lists on these workbooks. That empty list then cascaded: `get_sensitivity_ranking`, `build_drivers_report`, `diagnose_workbook`, `audit_model` — all empty too.
  - New `bridge/name_parser.py` exposes `extract_vose_first_arg(formula, wrapper)` which classifies the first argument as `LiteralName` (string literal) or `CellRefName(sheet, cell)` (cell reference, with optional sheet qualifier). Supports same-sheet refs (`A5`), absolute refs (`$A$5`), sheet-qualified (`Sheet1!A5`), and quoted-sheet (`'Sheet with spaces'!B12`).
  - New `_resolve_vose_name` helper on `ModelRiskBridge` resolves a `CellRefName` to its actual name by reading the target cell via `ExcelBridge.get_cell`.
  - Wired through all four affected methods. 18 new tests in `test_name_parser.py` cover literal/cell-ref/unrecognised forms plus exact-wrapper-name matching.
  - Regression sentinel: `AB123` no longer parses as `sheet="A", col="B", row="123"` — the sheet-prefix branch of the regex now requires the `!` separator, so multi-letter columns are unambiguous.
- **`get_samples` no longer returns each sample wrapped in an MCP content-block dict.** FastMCP serialises bare `list[T]` returns by expanding each element into its own `{"type":"text","text":"<value>"}` content block, which made the response unusable to LLMs (they saw N opaque text blobs instead of one array of floats). Return type changed to a single dict envelope `{"output_name": ..., "sample_count": N, "samples": [...]}` so FastMCP sees one structured payload and serialises it once.

### Where it slots in

These two fixes together restore most of the read-side tool surface against real workbooks. Before alpha.14, a user with a typical cell-ref-named ModelRisk model would see "no inputs found" everywhere; after alpha.14 the scanner finds them, the sensitivity tools can rank them, and the report builders can describe them.

## [0.3.0-alpha.13] — 2026-05-22

A single-sheet uncertainty-drivers report — narrower than the executive dashboard, but with auto-generated narrative that explains what the tornado chart actually means.

### Added

- **`build_drivers_report(output_name, title?, subtitle?, sheet_name?, workbook_name?)`** — new MCP tool. Drops a single sheet onto the workbook with:
  - **Title band** — "Uncertainty Drivers — \<Output>" + run metadata.
  - **KEY FINDINGS** (3-5 auto-generated plain-English bullets) — names the dominant driver with direction language ("higher widget cost lowers NPV"), quantifies top-N variance share (rough Spearman r² approximation), flags concentration ("Risk is concentrated — most uncertainty from a small number of inputs") vs diffuse profiles, lists negligible inputs the decision-maker can safely deprioritise.
  - **Tornado chart** — full-prominence, sorted, axis-inverted (largest driver at top).
  - **Driver ranking table** — every input with correlation, |r|, approximate variance share. |r| cell coloured by strength tier (dark-red strong, orange medium, gray weak).
  - **HOW TO READ THIS CHART** panel — three short paragraphs explaining Spearman correlation, what bar magnitudes mean, and how to interpret variance share. Written for stakeholders who don't know what r means.
  - **RECOMMENDED ACTIONS** — three tiers: focus mitigation (|r| ≥ 0.4), monitor (0.2 ≤ |r| < 0.4), deprioritise (|r| < 0.2). Each tier lists the inputs that landed in it.
- `bridge/reports.py::DriversReportBuilder` — the new report builder. Shares the title-band styling + helper functions with `ExecutiveReportBuilder` (both live in the same module).
- Narrative helpers: `_strength_label`, `_concentration_label`, `_variance_share`, `_compose_findings`, `_compose_recommendations`, `_drivers_headline`. The narrative is deterministic from the data — same input always produces the same wording.

### Where it slots in

The executive report is the broader dashboard ("here's everything about the simulation"). The drivers report is the focused deliverable ("here's what matters and what to do about it"). Use case: a decision-maker asks "what should I worry about?" — `build_drivers_report` gives them a one-page answer naming the input, the magnitude, the direction, and the recommended action tier.

### Tests

8 new tests in `test_reports.py::TestDriversReportBuilder` covering title placement, findings name+direction generation, top-N variance share text, driver-table population, recommendations-tier assignment, empty-sensitivity edge case, concentration label classification, sheet replacement.

Tool count grows to 39.

## [0.3.0-alpha.12] — 2026-05-22

Adds the headline feature for end-user testing this week: a single-sheet executive-report builder that the LLM can produce in response to "create a report for a decision-maker."

### Added

- **`build_executive_report(primary_output, title?, subtitle?, secondary_outputs?, contingency_percentile=0.90, top_drivers=5, sheet_name="Executive_Report", workbook_name?)`** — new MCP tool that drops a one-sheet decision-maker dashboard onto the workbook. Idempotent; re-running replaces the sheet. The layout:
  - **Title band** (rows 1-2) — large, on a deep-navy background.
  - **Headline numbers** (rows 5-6) — mean / P5 / P50 / P90 (configurable) / stdev, big bold colored fonts. P5 in green (downside), P-high in red (upside risk), stdev color shifts amber/red as CV (coefficient of variation) rises.
  - **Charts band** (~rows 9-24) — side-by-side: histogram + cumulative overlay of the primary output's samples on the left, tornado mini of the top N sensitivity drivers on the right.
  - **Statistics table** (rows 26+) — full mean/stdev/P5/P50/P95/CV per output (primary first, then secondary). CV values colored by volatility tier.
  - **Risk callouts** (below stats) — auto-generated plain-English sentences for the decision-maker: "90% confident X lands between [A, B]", "Tail risk: PN is Y% above mean", "Primary driver: <input> (r = ±0.65) — focus mitigation here". Volatility callout fires only above the CV thresholds.
- `bridge/reports.py::ExecutiveReportBuilder` — the report orchestrator. Layout constants are class attributes so a redesign is one edit. All formatting wrapped in best-effort try/except so a COM hiccup on a single colour set doesn't tank the report.
- New chart variant: histogram + cumulative-overlay on a single chart object (column chart for counts, line on a secondary axis for cumulative %). Reuses the `TornadoChartWriter` pattern.

### Tests

348 unit tests pass (was 335): +13 new tests for the report builder covering title/subtitle placement, headline cells, secondary output rows in the stats table, callout generation from data, sheet-replacement idempotence, empty-samples / empty-sensitivity edge cases, high-CV volatility callouts, headline-summary string format, the `default_subtitle` helper, and MCP-tool passthrough.

Tool count grows to 38.

## [0.3.0-alpha.11] — 2026-05-21

Five real bugs found by a real end-user testing session. All shipped as fixes; one is a critical correctness bug that silently broke every list-scan against real Excel (unit tests passed because the fakes returned lists where real xlwings returns tuples).

### Fixed

- **CRITICAL: list-scan collapse against real Excel.** `ExcelBridge._as_2d` only accepted `list` for value normalisation, but xlwings on Windows returns `Range.formula` as a tuple of tuples (raw COM SAFEARRAY) — only the `.value` accessor wraps in lists. When a workbook had two or more cells, the formula payload arrived as a tuple, `_as_2d` treated it as a scalar, and the whole row's formulas were string-cast into one fake "cell". The regex `_VOSE_INPUT_RE` found the first match in that string and we yielded exactly one record instead of many — silently losing all but one input across `list_modelrisk_inputs`, `list_modelrisk_outputs`, `get_workbook_summary`, `find_hard_coded_inputs`, `audit_model`, and most importantly `run_simulation`'s input registration. Unit tests didn't catch this because the fake Excel returned lists. Fix: `_as_2d` accepts both `list` and `tuple` at every nesting level and normalises to lists. Pinned by 6 new tests in `test_excel_bridge.py::TestAs2dTupleHandling` including a regression test that mirrors the exact production failure mode.
- **`workbook_path` returns `""` for unsaved workbooks instead of the bare name.** `Workbook.FullName` returns just the workbook's name (`"Book3"`) when the workbook has never been saved. Previously we propagated that string as the `path` field, which misled downstream code that treated it as a filesystem location (a `.vmrs` save targeting `<path>/<book>.vmrs` would resolve to a relative path and land in the user's cwd). Now we detect missing path separators and report empty path explicitly.
- **`run_simulation` silently dropped `iterations` parameter.** Callers using the natural English term ("run 5000 iterations") rather than ModelRisk's UI term (`samples`) saw their argument silently ignored — the default 1000 ran instead. Now `iterations` is an explicit parameter alias for `samples`; both work. Loudly raises `ValueError` if both are passed with conflicting values (silent drops are exactly the class of bug we're fixing here).
- **`RunSimulationResult` no longer duplicates `samples` as `iterations`** in its response shape. Previously both fields appeared with identical values — confusing, and an attractive nuisance for callers passing `iterations` as input (which got silently dropped). The canonical name is now `samples` for both input and output.

### Notes on adjacent reports

Several other issues raised in the same session were actually already fixed in earlier alphas:

- `wrap_with_output` refusing to wrap non-Vose formulas: fixed in alpha.10.
- No `save_workbook_as` / no generic write tool: both shipped in alpha.10 as `save_workbook_as` and `write_formula`.
- OneDrive `get_active_workbook` hard-fail: fixed in alpha.4.
- MRService.dll activation error message clarity: moot since alpha.5 ships a bundled activation key.

Two complaints referenced tools that no longer exist in the v0.3 architecture:

- `use_vba_helper_for_simulation`: deleted in v0.3.0-alpha.1 when we pivoted from the ATL VBA-helper approach to the XLL command surface.
- `ensure_modelrisk_active` with its bitness-mismatch hypothesis: deleted in v0.3.0-alpha.1 along with the rest of the COM-Dispatch diagnostic apparatus.

If a fresh Claude session is still seeing these, it's pulling from training-data documentation of the older architecture, not from the live tool list. The MCP server's actual `tools/list` returns 37 tools, none of which match those names.

### Tests

333 unit tests pass (was 328): +6 covering the tuple-vs-list path through `_as_2d`, plus 1 for the unsaved-workbook path-degradation case.

## [0.3.0-alpha.10] — 2026-05-21

Two new building tools that fill the gap surfaced when Claude tried to build a Monte Carlo model from scratch end-to-end. Previously: no way to write a non-Vose formula (`=A1*B1`, `=SUM(...)`, `=IF(...)`) and no way to save the workbook to a path. The "build a tiny test model" prompt couldn't be completed without manual user steps.

### Added

- **`write_formula(workbook, sheet, cell, formula, allow_overwrite=False, dry_run=True)`** — single MCP tool for writing arbitrary formulas / literal values into a cell. Use for wiring inputs into outputs (`=A1*B1`), aggregations (`=SUM(B1:B10)`), conditional logic, or anything else not covered by the Vose-specific building tools. Safety: empty cells write freely; non-empty cells require `allow_overwrite=True` (protects both user-written formulas and prior Vose distributions). Defaults to `dry_run=True` like every other building tool. Adds a leading `=` automatically for formula-shaped input; numeric literals pass through unchanged.
- **`save_workbook_as(workbook, path, overwrite=False)`** — explicit-path save. Distinct from the user's Ctrl+S — the server still never calls `Workbook.Save()` implicitly. Only fires when the caller named a target file. Refuses to overwrite an existing file unless `overwrite=True`. Validates the target has an Excel extension. Returns the resolved absolute path that was written.

### Why this matters

Without these, the "build a model from scratch" workflow had a dead-end: Claude could generate distributions and wrap inputs/outputs around existing content, but couldn't put the existing content there to wrap. And it couldn't save the result so a future session could read the `.vmrs`. End users hit this on the very first "build me a test model" prompt — Claude correctly flagged the missing surface upfront rather than producing a partial result.

### Bridge changes

- `ExcelBridge.save_workbook_as(workbook, path, overwrite=False) -> str` — thin wrapper over xlwings' `book.save(path)` with the file-existence + extension safety checks. Raises `CellReferenceError` on refused overwrites or bad paths.

### Tests

Brings total to 327 unit tests (was 320 in alpha.9). +7 new building-tool tests covering:
- dry-run-by-default
- commit to empty cell
- leading `=` auto-prepend on formula-shaped input vs. numeric-literal pass-through
- refuse-to-overwrite-non-empty (default safety)
- `allow_overwrite=True` lets caller clobber
- the typical "wire-then-wrap" workflow (`write_formula` → `wrap_with_output`)
- `save_workbook_as` passthrough + overwrite flag wiring

Tool count grows to 37 (was 35).

## [0.3.0-alpha.9] — 2026-05-21

End-user install friction drops sharply: a single `modelrisk-mcp install` command now wires the server into every detected MCP client config, with backups and per-client dry-run friendly output. README also documents the zero-install `uvx` route for users who already have `uv` set up.

### Added

- New CLI subcommand: `modelrisk-mcp install`. Detects Claude Desktop (`%APPDATA%\Claude\claude_desktop_config.json`) and Claude Code (`~/.claude/settings.json`), backs up each existing config with a timestamped `.bak.` suffix, merges in the `modelrisk` server entry without clobbering other servers, and reports per-client status. Uses the absolute path to the installed `modelrisk-mcp` exe so the registration works even when `Scripts/` isn't on PATH for Claude's spawned subprocess. Flags: `--name` (custom server key for side-by-side dev/prod installs), `--force` (overwrite an existing entry under the same name).
- Reverse subcommand: `modelrisk-mcp uninstall`. Removes the entry idempotently — reports `skipped` if it isn't there to begin with.
- New module `src/modelrisk_mcp/install.py` holds the config-mangling logic; `__main__.py` provides the argparse glue. The legacy "no subcommand → run the server" behaviour is preserved (existing `claude_desktop_config.json` entries like `"command": "modelrisk-mcp"` keep working unchanged).
- README "Wire into Claude Desktop" section rewritten to present three options ordered by friction: `modelrisk-mcp install` (one command), `uvx modelrisk-mcp` (zero install if you have `uv`), and the hand-edit JSON snippet (last resort).

### Tests

320 unit tests pass (was 303): +17 for the install module covering create-on-missing, merge-into-existing, custom server names, force-overwrite, idempotent uninstall, malformed-JSON refusal, the CLI dispatch back to the install entry point, and the backward-compat "no subcommand defaults to serve" path.

## [0.3.0-alpha.8] — 2026-05-21

End-to-end integration test passed live for the first time, surfacing one operational caveat now documented.

### Verified

- All 7 integration tests in `tests/integration/test_e2e_run_simulation.py` pass against a real Excel + ModelRisk XLL + MRService.dll round-trip (18.85s total). The empirical moments of `Y = 2 * N(0, 1)` over 1000 iterations match the analytic moments inside the documented tolerance bands. All v0.3 read-path tools (`list_vmrs_variables`, `get_samples`, `diagnose_workbook`) work end-to-end.

### Documented

- **Launch order caveat in the README's "Known caveats" section.** Excel must be running interactively (Start menu / taskbar) before the MCP server tries to drive `run_simulation`. When Excel is launched programmatically by an automation client, ModelRisk's XLL skips part of its `xlAutoOpen` initialisation — XLL functions still register as worksheet UDFs (so cell formulas work), but XLL commands (`VoseStartSimulCustom12` etc.) never get added to Excel's `Application.Run` table, and the simulation pipeline depends on those commands. This is a ModelRisk XLL behaviour, not a bug in this server, but the launch-order requirement is now explicit. The integration test discovered this; the earlier real-workbook test by hand worked because Excel had been open interactively.

## [0.3.0-alpha.7] — 2026-05-21

Hotfix for the MCP Registry publish step that failed in 0.3.0-alpha.6.

### Fixed

- `server.json::description` now fits the MCP Registry's 100-character limit. 0.3.0-alpha.6 shipped to PyPI cleanly but the registry validator rejected the entry with `expected length <= 100` (the example in the SDK docs happened to be 67 chars so the limit wasn't visible). Trimmed to 94 chars.
- Adds the v0.3 integration test scaffold (`tests/integration/test_e2e_run_simulation.py`) that landed between the alpha.5 and alpha.6 tags — covered in the changelog now that it's published.

### Added (previously merged on `main`, just not part of an earlier tag)

- `tests/integration/test_e2e_run_simulation.py` — 7 gated tests that exercise the real Excel + ModelRisk XLL + MRService.dll round-trip via a programmatically-created 2-cell workbook. Asserts the empirical moments of `Y = 2 * N(0, 1)` match the analytic ones and that all the v0.3 read-path tools (`list_vmrs_variables`, `get_samples`, `diagnose_workbook`) work end-to-end.

## [0.3.0-alpha.6] — 2026-05-21

Registers the server with the official [MCP Registry](https://registry.modelcontextprotocol.io/) so users can discover it through the canonical channel (and aggregator clients can index it).

### Added

- `server.json` at repo root — MCP Registry metadata. Server name: `io.github.vosesoftware/modelrisk-mcp`. Declares the PyPI package as the canonical artifact for `transport: stdio` clients.
- `<!-- mcp-name: io.github.vosesoftware/modelrisk-mcp -->` marker in the README. The MCP Registry verifies ownership by fetching the PyPI package's README (which is the `long_description` baked into the wheel) and looking for this string. It's a Markdown comment so it stays invisible in rendered views.
- New release-pipeline job `publish-mcp-registry`. Runs after `publish-pypi` (so PyPI has the new README with the marker before the registry verifier looks for it), waits 60s for PyPI's CDN, then authenticates via GitHub OIDC and runs `mcp-publisher publish`. No tokens to manage — same trust model as PyPI trusted publishing.

### Changed

- `docs/community-submission.md` rewritten. The old draft targeted the `modelcontextprotocol/servers` README's "Community Servers" section, which has been retired (upstream now redirects all server submissions to the MCP Registry). New doc walks through the actual automated flow, ownership verification, and emergency-manual-publish procedure.

## [0.3.0-alpha.5] — 2026-05-21

First publish-ready release. Adds the tornado chart writer, fixes a real `discover_inputs` scoring bug, closes the security verification loop on the obfuscated activation key, and refreshes internal docs to match v0.3 architecture.

### Added

- `create_tornado_chart(output_name, workbook_name?, sheet_name?)` — renders a SensitivityRanking as a native Excel BarClustered chart on a new sheet (`Tornado_<output_name>` by default, truncated to Excel's 31-char limit). Sheet has a sorted data table (Input | Spearman correlation | |corr| sort key | Regression coefficient) plus the chart with inverted category axis so the largest-magnitude input is at the top — the tornado convention. Idempotent: existing sheets with the target name are replaced, so it's safe to re-run after each new simulation.
- `bridge/charts.py::TornadoChartWriter` — first member of the chart-writer family. Future siblings: RiskProfileChartWriter (cumulative + density), HistogramWriter, ScenarioComparisonWriter.
- `scripts/scan_exe_for_key.py` — paranoid scan of a built PyInstaller exe for every encoding of the plain activation key (ASCII decimal, UTF-16 LE wide string, little-endian int64 bytes, big-endian int64 bytes, 8-byte compact form, hex string both cases, and composite first4+last4 across printable runs). Exits non-zero on any hit; wired into `release.yml` as a release-blocker step before the PyPI upload so a regression in the obfuscation can't ship.

### Fixed

- `discover_inputs` no longer over-scores zero-valued cells. The `value not in (0, 1)` exclusion was guarding only the multiple-of-10 bonus; the multiple-of-100 and multiple-of-1000 bonuses still fired for `value=0` because `0 % n == 0`. Result: a cell holding 0 scored 2.0, identical to a cell holding 100 — flags tied with real scenario assumptions. Also added an explicit `not isinstance(value, bool)` guard so `False` cells don't take the same code path (Python's `isinstance(True, int)` is True).

### Verified

- v0.3.0-alpha.5 builds — wheel (~17 KB), sdist (~140 KB), and PyInstaller exe (~39 MB) — and the exe scan returns clean across every encoding tested. Even the obfuscated base85 blobs from `_keymat.py` aren't directly findable in the binary because PyInstaller compresses bundled Python into a PYZ archive; the source-level base85 strings become compiled `.pyc` bytecode constants. A reverse-engineer would need to extract the PYZ, decompress the bytecode, and reimplement the XOR decoder by hand. Casual `strings`-based extraction yields nothing actionable.

### Docs

- `docs/architecture.md` and `docs/com-surface.md` rewritten to match the v0.3 stack (MRService.dll via ctypes + XLL `Application.Run` for simulation kickoff; no ATL CoClass dispatch).

### Tests

303 unit tests pass (was 290 at alpha.4): +11 chart writer tests + 2 regression tests for the discover_inputs zero-value scoring fix.

## [0.3.0-alpha.4] — 2026-05-21

Four feature additions in one batch: read-path completeness, a one-call session-start tool, deterministic scenario sweeps, and 5 new audit rules. Tool count grows 30 → 34, audit rules 6 → 11. No breaking changes.

### Added

- `list_vmrs_variables(workbook_name?)` — enumerates VoseInput / VoseOutput names that exist in the active `.vmrs`. Workbook drives the candidate list (the SDK exposes no name-enumeration call against a `.vmrs` directly — `MRVarsGetModelVarsNames` takes a different ModelID and access-violates against `OpenSimulationModel`'s pointer).
- `get_samples(output_name, max_n=10000, workbook_name?)` — raw per-iteration sample array for one variable. Unlocks custom histograms, arbitrary percentiles, and any downstream analysis the aggregate-only `get_simulation_results` couldn't cover. Capped at 10 000 by default to stay under MCP-wire budget; configurable up to 1 000 000.
- `diagnose_workbook(workbook_name?)` — one-call session-start introspection. Returns Excel connection state, MRService activation, active workbook + sheets, input/output/distribution/formula counts, sibling `.vmrs` path + modification timestamp, audit-log location, and an `issues[]` list with human-readable strings flagging anything that would block downstream operations. Saves 4-5 individual tool calls per session. Short-circuits cleanly on Excel-not-reachable.
- `run_scenarios(sheet, cell, values, samples?, seed?, workbook_name?)` — sweep an input cell across deterministic values, running a full simulation at each, returning per-output P5/P50/P95/mean for every scenario. The cell's original formula is always restored afterwards, even when a scenario raises mid-sweep. `values` capped at 1-20 entries to prevent runaway compute.
- 5 new audit rules:
  - **VOSE-007** `risk_event_degenerate_probability` — `VoseRiskEvent` with literal probability of 0 or 1 (wrapper is degenerate).
  - **VOSE-008** `voseoutput_missing_name` — `VoseOutput()` with no name or empty-string name.
  - **VOSE-009** `duplicate_output_names` — same `VoseOutput("X")` declared on multiple cells.
  - **VOSE-010** `input_wrapper_without_distribution` — `VoseInput` wrapper but no distribution → input won't vary across iterations.
  - **VOSE-011** `high_volatility_normal_positive_mean` — `VoseNormal(mu, sigma)` with `mu > 0` and `sigma > mu/2` (~16% negative samples; lognormal probably wanted).
- `docs/authoring-audit-rules.md` — guide for adding new rules. Three-file pattern (YAML + detector + tests), worked example using VOSE-007, when-not-to-add-a-rule discussion.

### Bridge / schema changes

- `bridge/results.py::ResultsReader.list_variables()` and `.get_samples()` — new methods consumed by the new MCP tools.
- `bridge/modelrisk.py::ModelRiskBridge.run_scenarios()` — orchestrates Excel write + sim + read with guaranteed cell-state restoration in a `try/finally`.
- New schemas in `schemas/results.py`: `ScenarioOutcome`, `ScenarioRun`, `ScenarioSweepResult`, `VmrsVariableEntry`.

### Tests

290 unit tests pass (was 256 at start of alpha.4 work). 34 new tests across the four features: 5 for read-path tools, 6 for `diagnose_workbook`, 6 for scenario sweeps, 17 for the new audit detectors (positive + negative cases per rule, plus threshold-boundary tests for VOSE-011).

### Notes

False-positive avoidance pattern adopted in the new audit rules: numeric-threshold rules skip cell-reference args (e.g. `VoseRiskEvent(B5, ...)` is NOT flagged because we can't statically know what's in B5). Static analysis should be conservative when it lacks information.

## [0.3.0-alpha.3] — 2026-05-21

End-user-frictionless activation + obfuscation so the bundled MRService.dll key isn't grep-able from the wheel, plus 50 new MCP-wrapper tests that were missing since the v0.3 refactor.

### Added

- Bundled MRService.dll activation. The DLL needs per-process activation to open `.vmrs` files; we now ship a Vose Software-owned key as a fallback in `bridge/mrservice.py::_activate()` so the read path works out of the box. Precedence: `MRSERVICE_ACTIVATION_KEY` env var → `MRSERVICE_ACTIVATION_KEY1`/`_2` env vars → bundled key. `MRSERVICE_DISABLE_BUNDLED_KEY=1` opts out.
- `bridge/_keymat.py` + `scripts/encode_activation_key.py` — XOR-encoded + base85-stored key material so the literal int never appears in shipped source or `strings` output. Algorithm is public (Kerckhoffs); only the value is secret.
- `tests/unit/test_tools_{reading,workflows,simulation,restore}_mocked.py` (50 tests) — MCP-tool wrappers were previously untested; these guard against bridge method-rename / kwarg-shape regressions that only surface at end-user runtime.
- CI guard `test_no_literal_in_package_sources` recursively scans every shipped `.py` for the decoded key's decimal form; fails if anyone re-inlines it.

### Verified

- Wheel and sdist scans for the plain activation key return zero hits.
- Bundled key activates the real MRService.dll end-to-end (smoke test in the bridge).

### Tests

256 unit tests pass (was 206). 50 new MCP-wrapper tests + 3 reworked activation tests.

## [0.3.0-alpha.2] — 2026-05-20

Adds programmatic simulation triggering via the XLL command surface (no ATL COM dispatch needed), graceful OneDrive path handling, and the missing `read_vmrs` / `set_active_vmrs` tools.

### Added

- `bridge/simulation.py` — `SimulationController` drives runs via plain `Application.Run("VoseStartSimulCustom12", options)` + `Application.Run("VoseGetDataSZ12", session, path)`, replicating exactly what the ATL does internally (`ModelRiskAtl/ModelRisk_Simulate.cpp:102`, `ModelRiskAtl/ModelRiskSimulationResults.cpp:1196`). No ATL Dispatch required, so cross-bitness automation works.
- `SimulationOptions` dataclass reproduces `CSimulationOptions::PackToStringList` (`SimulationObj.cpp:94`) — `[Key]:Value` lines in exact field order.
- New MCP tool `run_simulation(workbook_name?, samples=1000, seed=1, save_to?)` — blocks until the sim finishes, saves the `.vmrs` next to the workbook by default, and auto-pins it as the active results source.
- New MCP tools `set_active_vmrs(path)` and `read_vmrs(path, output_names?)` — formerly referenced in error messages but not registered.

### Changed

- `ExcelBridge.get_active_workbook()` / `list_workbooks()` degrade gracefully when xlwings can't resolve OneDrive paths (`ONEDRIVE_COMMERCIAL_WIN` not set). Returns name-only `WorkbookInfo` with empty path; downstream name-based COM operations still work.
- `SimulationController` falls back to saving in the user's Desktop folder when the workbook's folder can't be resolved (the OneDrive case).
- `ModelRiskBridge.run_simulation()` calls `ResultsReader.set_active_vmrs()` automatically after a successful run, so the next `get_simulation_results` call doesn't need a sibling-discovery step.

### Tests

201 unit tests pass (was 182). New coverage: OneDrive path fallback (5), SimulationController options packing + Application.Run shape + session-name format + failure modes (14).

## [0.2.0-alpha.2] — 2026-05-20

Auto-activates the ModelRisk add-in inside Excel before reporting the COM surface unreachable. Closes the "modelrisk_loaded: false even though ModelRisk is installed" footgun.

### Added

- `ExcelBridge.list_com_addins() / list_excel_addins()` — enumerate Excel's COMAddIns and AddIns collections as plain dicts.
- `ExcelBridge.enable_com_addin(predicate) / enable_excel_addin(predicate)` — flip `.Connect=True` / `.Installed=True` on matching entries, return the names of those actually flipped. Idempotent; no-op on already-on entries.
- `ModelRiskBridge.ensure_modelrisk_active()` — scans both collections, enables any entry whose description / progid / name mentions ModelRisk or Vose, retries Dispatch, returns a diagnostic dict (`com_addins_enabled`, `excel_addins_enabled`, `com_addins_seen`, `excel_addins_seen`, `modelrisk_dispatchable`).
- New MCP tool `ensure_modelrisk_active` — explicit invocation for debugging "COM unreachable" reports.
- Simulation tools (`set_simulation_settings`, `run_simulation`) now call `ensure_modelrisk_active` transparently before touching COM. The LLM no longer needs to ask the user to manually load the add-in.

### Changed

- `ModelRiskBridge.is_modelrisk_loaded()` now attempts auto-activation if the first Dispatch fails. Returns True iff Dispatch works after activation.

### Notes

If auto-activation can't find a ModelRisk add-in to flip on, the diagnostic surfaces every COM and Excel add-in it *did* see — useful for ruling out bitness mismatches and broken installs.

## [0.2.0-alpha.1] — 2026-05-20

HTTP transport — unblocks Claude for Excel and other remote MCP clients that can't spawn local subprocesses.

### Added

- `--transport={stdio,streamable-http,sse}` CLI flag in `__main__.py`. stdio remains the default for backwards compatibility.
- `--host`, `--port`, `--mount-path` flags for HTTP transports. Defaults to `127.0.0.1:8000`.
- Bearer-token middleware (`http_auth.py`) — required on non-loopback HTTP binds, recommended even on loopback. Constant-time comparison via `hmac.compare_digest`. Token via `--token` or `MODELRISK_MCP_TOKEN` env var.
- `docs/claude-for-excel.md` — wiring guide covering the Office.js sandbox / COM-bridge architectural story.
- README section on HTTP transport with the strategic framing.
- `tests/unit/test_cli.py` + `tests/unit/test_http_auth.py` — 14 new tests covering parser defaults, middleware reject/accept paths, RFC 6750 case-insensitivity of the Bearer keyword.

### Changed

- Spec §2.2 — "Hosting the MCP server remotely" no longer a non-goal.
- README front-page table picks up a "Works with Claude for Excel" implication (no UI change needed — Claude for Excel was always in the compatible-clients list).

## [0.1.0] — 2026-05-20

Public v0.1 release. Repository goes public at this tag.

### Added

- Public-ready README leading with the strategic narrative (open, agentic, methodology-aware), feature comparison table, and full Safety by design section.
- `docs/demo-script.md` — beat-by-beat recording guide for the README demo GIF.
- `docs/community-submission.md` — drafted PR text for the modelcontextprotocol/servers directory.
- Spec doc updated to v1.4 with the per-phase completion record.

### Changed

- `Development Status` classifier moved to `5 - Production/Stable`.

## [0.1.0-rc.1] — 2026-05-20

First release candidate. Feature surface matches the v0.1 spec. PyPI publication path and standalone `.exe` build pipeline both verified locally.

### Added

- Final `pyproject.toml` metadata: full PyPI classifiers, project URLs (Changelog, Documentation, Vose Software), wheel `force-include` for the four packaged YAML/JSON data files, sdist include list.
- `.github/workflows/release.yml` — tag-triggered build of wheel, sdist, and standalone Windows `.exe`; uploads to PyPI via trusted publishing and to GitHub Releases.
- `CHANGELOG.md` — this file. Phase-by-phase history.
- Updated `modelrisk_mcp.spec` — PyInstaller bundle now declares hidden imports for every tool/resource/prompt module and ships the catalogue + rule YAML files alongside the `.exe`.

## [0.1.0-beta.1] — 2026-05-20 (commit `d4431bc`)

Phase 5 — workflows, resources, prompts.

### Added

- 4 workflow tools (`tools/workflows.py`): `propose_distributions_for_inputs`, `discover_inputs`, `audit_model`, `generate_executive_summary`.
- Audit engine (`audit/engine.py`, `audit/rules.py`) with 6 detectors mapped to rules in `data/audit_rules.yaml`.
- 7 resources under `modelrisk://`: `functions`, `functions/{name}`, `distributions`, `methodology`, `workbook/current`, `workbook/current/sheet/{name}`, `audit-rules`.
- 5 prompt templates (slash commands): `/build-risk-model`, `/audit-model`, `/interpret-results`, `/add-uncertainty`, `/import-legacy-model`.
- `data/distributions.yaml` — keyword-matched distribution selection guide.

## [0.1.0-alpha.3] — 2026-05-20 (commit `464c5f3`)

Phase 4 — simulation control.

### Added

- `bridge/simulation.py` with `SimulationController` and `SimulationCom` Protocol abstracting `ModelRiskSimulation` + `ModelRiskSimulationSettings`.
- 4 simulation tools: `set_simulation_settings`, `run_simulation`, `stop_simulation` (registered stub raising `SimulationNotAvailableError`), `get_simulation_status` (in-process polling fallback).
- `seed=42` auto-flips `use_fixed_seed=True` when the user doesn't pass it explicitly.

## [0.1.0-alpha.2] — 2026-05-20 (commit `2cc7e6b`)

Phase 3 — building tools and `restore_cell`.

### Added

- 10 building tools (`tools/building.py`): `insert_distribution`, `wrap_with_input`, `wrap_with_output`, `replace_constant_with_distribution`, `fit_distribution_to_data`, `create_aggregate_mc`, `create_risk_event`, `create_time_series`, `create_copula`, `set_named_range`. Every tool defaults `dry_run=True`.
- `restore_cell` MCP tool (`tools/restore.py`) — restores any cell from the audit log.
- `ModelRiskBridge.safe_write_cell` — every cell write goes through this and gets: writer-mutex acquisition, non-Vose-formula refusal via tokenised detection, audit-log append.
- `ExcelBridge.set_named_range` + `ExcelBridge.undo` for the Excel-undo-stack acceptance test.

## [0.1.0-alpha.1] — 2026-05-20 (commit `793a17c`)

Phase 2 — reading tools.

### Added

- 12 reading tools (`tools/reading.py`): `list_open_workbooks`, `get_active_workbook`, `get_workbook_summary`, `list_modelrisk_inputs`, `list_modelrisk_outputs`, `list_distributions`, `get_cell`, `read_range`, `get_simulation_results`, `get_correlation_matrix`, `get_sensitivity_ranking`, `find_hard_coded_inputs`.
- `bridge/results.py` — `ResultsReader` wrapping `ISimVariable.GetMean/Percentile/StDev/...`. Pearson + Spearman correlation and Spearman-based tornado computed in numpy from `GetSamples()`.
- `docs/installation.md`, `docs/claude-desktop.md`, `docs/claude-code.md`.

## [0.0.2] — 2026-05-19 (commit `c4fe8b5`)

Phase 1 — bridge layer, safety mechanisms, function catalogue.

### Added

- 1417-entry function catalogue (`data/functions.json`) extracted from the ModelRisk IDL + XLL header.
- `bridge/{catalogue,excel,formulas,modelrisk,progids}.py` — three-layer architecture.
- `safety.py` — tokenised `is_vose_formula` detector, bulk-write guard, audit-log appender, Windows-named-mutex `WriterMutex`.
- Pydantic v2 schemas for every tool input/output (`schemas/{workbook,distributions,results}.py`).
- `data/optional_overrides.yaml` — flips `VoseModPERT.gamma` (and similar) to optional with their documented defaults.
- Gated integration test infrastructure (`tests/integration/`) — skips cleanly when Excel isn't running.

## [0.0.1] — 2026-05-19 (commit `3741386`)

Phase 0 — scaffold.

### Added

- Empty FastMCP server that responds to `initialize` and returns an empty `tools/list`.
- `pyproject.toml`, `LICENSE` (MIT), `.gitignore`, `.python-version`.
- `.github/workflows/ci.yml` — ruff + mypy + pytest on Windows × Python 3.11/3.12/3.13.
- `scripts/spike_com_surface.py` — probes ModelRisk's COM surface and writes `docs/com-surface.md`.
- `modelrisk_mcp.spec` — PyInstaller spec used for the Phase 0 smoke build.
