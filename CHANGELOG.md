# Changelog

All notable changes to ModelRisk MCP. Follows [Keep a Changelog](https://keepachangelog.com/) and [Semantic Versioning](https://semver.org/).

## [Unreleased]

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
