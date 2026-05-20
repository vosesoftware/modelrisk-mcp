# Changelog

All notable changes to ModelRisk MCP. Follows [Keep a Changelog](https://keepachangelog.com/) and [Semantic Versioning](https://semver.org/).

## [Unreleased]

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
