# Changelog

All notable changes to ModelRisk MCP. Follows [Keep a Changelog](https://keepachangelog.com/) and [Semantic Versioning](https://semver.org/).

## [Unreleased]

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
