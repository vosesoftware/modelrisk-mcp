# ModelRisk MCP

<!-- mcp-name: io.github.vosesoftware/modelrisk-mcp -->

**An open Model Context Protocol server for [Vose Software's ModelRisk](https://www.vosesoftware.com).**

Use it with Claude Desktop, Claude Code, Cursor, Zed, or any MCP-compliant client to read, build, fit, and run Monte Carlo risk models in Excel from a conversation.

> ModelRisk MCP is an open MCP server on the standard Anthropic Model Context Protocol. No proprietary layer, no lock-in. The 1417-entry function catalogue, the Vose methodology principles, and the audit rule set are all included in the package — and editable.

**Stable: `0.4.1`** — on the mcp 2.0 SDK; the server starts ModelRisk itself when no Excel is running (auto-launch + add-in activation); programmatic simulation via `run_simulation` (XLL command surface, no fragile COM dispatch); `.vmrs` results read via the official ModelRisk SDK, with an automatic in-session fallback when MRService.dll is absent; activation ships bundled so no environment configuration is required. 59 tools across reading, building, fitting, simulation, scenario-sweep, restore, charting, audit, **a quantitative analysis & decision layer** (distribution calculator, tail risk, fit-and-rank, fast aggregates, correlation matrix, copula & time-series fitting, distribution comparison, backtest, uncertainty decomposition, reverse stress test), and workflow surfaces.

---

## Purpose

**Monte Carlo risk modelling is the right way to reason about an uncertain future — and almost nobody does it.**

The mathematics has been settled for decades: when your inputs are uncertain, you don't reason with a single "best guess" number, you reason with the *distribution of outcomes*. A point estimate ("Q3 revenue will be $4.2M") hides exactly the thing a decision-maker needs to know — how wrong could it be, how bad is the downside, which assumption is driving the risk. ModelRisk has given Excel users a rigorous, validated engine for this for years.

The barrier was never the maths. It was the **friction**:

- The discipline carries a real skill curve — 1,400+ functions, the right distribution family for each input, the methodology traps (fit without parameter uncertainty, correlations ignored, risk events modelled as `p × impact`) that quietly produce confident-but-wrong answers.
- Setting a model up by hand is slow and fiddly — wrapping inputs, naming outputs, wiring copulas and time-series, formatting the report.
- So under deadline pressure, most teams fall back to a single-point spreadsheet and a gut-feel range. The rigorous tool sits unused.

**This server removes that friction.** Large language models can now drive domain tools through the open Model Context Protocol. ModelRisk MCP puts the whole ModelRisk surface — build, fit, simulate, audit, interpret — behind a conversation. You describe the problem in plain language; Claude proposes the right distributions, wires the structure, runs the simulation, reads the tail, and writes the report. The expertise moves into the loop; the friction drops to a sentence.

It is built on five deliberate principles:

1. **Open, not locked-in.** A server on the standard Anthropic MCP — works with Claude Desktop, Claude Code, Claude for Excel, Cursor, Zed, and any compliant client. MIT-licensed. The function catalogue, methodology, and audit rules ship in the package and are editable. No proprietary connector, no vendor cage.
2. **Methodology-grounded, not just mechanically capable.** The server is opinionated about *correct* practice. Fits default to `uncertainty=TRUE`; risk events use the bimodal `VoseRiskEvent`; the 13-rule audit encodes the mistakes Vose practitioners have seen across decades of consulting. It won't just do what you ask — it'll steer you toward what's right.
3. **Local-only, no telemetry.** Everything runs on your machine against your Excel and your ModelRisk install. No data leaves your computer; the activation key is bundled and offline.
4. **Excel stays the model.** No re-platforming, no shadow tooling. The workbook *is* the model — versionable in Git, openable by anyone with Excel + ModelRisk, reproducible by re-running one tool call.
5. **Safe by design.** Every write previews first (`dry_run=True`), lands in Excel's undo stack, and is logged. The server can modify your workbook — and does so under nine layered safeguards (see [Safety by design](#safety-by-design)).

The goal is simple: make defensible, quantitative risk analysis something you reach for by default — because it's now no harder than asking.

---

## What this does

This server turns Claude (or any MCP client) into a methodology-aware co-pilot for ModelRisk. It can:

- **Build** new Monte Carlo models from a description — insert distributions, fit families to data, build aggregates, copulas, time-series, risk events.
- **Run** simulations from the conversation. `run_simulation` triggers the same XLL command the ribbon "Simulate" button uses, blocks until the run finishes, saves a `.vmrs` next to the workbook, and auto-pins it as the results source.
- **Read** model structure and per-iteration results — inputs, outputs, percentiles, correlation matrices, tornado rankings — directly from `.vmrs` files via ModelRisk's official SDK (MRService.dll). No COM dispatch fragility.
- **Audit** a workbook against Vose's methodology rules and propose fixes.
- **Interpret** results into structured executive summaries with contingency analysis.

Every formula written to Excel is validated against the ModelRisk function catalogue first — there's no path to a hallucinated function name reaching your workbook.

**See the [user manual](docs/user-manual.md)** for a walkthrough of the nine things you can do, a realistic end-to-end example, and what the server explicitly does and doesn't do. New to Monte Carlo or to the ModelRisk MCP toolchain? Start with the **[15-minute quick-start tutorial](docs/quick-start.md)**, then pick a **[walk-through scenario](docs/scenarios.md)** matching your problem (budgets, data fitting, loss aggregation, correlated inputs, stress tests, model audits); unfamiliar with a term, see the **[glossary](docs/glossary.md)**.

---

## Feature comparison

| Capability | ModelRisk MCP | Closed alternatives |
|---|:---:|:---:|
| Read model structure (inputs, outputs, distributions) | ✓ | ✓ |
| Read simulation results, percentiles, sensitivity | ✓ | ✓ |
| Insert distributions into cells | ✓ | — |
| Fit distributions from data | ✓ | — |
| Build aggregate (frequency × severity) models | ✓ | — |
| Build copulas / correlated inputs | ✓ | — |
| Build time-series stochastic processes | ✓ | — |
| Run simulations from the conversation | ✓ | — |
| Audit model for common methodology mistakes | ✓ | — |
| Works with Claude Desktop / Code / Cursor / Zed / any MCP client | ✓ | — |
| Open source, MIT licensed | ✓ | — |
| Local-only, no telemetry | ✓ | varies |
| Default-safe (dry-run preview before every write) | ✓ | n/a |

---

## Install

### Prerequisites

- Windows 10 or 11, 64-bit
- Excel 2019 or newer with the ModelRisk add-in installed and loaded — **the latest ModelRisk from [vosesoftware.com](https://www.vosesoftware.com) is recommended**: the current release includes MRService.dll, which the results-reading tools need (installers up to 9.1.x did not ship it)
- One of:
  - Python 3.11+ (recommended via [`uv`](https://docs.astral.sh/uv/))
  - Or the standalone `modelrisk-mcp.exe` from the [latest release](https://github.com/vosesoftware/modelrisk-mcp/releases/latest) — no Python knowledge required

**Activation (results reading):** None required — MRService.dll (the SDK that reads `.vmrs` files) ships with the current ModelRisk release and is activated automatically by a bundled offline key. If the server reports "MRService.dll not found", update ModelRisk to the latest version from [vosesoftware.com](https://www.vosesoftware.com) — installers up to 9.1.x did not include the DLL. Two env overrides exist for non-standard setups (both only affect *reading* `.vmrs` results; simulations run regardless):

- `MRSERVICE_DLL_PATH` — full path to an MRService.dll if it isn't in the standard ModelRisk folder. Must be version 7.3.2.1 or newer (the bundled key covers 7.3.2.1–9.2.2.1); older copies shipped with other Vose products (e.g. Tamara's 7.1.x) are too old and are rejected with a version diagnosis. Its sibling DLLs resolve from that folder, so point at it in place.
- `MRSERVICE_ACTIVATION_KEY` — your own activation key, to override the bundled one for any MRService version.

### Recommended: auto-wire from PyPI

```powershell
pip install modelrisk-mcp
modelrisk-mcp install      # adds the server to every MCP app it finds
```

Then **restart the apps it lists**. `modelrisk-mcp install` finds Claude Desktop (including the Microsoft Store build), Claude Code, Cursor, VS Code, Windsurf, Gemini CLI and LM Studio, backs up each settings file and adds the server beside anything already there. `modelrisk-mcp status` shows where each app stands; `--client cursor` (repeatable) limits install or uninstall to the apps you name. (Excel + a licensed ModelRisk add-in are still required — that's the product itself.)

### One-click `.mcpb` (Claude Desktop Extension)

Download **`modelrisk-mcp.mcpb`** from the [latest release](https://github.com/vosesoftware/modelrisk-mcp/releases/latest), open it (Claude Desktop → **Settings → Extensions → Install Extension…**), then restart Claude. The bundle ships the server as a standalone exe — no Python, no config edit.

> ⚠️ **Known issue (Claude Desktop, latest Windows MSIX builds, e.g. 1.12603.x):** the Extensions installer can silently do nothing when you pick a `.mcpb` — no error, no install. This is a **Claude Desktop installer bug** (it fails before logging), not a problem with the bundle (it validates with `mcpb` and installs fine once the client is fixed). **Until Anthropic patches it, use the `pip install` + `modelrisk-mcp install` path above** — it bypasses the Extensions UI entirely.

### From PyPI

```powershell
pip install modelrisk-mcp
```

### From source

```powershell
git clone https://github.com/vosesoftware/modelrisk-mcp
cd modelrisk-mcp
uv sync
uv run python -m modelrisk_mcp     # speaks MCP over stdio
```

### Standalone `.exe`

Download `modelrisk-mcp.exe` from [Releases](https://github.com/vosesoftware/modelrisk-mcp/releases/latest), drop it where you want to keep it, then run `.\modelrisk-mcp.exe install` from a terminal in that folder. It registers that exe, at that path, with Claude Desktop and Claude Code (0.4.1 and later; older exes wrote an entry that could not start). To wire it by hand instead, see [docs/claude-desktop.md](docs/claude-desktop.md).

---

## Wire into Claude Desktop

Three options, simplest first.

### One-command auto-wire (recommended)

```powershell
pip install modelrisk-mcp
modelrisk-mcp install
```

`modelrisk-mcp install` detects every installed MCP client (Claude Desktop, Claude Code), backs up its existing config, and adds the `modelrisk` server entry — preserving any other servers you already have configured. Output looks like:

```
  + Claude Desktop   added    C:\Users\you\AppData\Roaming\Claude\claude_desktop_config.json
      Registered 'modelrisk' -> {'command': 'C:\\...\\Scripts\\modelrisk-mcp.exe'}
      backup: ...claude_desktop_config.json.bak.20260521-153000

Restart Claude Desktop / Claude Code to pick up the new server.
```

To undo: `modelrisk-mcp uninstall`. To register a second instance with a different name (e.g. dev and prod side-by-side): `modelrisk-mcp install --name=modelrisk-dev`.

### Zero-install via `uvx` (if you already use `uv`)

If you have [`uv`](https://docs.astral.sh/uv/) installed, you can skip the `pip install` step entirely. Just add to `%APPDATA%\Claude\claude_desktop_config.json` directly:

```json
{
  "mcpServers": {
    "modelrisk": {
      "command": "uvx",
      "args": ["modelrisk-mcp"]
    }
  }
}
```

`uvx` downloads `modelrisk-mcp` into an ephemeral cache on first run and updates automatically when new versions hit PyPI.

### Manual JSON edit (if you must)

Open `%APPDATA%\Claude\claude_desktop_config.json` and add the entry by hand:

```json
{
  "mcpServers": {
    "modelrisk": {
      "command": "C:/path/to/modelrisk-mcp.exe"
    }
  }
}
```

Use the absolute path to the `.exe` you downloaded from the [latest release](https://github.com/vosesoftware/modelrisk-mcp/releases/latest), or `"command": "python", "args": ["-m", "modelrisk_mcp"]` if you `pip install`ed.

---

After any of the three, **restart Claude Desktop** so it spawns the MCP server subprocess. The ModelRisk tools appear under the connections icon. Full guide: [docs/claude-desktop.md](docs/claude-desktop.md). Claude Code setup: [docs/claude-code.md](docs/claude-code.md).

---

## Claude for Excel — not currently reachable (read this before trying)

**Current Claude for Excel builds cannot connect to this server, and no local
workaround exists.** Earlier versions of this README described a
Settings → Connectors flow; that was wrong for shipping builds and has been
withdrawn (verified against a 2026-07 build; independently confirmed by a
field report).

What we verified:

- The Claude for Excel add-in has **no local Connectors panel**. It resolves
  connectors from your **claude.ai account**, and claude.ai custom connectors
  are **remote MCP servers fetched by Anthropic's infrastructure** — not by
  your device. A loopback URL (`http://127.0.0.1:…`) is therefore unreachable
  **by construction**, and the dialog's OAuth fields don't match this server's
  bearer-token auth.
- Local **stdio** servers (the Claude Desktop kind) never appear in the
  claude.ai / Claude Desktop "Connectors" panel. Their absence there is
  normal, not a failed install — Claude Desktop lists them separately under
  local MCP servers.
- **Do not expose this server through a public tunnel to work around this.**
  It can write formulas into, simulate, and save your workbooks; putting that
  behind a single bearer token on a public URL is a materially different risk
  posture than the loopback bind these docs describe, and we don't support it.

**Use Claude Desktop or Claude Code instead** (local stdio — the supported,
tested path; `modelrisk-mcp install` wires both). The HTTP transport below
remains useful for LAN/self-hosted setups with MCP clients that genuinely
run on your machine:

```powershell
$env:MODELRISK_MCP_TOKEN = [Guid]::NewGuid().ToString("N") * 2
modelrisk-mcp --transport=streamable-http --port=8000 --token=$env:MODELRISK_MCP_TOKEN
```

Details and the full analysis: [docs/claude-for-excel.md](docs/claude-for-excel.md).

---

## First conversation

Open a workbook in Excel that has at least one Vose function — even a single `=VoseNormal(0,1)`. Then in Claude:

> Summarise the active workbook's risk model — inputs, outputs, distributions.

Or jump straight into building:

> /build-risk-model

This walks through 9 steps, from identifying outputs through running the simulation and interpreting results. See [the slash-command catalogue](src/modelrisk_mcp/prompts) for the other workflows.

### Demo gallery

Prefer a guided tour? **[examples/demos](examples/demos)** contains fifteen hands-on demo
workbooks — one per feature area, each a simple real-life use case (store-expansion NPV,
bakery demand fitting, correlated construction costs, insurance aggregate loss, reverse
stress testing, board reporting…). Every workbook's README sheet gives the exact prompts
to type to Claude, step by step. Start with `00_START_HERE.xlsx`.

---

## Safety by design

The server can both read *and* modify your workbook — that's the central differentiator. We make that safe with nine layered mechanisms (spec §11):

1. **`dry_run=True` is the default** on every building tool. Claude must explicitly pass `dry_run=False` to commit. Previewing comes free; a forgotten flag becomes a preview, never an overwrite.
2. **Every write lands in Excel's native undo stack.** `Ctrl+Z` works exactly as you'd expect.
3. **Bulk-write guard.** Tools writing >50 cells in one call require explicit `confirm_bulk=True`. Time-series and copula tools — which write contiguous ranges by design — are exempt.
4. **No automatic saves.** The server never calls `Workbook.Save()`. You control `Ctrl+S`.
5. **No overwriting non-Vose formulas.** A formula-tokenised detector (not a substring check) refuses to overwrite a cell whose existing formula uses non-Vose functions. The one tool explicitly allowed to do this is `replace_constant_with_distribution`, by design.
6. **Audit log** of every write in `%LOCALAPPDATA%\VoseSoftware\modelrisk-mcp\writes.log` — timestamp, cell, before/after formulas, before value. JSONL, append-only.
7. **Read-only mode.** Launch with `--read-only` to disable every building/simulation tool.
8. **Single-writer mutex.** Two MCP server instances can't drive the same Excel concurrently — the second instance raises `ConcurrentWriterError` on any building tool call.
9. **Restore from audit log.** The `restore_cell` tool reads writes.log and rewrites the pre-write formula — even after Excel's undo stack has been cleared.

---

## What's inside

- **57 tools** — 12 reading, 14 building (incl. `create_aggregate` for FFT / Panjer / MC frequency-severity methods), 5 simulation (incl. `run_simulation`, `run_scenarios`, `get_samples`, `restore_cell`, `restore_deterministic_state`), 11 analysis (`compute_distribution`, `fit_and_rank_distributions`, `fit_time_series`, `fit_copula_to_data`, `reverse_stress_test`, `get_tail_risk`, `fit_tail`, `compute_correlation_matrix`, `compare_distributions`, `backtest_output`, `decompose_uncertainty`), 12 workflow / reporting (incl. `plan_risk_model`, `build_model_from_brief`, `fit_all_data_and_wire`, `audit_model`, `diagnose_workbook`, `create_tornado_chart`, `create_histogram_chart`, `create_cdf_chart`, `build_drivers_report`, `build_executive_report`, `generate_executive_summary`, `save_workbook_as`), 3 VMRS (`read_vmrs`, `set_active_vmrs`, `list_vmrs_variables`)
- **5 resources** — `modelrisk://functions`, `modelrisk://distributions`, `modelrisk://methodology`, `modelrisk://workbook/current`, `modelrisk://audit-rules`
- **5 slash-command prompts** — `/build-risk-model`, `/audit-model`, `/interpret-results`, `/add-uncertainty`, `/import-legacy-model`
- **1417-entry function catalogue** extracted directly from the ModelRisk IDL + XLL header
- **17 audit rules** — 13 Monte-Carlo-methodology (VOSE-001 … VOSE-013) + 4 spreadsheet-integrity (SS-001 … SS-004) — with editable severity in `data/audit_rules.yaml`; add your own with `docs/authoring-audit-rules.md`
- **Methodology-grounded distribution selection guide** in `data/distributions.yaml`

---

## Methodology

The server is opinionated about Monte Carlo methodology — fetch `modelrisk://methodology` from any MCP client to read the 8 core principles. Highlights:

- Every uncertain input is a distribution. Treating a noisy input as deterministic understates total uncertainty by exactly the amount it could swing.
- Distribution fits use `uncertainty=TRUE`. Carry parameter uncertainty through the simulation; don't pretend the best-fit parameters are exact.
- Risk events use `VoseRiskEvent`, not `probability * impact`. The bimodal nature matters.
- Correlated inputs use copulas. Independent inputs that are actually correlated produce artificially tight outputs.

---

## Architecture

Three internal layers plus two external integration paths:

```
┌──────────────────────────────────┐
│  MCP client                      │
│  (Claude Desktop, Code, etc.)    │
└────────────────┬─────────────────┘
                 │ JSON-RPC / stdio (or HTTP)
                 ▼
┌──────────────────────────────────┐
│  MCP server layer (tools,        │
│   prompts)                       │
├──────────────────────────────────┤
│  ModelRiskBridge (domain)        │
│  + SimulationController          │
│  + ResultsReader                 │
├──────────────────────────────────┤
│  ExcelBridge      MrServiceBridge│
│  (xlwings)        (ctypes)       │
└──────┬───────────────────┬───────┘
       │ Application.Run   │ MRLIB_*
       │ + cell I/O        │ (read .vmrs)
       ▼                   ▼
┌──────────────┐   ┌──────────────────┐
│ Excel +      │   │  MRService.dll   │
│ ModelRisk XLL│   │  (SDK)           │
└──────────────┘   └──────────────────┘
```

Two integration paths, each chosen for what it does best:

- **Builds + simulation trigger** → Excel COM via xlwings, plus `Application.Run("VoseStartSimulCustom12", …)` for the simulation kickoff. Mirrors what the ModelRisk ATL does internally; bypasses the fragile ATL CoClass Dispatch surface that doesn't expose IDispatch.
- **Results read** → MRService.dll directly via ctypes. Vose's official SDK opens `.vmrs` files, returns sample arrays, computes statistics. No COM round-trips per output; per-iteration sample arrays available for arbitrary downstream analysis.

More: [docs/architecture.md](docs/architecture.md), [docs/com-surface.md](docs/com-surface.md).

## Known caveats

- **Starting ModelRisk is automatic (since 0.3.2).** If no Excel is running when a tool is called, the server starts an attachable Excel and registers the ModelRisk XLL into it, so Vose functions resolve and simulations run. If Excel is already up but the add-in isn't live (e.g. ModelRisk's "Start with Excel" is off), the server auto-activates it before simulating, or returns a clear instruction if it can't. Disable auto-launch with `MODELRISK_AUTO_LAUNCH=0` if you'd rather manage Excel yourself. (Earlier versions required you to open Excel + ModelRisk by hand first and could fail with an opaque "macro may not be available".)
- **OneDrive-hosted workbooks**: xlwings can fail to resolve the workbook's full path without `ONEDRIVE_COMMERCIAL_WIN` set. The bridge degrades gracefully — name-based operations still work, and `run_simulation` defaults the `.vmrs` save location to the user's Desktop when the workbook folder can't be resolved.
- **Active simulation results**: `get_simulation_results` reads from the `.vmrs` file produced by the most recent `run_simulation` call, or the most recent sibling `.vmrs` next to the workbook. Use `set_active_vmrs(path)` or `read_vmrs(path)` to point at a specific file.
- **Results reading on older ModelRisk installs**: re-opening a saved `.vmrs` file requires MRService.dll, which ships with the current ModelRisk release but was not included in installers up to 9.1.x — update ModelRisk from [vosesoftware.com](https://www.vosesoftware.com) for the full experience. Without the DLL, the results tools **fall back automatically to reading the live Excel session** (via temporary `VoseSimValue` scratch formulas, sample-identical to the `.vmrs`, `source: "in-session"` in the response). The fallback is session-bound: it serves the most recent run in the Excel session that ran it, can't re-open `.vmrs` files after Excel closes, and is refused in `--read-only` mode (it needs scratch-cell writes; the DLL path is the write-free reader).

---

## License

MIT. See [LICENSE](LICENSE).

---

## Documentation

| Doc | What it's for |
|---|---|
| [Quick-start tutorial](docs/quick-start.md) | 15 minutes, zero to your first simulation |
| [Walk-through scenarios](docs/scenarios.md) | Six problem-shaped recipes — budgets, data fitting, loss aggregation, correlated inputs, stress tests, audits |
| [User manual](docs/user-manual.md) | The nine things you can do, in depth; what the server does and doesn't do |
| [Methodology](docs/methodology.md) | The principles behind every model, each tied to the audit rule that enforces it |
| [Knowledge base](docs/knowledge-base.md) | Risk-analysis guidance distilled from the ModelRisk Help — served to the LLM at build time (`modelrisk://knowledge`) |
| [Distribution selection](docs/distribution-selection.md) | Which distribution for which uncertain quantity |
| [Modeling patterns](docs/modeling-patterns.md) | Techniques — frequency-severity, correlation, common random numbers, time-series choice |
| [Glossary](docs/glossary.md) | Monte Carlo + MCP vocabulary for non-statisticians |
| [Installation](docs/installation.md) | Full install + activation detail |
| [Claude Desktop setup](docs/claude-desktop.md) · [Claude Code setup](docs/claude-code.md) · [Claude for Excel setup](docs/claude-for-excel.md) | Per-client wiring |
| [Chart style guide](docs/chart-style-guide.md) | The native-Excel-chart styling ruleset the reports follow |
| [Authoring audit rules](docs/authoring-audit-rules.md) | Extend the 13-rule audit set with your own |
| [Architecture](docs/architecture.md) | Internal layers + the two integration paths |

---

## Links

- **Vose Software**: <https://www.vosesoftware.com>
- **ModelRisk product page**: <https://www.vosesoftware.com/products/modelrisk/>
- **Source**: <https://github.com/vosesoftware/modelrisk-mcp>
- **Releases**: <https://github.com/vosesoftware/modelrisk-mcp/releases>
- **Issues**: <https://github.com/vosesoftware/modelrisk-mcp/issues>
- **Changelog**: [CHANGELOG.md](CHANGELOG.md)
- **MCP spec**: <https://modelcontextprotocol.io>
