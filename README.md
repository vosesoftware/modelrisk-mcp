# ModelRisk MCP

<!-- mcp-name: io.github.vosesoftware/modelrisk-mcp -->

**An open Model Context Protocol server for [Vose Software's ModelRisk](https://www.vosesoftware.com).**

Use it with Claude Desktop, Claude Code, Cursor, Zed, or any MCP-compliant client to read, build, fit, and run Monte Carlo risk models in Excel from a conversation.

> ModelRisk MCP is an open MCP server on the standard Anthropic Model Context Protocol. No proprietary layer, no lock-in. The 1417-entry function catalogue, the Vose methodology principles, and the audit rule set are all included in the package — and editable.

**Stable: `0.3.0`** — programmatic simulation via the `run_simulation` tool wired end-to-end (XLL command surface, no fragile COM dispatch); `.vmrs` results read via the official ModelRisk SDK; activation ships bundled so no environment configuration is required. 40 tools across reading, building, simulation, scenario-sweep, restore, charting, audit, and workflow surfaces.

---

## What this does

This server turns Claude (or any MCP client) into a methodology-aware co-pilot for ModelRisk. It can:

- **Build** new Monte Carlo models from a description — insert distributions, fit families to data, build aggregates, copulas, time-series, risk events.
- **Run** simulations from the conversation. `run_simulation` triggers the same XLL command the ribbon "Simulate" button uses, blocks until the run finishes, saves a `.vmrs` next to the workbook, and auto-pins it as the results source.
- **Read** model structure and per-iteration results — inputs, outputs, percentiles, correlation matrices, tornado rankings — directly from `.vmrs` files via ModelRisk's official SDK (MRService.dll). No COM dispatch fragility.
- **Audit** a workbook against Vose's methodology rules and propose fixes.
- **Interpret** results into structured executive summaries with contingency analysis.

Every formula written to Excel is validated against the ModelRisk function catalogue first — there's no path to a hallucinated function name reaching your workbook.

**See the [user manual](docs/user-manual.md)** for a walkthrough of the eight things you can do, a realistic end-to-end example, and what the server explicitly does and doesn't do. New to Monte Carlo or to the ModelRisk MCP toolchain? Start with the **[15-minute quick-start tutorial](docs/quick-start.md)**; unfamiliar with a term, see the **[glossary](docs/glossary.md)**.

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
- Excel 2019 or newer with the ModelRisk add-in installed and loaded
- One of:
  - Python 3.11+ (recommended via [`uv`](https://docs.astral.sh/uv/))
  - Or the standalone `modelrisk-mcp.exe` from the [latest release](https://github.com/vosesoftware/modelrisk-mcp/releases/latest) — no Python knowledge required

**Activation:** None required. MRService.dll (the SDK that reads `.vmrs` files) is activated automatically by a bundled offline key. Set `MRSERVICE_ACTIVATION_KEY` only if you want to override the default with your own.

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

Download `modelrisk-mcp.exe` from [Releases](https://github.com/vosesoftware/modelrisk-mcp/releases/latest), drop it anywhere on disk, and point Claude Desktop at it. See [docs/claude-desktop.md](docs/claude-desktop.md).

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

## Wire into Claude for Excel (HTTP transport)

Claude for Excel runs inside an Office.js sandbox and can't spawn subprocesses, so it talks to MCP servers over HTTP. Start the server in HTTP mode:

```powershell
$env:MODELRISK_MCP_TOKEN = [Guid]::NewGuid().ToString("N") * 2
modelrisk-mcp --transport=streamable-http --port=8000 --token=$env:MODELRISK_MCP_TOKEN
```

Then in Claude for Excel: Settings → Connectors → Add MCP server, URL `http://127.0.0.1:8000/mcp`, paste the token. Full guide: [docs/claude-for-excel.md](docs/claude-for-excel.md).

**Why this is interesting:** Claude for Excel's sandbox can't reach Excel's COM surface or the ModelRisk ribbon on its own. ModelRisk MCP runs outside the sandbox and bridges that gap — Claude for Excel can do things via this server it structurally can't do otherwise (run simulations, dispatch ModelRisk COM, write distributions through the safety pipeline).

---

## First conversation

Open a workbook in Excel that has at least one Vose function — even a single `=VoseNormal(0,1)`. Then in Claude:

> Summarise the active workbook's risk model — inputs, outputs, distributions.

Or jump straight into building:

> /build-risk-model

This walks through 9 steps, from identifying outputs through running the simulation and interpreting results. See [the slash-command catalogue](src/modelrisk_mcp/prompts) for the other workflows.

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

- **40 tools** — 12 reading, 13 building, 5 simulation (incl. `run_simulation`, `run_scenarios`, `get_samples`, `restore_cell`, `restore_deterministic_state`), 7 workflow / reporting (incl. `audit_model`, `diagnose_workbook`, `create_tornado_chart`, `build_drivers_report`, `build_executive_report`, `generate_executive_summary`, `save_workbook_as`), 3 VMRS (`read_vmrs`, `set_active_vmrs`, `list_vmrs_variables`)
- **5 resources** — `modelrisk://functions`, `modelrisk://distributions`, `modelrisk://methodology`, `modelrisk://workbook/current`, `modelrisk://audit-rules`
- **5 slash-command prompts** — `/build-risk-model`, `/audit-model`, `/interpret-results`, `/add-uncertainty`, `/import-legacy-model`
- **1417-entry function catalogue** extracted directly from the ModelRisk IDL + XLL header
- **13 audit rules** (VOSE-001 through VOSE-013) with editable severity in `data/audit_rules.yaml`; add your own with `docs/authoring-audit-rules.md`
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
│  FastMCP layer (tools, resources,│
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

- **Launch order: Excel first, MCP server second.** Excel must already be running interactively (started from the Start menu, taskbar, or by double-clicking an `.xlsx`) before the MCP server tries to drive `run_simulation`. When Excel is launched programmatically by an automation client, ModelRisk's XLL skips part of its `xlAutoOpen` initialisation — the XLL functions still register as worksheet UDFs, but the XLL commands (`VoseStartSimulCustom12` etc.) never get added to Excel's `Application.Run` table. The simulation pipeline depends on those commands. If `run_simulation` fails with "macro may not be available", restart Excel by hand and try again.
- **OneDrive-hosted workbooks**: xlwings can fail to resolve the workbook's full path without `ONEDRIVE_COMMERCIAL_WIN` set. The bridge degrades gracefully — name-based operations still work, and `run_simulation` defaults the `.vmrs` save location to the user's Desktop when the workbook folder can't be resolved.
- **Active simulation results**: `get_simulation_results` reads from the `.vmrs` file produced by the most recent `run_simulation` call, or the most recent sibling `.vmrs` next to the workbook. Use `set_active_vmrs(path)` or `read_vmrs(path)` to point at a specific file.

---

## License

MIT. See [LICENSE](LICENSE).

---

## Links

- **Vose Software**: <https://www.vosesoftware.com>
- **ModelRisk product page**: <https://www.vosesoftware.com/products/modelrisk/>
- **Source**: <https://github.com/vosesoftware/modelrisk-mcp>
- **Releases**: <https://github.com/vosesoftware/modelrisk-mcp/releases>
- **Issues**: <https://github.com/vosesoftware/modelrisk-mcp/issues>
- **Changelog**: [CHANGELOG.md](CHANGELOG.md)
- **MCP spec**: <https://modelcontextprotocol.io>
