# Installation

## Prerequisites

- **Windows 10 or 11, 64-bit.** ModelRisk is Windows-only.
- **Microsoft Excel 2019 or newer** with the ModelRisk add-in installed and loaded. Open Excel and confirm the **ModelRisk** ribbon tab is present before continuing.
- One of:
  - **Python 3.11, 3.12, or 3.13** (recommended via [`uv`](https://docs.astral.sh/uv/) so it manages Python for you), or
  - The standalone `modelrisk-mcp.exe` from the [latest release](https://github.com/vosesoftware/modelrisk-mcp/releases/latest) — no Python knowledge required.

**Activation.** None required. MRService.dll (the SDK that reads `.vmrs` files) is activated automatically by a bundled offline key. Set `MRSERVICE_ACTIVATION_KEY` only if you want to override the default with your own.

## From PyPI

```powershell
pip install modelrisk-mcp
modelrisk-mcp install        # auto-wires Claude Desktop + Claude Code
```

`modelrisk-mcp install` detects every installed MCP client, backs up its existing config, and adds the `modelrisk` server entry — preserving any other servers you already have. To undo: `modelrisk-mcp uninstall`.

## Standalone Windows executable

Download `modelrisk-mcp.exe` from the [latest release](https://github.com/vosesoftware/modelrisk-mcp/releases/latest) and put it anywhere on disk. It bundles Python and every dependency. Point Claude Desktop at it manually (see [docs/claude-desktop.md](claude-desktop.md)) — the `install` CLI is only available with the Python package.

## Zero-install via `uvx`

If you already use [`uv`](https://docs.astral.sh/uv/), skip the install entirely and reference `uvx modelrisk-mcp` from your MCP client config. `uvx` downloads the package into an ephemeral cache on first run and updates automatically when new versions hit PyPI. See the README for the JSON snippet.

## From source

```powershell
git clone https://github.com/vosesoftware/modelrisk-mcp
cd modelrisk-mcp
uv sync
uv run python -m modelrisk_mcp     # speaks MCP over stdio
uv run mcp dev src/modelrisk_mcp/__main__.py   # for the MCP Inspector
```

## Verifying the install

```powershell
modelrisk-mcp --help
```

If you see the transport / port / token options, the entry point is wired correctly. Beyond that, the fastest way to verify the full server works is to point MCP Inspector at it (`uv run mcp dev src/modelrisk_mcp/__main__.py`) and check that `tools/list` returns 44 tools — `list_open_workbooks`, `get_active_workbook`, `audit_model`, `run_simulation`, etc.

If `run_simulation` fails with "macro may not be available", make sure Excel was started interactively (taskbar / Start menu / double-click `.xlsx`) *before* the MCP server tried to drive it. The ModelRisk XLL skips part of its initialisation when Excel is launched programmatically; see the **Known caveats** section in the README.
