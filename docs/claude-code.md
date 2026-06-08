# Wiring ModelRisk MCP into Claude Code

Claude Code (the CLI / IDE extension) connects to MCP servers via the same `claude_desktop_config.json` mechanism plus a project-local `.mcp.json` if you want repo-scoped configuration.

## One-command setup (recommended)

```powershell
pip install modelrisk-mcp
modelrisk-mcp install
```

`modelrisk-mcp install` detects both Claude Desktop *and* Claude Code, backs up each config, and adds the `modelrisk` server entry to both. Restart Claude Code afterwards. To undo: `modelrisk-mcp uninstall`.

## Manual global config

Claude Code reads the same Windows config as Claude Desktop:

```
%APPDATA%\Claude\claude_desktop_config.json
```

Add a `modelrisk` entry the same way you would for Claude Desktop — see [docs/claude-desktop.md](claude-desktop.md) for the three options (auto-install, standalone `.exe`, source checkout).

## Project-local config

If you want ModelRisk to be available only inside a specific Excel-modelling project, drop a `.mcp.json` at the project root:

```json
{
  "mcpServers": {
    "modelrisk": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "C:/Users/you/source/repos/modelrisk-mcp",
        "python",
        "-m",
        "modelrisk_mcp"
      ]
    }
  }
}
```

When you open the project in Claude Code, you'll be asked once whether to trust the local server. After approval, it's available for that workspace.

## Verifying with the MCP Inspector

From a checkout of the modelrisk-mcp repo:

```powershell
uv run mcp dev src/modelrisk_mcp/__main__.py
```

Opens the MCP Inspector at <http://localhost:5173>. You should see **50 tools** listed under `tools/list` — `list_open_workbooks`, `get_active_workbook`, `audit_model`, `run_simulation`, `build_drivers_report`, and so on. Every description starts with `"ModelRisk: ..."`.

## Slash commands

ModelRisk MCP ships 5 prompt templates that Claude Code surfaces as slash commands:

- `/build-risk-model` — guided 9-step workflow from outputs through running the sim
- `/audit-model` — run all 13 audit rules and report findings with suggested fixes
- `/interpret-results` — convert percentiles into an executive narrative
- `/add-uncertainty` — replace hard-coded constants with appropriate distributions
- `/import-legacy-model` — migrate a deterministic Excel model to ModelRisk

Type `/` in Claude Code to discover them.

## Tips

- **Keep Excel and the workbook open** while you talk to Claude. The server attaches to a running Excel instance; it doesn't launch one for you.
- **One MCP client at a time** drives Excel. The writer mutex prevents two clients from racing; the second will get `ConcurrentWriterError`.
- **Audit log.** Every write Claude makes lands in `%LOCALAPPDATA%\VoseSoftware\modelrisk-mcp\writes.log` with before/after formulas. Useful both for debugging and as the data source for the `restore_cell` tool — which can revert a write even after Excel's undo stack has been cleared.
