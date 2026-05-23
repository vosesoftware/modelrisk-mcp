# Wiring ModelRisk MCP into Claude Desktop

Claude Desktop discovers MCP servers from its `claude_desktop_config.json` file. Once configured, the ModelRisk server's tools appear in Claude's "Connected" menu and can be invoked directly from a chat.

## Locate the config file

On Windows, Claude Desktop reads its config from:

```
%APPDATA%\Claude\claude_desktop_config.json
```

In PowerShell:

```powershell
notepad "$env:APPDATA\Claude\claude_desktop_config.json"
```

If the file doesn't exist, create it with the empty skeleton:

```json
{
  "mcpServers": {}
}
```

## Add the ModelRisk server

Three options — pick one.

### Option A — One command, `modelrisk-mcp install` (recommended)

```powershell
pip install modelrisk-mcp
modelrisk-mcp install
```

`modelrisk-mcp install` finds the config file at the path above, backs it up (`*.bak.<timestamp>`), and adds the `modelrisk` entry — leaving any other servers you have configured intact. Output:

```
  + Claude Desktop   added    C:\Users\you\AppData\Roaming\Claude\claude_desktop_config.json
      Registered 'modelrisk' -> {'command': 'C:\\...\\Scripts\\modelrisk-mcp.exe'}
```

To undo: `modelrisk-mcp uninstall`. To register a second instance under a different name (e.g. dev vs prod side-by-side): `modelrisk-mcp install --name=modelrisk-dev`.

### Option B — Standalone `.exe`

Download `modelrisk-mcp.exe` from the [latest release](https://github.com/vosesoftware/modelrisk-mcp/releases/latest) — no Python required. Add this entry under `mcpServers` using the absolute path:

```json
{
  "mcpServers": {
    "modelrisk": {
      "command": "C:/path/to/modelrisk-mcp.exe"
    }
  }
}
```

### Option C — Run from a source checkout (development)

Replace `C:/Users/you/source/repos/modelrisk-mcp` with the actual path to your clone:

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

You can also pass `--read-only` in the `args` list to disable every building/simulation tool — useful for sensitive workflows where you only want Claude to inspect:

```json
"args": ["-m", "modelrisk_mcp", "--read-only"]
```

## Restart Claude Desktop

Quit Claude Desktop fully and re-open it. The ModelRisk tools should now be listed when you click the connections icon in the conversation.

## First conversation

Open a workbook in Excel that has at least one Vose function in it (any `=VoseNormal(...)` will do). In Claude Desktop, ask:

> What ModelRisk inputs and outputs are in the active workbook?

Claude will call `get_active_workbook`, `list_modelrisk_inputs`, and `list_modelrisk_outputs`, and report the results.

## Troubleshooting

| Symptom | Cause |
|---|---|
| Tools don't appear in Claude after restart | Config syntax error or wrong path. Open the file in VS Code; it'll flag JSON issues. |
| `ExcelNotRunningError` | Excel isn't open. Launch Excel before invoking ModelRisk tools. |
| `ModelRiskNotLoadedError` | The ModelRisk add-in isn't loaded in Excel. In Excel, go to *File → Options → Add-ins* and confirm ModelRisk is enabled. |
| Tools appear but every call hangs | An older MCP server instance may still hold the writer mutex. Quit Claude Desktop and any other MCP client, then retry. |

## What the server does NOT do

- It does not save your workbook automatically. You always control `Ctrl+S`.
- It defaults every building tool to `dry_run=True` — Claude must explicitly commit a change.
- It runs entirely on your machine. No data leaves your computer.
