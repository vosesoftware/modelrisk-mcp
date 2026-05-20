# Wiring ModelRisk MCP into Claude Code

Claude Code (the CLI / IDE extension) connects to MCP servers via the same `claude_desktop_config.json` mechanism plus a project-local `.mcp.json` if you want repo-scoped configuration.

## Global config

Same as [Claude Desktop's config](claude-desktop.md) — Claude Code reads the same file on Windows:

```
%APPDATA%\Claude\claude_desktop_config.json
```

Add the `modelrisk` entry from the Claude Desktop guide.

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

Opens the MCP Inspector at <http://localhost:5173>. You should see the 12 reading tools listed under `tools/list`, each with description starting `"ModelRisk: ..."`.

## Slash commands and prompts

Slash commands (`/build-risk-model`, `/audit-model`, `/import-legacy-model`, etc.) land in Phase 5. Until then, Claude Code can invoke any of the reading tools directly from a regular conversation.

## Tips

- **Keep Excel and the workbook open** while you talk to Claude. The server attaches to a running Excel instance; it doesn't launch one for you.
- **One MCP client at a time** drives Excel. The writer mutex prevents two clients from racing; the second will get `ConcurrentWriterError`.
- **Audit log.** Every write Claude makes (Phase 3+) lands in `%LOCALAPPDATA%\VoseSoftware\modelrisk-mcp\writes.log` with before/after formulas. Useful both for debugging and for the `restore_cell` tool (Phase 3).
