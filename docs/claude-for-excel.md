# Wiring ModelRisk MCP into Claude for Excel

Claude for Excel runs inside an Office.js iframe sandboxed *inside* Excel itself. It can't spawn local subprocesses (the way Claude Desktop launches a stdio MCP server), so it talks to MCP servers over **HTTP** instead. ModelRisk MCP serves both stdio and HTTP transports.

This is also where the architectural payoff is: Office.js can't reach Excel's COM surface or ModelRisk's ribbon. ModelRisk MCP, running outside the sandbox, can — which means Claude for Excel can do things via this server that it structurally can't do on its own.

## Prerequisites

- Excel 2019+ with ModelRisk loaded
- Claude for Excel installed (Microsoft AppSource → "Claude for Excel" by Anthropic)
- ModelRisk MCP installed: `pip install modelrisk-mcp` or the standalone `.exe`

## 1. Start the server in HTTP mode

Open a PowerShell window. Generate a token first — anything random and >=24 characters. PowerShell built-in:

```powershell
$env:MODELRISK_MCP_TOKEN = [Guid]::NewGuid().ToString("N") + [Guid]::NewGuid().ToString("N")
$env:MODELRISK_MCP_TOKEN
```

Copy the printed token — you'll paste it into Claude for Excel in step 2. Then start the server:

```powershell
modelrisk-mcp --transport=streamable-http --port=8000 --token=$env:MODELRISK_MCP_TOKEN
```

You should see:

```
INFO:     Uvicorn running on http://127.0.0.1:8000 (Press CTRL+C to quit)
```

Leave the window open while you work. Closing it stops the server.

> **Why a token?** The HTTP endpoint binds to `127.0.0.1` (loopback) by default, so only processes on your machine can reach it. But *any* process on your machine could otherwise hit it and drive Excel — including malicious ones. A bearer token shuts that down. If you bind to a non-loopback host, the token isn't optional.

## 2. Add the connector in Claude for Excel

Open Excel, open Claude for Excel (the side-panel icon), then:

1. Go to **Settings → Connectors** (or the equivalent in your current Claude for Excel build).
2. Click **Add MCP server**.
3. Fill in:
   - **Name:** `modelrisk` (or anything you'll recognise)
   - **URL:** `http://127.0.0.1:8000/mcp`
   - **Authentication:** Bearer token
   - **Token:** paste the token from step 1
4. Save.

Claude for Excel should report the connection as live and show the 40 ModelRisk tools.

## 3. First conversation

In Claude for Excel:

> Summarise the active workbook's risk model — inputs, outputs, distributions.

Or jump straight in:

> /build-risk-model

The same tool surface and same prompts as Claude Desktop. See [docs/demo-script.md](demo-script.md) for the headline workflow.

## Lifecycle tips

- **Token in another shell.** If you need to share the token with a second Claude for Excel session or paste it into a different config, save it to a file once and read it back. `Get-Clipboard` after the env-var line above also works.
- **Stopping the server.** `Ctrl+C` in the PowerShell window. The Claude for Excel connector will go red until you start it again.
- **Auto-start on boot.** Wrap the launch command in a Windows scheduled task running at logon, or a Start Menu shortcut. Keep the token out of source control.
- **Two-machine setup (advanced).** Bind to `0.0.0.0`, expose the chosen port through your firewall, and use a strong token. Watch the security model carefully — this server can write to your Excel.

## Troubleshooting

### "Connection refused" or "ERR_CONNECTION_RESET"

The server isn't running, or it's bound to a port Claude for Excel can't reach. Check the PowerShell window — uvicorn prints `Uvicorn running on http://127.0.0.1:8000` when it's healthy.

### 401 Unauthorised

The token Claude for Excel is sending doesn't match the one the server expects. Re-paste it. Tokens are case-sensitive.

### Tools listed but every call hangs

The server is running but can't reach Excel. Confirm Excel is open and ModelRisk is loaded. Try the standalone `modelrisk-mcp --transport=stdio` with Claude Desktop first — that's a simpler topology to debug.

### Concurrent-writer errors

If Claude Desktop is running the same server over stdio at the same time, the writer mutex will reject one of them. Pick one client per session, or run the HTTP server with a different mutex name (advanced — see `src/modelrisk_mcp/safety.py`).

### "version": "1.27.1" in serverInfo

That's the FastMCP library version FastMCP currently reports, not our package version. Cosmetic, doesn't affect behaviour. The actual server is identifiable as `"name": "modelrisk-mcp"`.

## Security model — important

ModelRisk MCP over HTTP is a **local-only, single-user** integration in its default configuration. The defaults:

- Loopback bind (`127.0.0.1`) — only your own machine
- Bearer token required if you change the bind to a non-loopback host
- No outbound network calls — the server doesn't phone home
- All writes still default to `dry_run=True`; the writer mutex still serialises commits; the audit log still records every change

Do **not** expose this server to the public internet. The MCP tool surface includes `replace_constant_with_distribution`, `run_simulation`, and `set_named_range` — anyone who can reach the endpoint with a valid token can drive your Excel. Treat the token like an API key.
