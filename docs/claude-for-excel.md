# Claude for Excel and ModelRisk MCP — the honest status

> **TL;DR: Claude for Excel cannot currently reach this server. Use Claude
> Desktop or Claude Code (local stdio) instead — `modelrisk-mcp install`
> wires both.** This page replaces earlier instructions that described a
> Settings → Connectors flow inside the add-in; those instructions were wrong
> for shipping builds and are withdrawn (verified 2026-07 against a current
> build, and independently confirmed by a field report).

## Why it can't work today

Claude for Excel is an Office.js add-in running in a browser WebView. Its
connector model has three properties that together rule out a local,
Excel-driving MCP server:

1. **No local Connectors panel.** The add-in's settings contain no "Add MCP
   server" entry. It resolves connectors from your **claude.ai account**
   connector list.
2. **claude.ai custom connectors are remote.** The "Add custom connector"
   dialog takes a **remote MCP server URL** plus OAuth credentials, and those
   connectors are **fetched by Anthropic's infrastructure, not by your
   device**. A loopback URL (`http://127.0.0.1:8000/mcp`) is unreachable *by
   construction* — no configuration makes it work. The OAuth fields also
   don't match this server's bearer-token auth.
3. **Local stdio servers never appear in the "Connectors" panel** on
   claude.ai or Claude Desktop. That's normal: the Connectors panel lists
   account-level remote connectors; local stdio servers are a separate,
   Desktop-only mechanism. Not seeing `modelrisk` there does **not** mean the
   install failed.

The server side is fine — the HTTP transport handshakes correctly and
enforces its bearer token (`401` without it). The limitation is entirely in
how the client resolves connectors.

## Do not tunnel around this

The only technically-possible route to Claude for Excel today would be
exposing the server through a public tunnel with OAuth in front. **We do not
recommend or support this.** This server can write formulas into your
workbooks, run simulations, and save files to disk. A publicly reachable
instance guarded by a single bearer token is a materially different security
posture from the loopback bind these docs are written around. If your
organisation genuinely needs an in-Excel path, contact Vose — that is a
product decision (a hosted connector with real auth), not a configuration
setting.

## What to use instead

| Client | Transport | Status |
|---|---|---|
| **Claude Desktop** | local stdio | ✅ Supported, tested — `modelrisk-mcp install` |
| **Claude Code** | local stdio | ✅ Supported, tested — `modelrisk-mcp install` (0.3.11+; earlier versions wrote the wrong config file) |
| Cursor / Zed / other local MCP clients | local stdio | ✅ Works; configure manually |
| An MCP client running on your own machine/LAN | `--transport=streamable-http` + bearer token | ✅ Works — keep the default loopback bind |
| **Claude for Excel** | — | ❌ Not reachable by current builds |

Everything this server does — building models, running simulations, reading
results, charts, reports — works identically from Claude Desktop and Claude
Code against the same Excel session you have open. Excel and ModelRisk are
driven either way; only the chat window lives elsewhere.

## HTTP transport (for clients that do run on your machine)

```powershell
$env:MODELRISK_MCP_TOKEN = [Guid]::NewGuid().ToString("N") * 2
modelrisk-mcp --transport=streamable-http --host=127.0.0.1 --port=8000 --token=$env:MODELRISK_MCP_TOKEN
```

- `POST /mcp` with `Authorization: Bearer <token>` → MCP initialize handshake.
- No token → `401`.
- Keep `--host=127.0.0.1` unless you fully understand the exposure of
  `0.0.0.0` on your network.
- Add `--read-only` for a first session against a model you care about
  (0.3.11+): reading and analysis work; writes, simulations and saves are
  refused with a clear error.

## If Anthropic's connector model changes

If a future Claude for Excel build adds local MCP support or a
device-fetched connector mode, this page will be updated and the client
re-tested before instructions are published. Watch the CHANGELOG.
