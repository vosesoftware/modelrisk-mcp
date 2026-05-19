# ModelRisk MCP

**An open Model Context Protocol server for Vose Software's ModelRisk Excel add-in.**

Use it with Claude Desktop, Claude Code, Cursor, Zed, or any MCP-compliant client to read, build, fit, and run Monte Carlo risk models in Excel from a conversation.

> Status: pre-release scaffold. v0.1.0 ships before 30 June 2026. See [the implementation spec](../../Downloads/modelrisk-mcp-server-spec.md) for the phased roadmap.

## What you can do

- **Read** an existing ModelRisk model: inputs, outputs, distributions, results, percentiles, sensitivity rankings.
- **Build** new models: insert distributions, fit from data, wrap inputs/outputs, build aggregates, copulas, time series, and risk events.
- **Run** simulations directly from the conversation.
- **Audit** a model against Vose methodology rules.
- **Interpret** results with structured executive summaries.

## Quick start (placeholder — full docs in Phase 6)

```powershell
# Local development install
uv sync
uv run python -m modelrisk_mcp
```

## License

MIT. See [LICENSE](LICENSE).

## Links

- Vose Software: <https://www.vosesoftware.com>
- Repository: <https://github.com/vosesoftware/modelrisk-mcp>
