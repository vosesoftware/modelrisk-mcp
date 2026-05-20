# Community submission — modelcontextprotocol/servers

The official MCP server directory lives at <https://github.com/modelcontextprotocol/servers>. Submitting ModelRisk MCP there gives the project discoverability inside the broader MCP ecosystem.

This file is the prepared text for the submission — open the PR after `v0.1.0` ships and the repo is public.

## Steps

1. Fork <https://github.com/modelcontextprotocol/servers>.
2. Clone your fork locally.
3. Edit `README.md` to add `ModelRisk MCP` to the "Community Servers" section (alphabetical order). The exact patch is below.
4. Commit, push, open the PR with the title and body below.

## README patch (alphabetical insertion under "Community Servers")

```markdown
- **[ModelRisk MCP](https://github.com/vosesoftware/modelrisk-mcp)** — Open MCP server for [Vose Software's ModelRisk](https://www.vosesoftware.com) Monte Carlo Excel add-in. Read, build, fit, and run risk models from a Claude conversation.
```

If the existing list uses a different one-liner format, match it.

## PR title

```
Add ModelRisk MCP to Community Servers
```

## PR body

```markdown
Adds [ModelRisk MCP](https://github.com/vosesoftware/modelrisk-mcp) under
Community Servers.

ModelRisk MCP is an open Model Context Protocol server for [Vose
Software's ModelRisk](https://www.vosesoftware.com), a Monte Carlo
risk-modelling add-in for Excel. It exposes 31 tools, 7 resources, and
5 slash-command prompts that let Claude (or any MCP client) read,
build, fit, and run Monte Carlo models in Excel from a conversation.

- Repository: https://github.com/vosesoftware/modelrisk-mcp
- License: MIT
- Runtime: Python 3.11+ on Windows (Excel + ModelRisk required)
- Distribution: PyPI (`pip install modelrisk-mcp`) and standalone `.exe`

Notable differentiators:

- Build vs read-only: most Excel-adjacent MCP servers I've seen only
  read; ModelRisk MCP also writes (insert distributions, fit from
  data, build aggregates/copulas/time-series, run simulations).
- Methodology-aware: the 1417-entry ModelRisk function catalogue,
  Vose's methodology principles, and a 6-rule audit engine all ship
  with the package so the LLM is grounded against hallucinated
  function names and common modelling mistakes.
- Safe by design: every building tool defaults to `dry_run=True`;
  every write lands in Excel's native undo stack; every commit goes
  through a writer mutex and an append-only audit log; non-Vose
  formulas are not overwritten without explicit opt-in.

Happy to adjust the entry format if a different shape suits the
directory better.
```

## Other places worth posting (separate channels)

- **Hacker News** — "Show HN: ModelRisk MCP — Open MCP server for Monte Carlo simulation in Excel" or similar. Lead with the strategic narrative and the demo GIF. Best window: weekday morning US Eastern.
- **r/excel + r/RiskManagement** — short post linking to the README and demo.
- **MCP community Discord / forums** — drop the link with one or two sentences.
- **LinkedIn** — Vose Software corporate post + personal posts from team members. The strategic narrative ("open vs proprietary connector") plays better on LinkedIn than the technical depth does.

Avoid being adversarial about the closed alternatives. The comparison table in the README already does that work without naming names; piling on doesn't add anything and risks looking petty.
