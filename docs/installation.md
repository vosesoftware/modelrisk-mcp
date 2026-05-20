# Installation

## Prerequisites

- **Windows 10 or 11, 64-bit.** ModelRisk is Windows-only.
- **Microsoft Excel 2019 or newer** with the ModelRisk add-in installed and loaded. Open Excel and confirm the **ModelRisk** ribbon tab is present before continuing.
- **Python 3.11, 3.12, or 3.13.** The recommended way is to install [`uv`](https://docs.astral.sh/uv/) which can manage Python versions for you.

## From PyPI (once published)

```powershell
pip install modelrisk-mcp
```

## From source

```powershell
git clone https://github.com/vosesoftware/modelrisk-mcp
cd modelrisk-mcp
uv sync --extra dev
uv run python -m modelrisk_mcp     # speaks MCP over stdio; for the inspector use:
uv run mcp dev src/modelrisk_mcp/__main__.py
```

## Standalone Windows executable (Phase 6)

Once we ship signed `.exe` releases, you'll be able to download a single binary from the GitHub releases page and point any MCP client at it — no Python knowledge required. Until then, use the source path above.

## Verifying the install

With Excel open, run the COM-surface spike script:

```powershell
uv run python scripts/spike_com_surface.py
```

It writes `docs/com-surface.md` with the live probe results: which ModelRisk ProgIDs Dispatch, which `ISimVariable` accessors are exposed, etc. If anything reports `NO`, the MCP server will still work for that endpoint's neighbours — but flag it to the team so we can prioritise.
