# Architecture

See the full implementation spec for the canonical version. This file is
a stub that grows during Phase 1 and beyond.

## Three layers

1. **FastMCP layer** (`server.py`, `tools/`, `resources/`, `prompts/`) —
   protocol surface. Pure Python. Stateless except for the registered
   FastMCP instance.
2. **ModelRiskBridge** (`bridge/modelrisk.py`) — domain logic. Owns the
   function catalogue, formula construction, simulation control, results
   parsing. Pure Python; no direct COM calls.
3. **ExcelBridge** (`bridge/excel.py`) — transport. The only layer that
   imports `xlwings` or `win32com`. Wraps xlwings for ordinary Excel
   operations and uses `win32com.client` directly for ModelRisk-specific
   COM endpoints.

See [com-surface.md](com-surface.md) for the COM availability matrix.
