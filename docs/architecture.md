# Architecture

Three internal layers plus two external integration paths. The split is
deliberate: ModelRisk MCP *writes* through Excel (builds + simulation
trigger) and *reads* through ModelRisk's official SDK directly. Each path
is picked for what it does best, and the boundaries are where every
prior attempt at a unified-COM design fractured.

## Component diagram

```
┌──────────────────────────────────┐
│  MCP client                      │
│  (Claude Desktop, Code, etc.)    │
└────────────────┬─────────────────┘
                 │ JSON-RPC / stdio (or HTTP)
                 ▼
┌──────────────────────────────────┐
│  FastMCP layer                   │  server.py, tools/, resources/, prompts/
│  — tools, resources, prompts     │
├──────────────────────────────────┤
│  ModelRiskBridge (domain)        │  bridge/modelrisk.py
│  + SimulationController          │  bridge/simulation.py
│  + ResultsReader                 │  bridge/results.py
├──────────────────────────────────┤
│  ExcelBridge      MrServiceBridge│  bridge/excel.py
│  (xlwings)        (ctypes)       │  bridge/mrservice.py
└──────┬───────────────────┬───────┘
       │ Application.Run   │ MRLIB_*
       │ + cell I/O        │ (read .vmrs)
       ▼                   ▼
┌──────────────┐   ┌──────────────────┐
│ Excel +      │   │  MRService.dll   │
│ ModelRisk XLL│   │  (SDK)           │
└──────────────┘   └──────────────────┘
```

## Layers

### FastMCP layer

Protocol surface. Pure Python. Stateless except for the registered
FastMCP instance. Tools, resources, and prompts live under
[`src/modelrisk_mcp/server.py`](../src/modelrisk_mcp/server.py),
[`tools/`](../src/modelrisk_mcp/tools), [`resources/`](../src/modelrisk_mcp/resources),
and [`prompts/`](../src/modelrisk_mcp/prompts). Nothing in this layer
talks to Excel or to MRService; it only calls into the domain layer.

### Domain layer

Owns the function catalogue, formula validation, simulation control,
and results parsing. Pure Python; no direct COM or ctypes calls.

- **`ModelRiskBridge`** ([`bridge/modelrisk.py`](../src/modelrisk_mcp/bridge/modelrisk.py))
  — top-level domain object. Composes `ExcelBridge`, `SimulationController`,
  `ResultsReader`, plus the `FunctionCatalogue`. Every tool entry point
  goes through this object.
- **`SimulationController`** ([`bridge/simulation.py`](../src/modelrisk_mcp/bridge/simulation.py))
  — drives Monte Carlo runs through the XLL command surface (no ATL
  dispatch). Packs `CSimulationOptions` into the `[Key]:Value` array
  shape `VoseStartSimulCustom12` expects, then calls
  `VoseGetDataSZ12` with the `h<hwnd>_SaveResultsToFile_<book>` session
  name to persist the `.vmrs`.
- **`ResultsReader`** ([`bridge/results.py`](../src/modelrisk_mcp/bridge/results.py))
  — opens `.vmrs` files via `MrServiceBridge`, resolves variable IDs,
  pulls sample arrays, computes statistics and percentiles via
  `MRLIB_CalcStatistics` / `MRLIB_CalcPercentilesArray`. Correlation
  matrices and tornado sensitivity are computed in Python (numpy) from
  the per-iteration sample arrays.

### Transport layer

The only layer that touches Excel or MRService.dll. Two bridges, two
external surfaces.

- **`ExcelBridge`** ([`bridge/excel.py`](../src/modelrisk_mcp/bridge/excel.py))
  — wraps xlwings for ordinary Excel I/O: workbook enumeration, cell
  and range reads, formula writes, named-range management, and add-in
  inspection. The simulation kickoff is *also* an Excel COM call (via
  the `app.api.Run` accessor under `SimulationController`), but it
  reaches the XLL through Excel's standard `Application.Run` surface —
  not through any ModelRisk COM dispatch.
- **`MrServiceBridge`** ([`bridge/mrservice.py`](../src/modelrisk_mcp/bridge/mrservice.py))
  — ctypes wrapper around `MRService.dll`. Discovers the DLL via the
  `MRSERVICE_DLL` env var or the standard install paths, activates it
  per-process with `MRLIB_SetOfflineActivationKey` (key bundled and
  reconstructed in [`bridge/_keymat.py`](../src/modelrisk_mcp/bridge/_keymat.py)),
  opens `.vmrs` files with `MRLIB_OpenSimulationModel`, and exposes a
  `VmrsHandle` that proxies the per-variable sample, statistics, and
  percentile calls.

## Two integration paths

Each path is chosen for what it does best, and each one bypasses
exactly the COM surface that doesn't work.

### Builds + simulation kickoff → Excel COM (xlwings + `Application.Run`)

Everything that mutates the workbook or asks Excel to do work goes
through Excel's own COM surface, attached via xlwings:

- Cell and range I/O, named ranges, formula writes, add-in inspection
  — direct xlwings calls.
- Simulation kickoff — `Application.Run("VoseStartSimulCustom12", options_array)`
  followed by `Application.Run("VoseGetDataSZ12", session_name, target_path)`.

`Application.Run` is a documented Excel COM endpoint that any registered
XLL command can be reached through. We replicate exactly what
`ModelRiskAtl`'s `IModelRiskSimulation::StartSimulation` and
`IModelRiskSimulationResults::SaveResultsToFile` do internally when
their methods fire — the XLL handlers route the `Application.Run` call
to `CSimulationsManager` the same way they route the ATL's internal
call. The version that *did not* work, and which we no longer attempt,
was dispatching the ATL CoClasses directly: `ModelRiskAtl.dll` declares
them as IDispatch in the IDL, but the runtime objects don't return
`IDispatch` from `QueryInterface`, so cross-process automation falls
over. The XLL command surface side-steps that entirely.

### Results read → MRService.dll (ctypes)

Vose's official SDK opens `.vmrs` files, returns sample arrays, and
computes statistics natively:

- `MRLIB_OpenSimulationModel(path, &model_ptr)` — opens a `.vmrs`.
  `.xlsx` is rejected; Excel + the XLL still has to run the simulation.
- `MRLIB_GetModelDataLength`, `MRLIB_GetModelData` — per-variable
  sample arrays.
- `MRLIB_CalcStatistics`, `MRLIB_CalcPercentilesArray` — population
  moments and arbitrary percentiles, computed in-DLL on numpy buffers.
- `MRLIB_GetModelVarID` — name-to-ID lookup keyed on the names already
  read out of Excel by `list_modelrisk_inputs` / `_outputs`.

No COM round-trips per output; the per-iteration sample arrays are
materialised in Python for any downstream analysis we want (correlation
matrices, tornados, custom percentiles, hand-rolled VaR/TVaR).

## What lives where

| Concern                     | File                          |
|---                          |---                            |
| MCP tool registration       | `tools/*.py`                  |
| MCP resources, prompts      | `resources/*.py`, `prompts/*.py` |
| Domain orchestration        | `bridge/modelrisk.py`         |
| Function catalogue load     | `bridge/catalogue.py`         |
| Formula construction        | `bridge/formulas.py`          |
| Simulation control          | `bridge/simulation.py`        |
| `.vmrs` reading + stats     | `bridge/results.py`           |
| xlwings + Excel COM         | `bridge/excel.py`             |
| MRService.dll ctypes        | `bridge/mrservice.py`         |
| Bundled activation key      | `bridge/_keymat.py`           |
| Safety rails (mutex, log)   | `safety.py`                   |
| Typed errors                | `errors.py`                   |

See [com-surface.md](com-surface.md) for the precise XLL commands and
`MRLIB_*` exports we depend on.
