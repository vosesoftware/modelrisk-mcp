# External surface

ModelRisk MCP talks to two external surfaces. This file is the
canonical record of which entry points we depend on, what they expect,
and where the dispatch lives in our codebase.

> **Historical note.** Earlier `v0.2` builds attempted to dispatch the
> ATL CoClasses (`ModelRisk.ModelRiskSimulation`,
> `ModelRisk.ModelRiskSimulationResults`,
> `ModelRisk.ModelRiskSimulationSettings`,
> `ModelRisk`) declared in `ModelRiskAtl/*.rgs`. The IDL claims
> IDispatch, but the runtime objects don't return `IDispatch` from
> `QueryInterface`, so cross-process `Dispatch()` from Python fails. We
> abandoned that path. We do **not** call into the ATL CoClasses; we do
> **not** ship `scripts/spike_com_surface.py` as a configuration step.
> If you see references to the ATL surface in older docs or commits,
> they predate `v0.3`.

## Surface 1: Excel + ModelRisk XLL — via `Application.Run`

xlwings attaches to a running Excel instance and gives us Excel's
standard COM Application surface. Builds and simulation kickoff go
through that — no ModelRisk-specific COM dispatch.

The XLL exports a handful of `VoseXxx` *commands* (not worksheet
functions) that the ribbon and the ATL use internally. Excel exposes
those to automation via `Application.Run(name, ...)`. We use exactly
two of them.

Both calls are dispatched from
[`bridge/simulation.py`](../src/modelrisk_mcp/bridge/simulation.py) via
the bridge's xlwings handle:

```python
app.api.Run("VoseStartSimulCustom12", options_2d)
app.api.Run("VoseGetDataSZ12", session_name, target_path)
```

### `VoseStartSimulCustom12(options_array)`

Starts a simulation on the active workbook. **Synchronous** —
`Application.Run` doesn't return until the simulation finishes.

`options_array` is a 1-row 2D SAFEARRAY of BSTRs, each entry a
`[Key]:Value` string. The exact order and key set mirrors
`CSimulationOptions::PackToStringList` (`ModelRiskAtl/SimulationObj.cpp`).
Reproduced by `SimulationOptions.to_string_list()`:

| Key                  | Type   | Notes                                                                          |
|---                   |---     |---                                                                              |
| `[N]`                | int    | Number of simulations (`sim_count`, default 1)                                 |
| `[Samples]`          | int    | Iterations per simulation                                                      |
| `[CntNames]`         | int    | Output filter count                                                            |
| `[name<i>]`          | string | Output filter names (omit all for "all outputs")                               |
| `[SeedFixed]`        | 0/1    | Use the supplied seed                                                          |
| `[SeedMultiplyType]` | int    | How seeds extend across multiple simulations                                   |
| `[CntSeeds]`         | int    | Number of seed values                                                          |
| `[seed<i>]`          | int    | One seed per simulation                                                        |
| `[RefreshExcel]`     | 0/1    | Refresh Excel during the run (default 0 — faster, no flicker)                 |
| `[RefreshRate]`      | int    | Refresh interval if enabled                                                    |
| `[StopOnOutputError]`| 0/1    | Abort on output error                                                          |
| `[ShowResultsAtEnd]` | 0/1    | Pop the results window when done (default 0 — we read `.vmrs` ourselves)      |
| `[HideProgressWindow]`| 0/1   | Suppress the progress dialog (default 1 for headless MCP)                     |
| `[MinSimBufferSize]` | int    | Minimum sample buffer                                                          |
| `[MacrosUsage]`      | int    | Macro hook mode                                                                |
| `[Macros<i>]`        | string | Per-phase macro names (start, before, after, end)                              |

XLL handler reference: `ModelRiskCloude/XllAddIn.cpp` declares the
export; `ModelRiskCloude/SimulationObj_VBA.cpp` dispatches on the
packed `[Key]:Value` array and forwards to `CSimulationsManager` —
the same path the ATL's `IModelRiskSimulation::StartSimulation` takes.

### `VoseGetDataSZ12(session_name, xlParam1, ...)`

Multi-purpose XLL data endpoint. The handler in
`ModelRiskCloude/SimulationObj_VBA.cpp:805` dispatches on the
*operation prefix* embedded in the session name:

```
h<hWndExcel>_<Operation>_<book_name>
```

`<hWndExcel>` is `Application.Hwnd`. `<Operation>` is one of the
strings the ATL uses internally; we only use **`SaveResultsToFile`**.

For the save case the handler reads `xlParam1` as the target file
path. When non-empty, the file dialog is skipped and the simulation
results are written to that path via
`CSimulationsManager::SaveWorkbookResults(sc, path)`. The handler
internally signals success through a memory-mapped file for the ATL's
benefit; from automation we just verify the file appeared on disk.

```python
session_name = f"h{app.api.Hwnd}_SaveResultsToFile_{book_name}"
app.api.Run("VoseGetDataSZ12", session_name, target_path)
```

Session-name composition reference: `ModelRiskAtl/ModelRiskSimulationResults.cpp:54`.

### Other Excel COM endpoints we touch

Ordinary Excel automation, via xlwings — no ModelRisk specifics:

- `Application.Books`, `Workbook.Name`, `Workbook.FullName`,
  `Workbook.Sheets`, `Worksheet.Range`, `Range.Formula`,
  `Range.Value`, `Range.NumberFormat`, `Names.Add`, `Application.Undo`,
  `Application.COMAddIns`, `Application.AddIns`.

All of these go through [`bridge/excel.py`](../src/modelrisk_mcp/bridge/excel.py).
The xlwings attach is the only Excel-process discovery we do; we never
launch Excel.

## Surface 2: MRService.dll — via ctypes

Vose's official SDK reads `.vmrs` files. ctypes binding lives in
[`bridge/mrservice.py`](../src/modelrisk_mcp/bridge/mrservice.py).
DLL discovery honours `$env:MRSERVICE_DLL` first, then the standard
install paths (`C:\Program Files\Vose Software\ModelRisk\MRService.dll`
etc.). Activation runs once per process via the bundled key in
[`bridge/_keymat.py`](../src/modelrisk_mcp/bridge/_keymat.py); the
`MRSERVICE_ACTIVATION_KEY` env var overrides it.

| Export                          | Signature (ctypes-shaped)                                                                                              | Use                                                                 |
|---                              |---                                                                                                                     |---                                                                  |
| `MRLIB_SetOfflineActivationKey`  | `bool(c_int64 key)`                                                                                                    | Per-process activation. Bundled key reconstructed in `_keymat`.    |
| `MRLIB_SetOfflineActivationKeyEx2`| `bool(...)`                                                                                                            | Ex2 form; fallback path if the single-key form is rejected.        |
| `MRLIB_OpenSimulationModel`      | `bool(POINTER(c_wchar) path, POINTER(c_longlong) &model_ptr)`                                                          | Open a `.vmrs`. `.xlsx` is rejected — Excel must run the sim.       |
| `MRLIB_CloseSimulationModel`     | `bool(c_longlong model_ptr)`                                                                                           | Release the handle. `VmrsHandle.__exit__` calls this.              |
| `MRLIB_GetModelDataLength`       | `int(c_longlong model_ptr, c_int sim_index)`                                                                           | Iteration count for a given simulation index.                       |
| `MRLIB_GetModelData`             | `int(model, sim, var_id, double* buf, int bufLen, int samplesToRead, bool checkFilter, int* nFilteredOut, int* nErr)` | Pull sample array for one variable. Filtered + error samples removed. |
| `MRLIB_GetModelVarID`            | `bool(model, sim, name, &var_id)`                                                                                      | Name → ID lookup; the names come from the Excel side first.        |
| `MRLIB_CalcStatistics`           | `int(double* data, int n, _, &mean, &min, &max, &var, &cofV, &stdev, &skew, &kurt)`                                    | Population moments in-DLL.                                          |
| `MRLIB_CalcPercentilesArray`     | `int(double* out, double* p, int p_size, double* data, int n, bool sorted, int* err)`                                  | Arbitrary percentiles in-DLL.                                       |
| `MRLIB_CalculateRiskRatio`       | `int(...)`                                                                                                              | Reserved; not currently consumed by the bridge.                     |

Per-variable correlation matrices and tornado sensitivity rankings are
*not* computed by MRService — they're done in Python (numpy) on the
sample arrays returned by `MRLIB_GetModelData`. See
[`bridge/results.py`](../src/modelrisk_mcp/bridge/results.py).

## What we do NOT use

For the avoidance of doubt:

- **ATL CoClasses** (`ModelRisk.ModelRiskSimulation`, `…Results`,
  `…Settings`, `ModelRisk`) — declared in `ModelRiskAtl/*.rgs`, but
  cross-process IDispatch fails. We don't `Dispatch` them. We don't
  attempt `CoCreateInstance`. They don't appear anywhere in the
  `v0.3` codebase.
- **`win32com.client`** — no longer imported by any bridge. Excel
  automation is exclusively xlwings (which uses `win32com` internally,
  but that's a transitive detail).
- **The ATL's IPC memory-mapped files / `Send_sz_to_ATL`** — the
  XLL handlers self-initialise the MMF when invoked via
  `Application.Run`, so we don't have to set anything up on the Python
  side. We confirm save success by checking that the file appears on
  disk.

## See also

- [architecture.md](architecture.md) — the layered view and how these
  surfaces fit together.
- [mrservice-spike.md](mrservice-spike.md) — the 2026-05-20 spike that
  validated the `.vmrs` read path and ruled out headless `.xlsx`.
- [`bridge/simulation.py`](../src/modelrisk_mcp/bridge/simulation.py)
  module docstring — the canonical pointer to the C++ source lines in
  ModelRiskAtl / ModelRiskCloude that we mirror.
- [`bridge/mrservice.py`](../src/modelrisk_mcp/bridge/mrservice.py)
  module docstring — verified DLL surface as of the same spike.
