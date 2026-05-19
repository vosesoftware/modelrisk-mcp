# ModelRisk COM surface

Canonical record of which ModelRisk COM endpoints are exposed by the
installed ModelRisk build. This file is overwritten by
`scripts/spike_com_surface.py` — run that on a Windows machine with
Excel + ModelRisk loaded to get live results.

## ProgIDs

Source: `.rgs` files under `ModelRisk_Project/VBAProject/ModelRiskAtl/`
in the ModelRisk source tree. Verified 2026-05-19.

| Purpose | ProgID | CLSID |
| --- | --- | --- |
| Distribution functions (`VoseNormal`, etc.) | `ModelRisk` | `{570013C9-8251-44CF-AF83-EDD333725537}` |
| Run a simulation | `ModelRisk.ModelRiskSimulation` | `{59530ADE-E690-4802-A6E4-890B72596310}` |
| Read results | `ModelRisk.ModelRiskSimulationResults` | `{B1EEBA78-BE81-4d37-8FEA-FC3AE14BE755}` |
| Read/write settings | `ModelRisk.ModelRiskSimulationSettings` | `{389CD5FB-F265-467e-A255-90C206CE7220}` |

All four share TypeLib `{ECC429DA-26E6-4D86-9B2D-1E14E0461749}`.

## Expected surface (pre-probe baseline)

These are what the IDL declares. The spike script confirms them on the
local install.

`ModelRisk.ModelRiskSimulation`:
- `StartSimulation()` — no parameters

`ModelRisk.ModelRiskSimulationSettings`:
- `Samples`, `Simulations`, `UseFixedSeed`, `Seed[idx]`, `MultipleSeedType`,
  `RefreshExcel`, `RefreshRate`, `StopOnOutputError`, `ShowResultsAtEnd`,
  `HideProgressWindow`

`ModelRisk.ModelRiskSimulationResults`:
- Collection navigation: `SimVariables()`, `SimInputs()`, `SimOutputs()`
- Direct: `GetVariablesCount`, `GetVarName`, `GetVarSamples`,
  `GetVarType`, `GetVarLocation`, `GetVarRange`
- Charts: `AddChart`, `SetChartType`, `GetChartsCount`, `GetChartType`,
  `GetChartName`, `GetChartVariables`, `AddChartForEmbeddedReport`
- Reporting: `ReportAllVariables`, `ReportSelectedCharts`
- Persistence: `SaveResultsToFile`, `LoadResultsFromFile`

`ISimVariable` (obtained from `SimOutputs().Item(i)` etc.):
- Confirmed: `GetName`, `GetRangeName`, `GetMean`, `GetVariance`,
  `GetStDev`, `GetSkewness`, `GetKurtosis`, `GetCofV`, `GetPercentile(p)`,
  `GetProbability(x)`, `GetSamples()`
- **Pending developer confirmation:** `GetCorrelation`, `GetTornado`.
  Until confirmed, `ResultsReader` computes both in Python from
  `GetSamples()` via numpy.

## Known gaps (ticketed)

- Native simulation cancellation
- Native `SimulationStatus` property

v0.1 ships `NotImplementedError` stubs for these, plus a polling
fallback for status.
