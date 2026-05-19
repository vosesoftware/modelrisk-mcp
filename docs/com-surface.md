# ModelRisk COM surface

This file is the canonical record of which ModelRisk COM endpoints are
available in the installed version vs which are ticketed for the
ModelRisk core team.

Populated by running `uv run python scripts/spike_com_surface.py` on a
Windows machine with Excel + ModelRisk loaded.

Expected starting state (confirmed by Vose on 2026-05-19):

| Endpoint | Available today | Notes |
| --- | --- | --- |
| `Application.Simulate(iterations, seed)` | YES | |
| `Application.Settings.Iterations` | YES | |
| `Application.Settings.Seed` | YES | |
| `Application.Settings.SamplingMethod` | YES | |
| `Application.Results.GetMean` | YES | |
| `Application.Results.GetPercentile` | YES | |
| `Application.Results.GetCorrelation` | YES | |
| `Application.Results.GetTornado` | YES | |
| `Application.Results.GetIterations` | YES | |
| `Application.StopSimulation()` | NO | Ticketed for ModelRisk core. v0.1 ships `NotImplementedError`. |
| `Application.SimulationStatus` (property) | NO | Same. Polling fallback in v0.1. |

The spike script overwrites this file with the live probe results.
