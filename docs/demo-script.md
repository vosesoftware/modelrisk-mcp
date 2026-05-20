# Demo recording script

Use this to record the README's headline demo: Claude turning a deterministic Excel cost model into a Monte Carlo simulation with `/build-risk-model`, all in under 2 minutes.

## Setup (before pressing record)

1. **Excel:** open a fresh workbook. In Sheet1, build a deterministic mini-model:

   | Cell | Content | Note |
   |---|---|---|
   | A1 | `Unit cost` | label |
   | B1 | `25` | the hard-coded input |
   | A2 | `Units` | label |
   | B2 | `1000` | the hard-coded input |
   | A3 | `Fixed cost` | label |
   | B3 | `50000` | the hard-coded input |
   | A4 | `Total cost` | label |
   | B4 | `=B1*B2+B3` | the deterministic output |

   Save as `demo-cost-model.xlsx`. Confirm B4 reads `75000`.

2. **Claude Desktop:** confirm the ModelRisk server is wired up. Run `/build-risk-model` once in a throwaway conversation to confirm Claude lists the right tool calls.

3. **Recording:** set the recorder window to 1280×720 (or 1920×1080 if you want crisper). Hide any other windows. Put Claude Desktop and Excel side-by-side, ~60/40 split favouring Claude.

4. **Tooling:** use any screen recorder that produces clean `.mp4` (OBS, ScreenToGif for animated GIF, ShareX). Aim for 5-7 MB to drop into the README directly.

## Recording (≈110 seconds total)

Beat-by-beat. The numbers in parens are stopwatch markers.

**(0:00 — 0:08) Opening shot.** Excel showing the deterministic model with `Total cost = 75000` clearly visible. Cursor hovers over B4 to emphasise the deterministic output. 

**(0:08 — 0:18) Open Claude.** Switch focus to Claude Desktop. Type:

```
/build-risk-model
```

then press Enter. Claude responds with its first prompt: asks about the decision and outputs.

**(0:18 — 0:35) Type the answer:**

```
I have a small cost model in the active Excel workbook. Total cost is in cell B4. I want a confidence range around it, given that unit cost and units are both uncertain. Fixed cost is known.
```

Claude calls `get_active_workbook` → `list_modelrisk_outputs` (returns `[]`) → `discover_inputs`. Watch the tool calls flow in the right pane.

**(0:35 — 0:50) Propose distributions.** Claude surfaces B1 (unit cost) and B2 (units) as hard-coded candidates and asks about ranges. Type:

```
Unit cost: best case 22, most likely 25, worst case 32.
Units: best case 800, most likely 1000, worst case 1300.
```

Claude calls `propose_distributions_for_inputs` (returns VoseModPERT for both, "three-point estimate" scenario) and previews the formulas with `dry_run=True`. Two preview boxes appear.

**(0:50 — 1:05) Commit.** Type:

```
Looks good — commit both.
```

Claude re-calls `replace_constant_with_distribution` with `dry_run=False` for B1 then B2. Excel updates live: B1 becomes `=VoseInput("Unit cost")+VoseModPERT(22,25,32,4)`, B2 becomes `=VoseInput("Units")+VoseModPERT(800,1000,1300,4)`. Then `wrap_with_output(name="Total cost")` on B4.

**(1:05 — 1:30) Run + interpret.** Claude calls `run_simulation(iterations=10000)`. Brief pause as the sim runs (in Excel — ModelRisk's progress dialog flashes if not hidden). Then `generate_executive_summary` returns markdown. Claude pastes the table.

Final frame: the markdown summary visible in Claude with the mean / P5 / P95 / sensitivity rankings, and Excel showing the now-stochastic B4 (a sampled value, not 75000).

**(1:30 — 1:40) Closing.** Tagline overlay: *"ModelRisk MCP — open Model Context Protocol server for Vose Software's ModelRisk. github.com/vosesoftware/modelrisk-mcp"*

## After recording

- Save as `docs/demo.gif` (animated GIF, optimised) AND `docs/demo.mp4` (for higher quality fallback).
- Drop the GIF into the README under a new `## Demo` section, above the comparison table.
- If the file is over 10 MB, trim to ≈1 minute or reduce frame rate.

## Variations worth recording later (separate files)

- **`/audit-model` demo** — open a workbook with deliberate methodology issues (a VoseFit without uncertainty, a distribution without a wrapper) and watch Claude walk through the findings. Good for the "audit" tab of the website.
- **`/import-legacy-model` demo** — open a workbook with `RiskNormal` / `RiskTriang` calls from another vendor's add-in. Claude maps them to Vose equivalents one cell at a time.

## What can go wrong while recording

- **Excel save dialog interrupts the flow.** Make sure auto-save is off and the workbook is already saved before recording.
- **ModelRisk progress dialog covers the view.** Either set `hide_progress_window=true` in `set_simulation_settings` first, OR position Excel so the dialog appears off-camera.
- **`get_simulation_results` returns empty.** ModelRisk needs the simulation to have *finished* and populated the Results Viewer. If it returns empty, run again — sometimes the dispatch happens before the run is fully complete.
