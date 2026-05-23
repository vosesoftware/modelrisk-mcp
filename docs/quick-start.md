# ModelRisk MCP — 15-minute Quick-start Tutorial

Zero to first simulation in fifteen minutes. By the end you'll have built a probabilistic NPV model from scratch, run 10,000 iterations, and read back the P10 / P50 / P90.

## What you need

- Windows 10 or 11, 64-bit
- Excel 2019+ with the ModelRisk add-in installed and the **ModelRisk** ribbon tab visible
- Claude Desktop installed (or Claude Code, or Claude for Excel)
- ~15 minutes

You don't need to know ModelRisk's syntax — Claude does. You don't need to know which distribution family fits demand — Claude will propose one.

---

## Step 1 — Install the server (2 min)

Open PowerShell and run:

```powershell
pip install modelrisk-mcp
modelrisk-mcp install
```

You'll see:

```
  + Claude Desktop   added    C:\Users\you\AppData\Roaming\Claude\claude_desktop_config.json
      Registered 'modelrisk' -> {'command': 'C:\\...\\Scripts\\modelrisk-mcp.exe'}
      backup: ...claude_desktop_config.json.bak.20260523-104517

Restart Claude Desktop / Claude Code to pick up the new server.
```

If you also have Claude Code installed, it gets the same treatment. The installer never clobbers an existing entry — your other MCP servers stay intact and the previous config is backed up beside the original.

**Restart Claude Desktop now.** The ModelRisk tools won't appear until the host process reloads its config.

---

## Step 2 — Open Excel with a simple model (3 min)

Open a fresh Excel workbook. Type this in:

| Cell | Value | What it is |
|---|---|---|
| `A1` | `Year 1 revenue` | label |
| `B1` | `1000000` | $1M baseline |
| `A2` | `Annual growth` | label |
| `B2` | `0.08` | 8% baseline |
| `A3` | `Discount rate` | label |
| `B3` | `0.10` | 10% baseline |
| `A4` | `Years` | label |
| `B4` | `5` | 5-year horizon |
| `A6` | `NPV` | label |
| `B6` | `=B1 * ((1-(1+B2)^B4*(1+B3)^-B4)/(B3-B2))` | annuity formula |

You should see `B6 ≈ 4,383,143`. That's the deterministic NPV — one point estimate, no uncertainty, no range.

![Deterministic NPV model with hard-coded inputs](img/01-before-deterministic.png)

Save the workbook somewhere you can find it (Desktop, anywhere). Keep it open in Excel.

> **Why this model?** It's the smallest interesting risk model: three uncertain inputs (revenue, growth, discount rate) and one summary output (NPV). Replace it with your own model later — the workflow is identical.

---

## Step 3 — Sanity-check the connection (1 min)

Open Claude Desktop. Type:

> What ModelRisk workbooks are currently open?

You should see Claude pause, call the `list_open_workbooks` and `get_active_workbook` tools, and report back with the name of the workbook you just saved.

If you see "I don't have access to any ModelRisk tools" or similar — the install didn't take. Double-check that Claude Desktop was restarted *after* `modelrisk-mcp install` ran.

---

## Step 4 — Wrap the constants as probabilistic inputs (4 min)

Tell Claude what's uncertain:

> The active workbook has three uncertain inputs: Year-1 revenue, annual growth, and the discount rate. Wrap each with an appropriate distribution. For revenue use a PERT (min 800k, most-likely 1M, max 1.3M). For growth, a normal centred on 8% with stdev 2%. For discount rate, a triangle (8%, 10%, 12%).

Claude will:

1. Call `propose_distributions_for_inputs` to validate the structures match the methodology guide.
2. Call `replace_constant_with_distribution` once per cell. Each is a **dry-run preview** first — you see what would happen, then Claude commits.
3. Wrap each output cell with `VoseOutput`.

After it's done, cell B1 should read something like `=VoseInput("Year1Revenue") * VosePERT(800000, 1000000, 1300000)`. The deterministic 1,000,000 is gone; in its place is a distribution that will be sampled 10,000 times.

![Same model after Claude wrapped the inputs with Vose distributions (View → Formulas turned on for clarity)](img/02-after-probabilistic.png)

If you make a mistake or want to undo: press `Ctrl+Z` in Excel. Every write the server makes lands in Excel's native undo stack.

---

## Step 5 — Wrap the output (1 min)

> Wrap cell B6 (the NPV) as a VoseOutput named "NPV (5y)".

Claude calls `wrap_with_output`. Cell B6 becomes:

```
=VoseOutput("NPV (5y)") + (B1 * ((1-(1+B2)^B4*(1+B3)^-B4)/(B3-B2)))
```

The `VoseOutput` wrapper is what tells ModelRisk to record this cell's value on every iteration of the simulation. Without it, the iterations run but nothing gets stored to look at later.

---

## Step 6 — Run the simulation (2 min)

> Run a 10,000-iteration simulation.

Claude calls `run_simulation`. ModelRisk's engine fires up, recalculates the workbook 10,000 times with new samples each iteration, and writes the results to a `.vmrs` file beside your workbook. Typical time: 5-15 seconds.

When it's done, Claude reports the run summary: iterations completed, outputs registered, file path written.

---

## Step 7 — Read the results (1 min)

> What are the P10, P50, and P90 of the NPV? What's the standard deviation? Which input drives the variance the most?

Claude calls `get_simulation_results` and `get_sensitivity_ranking`. You'll see something like:

```
NPV (5y):
  Mean:   4,206,150
  Stdev:    742,300
  P10:    3,310,000
  P50:    4,182,000
  P90:    5,158,000

Sensitivity ranking (% of output variance):
  1. Year1Revenue       58%
  2. AnnualGrowth       29%
  3. DiscountRate       13%
```

Compare to your deterministic answer (4,383,143). The mean is roughly the same, but you now know:
- There's a 10% chance NPV comes in below 3.3M.
- The biggest driver of variance is the revenue assumption — focus there if you want to tighten the range.

![Post-simulation Results sheet — KPI table, percentiles, sensitivity ranking, and executive narrative](img/03-results-summary.png)

---

## Step 8 — Generate a report (1 min)

> Build the executive report.

Claude calls `build_executive_report`. A new sheet appears in the workbook with: a KPI tile (mean ± P10–P90), a histogram with cumulative distribution overlay, a percentile table, a tornado chart of the top drivers, and a one-paragraph narrative.

You can paste the report sheet into a Word doc, a slide deck, or just share the workbook directly. Everything is reproducible — `run_simulation` against this same workbook will give the same percentiles (within Monte Carlo noise) tomorrow, next month, next year.

---

## What you just did

In 15 minutes you:

1. Installed the server (one PowerShell line)
2. Built a deterministic Excel NPV model
3. Wrapped 3 hard-coded inputs as ModelRisk distributions
4. Wrapped the output for tracking
5. Ran 10,000 iterations
6. Read back percentiles + sensitivity
7. Generated an executive report

No ModelRisk syntax memorisation. No clicking through ribbon dialogs. No manual chart formatting.

## Where to go from here

- **Replace the NPV model** with your own — revenue forecast, project budget, insurance reserve, capacity plan. Same conversation pattern.
- **Add data fitting** — if you have historical data, ask Claude to fit a distribution to it instead of providing parameters. See the [user manual §3](user-manual.md#3-fit-distributions-to-historical-data).
- **Build correlated inputs** — copulas link multiple inputs into a single dependency structure. See the [user manual §4](user-manual.md#4-build-correlated-multi-period-or-aggregate-structures).
- **Audit an inherited model** — when someone else's workbook lands on your desk, `audit_model` flags the common mistakes. See [user manual §5](user-manual.md#5-audit-a-model-for-common-mistakes).
- **Scenarios** — run the sim across multiple settings of one input. Try `Run scenarios sweeping the discount rate from 6% to 14% in 1% steps.`
- **Slash commands** — type `/build-risk-model`, `/audit-model`, `/interpret-results`, `/add-uncertainty`, or `/import-legacy-model` in Claude Desktop to launch guided workflows.

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| Claude says "I don't have access to those tools" | Claude Desktop didn't reload the config | Fully quit Claude Desktop (right-click tray icon → Quit) and re-open |
| `run_simulation` fails with "macro may not be available" | Excel was started programmatically | Restart Excel from the Start menu / taskbar / `.xlsx` double-click |
| Cells show `#NAME?` instead of values | ModelRisk add-in isn't loaded | Excel → File → Options → Add-ins → Manage Excel Add-ins → check ModelRisk |
| Two simultaneous calls produce `ConcurrentWriterError` | Two MCP clients trying to drive the same Excel | Pick one client per session; close the other |
| Workbook path comes back empty | OneDrive workbook without `ONEDRIVE_COMMERCIAL_WIN` env var | Set the env var, or save a local copy first |

For anything else, see [issues](https://github.com/vosesoftware/modelrisk-mcp/issues).
