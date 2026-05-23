# ModelRisk MCP — Launch Video Script

A 5–6 minute video for the v0.3.0 public launch: posted to LinkedIn / X / HN / the GitHub README hero spot. Audience: risk modellers who use Excel daily; finance / FP&A teams; consultants; people who've heard "Monte Carlo in Excel" but never tried it.

> **Companion short**: `docs/demo-script.md` is a 90-second screen recording for the README. This file is the longer launch piece.

---

## Production notes

- **Runtime target**: 5:30–6:00 final cut.
- **Aspect**: 16:9, 1920×1080 master. Crop to 1:1 for social if needed.
- **Recording stack**: OBS for screen + cam (PiP bottom-right, optional), Audacity for any voiceover, DaVinci Resolve for cut.
- **Capture rig**: Excel + Claude Desktop side-by-side at 60/40 split favouring Claude. ModelRisk progress dialog positioned off-frame or use `hide_progress_window=true`.
- **Pace**: deliberate, not rushed. Every numeric output stays on screen long enough to read.
- **Lower thirds**: white text, no logos, used sparingly (intro, milestones, CTA).
- **Music**: light, instrumental, ducked to −18 dB under voiceover. Optional — works without music.
- **Cursor**: large, smooth-tracking. Use a tool like Cursorcerer if your recorder doesn't smooth it.
- **Voiceover**: confident, expert, not salesy. Cadence ~150 wpm. Pauses between beats.

---

## Beat sheet

| Time | Beat | Visual | Audio |
|---|---|---|---|
| 0:00–0:15 | Cold open: the problem | Excel cell B4 showing `$4,383,143` | "Most Monte Carlo models in Excel start the same way…" |
| 0:15–0:45 | The protagonist | Side-by-side Excel + Claude Desktop | "You're a risk modeller. You know your domain cold, but you don't want to memorise 1,400 ModelRisk functions." |
| 0:45–1:15 | Install + first conversation | PowerShell `pip install` → Claude Desktop | "One command. Claude Desktop is wired up." |
| 1:15–3:00 | End-to-end workflow | Live conversation, formulas filling in | The whole construction-overrun example, sped up to 60-second highlights |
| 3:00–4:00 | What else can it do | Quick montage of audit, fit, copula, time-series | Voiceover hits the headlines |
| 4:00–5:00 | Why this matters | Slate cards | "Open standard. Local-only. Methodology-grounded. Safe by default." |
| 5:00–5:30 | CTA | URL, install command, GitHub link | "Install today: `pip install modelrisk-mcp`." |
| 5:30–6:00 | Outro card | Logo + URL | (silence or music tail) |

---

## Full script

### Scene 1 — Cold open (0:00 — 0:15)

**Visual**: clean Excel view, single cell `B4` highlighted, value `$4,383,143`. Slow zoom-in over 3 seconds. Title card fades in below: **"This is your NPV. It's wrong."**

**Voiceover (calm, paced)**:
> "Most Monte Carlo models in Excel start the same way. A single point estimate for net present value. Maybe four million dollars. The board sees that number and treats it as the truth."

**Cut to**: same workbook, but now showing a deterministic baseline next to a stochastic one. The single-point estimate gets replaced by a range: `P10 $3.31M → P50 $4.18M → P90 $5.16M`.

**Voiceover**:
> "But you know the inputs aren't certain. They never are. The question isn't *what will NPV be* — it's *how much could you be wrong, and which input is hurting you the most?*"

---

### Scene 2 — The protagonist (0:15 — 0:45)

**Visual**: split screen. Left: a cluttered SharePoint folder full of `.xlsx` files. Right: a person at a laptop, Claude Desktop open. The Claude window shows a fresh conversation.

**Voiceover**:
> "You're a risk modeller, or a finance analyst, or a consultant. You live in Excel. You know your domain cold. What you don't want is to memorise 1,400 ModelRisk functions, or click through ribbon dialogs for every distribution choice, or read three textbooks on copulas before next week's board pack."

**Cut to**: zoom into the Claude conversation. Empty composer.

**Voiceover**:
> "What you want is to describe the problem and have the modelling happen around you. That's what ModelRisk MCP is."

**Lower third fades in**: *"ModelRisk MCP — open Model Context Protocol server for Vose Software's ModelRisk."*

---

### Scene 3 — Install + first conversation (0:45 — 1:15)

**Visual**: clean PowerShell window. Type with realistic speed:

```powershell
pip install modelrisk-mcp
modelrisk-mcp install
```

**On screen**: the install output flashes by:

```
  + Claude Desktop   added    C:\Users\you\AppData\Roaming\Claude\…
  + Claude Code      added    C:\Users\you\.claude\settings.json

Restart Claude Desktop / Claude Code to pick up the new server.
```

**Voiceover**:
> "Two commands. The installer auto-wires Claude Desktop, Claude Code, and any MCP-compatible client. Backs up your existing config. Doesn't touch your other servers."

**Cut to**: Claude Desktop, fresh window. Type into the composer:

> Summarise the active workbook's risk model — inputs, outputs, distributions.

Claude responds, tool calls flickering in the right pane: `get_active_workbook`, `list_modelrisk_inputs`, `list_modelrisk_outputs`, `list_distributions`. The model summary appears.

**Voiceover (over the response)**:
> "Forty tools, five resources, five prompts. Reading, building, simulating, auditing, reporting. The first call is read-only — Claude can't write to your workbook without your say-so."

---

### Scene 4 — End-to-end workflow (1:15 — 3:00)

The headline beat — the longest scene. Walk through the construction cost-overrun example from the user manual at speed-up 1.5×.

**Visual**: Excel + Claude side by side. Claude conversation visible.

**On-screen (Claude composer)**:
> Look at the BaselineBudget.xlsx workbook. Identify the line items where overruns are the biggest risk drivers based on standard construction-project methodology.

**Voiceover (over the response stream)**:
> "Claude reads the workbook. Calls `find_hard_coded_inputs`. Looks at the cell labels — labour escalation, materials, permit delays, weather days, scope creep. Knows the methodology. Suggests which to model as bare distributions, which as risk events, which as a copula."

**Cut**: Excel cells get highlighted as Claude proposes them. Tool calls visible in Claude's right pane.

**On-screen (next message)**:
> Here's actual overrun data from our last 30 projects on the History sheet, column F. Fit a distribution and use it as the overall escalation factor in cell B45.

**Voiceover**:
> "Got real data? Hand it over. `fit_distribution_to_data` returns the best-fit lognormal with parameter uncertainty enabled — that's the difference between *we're 90% sure demand is 800 to 1200* and *the best-fit point estimate is 1000*, full stop. The second number is wrong. The server defaults to the correct one."

**Cut**: Excel cell B45 changes to:
```
=VoseInput("EscalationFactor")+VoseLognormal(0.08, 0.05, uncertainty=TRUE)
```

**On-screen (next message)**:
> For weather days and permit delays use VoseRiskEvent. Probabilities 0.4 and 0.15. Impact: triangle (5,10,20) and triangle (30,60,120) days. Convert days to dollars at $50K per day.

**Visual**: cells fill in. The Vose formulas appear with the correct bimodal `VoseRiskEvent` wrapper.

**Voiceover**:
> "Risk events use `VoseRiskEvent`, not probability times impact. The bimodal structure matters — it captures the fact that the event either fires or it doesn't, and *if* it fires, you draw from the impact distribution. That's the right shape. The wrong shape would smear the risk across every iteration and quietly understate the tail."

**Cut**: Claude composer.

**On-screen**:
> Run a 20,000-iteration sim. Then tell me the P90 cost, the P50, and which three drivers explain the most variance.

**Visual**: brief pause as ModelRisk runs the sim. Progress bar in Excel (or hidden, depending on staging). Then Claude's response renders.

**On-screen (Claude response)**:
```
NPV (5y) [Construction overrun model]:
  Mean:    $42.3M
  P50:     $42.3M
  P90:     $48.7M

Sensitivity ranking (% of output variance):
  1. EscalationFactor    54%
  2. PermitDelay         22%
  3. WeatherEvent        11%
```

**Voiceover**:
> "Twenty thousand iterations in under a minute. Mean $42.3 million. There's a ten percent probability the final cost lands above $48.7 million. The dominant driver is escalation — over half the variance — so that's the assumption to tighten first, not weather or permits."

**Cut to**: Claude composer.

**On-screen**:
> Build the executive report.

**Visual**: a new sheet appears in Excel with a multi-tile layout: KPI card, histogram with cumulative overlay, percentile table, tornado, executive narrative.

**Voiceover**:
> "Executive report — three sheets, corporate styling, ready for the steering committee. No manual chart formatting. The whole sequence took about an hour real-time. Doing it in raw ModelRisk would take a day. Doing it as a single point estimate would take twenty minutes and give you nothing."

---

### Scene 5 — What else can it do (3:00 — 4:00)

Fast montage. Each capability gets ~10 seconds of screen time.

**Visual + voiceover, each ~10s**:

1. **Audit**: `audit_model` called; findings flow into the Claude pane. "Ran inherited a workbook and need to know what's wrong? Thirteen audit rules — Vose's own methodology canon — encoded as detectors. Distribution without wrapper, fit without uncertainty, errored cell inside a VosePERT, wrong arg count. Catches what you'd miss on a manual read."

2. **Fit**: cell containing `=VoseLognormalFit(B5:B35, TRUE)` highlights. "Fit any of the twenty-plus distribution families to historical data — with parameter uncertainty enabled by default."

3. **Copulas**: cells filling in showing `=VoseCopulaBiClayton(0.4, 1)` linked to two `VoseInput`s. "Link inputs that correlate in real life. Gaussian, Student t, Clayton, Frank, Gumbel — pick what matches your data."

4. **Time series**: a column of cells showing `=VoseTimeMR(…)`. "Prices, rates, populations — anything with momentum across periods. Mean-reverting, GBM, jump-diffusion, AR(1)."

5. **Aggregates / risk events / scenarios**: quick flash of each.

**Voiceover (over the montage)**:
> "Audit. Fit. Copulas. Time series. Aggregates. Risk events. Scenarios. Every tool in the box, accessible from a conversation. No syntax reference required."

---

### Scene 6 — Why this matters (4:00 — 5:00)

**Visual**: clean slate cards, one per principle. Black text on white. Five seconds each.

1. **Open Model Context Protocol — no lock-in.** Slate.
   > "ModelRisk MCP speaks the open Model Context Protocol. Works with Claude Desktop, Claude Code, Claude for Excel, Cursor, Zed — any MCP-compatible client. MIT-licensed. Fork it."

2. **Local-only. No telemetry.** Slate.
   > "Everything runs on your machine. No data leaves your computer. The activation key is bundled and offline. Your model and your data stay where they are."

3. **Methodology-grounded.** Slate showing the 8 Vose principles.
   > "Distribution choices follow Vose's methodology canon. Fits default to `uncertainty=TRUE`. Risk events use the bimodal wrapper, not the wrong-but-common `p × impact` shortcut."

4. **Safe by default.** Slate listing the nine safety mechanisms.
   > "Every building tool defaults to dry-run. Every write lands in Excel's undo stack. Bulk writes need explicit confirmation. The writer mutex prevents two clients racing. Every commit gets logged. You can revert a change even after Excel's undo stack is empty."

5. **Excel stays the model.** Slate.
   > "No re-platforming, no shadow tooling. The workbook *is* the model — versionable in Git, openable by anyone with Excel and ModelRisk, reproducible by re-running the same simulation call. You don't trade away your infrastructure to get LLM-driven Monte Carlo."

---

### Scene 7 — CTA (5:00 — 5:30)

**Visual**: clean card with the install command in a big monospace font:

```
pip install modelrisk-mcp
modelrisk-mcp install
```

Below, in smaller text: **github.com/vosesoftware/modelrisk-mcp**

**Voiceover**:
> "Install today. Open source, MIT licensed, on PyPI now. Documentation, quick-start tutorial, glossary, full user manual at the GitHub repo. From a deterministic Excel cell to a defensible probabilistic model — start the conversation."

---

### Scene 8 — Outro (5:30 — 6:00)

**Visual**: full-screen logo, URL, and Vose Software lockup. Hold for 5 seconds. Music tail.

**No voiceover**.

---

## On-camera variant (optional, if you want a presenter)

Same beats, but Scene 2 and Scene 6 swap to a piece-to-camera with the presenter visible. Adds warmth at the cost of about 30 seconds of total runtime. Recommend filming both and A/B testing.

Presenter dress: business casual. White or light-grey background. Soft front-light, no harsh shadows. Standing or sitting both work; sitting reads more "expert", standing reads more "energetic".

## Variations worth recording later

- **`/audit-model` deep dive (3 min)** — open a workbook with deliberate methodology issues (VoseFit without uncertainty, distribution without wrapper, errored cell inside a VosePERT, wrong arity). Walk through each finding and Claude's proposed fix. Good fit for the "audit" landing page section.
- **`/import-legacy-model` demo (2 min)** — open a workbook with `RiskNormal` / `RiskTriang` from a competitor add-in. Claude maps them to Vose equivalents one cell at a time. Plays to the "lift-and-shift" narrative.
- **Claude for Excel (Office.js) demo (3 min)** — HTTP transport in action. Talk to the same MCP server from inside Excel itself. Useful for the integration partners.
- **Audit on a real customer model (5 min)** — coordinate with a customer to record a real audit pass. Highest-value piece of social proof we can produce.

## Pre-flight checklist before recording

- [ ] Restart Claude Desktop so the latest config loads
- [ ] Restart Excel from the Start menu (not programmatically) so the XLL's `xlAutoOpen` runs cleanly
- [ ] Set `hide_progress_window=true` OR position Excel so ModelRisk's progress dialog appears off-frame
- [ ] Disable Windows + Excel notifications (Focus Assist on, Excel auto-save off)
- [ ] Confirm the demo workbook is saved on disk before pressing record — no save dialog mid-take
- [ ] Run the whole sequence end-to-end once as a dry-run; check no `#NAME?` cells, no `ConcurrentWriterError`
- [ ] Capture room-tone for 10 seconds at the start in case the editor needs ambient fill
- [ ] Record one safety take of the cold open (Scene 1) — that's the shot that has to land

## Post-production checklist

- [ ] Cut to ≤6:00 final runtime
- [ ] Burn-in subtitles for accessibility (auto-generate then hand-correct)
- [ ] Render two masters: 1920×1080 (YouTube, GitHub, LinkedIn) and 1080×1080 (square, X / IG)
- [ ] Render two thumbnail candidates — one with a face, one with the headline metric ("$4.18M ± $0.74M")
- [ ] Upload to YouTube as unlisted first; share internally for review
- [ ] Update README hero spot with the embed once final
- [ ] Cross-post to LinkedIn (Vose corporate + personal), HN, MCP community Discord, r/excel + r/RiskManagement
