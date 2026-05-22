"""Executive report generator — single-sheet decision-maker dashboards.

End users want one page that tells the whole story: headline numbers,
the right two charts, a stats table, and risk callouts framed for a
decision-maker rather than a statistician. This module produces that
artefact.

Architecture:

- `ExecutiveReportBuilder` orchestrates a single-sheet layout. Adapts
  to the model's shape (number of outputs, number of inputs, magnitude
  of the headline numbers).
- Composes existing primitives (SensitivityRanking, SimulationResult,
  raw samples) into a coherent layout. No new statistical work.
- Native Excel charts via xlwings + raw COM for fine control over
  axis labels, colours, number formats.
- All formatting goes through best-effort try/except so a COM hiccup
  on a single colour set never tanks the whole report.

Layout (single sheet):
  Rows 1-3:    Title band
  Rows 5-7:    Headline numbers (mean / P5 / P50 / P95) — large, colored
  Rows 9-24:   Side-by-side charts — histogram (left) + tornado (right)
  Rows 26-34:  Statistics table for primary + secondary outputs
  Rows 36-42:  Risk callouts (auto-generated decision framing)

The exact row positions are constants on the class so they can shift
if a model has many secondary outputs (the stats table grows downward
and pushes the callouts).
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import datetime
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from modelrisk_mcp.schemas.results import (
        SensitivityRanking,
        SimulationResult,
    )


# Excel constants used via raw COM.
_XL_BAR_CLUSTERED = 57
_XL_COLUMN_CLUSTERED = 51
_XL_COLUMNS = 2          # PlotBy argument for SetSourceData (series are columns)
_XL_SHEET_VERY_HIDDEN = 2  # Worksheet.Visible value — unreachable from UI


# Single workbook-scoped sheet that holds the staging data for every
# chart in every report. xlSheetVeryHidden so the user can't see it
# from the tab strip and can't unhide it via right-click — only
# `Visible = -1` from VBA or another build_*_report call will surface
# it again. Each builder owns a distinct column block (see the
# `_HELPER_LAYOUT_*` constants on each class) to avoid stomping on its
# sibling's data when both reports coexist in the same workbook.
_HELPER_SHEET_NAME = "_ModelRiskReports"


def _last_visible_sheet(book: Any) -> Any:
    """Return the last sheet that the user can see in the tab strip.

    Bug #24 (alpha.21): `book.sheets[-1]` includes very-hidden
    sheets like `_ModelRiskReports`. Trying to add a new sheet
    `after=` a very-hidden one fails with "Move method of Worksheet
    class failed" — Excel refuses. We need a visible anchor."""
    last_visible: Any = None
    for sheet in book.sheets:
        try:
            api = sheet.api
            # xlSheetVisible = -1. Both xlSheetHidden (0) and
            # xlSheetVeryHidden (2) count as "not visible" for the
            # purpose of being a valid Move anchor.
            visible = int(getattr(api, "Visible", -1))
        except Exception:
            visible = -1
        if visible == -1:
            last_visible = sheet
    # Fallback to the first sheet if literally every sheet is hidden
    # (shouldn't happen — Excel requires at least one visible sheet —
    # but defensive code that never asserts is the rule here).
    return last_visible if last_visible is not None else book.sheets[0]


# Color palette (RGB-as-integer for xlwings api.Interior.Color).
# Using BGR encoding — Excel COM convention.
def _rgb(r: int, g: int, b: int) -> int:
    return r + (g << 8) + (b << 16)


_COLOR_TITLE_BG = _rgb(30, 60, 110)      # deep navy
_COLOR_TITLE_FG = _rgb(255, 255, 255)    # white
_COLOR_HEADLINE_NEUTRAL = _rgb(30, 30, 30)
_COLOR_HEADLINE_GOOD = _rgb(20, 130, 50)
_COLOR_HEADLINE_WARN = _rgb(200, 110, 20)
_COLOR_HEADLINE_RISK = _rgb(190, 30, 30)
_COLOR_BAND_LIGHT = _rgb(240, 243, 248)   # very light blue-gray
_COLOR_DRIVER_STRONG = _rgb(160, 30, 30)
_COLOR_DRIVER_MEDIUM = _rgb(220, 130, 40)
_COLOR_DRIVER_WEAK = _rgb(140, 140, 140)


# ---------------------------------------------------------------------------
# Result type
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ExecutiveReportResult:
    """What was built. Returned to the MCP layer so the response carries
    a summary the LLM can quote."""

    sheet_name: str
    primary_output: str
    secondary_outputs: tuple[str, ...]
    chart_count: int
    callout_count: int
    headline_summary: str  # one-line text the LLM can read aloud


# ---------------------------------------------------------------------------
# Builder
# ---------------------------------------------------------------------------


class ExecutiveReportBuilder:
    """Renders a single-sheet decision-maker dashboard.

    Composition over inheritance — this class consumes already-computed
    SimulationResult + SensitivityRanking objects and raw sample arrays.
    The caller (ModelRiskBridge) handles data gathering.
    """

    # Layout constants — all in one place so a redesign is one edit.
    TITLE_ROW = 1
    SUBTITLE_ROW = 2
    HEADLINE_LABEL_ROW = 5
    HEADLINE_VALUE_ROW = 6
    CHART_BAND_TOP = 9
    CHART_BAND_HEIGHT = 16  # rows
    STATS_TABLE_TOP = 26    # bumped if charts grow
    CALLOUT_TOP_PADDING = 2  # rows below stats table

    @staticmethod
    def build(
        book: Any,
        *,
        sheet_name: str,
        title: str,
        subtitle: str,
        primary_output: str,
        primary_result: SimulationResult,
        primary_samples: list[float],
        sensitivity: SensitivityRanking,
        secondary_results: list[SimulationResult] | None = None,
        contingency_percentile: float = 0.90,
        top_drivers: int = 5,
    ) -> ExecutiveReportResult:
        """Render the report. Idempotent — if `sheet_name` already
        exists it's replaced."""
        ExecutiveReportBuilder._remove_existing_sheet(book, sheet_name)
        # Anchor after the last VISIBLE sheet — `book.sheets[-1]` would
        # be `_ModelRiskReports` (xlSheetVeryHidden) once that helper
        # exists, and Excel refuses to move new sheets after very-
        # hidden ones (bug #24).
        sheet = book.sheets.add(sheet_name, after=_last_visible_sheet(book))

        # Adapt column widths once, before writing — the user shouldn't
        # need to manually resize anything.
        ExecutiveReportBuilder._set_column_widths(sheet)

        ExecutiveReportBuilder._write_title_band(sheet, title, subtitle)
        ExecutiveReportBuilder._write_headline_numbers(
            sheet, primary_output, primary_result, contingency_percentile,
        )
        chart_count = ExecutiveReportBuilder._add_charts(
            book,
            sheet,
            primary_output=primary_output,
            primary_samples=primary_samples,
            sensitivity=sensitivity,
            top_drivers=top_drivers,
        )
        stats_end_row = ExecutiveReportBuilder._write_stats_table(
            sheet,
            primary_result=primary_result,
            secondary_results=secondary_results or [],
        )
        callouts = ExecutiveReportBuilder._compose_callouts(
            primary_output=primary_output,
            primary_result=primary_result,
            sensitivity=sensitivity,
            contingency_percentile=contingency_percentile,
        )
        ExecutiveReportBuilder._write_callouts(
            sheet, callouts, top_row=stats_end_row + ExecutiveReportBuilder.CALLOUT_TOP_PADDING,
        )

        # Try to activate the new sheet so it's the one the user sees.
        try:
            sheet.activate()
        except Exception:
            pass

        return ExecutiveReportResult(
            sheet_name=sheet_name,
            primary_output=primary_output,
            secondary_outputs=tuple(
                r.output_name for r in (secondary_results or [])
            ),
            chart_count=chart_count,
            callout_count=len(callouts),
            headline_summary=_headline_summary(
                primary_output, primary_result, contingency_percentile,
            ),
        )

    # ----- composition ----------------------------------------------------

    @staticmethod
    def _set_column_widths(sheet: Any) -> None:
        """Layout: narrow gutter at A, content B-G, mid-gutter at H,
        secondary content I-L, narrow gutter at M. The A and M gutters
        give the report breathing room when printed or screenshotted —
        the prior 'labels in column A' layout looked cramped.

        Hand-tuned widths; the stats-table CV cell was overflowing
        as `####` at width 14, hence the bumps."""
        widths = {
            "A": 2,   # left gutter
            "B": 18,  # primary labels / Mean
            "C": 16, "D": 16, "E": 16, "F": 16,  # P5 / P50 / P-hi / StDev
            "G": 16,  # stats-table CV
            "H": 4,   # mid gutter
            "I": 22,  # output name / callouts / driver labels
            "J": 14, "K": 14, "L": 14,
            "M": 2,   # right gutter
        }
        for col, w in widths.items():
            try:
                sheet.range(f"{col}1").column_width = w
            except Exception:
                pass

    @staticmethod
    def _write_title_band(sheet: Any, title: str, subtitle: str) -> None:
        title_row = ExecutiveReportBuilder.TITLE_ROW
        sub_row = ExecutiveReportBuilder.SUBTITLE_ROW
        # Title text lives in B (the first content column post-gutter).
        sheet.range(f"B{title_row}").value = title
        sheet.range(f"B{sub_row}").value = subtitle
        try:
            # Merge across B:L for both rows so the band looks unified.
            # The A and M columns stay as narrow gutters so the band
            # gets visual breathing room on both sides.
            sheet.range(f"B{title_row}:L{title_row}").merge()
            sheet.range(f"B{sub_row}:L{sub_row}").merge()
            band = sheet.range(f"B{title_row}:L{sub_row}")
            band.api.Interior.Color = _COLOR_TITLE_BG
            band.api.Font.Color = _COLOR_TITLE_FG
            sheet.range(f"B{title_row}").api.Font.Size = 18
            sheet.range(f"B{title_row}").api.Font.Bold = True
            sheet.range(f"B{sub_row}").api.Font.Size = 11
            sheet.range(f"B{title_row}").api.HorizontalAlignment = -4108  # xlCenter
            sheet.range(f"B{sub_row}").api.HorizontalAlignment = -4108
            sheet.range(f"B{title_row}").row_height = 28
            sheet.range(f"B{sub_row}").row_height = 18
        except Exception:
            pass  # formatting is best-effort

    @staticmethod
    def _write_headline_numbers(
        sheet: Any,
        primary_output: str,
        result: SimulationResult,
        contingency_percentile: float,
    ) -> None:
        label_row = ExecutiveReportBuilder.HEADLINE_LABEL_ROW
        value_row = ExecutiveReportBuilder.HEADLINE_VALUE_ROW

        p5 = result.percentiles.get(0.05, result.min)
        p50 = result.percentiles.get(0.50, result.mean)
        p_high = result.percentiles.get(
            contingency_percentile,
            result.percentiles.get(0.95, result.max),
        )
        p_high_label = f"P{int(contingency_percentile * 100)}"

        # Labels in row `label_row` — shifted one column right post-
        # alpha.20 so A stays as a narrow gutter.
        labels = [
            ("B", "MEAN"),
            ("C", "P5 (low)"),
            ("D", "P50 (median)"),
            ("E", f"{p_high_label} (high)"),
            ("F", "STDEV"),
        ]
        for col, text in labels:
            sheet.range(f"{col}{label_row}").value = text
        # Values in row `value_row`
        values = [
            ("B", result.mean, _COLOR_HEADLINE_NEUTRAL),
            ("C", p5, _COLOR_HEADLINE_GOOD),
            ("D", p50, _COLOR_HEADLINE_NEUTRAL),
            ("E", p_high, _COLOR_HEADLINE_RISK),
            ("F", result.stdev, _color_for_cv(result.mean, result.stdev)),
        ]
        for col, val, color in values:
            sheet.range(f"{col}{value_row}").value = val
            try:
                sheet.range(f"{col}{value_row}").api.Font.Color = color
            except Exception:
                pass

        try:
            sheet.range(f"B{label_row}:F{label_row}").api.Font.Bold = True
            sheet.range(f"B{label_row}:F{label_row}").api.Font.Size = 10
            sheet.range(f"B{label_row}:F{label_row}").api.Font.Color = _rgb(100, 100, 100)
            sheet.range(f"B{value_row}:F{value_row}").api.Font.Size = 16
            sheet.range(f"B{value_row}:F{value_row}").api.Font.Bold = True
            sheet.range(f"B{value_row}:F{value_row}").api.NumberFormat = "#,##0.00"
            sheet.range(f"B{value_row}").row_height = 24
            # Light divider band underneath, spanning the content
            # area (B..L) but not the gutters.
            sheet.range(
                f"B{value_row + 1}:L{value_row + 1}"
            ).api.Interior.Color = _COLOR_BAND_LIGHT
        except Exception:
            pass

        # Output name in column I so it's visible across the gutter
        # without crowding the headline numbers.
        sheet.range(f"I{label_row}").value = "OUTPUT"
        sheet.range(f"I{value_row}").value = primary_output
        try:
            sheet.range(f"I{label_row}").api.Font.Bold = True
            sheet.range(f"I{label_row}").api.Font.Size = 10
            sheet.range(f"I{label_row}").api.Font.Color = _rgb(100, 100, 100)
            sheet.range(f"I{value_row}").api.Font.Bold = True
            sheet.range(f"I{value_row}").api.Font.Size = 14
        except Exception:
            pass

    # Helper-sheet block ownership: the executive report parks histogram
    # data in A:C and tornado data in E:F. The drivers report uses a
    # separate block (I:J) so the two can coexist without stomping on
    # each other.
    _HELPER_HISTOGRAM_COL = "A"
    _HELPER_TORNADO_COL = "E"

    @staticmethod
    def _add_charts(
        book: Any,
        sheet: Any,
        *,
        primary_output: str,
        primary_samples: list[float],
        sensitivity: SensitivityRanking,
        top_drivers: int,
    ) -> int:
        """Histogram + cumulative on the left, tornado on the right.
        Returns count of charts actually created (0-2). Staging data
        for each chart goes on the hidden helper sheet — see
        `_get_or_create_helper_sheet`."""
        chart_count = 0
        helper = _get_or_create_helper_sheet(book)
        if helper is None:
            return 0

        # 1. Histogram of the primary output's samples.
        hist_col = ExecutiveReportBuilder._HELPER_HISTOGRAM_COL
        if primary_samples:
            try:
                _clear_helper_block(
                    helper, hist_col, columns=3, rows=_HISTOGRAM_BINS + 2,
                )
                _write_histogram_data(
                    helper, primary_samples,
                    anchor_col=hist_col, anchor_row=1,
                )
                chart_count += _add_histogram_chart(
                    sheet,
                    helper=helper,
                    helper_anchor_col=hist_col,
                    bin_count=_HISTOGRAM_BINS,
                    title=f"Distribution — {primary_output}",
                    # Shifted right by ~16pt post-alpha.20 to align
                    # with column B (content starts after gutter A).
                    left=26, top=180,
                    width=400, height=240,
                )
            except Exception:
                pass

        # 2. Tornado of the top N drivers.
        tornado_col = ExecutiveReportBuilder._HELPER_TORNADO_COL
        entries = sorted(
            sensitivity.entries, key=lambda e: abs(e.correlation), reverse=True
        )[:top_drivers]
        if entries:
            try:
                _clear_helper_block(
                    helper, tornado_col, columns=2, rows=top_drivers + 2,
                )
                _write_tornado_data(
                    helper, entries,
                    anchor_col=tornado_col, anchor_row=1,
                )
                chart_count += _add_tornado_chart(
                    sheet,
                    helper=helper,
                    helper_anchor_col=tornado_col,
                    driver_count=len(entries),
                    title=f"Top drivers — {primary_output}",
                    left=440, top=180,
                    width=360, height=240,
                )
            except Exception:
                pass

        return chart_count

    @staticmethod
    def _write_stats_table(
        sheet: Any,
        *,
        primary_result: SimulationResult,
        secondary_results: list[SimulationResult],
    ) -> int:
        """Returns the last-used row so callouts can be positioned below.

        Columns post-alpha.20: B (Output) C (Mean) D (StDev) E (P5)
        F (P50) G (P95) H (CV). Column A stays as the narrow gutter."""
        top = ExecutiveReportBuilder.STATS_TABLE_TOP
        header = ["Output", "Mean", "StDev", "P5", "P50", "P95", "CV"]
        for i, label in enumerate(header):
            col = chr(ord("B") + i)  # B..H
            sheet.range(f"{col}{top}").value = label
        try:
            sheet.range(f"B{top}:H{top}").api.Font.Bold = True
            sheet.range(f"B{top}:H{top}").api.Interior.Color = _COLOR_BAND_LIGHT
            sheet.range(f"B{top}:H{top}").api.Borders(9).LineStyle = 1  # bottom border
        except Exception:
            pass

        all_results = [primary_result, *secondary_results]
        for i, r in enumerate(all_results, start=top + 1):
            cv = r.stdev / r.mean if r.mean != 0 else float("nan")
            sheet.range(f"B{i}").value = r.output_name
            sheet.range(f"C{i}").value = r.mean
            sheet.range(f"D{i}").value = r.stdev
            sheet.range(f"E{i}").value = r.percentiles.get(0.05, r.min)
            sheet.range(f"F{i}").value = r.percentiles.get(0.50, r.mean)
            sheet.range(f"G{i}").value = r.percentiles.get(0.95, r.max)
            sheet.range(f"H{i}").value = cv if math.isfinite(cv) else None
            try:
                sheet.range(f"C{i}:G{i}").api.NumberFormat = "#,##0.00"
                sheet.range(f"H{i}").api.NumberFormat = "0.000"
                # Subtle alternating-row tint for readability.
                if (i - top) % 2 == 0:
                    sheet.range(f"B{i}:H{i}").api.Interior.Color = _COLOR_BAND_LIGHT
                # Highlight high-CV rows in amber/red.
                if math.isfinite(cv):
                    if cv > 0.5:
                        sheet.range(f"H{i}").api.Font.Color = _COLOR_HEADLINE_RISK
                        sheet.range(f"H{i}").api.Font.Bold = True
                    elif cv > 0.2:
                        sheet.range(f"H{i}").api.Font.Color = _COLOR_HEADLINE_WARN
                        sheet.range(f"H{i}").api.Font.Bold = True
            except Exception:
                pass
        return top + len(all_results)

    @staticmethod
    def _compose_callouts(
        *,
        primary_output: str,
        primary_result: SimulationResult,
        sensitivity: SensitivityRanking,
        contingency_percentile: float,
    ) -> list[str]:
        """Auto-generated decision-maker callouts. Each is a single
        sentence framed in plain English."""
        callouts: list[str] = []
        p5 = primary_result.percentiles.get(0.05, primary_result.min)
        p95 = primary_result.percentiles.get(0.95, primary_result.max)
        p_hi_pct = int(contingency_percentile * 100)
        p_hi = primary_result.percentiles.get(
            contingency_percentile, primary_result.max,
        )
        mean = primary_result.mean

        # Range statement
        callouts.append(
            f"90% confidence: {primary_output} lands between "
            f"{p5:,.2f} (P5) and {p95:,.2f} (P95). Mean is {mean:,.2f}."
        )

        # Tail risk statement
        if p_hi > mean and mean != 0:
            tail_pct = (p_hi - mean) / abs(mean) * 100
            callouts.append(
                f"Tail risk: P{p_hi_pct} is {p_hi:,.2f}, "
                f"{tail_pct:+.0f}% above the mean. Plan capacity to handle this."
            )

        # Volatility statement
        if primary_result.mean != 0:
            cv = primary_result.stdev / abs(primary_result.mean)
            if cv > 0.5:
                callouts.append(
                    f"High volatility: CV is {cv:.2f}. Outcomes are "
                    "highly dispersed — sensitivity analysis is essential before committing."
                )
            elif cv > 0.2:
                callouts.append(
                    f"Moderate volatility: CV is {cv:.2f}. Range of "
                    "outcomes is wide enough to warrant a contingency reserve."
                )

        # Top driver statement
        if sensitivity.entries:
            top = sensitivity.entries[0]
            direction = "raises" if top.correlation > 0 else "lowers"
            magnitude = (
                "strongly"
                if abs(top.correlation) > 0.7
                else "moderately"
                if abs(top.correlation) > 0.3
                else "weakly"
            )
            callouts.append(
                f"Primary driver: {top.input_name} {magnitude} {direction} "
                f"{primary_output} (r = {top.correlation:+.2f}). "
                "Focus mitigation here."
            )

        return callouts

    @staticmethod
    def _write_callouts(
        sheet: Any, callouts: list[str], *, top_row: int,
    ) -> None:
        """Callouts live in column B onwards; column A is the gutter."""
        if not callouts:
            return
        sheet.range(f"B{top_row}").value = "RISK CALLOUTS"
        try:
            sheet.range(f"B{top_row}").api.Font.Bold = True
            sheet.range(f"B{top_row}").api.Font.Size = 11
            sheet.range(f"B{top_row}").api.Font.Color = _rgb(100, 100, 100)
        except Exception:
            pass
        for i, callout in enumerate(callouts, start=top_row + 1):
            sheet.range(f"B{i}").value = f"•  {callout}"
            try:
                # Merge across the content band (B..L) so long callouts
                # wrap nicely. A and M stay as gutters.
                sheet.range(f"B{i}:L{i}").merge()
                sheet.range(f"B{i}").api.WrapText = True
                sheet.range(f"B{i}").row_height = 24
            except Exception:
                pass

    @staticmethod
    def _remove_existing_sheet(book: Any, name: str) -> None:
        for sheet in list(book.sheets):
            if sheet.name == name:
                try:
                    sheet.delete()
                except Exception:
                    pass
                return


# ---------------------------------------------------------------------------
# Histogram helpers
# ---------------------------------------------------------------------------


_HISTOGRAM_BINS = 30


# ---------------------------------------------------------------------------
# Hidden helper sheet + chart binding
# ---------------------------------------------------------------------------


def _get_or_create_helper_sheet(book: Any) -> Any:
    """Return the workbook-scoped hidden helper sheet, creating it (and
    marking it `xlSheetVeryHidden`) on first use.

    Why a hidden helper sheet at all: the histogram / tornado charts
    need a tabular source range. Putting that data on the visible
    report sheet (the previous design's M:Q columns) bled into the
    user's print/scroll experience even after `EntireColumn.Hidden`,
    and broke down further when the binding step itself silently
    failed (bug #18). The helper sheet hard-isolates the staging data
    from the user-visible report so the visible sheet contains only
    title, headline numbers, the rendered charts, the stats table, and
    callouts. Nothing else."""
    for sheet in book.sheets:
        if sheet.name == _HELPER_SHEET_NAME:
            return sheet
    helper = book.sheets.add(
        _HELPER_SHEET_NAME, after=_last_visible_sheet(book),
    )
    try:
        helper.api.Visible = _XL_SHEET_VERY_HIDDEN
    except Exception:
        # Falls back to a normal sheet (still hidden behind the report
        # tab order) — the report still renders correctly even if the
        # helper is visible in the tab strip.
        pass
    return helper


def _clear_helper_block(
    helper: Any, anchor_col: str, columns: int, rows: int,
) -> None:
    """Wipe a staging block on the helper sheet before re-writing it.

    Without this, a second `build_*_report` call could leave stale
    rows below the new data, and Excel's chart binding would silently
    extend to the longer range (mixing fresh + stale)."""
    last_col = chr(ord(anchor_col) + columns - 1)
    try:
        helper.range(f"{anchor_col}1:{last_col}{rows}").clear_contents()
    except Exception:
        pass


def _bind_chart_to_range(
    chart: Any,
    source_range: Any,
    *,
    plot_by: int = _XL_COLUMNS,
) -> bool:
    """Bind `chart` to `source_range` via the COM SetSourceData call.

    Bug #18 (alpha.15): xlwings' `Chart.set_source_data(range)` was
    observed to leave the chart's `SeriesCollection(1).Formula` empty
    on real workbooks even though no exception was raised — the chart
    rendered briefly while Excel auto-populated a default series, then
    went blank because the bind step silently failed. Calling
    `chart_api.SetSourceData(source.api, PlotBy)` directly via COM
    works around it.

    Returns True iff the binding produces a non-empty
    SeriesCollection(1).Formula afterwards (our "the bind actually
    stuck" probe). False signals the caller can attempt a fallback
    or surface a warning."""
    chart_api = chart.api[1] if isinstance(chart.api, tuple) else chart.api
    source_api = source_range.api
    try:
        chart_api.SetSourceData(Source=source_api, PlotBy=plot_by)
    except Exception:
        return False
    # Verification probe — read the formula back. If it's empty,
    # the bind silently dropped (the bug #18 symptom) and we report
    # failure so the caller knows. If the probe itself errors (some
    # COM proxies don't expose .Formula cleanly), assume success
    # rather than reporting a false negative.
    try:
        formula = str(chart_api.SeriesCollection(1).Formula or "")
        return bool(formula.strip())
    except Exception:
        return True


def _write_histogram_data(
    sheet: Any, samples: list[float], *, anchor_col: str, anchor_row: int,
) -> None:
    """Compute histogram counts and write bin centres / counts /
    cumulative as three columns starting at anchor on the provided
    sheet (typically the hidden helper sheet)."""
    if not samples:
        return
    lo, hi = min(samples), max(samples)
    if hi == lo:
        hi = lo + 1.0
    bin_w = (hi - lo) / _HISTOGRAM_BINS
    counts = [0] * _HISTOGRAM_BINS
    for s in samples:
        idx = min(_HISTOGRAM_BINS - 1, max(0, int((s - lo) / bin_w)))
        counts[idx] += 1
    # Write header in anchor row, data starting next row.
    headers = ["Bin", "Count", "Cumulative %"]
    for i, h in enumerate(headers):
        col = chr(ord(anchor_col) + i)
        sheet.range(f"{col}{anchor_row}").value = h
    cumulative = 0
    total = len(samples)
    for j in range(_HISTOGRAM_BINS):
        bin_centre = lo + (j + 0.5) * bin_w
        cumulative += counts[j]
        row = anchor_row + 1 + j
        sheet.range(f"{anchor_col}{row}").value = bin_centre
        sheet.range(f"{chr(ord(anchor_col) + 1)}{row}").value = counts[j]
        sheet.range(f"{chr(ord(anchor_col) + 2)}{row}").value = cumulative / total


def _add_histogram_chart(
    sheet: Any,
    *,
    helper: Any,
    helper_anchor_col: str,
    bin_count: int,
    title: str,
    left: int, top: int, width: int, height: int,
) -> int:
    """Add a histogram (column) + cumulative (line) chart on `sheet`,
    bound to the helper-sheet range starting at `helper_anchor_col`.
    Returns 1 on success (and the binding actually stuck), 0 if the
    chart object couldn't be created or the COM bind failed.

    Layout assumption: helper sheet has a 3-column block at
    `helper_anchor_col` — bin centres in col[0], counts in col[1],
    cumulative % in col[2], all with a header row at row 1. The chart
    plots Count + Cumulative as data series and uses Bin as the
    X-axis categories.

    Bug #18b (alpha.20) — prior versions used `SetSourceData` against
    the whole 3-column range, which made Excel turn EVERY column into
    a data series. The visible result was a chart with three labelled
    series ("Bin", "Count", "Cumulative %") where Bin's numeric values
    were being plotted as bars instead of forming the X-axis. Fix:
    bind only the Count + Cumulative columns, then explicitly assign
    XValues on both series to point at the Bin column."""
    try:
        chart = sheet.charts.add(left=left, top=top, width=width, height=height)
    except Exception:
        return 0
    bin_col = helper_anchor_col
    count_col = chr(ord(helper_anchor_col) + 1)
    cum_col = chr(ord(helper_anchor_col) + 2)
    # Bind to ONLY the Count + Cumulative columns. The Bin column
    # becomes the category axis below, not a data series.
    data_range = helper.range(
        f"{count_col}1:{cum_col}{1 + bin_count}"
    )
    try:
        chart_api = chart.api[1] if isinstance(chart.api, tuple) else chart.api
        chart_api.ChartType = _XL_COLUMN_CLUSTERED
    except Exception:
        pass
    bound = _bind_chart_to_range(chart, data_range, plot_by=_XL_COLUMNS)
    if not bound:
        return 0
    # Assign bin centres as the X-axis for both series. Without this
    # the chart's X-axis labels would be 1..N (the row index) instead
    # of the bin midpoints — which is what the prior version showed.
    x_range = helper.range(f"{bin_col}2:{bin_col}{1 + bin_count}")
    try:
        chart_api = chart.api[1] if isinstance(chart.api, tuple) else chart.api
        chart_api.HasTitle = True
        chart_api.ChartTitle.Text = title
        for i in (1, 2):  # series 1 = Count, series 2 = Cumulative %
            try:
                chart_api.SeriesCollection(i).XValues = x_range.api
            except Exception:
                pass
        # Make the cumulative series a line on a secondary axis.
        try:
            series = chart_api.SeriesCollection(2)
            series.ChartType = 4  # xlLine
            series.AxisGroup = 2  # secondary axis
        except Exception:
            pass
        chart.name = f"Histogram_{title[:20]}"[:31]
    except Exception:
        pass
    return 1


def _write_tornado_data(
    sheet: Any, entries: list[Any], *, anchor_col: str, anchor_row: int,
) -> None:
    """Two-column layout: name, correlation. Written to the helper
    sheet at the given anchor; the chart binds to this range."""
    sheet.range(f"{anchor_col}{anchor_row}").value = "Input"
    sheet.range(f"{chr(ord(anchor_col) + 1)}{anchor_row}").value = "Correlation"
    for i, e in enumerate(entries, start=anchor_row + 1):
        sheet.range(f"{anchor_col}{i}").value = e.input_name
        sheet.range(f"{chr(ord(anchor_col) + 1)}{i}").value = e.correlation


def _add_tornado_chart(
    sheet: Any,
    *,
    helper: Any,
    helper_anchor_col: str,
    driver_count: int,
    title: str,
    left: int, top: int, width: int, height: int,
) -> int:
    """Add a horizontal-bar tornado chart on `sheet`, bound to the
    helper-sheet range starting at `helper_anchor_col`. Returns 1 on
    success, 0 if either the chart creation or the COM bind failed."""
    try:
        chart = sheet.charts.add(left=left, top=top, width=width, height=height)
    except Exception:
        return 0
    last_col = chr(ord(helper_anchor_col) + 1)
    source_range = helper.range(
        f"{helper_anchor_col}1:{last_col}{1 + driver_count}"
    )
    try:
        chart_api = chart.api[1] if isinstance(chart.api, tuple) else chart.api
        chart_api.ChartType = _XL_BAR_CLUSTERED
    except Exception:
        pass
    bound = _bind_chart_to_range(chart, source_range, plot_by=_XL_COLUMNS)
    if not bound:
        return 0
    try:
        chart_api = chart.api[1] if isinstance(chart.api, tuple) else chart.api
        chart_api.HasTitle = True
        chart_api.ChartTitle.Text = title
        # Reverse plot order so the strongest driver is at the top
        # (tornado convention).
        try:
            category_axis = chart_api.Axes(1)
            category_axis.ReversePlotOrder = True
        except Exception:
            pass
        chart.name = f"Tornado_{title[:20]}"[:31]
    except Exception:
        pass
    return 1


# ---------------------------------------------------------------------------
# Adaptive color logic
# ---------------------------------------------------------------------------


def _color_for_cv(mean: float, stdev: float) -> int:
    """CV-based volatility color for the stdev cell."""
    if mean == 0:
        return _COLOR_HEADLINE_NEUTRAL
    cv = stdev / abs(mean)
    if cv > 0.5:
        return _COLOR_HEADLINE_RISK
    if cv > 0.2:
        return _COLOR_HEADLINE_WARN
    return _COLOR_HEADLINE_GOOD


def _headline_summary(
    primary_output: str,
    result: SimulationResult,
    contingency_percentile: float,
) -> str:
    """One-liner the LLM can quote back to the user."""
    p_hi = result.percentiles.get(contingency_percentile, result.max)
    p_hi_label = f"P{int(contingency_percentile * 100)}"
    return (
        f"{primary_output}: mean {result.mean:,.2f}, "
        f"stdev {result.stdev:,.2f}, {p_hi_label} {p_hi:,.2f}"
    )


# ---------------------------------------------------------------------------
# Time helper — used by the default subtitle
# ---------------------------------------------------------------------------


def default_subtitle(samples: int, seed: int | None = None) -> str:
    parts = [
        f"{samples:,} iterations",
        datetime.now().strftime("%Y-%m-%d"),
    ]
    if seed is not None:
        parts.insert(1, f"seed {seed}")
    return " · ".join(parts)


# ---------------------------------------------------------------------------
# Drivers report — narrower than ExecutiveReport, focused on
# sensitivity. Includes auto-generated plain-English narrative
# explaining *which* inputs matter and *how to read* the tornado.
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class DriversReportResult:
    """Returned to the MCP layer so the LLM can quote the headline
    finding without re-reading the sheet."""

    sheet_name: str
    output_name: str
    drivers_analyzed: int
    top_driver: str | None
    top_correlation: float | None
    concentration: str  # "concentrated" | "moderate" | "diffuse"
    headline_finding: str  # one-line summary the LLM can quote


class DriversReportBuilder:
    """Single-sheet sensitivity report that explains the tornado in
    plain English.

    The added value over `TornadoChartWriter` is the narrative:
    auto-generated key findings, a 'how to read this chart' panel,
    and tiered recommendations. Designed for decision-makers who
    don't know what Spearman correlation means.
    """

    TITLE_ROW = 1
    SUBTITLE_ROW = 2
    FINDINGS_HEADER_ROW = 4
    FINDINGS_FIRST_ROW = 5
    CHART_BAND_TOP_ROW = 11
    TABLE_HEADER_ROW = 11
    TABLE_DATA_ROW = 12
    EXPLAIN_HEADER_ROW = 33
    RECOMMEND_HEADER_ROW = 40

    @staticmethod
    def build(
        book: Any,
        *,
        sheet_name: str,
        output_name: str,
        sensitivity: SensitivityRanking,
        iterations: int,
        title: str | None = None,
        subtitle: str | None = None,
    ) -> DriversReportResult:
        """Render the drivers report. Idempotent — replaces an
        existing sheet of the same name."""
        ExecutiveReportBuilder._remove_existing_sheet(book, sheet_name)
        sheet = book.sheets.add(sheet_name, after=_last_visible_sheet(book))

        DriversReportBuilder._set_column_widths(sheet)

        effective_title = title or f"Uncertainty Drivers — {output_name}"
        effective_subtitle = subtitle or (
            f"Sensitivity analysis · {iterations:,} iterations · "
            f"{datetime.now().strftime('%Y-%m-%d')}"
        )
        DriversReportBuilder._write_title_band(
            sheet, effective_title, effective_subtitle,
        )

        # Sort entries by |correlation| descending so all downstream
        # logic sees the strongest-first ordering.
        entries = sorted(
            sensitivity.entries,
            key=lambda e: abs(e.correlation),
            reverse=True,
        )

        findings = _compose_findings(output_name, entries)
        DriversReportBuilder._write_findings(sheet, findings)

        DriversReportBuilder._write_tornado_chart(book, sheet, output_name, entries)
        DriversReportBuilder._write_driver_table(sheet, entries)

        DriversReportBuilder._write_chart_explanation(sheet)
        recommendations = _compose_recommendations(entries)
        DriversReportBuilder._write_recommendations(sheet, recommendations)

        try:
            sheet.activate()
        except Exception:
            pass

        top = entries[0] if entries else None
        headline = _drivers_headline(output_name, entries)
        return DriversReportResult(
            sheet_name=sheet_name,
            output_name=output_name,
            drivers_analyzed=len(entries),
            top_driver=top.input_name if top else None,
            top_correlation=top.correlation if top else None,
            concentration=_concentration_label(entries),
            headline_finding=headline,
        )

    # ----- composition ----------------------------------------------------

    @staticmethod
    def _set_column_widths(sheet: Any) -> None:
        """alpha.22 layout: narrow gutters at A and M, content in B-L.
        Driver table lives at H-K (was G-J). Match the executive
        report's gutter pattern so the two reports feel like siblings."""
        widths = {
            "A": 2,   # left gutter
            "B": 24,  # findings + recommendation labels
            "C": 14, "D": 14, "E": 14, "F": 14,
            "G": 4,   # mid gutter
            "H": 22,  # driver-table Input column
            "I": 14, "J": 14, "K": 14, "L": 14,
            "M": 2,   # right gutter
        }
        for col, w in widths.items():
            try:
                sheet.range(f"{col}1").column_width = w
            except Exception:
                pass

    @staticmethod
    def _write_title_band(sheet: Any, title: str, subtitle: str) -> None:
        sheet.range("B1").value = title
        sheet.range("B2").value = subtitle
        try:
            # Band spans B:L; A and M stay as narrow gutters.
            sheet.range("B1:L1").merge()
            sheet.range("B2:L2").merge()
            band = sheet.range("B1:L2")
            band.api.Interior.Color = _COLOR_TITLE_BG
            band.api.Font.Color = _COLOR_TITLE_FG
            sheet.range("B1").api.Font.Size = 18
            sheet.range("B1").api.Font.Bold = True
            sheet.range("B2").api.Font.Size = 11
            sheet.range("B1").api.HorizontalAlignment = -4108
            sheet.range("B2").api.HorizontalAlignment = -4108
            sheet.range("B1").row_height = 28
            sheet.range("B2").row_height = 18
        except Exception:
            pass

    @staticmethod
    def _write_findings(sheet: Any, findings: list[str]) -> None:
        sheet.range(f"B{DriversReportBuilder.FINDINGS_HEADER_ROW}").value = (
            "KEY FINDINGS"
        )
        try:
            r = DriversReportBuilder.FINDINGS_HEADER_ROW
            sheet.range(f"B{r}:L{r}").api.Font.Bold = True
            sheet.range(f"B{r}").api.Font.Size = 12
            sheet.range(f"B{r}").api.Font.Color = _rgb(100, 100, 100)
        except Exception:
            pass
        for i, finding in enumerate(findings, start=DriversReportBuilder.FINDINGS_FIRST_ROW):
            sheet.range(f"B{i}").value = f"•  {finding}"
            try:
                sheet.range(f"B{i}:L{i}").merge()
                sheet.range(f"B{i}").api.WrapText = True
                sheet.range(f"B{i}").row_height = 28
            except Exception:
                pass

    # Helper-sheet block ownership: drivers tornado lives at I:J on the
    # hidden helper sheet (executive report uses A:C and E:F).
    _HELPER_TORNADO_COL = "I"

    @staticmethod
    def _write_tornado_chart(
        book: Any, sheet: Any, output_name: str, entries: list[Any],
    ) -> None:
        """Render the prominent driver tornado on `sheet` with its
        source data on the hidden helper sheet (bug #19 — staging
        data must not leak onto the user-visible report). Binding
        goes through `_bind_chart_to_range` so the connection is
        verified (bug #18)."""
        if not entries:
            return
        helper = _get_or_create_helper_sheet(book)
        if helper is None:
            return
        anchor = DriversReportBuilder._HELPER_TORNADO_COL
        try:
            _clear_helper_block(
                helper, anchor, columns=2, rows=len(entries) + 2,
            )
            _write_tornado_data(
                helper, entries, anchor_col=anchor, anchor_row=1,
            )
            _add_tornado_chart(
                sheet,
                helper=helper,
                helper_anchor_col=anchor,
                driver_count=len(entries),
                title=f"What moves {output_name}",
                # Shifted right ~16pt post-alpha.22 to align with the
                # column-B content start (A is now the gutter).
                left=26, top=210, width=440, height=320,
            )
        except Exception:
            pass

    @staticmethod
    def _write_driver_table(sheet: Any, entries: list[Any]) -> None:
        # alpha.22 layout: table moves from G:J to H:K so it lives
        # past the mid-gutter at column G.
        header_row = DriversReportBuilder.TABLE_HEADER_ROW
        headers = ["Input", "Correlation (r)", "|r|", "Variance share"]
        for i, label in enumerate(headers):
            col = chr(ord("H") + i)
            sheet.range(f"{col}{header_row}").value = label
        try:
            sheet.range(f"H{header_row}:K{header_row}").api.Font.Bold = True
            sheet.range(
                f"H{header_row}:K{header_row}"
            ).api.Interior.Color = _COLOR_BAND_LIGHT
        except Exception:
            pass

        for idx, e in enumerate(entries, start=DriversReportBuilder.TABLE_DATA_ROW):
            sheet.range(f"H{idx}").value = e.input_name
            sheet.range(f"I{idx}").value = e.correlation
            sheet.range(f"J{idx}").value = abs(e.correlation)
            sheet.range(f"K{idx}").value = _variance_share(e.correlation)
            try:
                sheet.range(f"I{idx}").api.NumberFormat = "0.000"
                sheet.range(f"J{idx}").api.NumberFormat = "0.000"
                sheet.range(f"K{idx}").api.NumberFormat = "0.0%"
                # Color the |r| cell by strength tier.
                col_strength = _driver_strength_color(abs(e.correlation))
                sheet.range(f"J{idx}").api.Font.Color = col_strength
                sheet.range(f"J{idx}").api.Font.Bold = True
            except Exception:
                pass

    @staticmethod
    def _write_chart_explanation(sheet: Any) -> None:
        row = DriversReportBuilder.EXPLAIN_HEADER_ROW
        sheet.range(f"B{row}").value = "HOW TO READ THIS CHART"
        try:
            sheet.range(f"B{row}").api.Font.Bold = True
            sheet.range(f"B{row}").api.Font.Size = 12
            sheet.range(f"B{row}").api.Font.Color = _rgb(100, 100, 100)
        except Exception:
            pass
        paragraphs = [
            (
                "Each bar shows how strongly one input moves the output. "
                "We use Spearman rank correlation: a value of +1 means the "
                "input perfectly pushes the output up; -1 means it pushes "
                "the output down; 0 means no effect."
            ),
            (
                "Bars further from zero matter more. A bar at +0.7 means "
                "this input is a strong upside driver; at -0.7 a strong "
                "downside driver. Bars near zero are noise — those inputs "
                "could change without meaningful impact on the result."
            ),
            (
                "The 'Variance share' column is roughly how much of the "
                "output's total variation each input alone explains. The "
                "top few drivers usually account for most of the variance; "
                "that's where mitigation has the biggest payoff."
            ),
        ]
        for i, paragraph in enumerate(paragraphs, start=row + 1):
            sheet.range(f"B{i}").value = paragraph
            try:
                sheet.range(f"B{i}:L{i}").merge()
                sheet.range(f"B{i}").api.WrapText = True
                sheet.range(f"B{i}").row_height = 32
                sheet.range(f"B{i}").api.VerticalAlignment = -4160  # xlTop
            except Exception:
                pass

    @staticmethod
    def _write_recommendations(
        sheet: Any, recommendations: dict[str, list[str]],
    ) -> None:
        row = DriversReportBuilder.RECOMMEND_HEADER_ROW
        sheet.range(f"B{row}").value = "RECOMMENDED ACTIONS"
        try:
            sheet.range(f"B{row}").api.Font.Bold = True
            sheet.range(f"B{row}").api.Font.Size = 12
            sheet.range(f"B{row}").api.Font.Color = _rgb(100, 100, 100)
        except Exception:
            pass
        tier_rows = [
            ("Focus mitigation on:", "focus", _COLOR_DRIVER_STRONG),
            ("Monitor:", "monitor", _COLOR_DRIVER_MEDIUM),
            ("Can be deprioritised:", "deprioritise", _COLOR_DRIVER_WEAK),
        ]
        for i, (label, key, color) in enumerate(tier_rows, start=row + 1):
            sheet.range(f"B{i}").value = label
            inputs = recommendations.get(key, [])
            value = ", ".join(inputs) if inputs else "(none)"
            sheet.range(f"C{i}").value = value
            try:
                sheet.range(f"B{i}").api.Font.Bold = True
                sheet.range(f"C{i}").api.Font.Color = color
                sheet.range(f"C{i}").api.Font.Bold = True
                sheet.range(f"C{i}:L{i}").merge()
                sheet.range(f"C{i}").api.WrapText = True
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Narrative-generation helpers
# ---------------------------------------------------------------------------


def _strength_label(corr: float) -> str:
    a = abs(corr)
    if a >= 0.7:
        return "dominant"
    if a >= 0.4:
        return "strong"
    if a >= 0.2:
        return "moderate"
    return "weak"


def _driver_strength_color(abs_corr: float) -> int:
    if abs_corr >= 0.4:
        return _COLOR_DRIVER_STRONG
    if abs_corr >= 0.2:
        return _COLOR_DRIVER_MEDIUM
    return _COLOR_DRIVER_WEAK


def _variance_share(corr: float) -> float:
    """Approximate fraction of output variance an input alone explains.
    Spearman r² is rank-based and isn't a literal variance decomposition,
    but it's a defensible decision-maker framing for relative importance."""
    return corr * corr


def _concentration_label(entries: list[Any]) -> str:
    """How concentrated the risk profile is. Top-3 share of total |r|
    is a robust signal:
      - >= 0.75  → concentrated (few drivers matter)
      - 0.5-0.75 → moderate
      - < 0.5    → diffuse (no single point of intervention)
    """
    if not entries:
        return "diffuse"
    total = sum(abs(e.correlation) for e in entries)
    if total == 0:
        return "diffuse"
    top3 = sum(abs(e.correlation) for e in entries[:3])
    share = top3 / total
    if share >= 0.75:
        return "concentrated"
    if share >= 0.5:
        return "moderate"
    return "diffuse"


def _compose_findings(
    output_name: str, entries: list[Any],
) -> list[str]:
    """Auto-generate 3-5 plain-English findings from the sensitivity
    ranking. Each is a single sentence framed for a decision-maker."""
    findings: list[str] = []
    if not entries:
        findings.append(
            f"No drivers were found for {output_name}. This is unusual — "
            "check that the workbook has VoseInput cells and that a "
            "simulation has been run."
        )
        return findings

    top = entries[0]
    strength = _strength_label(top.correlation)
    direction_phrase = (
        "raises" if top.correlation > 0 else "lowers"
    )
    findings.append(
        f"The {strength} driver of {output_name} is "
        f"{top.input_name} (r = {top.correlation:+.2f}). "
        f"A higher {top.input_name} {direction_phrase} {output_name}."
    )

    # Top-N coverage of variance share.
    top_n = entries[:3]
    coverage = sum(_variance_share(e.correlation) for e in top_n) * 100
    if len(top_n) > 1:
        names = ", ".join(e.input_name for e in top_n)
        findings.append(
            f"The top {len(top_n)} drivers together account for "
            f"approximately {coverage:.0f}% of {output_name}'s variance: "
            f"{names}."
        )

    # Concentration framing
    concentration = _concentration_label(entries)
    if concentration == "concentrated":
        findings.append(
            "Risk is concentrated — most of the uncertainty comes from a "
            "small number of inputs. Mitigation effort can focus narrowly "
            "with high payoff."
        )
    elif concentration == "diffuse":
        findings.append(
            "Risk is diffuse — no single input dominates, so no single "
            "intervention will substantially narrow the outcome range. "
            "Consider portfolio-level risk management rather than "
            "input-by-input mitigation."
        )

    # Weak-input callout — if any input has |r| < 0.15, name them so
    # decision-makers know what they can safely ignore.
    weak = [e for e in entries if abs(e.correlation) < 0.15]
    if weak and len(weak) < len(entries):
        weak_names = ", ".join(e.input_name for e in weak[:3])
        more = f" and {len(weak) - 3} more" if len(weak) > 3 else ""
        findings.append(
            f"Several inputs have negligible influence on {output_name} "
            f"(|r| < 0.15): {weak_names}{more}. These can be deprioritised "
            "in scenario planning."
        )

    return findings


def _compose_recommendations(
    entries: list[Any],
) -> dict[str, list[str]]:
    """Tier inputs into focus / monitor / deprioritise based on
    |correlation|."""
    focus, monitor, deprioritise = [], [], []
    for e in entries:
        a = abs(e.correlation)
        if a >= 0.4:
            focus.append(e.input_name)
        elif a >= 0.2:
            monitor.append(e.input_name)
        else:
            deprioritise.append(e.input_name)
    return {
        "focus": focus,
        "monitor": monitor,
        "deprioritise": deprioritise,
    }


def _drivers_headline(output_name: str, entries: list[Any]) -> str:
    """One-line summary for the LLM to quote."""
    if not entries:
        return f"No drivers identified for {output_name}."
    top = entries[0]
    strength = _strength_label(top.correlation)
    return (
        f"{top.input_name} is the {strength} driver of {output_name} "
        f"(r = {top.correlation:+.2f}); {len(entries)} inputs analyzed."
    )


__all__ = [
    "DriversReportBuilder",
    "DriversReportResult",
    "ExecutiveReportBuilder",
    "ExecutiveReportResult",
    "default_subtitle",
]
