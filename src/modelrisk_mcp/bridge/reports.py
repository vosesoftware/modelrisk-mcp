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
        sheet = book.sheets.add(sheet_name, after=book.sheets[-1])

        # Adapt column widths once, before writing — the user shouldn't
        # need to manually resize anything.
        ExecutiveReportBuilder._set_column_widths(sheet)

        ExecutiveReportBuilder._write_title_band(sheet, title, subtitle)
        ExecutiveReportBuilder._write_headline_numbers(
            sheet, primary_output, primary_result, contingency_percentile,
        )
        chart_count = ExecutiveReportBuilder._add_charts(
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
        """Wide enough for monetary numbers + labels. Tuned by hand."""
        widths = {
            "A": 18,  # labels
            "B": 16, "C": 16, "D": 16, "E": 16,  # headline / stats numbers
            "F": 4,   # gutter
            "G": 22,  # callouts / driver labels
            "H": 14,
            "I": 14,
            "J": 14,
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
        sheet.range(f"A{title_row}").value = title
        sheet.range(f"A{sub_row}").value = subtitle
        try:
            # Merge across A:J for both rows so the band looks unified.
            sheet.range(f"A{title_row}:J{title_row}").merge()
            sheet.range(f"A{sub_row}:J{sub_row}").merge()
            band = sheet.range(f"A{title_row}:J{sub_row}")
            band.api.Interior.Color = _COLOR_TITLE_BG
            band.api.Font.Color = _COLOR_TITLE_FG
            sheet.range(f"A{title_row}").api.Font.Size = 18
            sheet.range(f"A{title_row}").api.Font.Bold = True
            sheet.range(f"A{sub_row}").api.Font.Size = 11
            sheet.range(f"A{title_row}").api.HorizontalAlignment = -4108  # xlCenter
            sheet.range(f"A{sub_row}").api.HorizontalAlignment = -4108
            sheet.range(f"A{title_row}").row_height = 28
            sheet.range(f"A{sub_row}").row_height = 18
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

        # Labels in row `label_row`
        labels = [
            ("A", "MEAN"),
            ("B", "P5 (low)"),
            ("C", "P50 (median)"),
            ("D", f"{p_high_label} (high)"),
            ("E", "STDEV"),
        ]
        for col, text in labels:
            sheet.range(f"{col}{label_row}").value = text
        # Values in row `value_row`
        values = [
            ("A", result.mean, _COLOR_HEADLINE_NEUTRAL),
            ("B", p5, _COLOR_HEADLINE_GOOD),
            ("C", p50, _COLOR_HEADLINE_NEUTRAL),
            ("D", p_high, _COLOR_HEADLINE_RISK),
            ("E", result.stdev, _color_for_cv(result.mean, result.stdev)),
        ]
        for col, val, color in values:
            sheet.range(f"{col}{value_row}").value = val
            try:
                sheet.range(f"{col}{value_row}").api.Font.Color = color
            except Exception:
                pass

        try:
            sheet.range(f"A{label_row}:E{label_row}").api.Font.Bold = True
            sheet.range(f"A{label_row}:E{label_row}").api.Font.Size = 10
            sheet.range(f"A{label_row}:E{label_row}").api.Font.Color = _rgb(100, 100, 100)
            sheet.range(f"A{value_row}:E{value_row}").api.Font.Size = 16
            sheet.range(f"A{value_row}:E{value_row}").api.Font.Bold = True
            sheet.range(f"A{value_row}:E{value_row}").api.NumberFormat = "#,##0.00"
            sheet.range(f"A{value_row}").row_height = 24
            # Light divider band underneath
            sheet.range(
                f"A{value_row + 1}:J{value_row + 1}"
            ).api.Interior.Color = _COLOR_BAND_LIGHT
        except Exception:
            pass

        # Output name in column G so it's visible without crowding.
        sheet.range(f"G{label_row}").value = "OUTPUT"
        sheet.range(f"G{value_row}").value = primary_output
        try:
            sheet.range(f"G{label_row}").api.Font.Bold = True
            sheet.range(f"G{label_row}").api.Font.Color = _rgb(100, 100, 100)
            sheet.range(f"G{value_row}").api.Font.Bold = True
            sheet.range(f"G{value_row}").api.Font.Size = 14
        except Exception:
            pass

    @staticmethod
    def _add_charts(
        sheet: Any,
        *,
        primary_output: str,
        primary_samples: list[float],
        sensitivity: SensitivityRanking,
        top_drivers: int,
    ) -> int:
        """Histogram + cumulative on the left, tornado on the right.
        Returns count of charts actually created (0-2)."""
        chart_count = 0
        # The hidden-data approach: write histogram bin/count pairs and
        # tornado driver/correlation pairs into rows far to the right,
        # then reference those ranges for chart source data.
        top = ExecutiveReportBuilder.CHART_BAND_TOP
        height = ExecutiveReportBuilder.CHART_BAND_HEIGHT

        # 1. Histogram of the primary output's samples.
        if primary_samples:
            try:
                _write_histogram_data(sheet, primary_samples, anchor_col="M", anchor_row=1)
                chart_count += _add_histogram_chart(
                    sheet,
                    bin_count=_HISTOGRAM_BINS,
                    title=f"Distribution — {primary_output}",
                    left=10, top=180,  # rough point coords; xlwings uses raw points
                    width=380, height=220,
                )
            except Exception:
                pass

        # 2. Tornado of the top N drivers.
        entries = sorted(
            sensitivity.entries, key=lambda e: abs(e.correlation), reverse=True
        )[:top_drivers]
        if entries:
            try:
                _write_tornado_data(sheet, entries, anchor_col="P", anchor_row=1)
                chart_count += _add_tornado_chart(
                    sheet,
                    driver_count=len(entries),
                    title=f"Top drivers — {primary_output}",
                    left=420, top=180,
                    width=340, height=220,
                )
            except Exception:
                pass

        # Hide the helper data columns so the user sees only the charts.
        try:
            sheet.range("M:Q").api.EntireColumn.Hidden = True
        except Exception:
            pass

        # Reserve vertical space so the stats table doesn't overlap the
        # chart band. (The charts are anchored by point coords, not
        # cells, but we still bump the rows so the print layout stays
        # clean.)
        _ = top, height
        return chart_count

    @staticmethod
    def _write_stats_table(
        sheet: Any,
        *,
        primary_result: SimulationResult,
        secondary_results: list[SimulationResult],
    ) -> int:
        """Returns the last-used row so callouts can be positioned below."""
        top = ExecutiveReportBuilder.STATS_TABLE_TOP
        header = ["Output", "Mean", "StDev", "P5", "P50", "P95", "CV"]
        for i, label in enumerate(header):
            col = chr(ord("A") + i)
            sheet.range(f"{col}{top}").value = label
        try:
            sheet.range(f"A{top}:G{top}").api.Font.Bold = True
            sheet.range(f"A{top}:G{top}").api.Interior.Color = _COLOR_BAND_LIGHT
            sheet.range(f"A{top}:G{top}").api.Borders(9).LineStyle = 1  # bottom border
        except Exception:
            pass

        all_results = [primary_result, *secondary_results]
        for i, r in enumerate(all_results, start=top + 1):
            cv = r.stdev / r.mean if r.mean != 0 else float("nan")
            sheet.range(f"A{i}").value = r.output_name
            sheet.range(f"B{i}").value = r.mean
            sheet.range(f"C{i}").value = r.stdev
            sheet.range(f"D{i}").value = r.percentiles.get(0.05, r.min)
            sheet.range(f"E{i}").value = r.percentiles.get(0.50, r.mean)
            sheet.range(f"F{i}").value = r.percentiles.get(0.95, r.max)
            sheet.range(f"G{i}").value = cv if math.isfinite(cv) else None
            try:
                sheet.range(f"B{i}:F{i}").api.NumberFormat = "#,##0.00"
                sheet.range(f"G{i}").api.NumberFormat = "0.000"
                # Highlight high-CV rows in amber/red.
                if math.isfinite(cv):
                    if cv > 0.5:
                        sheet.range(f"G{i}").api.Font.Color = _COLOR_HEADLINE_RISK
                    elif cv > 0.2:
                        sheet.range(f"G{i}").api.Font.Color = _COLOR_HEADLINE_WARN
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
        if not callouts:
            return
        sheet.range(f"A{top_row}").value = "RISK CALLOUTS"
        try:
            sheet.range(f"A{top_row}").api.Font.Bold = True
            sheet.range(f"A{top_row}").api.Font.Size = 11
            sheet.range(f"A{top_row}").api.Font.Color = _rgb(100, 100, 100)
        except Exception:
            pass
        for i, callout in enumerate(callouts, start=top_row + 1):
            sheet.range(f"A{i}").value = f"•  {callout}"
            try:
                # Merge across the row so long callouts wrap nicely.
                sheet.range(f"A{i}:J{i}").merge()
                sheet.range(f"A{i}").api.WrapText = True
                sheet.range(f"A{i}").row_height = 24
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


def _write_histogram_data(
    sheet: Any, samples: list[float], *, anchor_col: str, anchor_row: int,
) -> None:
    """Compute histogram counts and write bin centres / counts /
    cumulative as three columns starting at anchor."""
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
    sheet: Any, *, bin_count: int, title: str,
    left: int, top: int, width: int, height: int,
) -> int:
    """Add a histogram (column) + cumulative (line) chart. Returns 1
    on success, 0 if the chart object couldn't be created."""
    chart = sheet.charts.add(left=left, top=top, width=width, height=height)
    chart.set_source_data(sheet.range(f"M1:O{1 + bin_count}"))
    try:
        chart_api = chart.api[1] if isinstance(chart.api, tuple) else chart.api
        chart_api.ChartType = _XL_COLUMN_CLUSTERED
        chart_api.HasTitle = True
        chart_api.ChartTitle.Text = title
        # Make the cumulative series a line on a secondary axis.
        try:
            series = chart_api.SeriesCollection(2)  # second series = Cumulative %
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
    """Two-column layout: name, correlation. Used by the tornado mini."""
    sheet.range(f"{anchor_col}{anchor_row}").value = "Input"
    sheet.range(f"{chr(ord(anchor_col) + 1)}{anchor_row}").value = "Correlation"
    for i, e in enumerate(entries, start=anchor_row + 1):
        sheet.range(f"{anchor_col}{i}").value = e.input_name
        sheet.range(f"{chr(ord(anchor_col) + 1)}{i}").value = e.correlation


def _add_tornado_chart(
    sheet: Any, *, driver_count: int, title: str,
    left: int, top: int, width: int, height: int,
) -> int:
    chart = sheet.charts.add(left=left, top=top, width=width, height=height)
    chart.set_source_data(sheet.range(f"P1:Q{1 + driver_count}"))
    try:
        chart_api = chart.api[1] if isinstance(chart.api, tuple) else chart.api
        chart_api.ChartType = _XL_BAR_CLUSTERED
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
        sheet = book.sheets.add(sheet_name, after=book.sheets[-1])

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

        DriversReportBuilder._write_tornado_chart(sheet, output_name, entries)
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
        widths = {
            "A": 24, "B": 14, "C": 14, "D": 14, "E": 14,
            "F": 4,
            "G": 20, "H": 12, "I": 12, "J": 12,
        }
        for col, w in widths.items():
            try:
                sheet.range(f"{col}1").column_width = w
            except Exception:
                pass

    @staticmethod
    def _write_title_band(sheet: Any, title: str, subtitle: str) -> None:
        sheet.range("A1").value = title
        sheet.range("A2").value = subtitle
        try:
            sheet.range("A1:J1").merge()
            sheet.range("A2:J2").merge()
            band = sheet.range("A1:J2")
            band.api.Interior.Color = _COLOR_TITLE_BG
            band.api.Font.Color = _COLOR_TITLE_FG
            sheet.range("A1").api.Font.Size = 18
            sheet.range("A1").api.Font.Bold = True
            sheet.range("A2").api.Font.Size = 11
            sheet.range("A1").api.HorizontalAlignment = -4108
            sheet.range("A2").api.HorizontalAlignment = -4108
            sheet.range("A1").row_height = 28
            sheet.range("A2").row_height = 18
        except Exception:
            pass

    @staticmethod
    def _write_findings(sheet: Any, findings: list[str]) -> None:
        sheet.range(f"A{DriversReportBuilder.FINDINGS_HEADER_ROW}").value = (
            "KEY FINDINGS"
        )
        try:
            r = DriversReportBuilder.FINDINGS_HEADER_ROW
            sheet.range(f"A{r}:J{r}").api.Font.Bold = True
            sheet.range(f"A{r}").api.Font.Size = 12
            sheet.range(f"A{r}").api.Font.Color = _rgb(100, 100, 100)
        except Exception:
            pass
        for i, finding in enumerate(findings, start=DriversReportBuilder.FINDINGS_FIRST_ROW):
            sheet.range(f"A{i}").value = f"•  {finding}"
            try:
                sheet.range(f"A{i}:J{i}").merge()
                sheet.range(f"A{i}").api.WrapText = True
                sheet.range(f"A{i}").row_height = 28
            except Exception:
                pass

    @staticmethod
    def _write_tornado_chart(
        sheet: Any, output_name: str, entries: list[Any],
    ) -> None:
        if not entries:
            return
        # Helper data goes far right (hidden later)
        _write_tornado_data(sheet, entries, anchor_col="P", anchor_row=1)
        try:
            chart = sheet.charts.add(
                left=10, top=210, width=420, height=320,
            )
            chart.set_source_data(
                sheet.range(f"P1:Q{1 + len(entries)}")
            )
            try:
                chart_api = (
                    chart.api[1] if isinstance(chart.api, tuple)
                    else chart.api
                )
                chart_api.ChartType = _XL_BAR_CLUSTERED
                chart_api.HasTitle = True
                chart_api.ChartTitle.Text = (
                    f"What moves {output_name}"
                )
                category_axis = chart_api.Axes(1)
                category_axis.ReversePlotOrder = True
            except Exception:
                pass
            try:
                chart.name = f"DriversTornado_{output_name[:15]}"[:31]
            except Exception:
                pass
        except Exception:
            pass
        try:
            sheet.range("P:Q").api.EntireColumn.Hidden = True
        except Exception:
            pass

    @staticmethod
    def _write_driver_table(sheet: Any, entries: list[Any]) -> None:
        # Header
        header_row = DriversReportBuilder.TABLE_HEADER_ROW
        headers = ["Input", "Correlation (r)", "|r|", "Variance share"]
        for i, label in enumerate(headers):
            col = chr(ord("G") + i)
            sheet.range(f"{col}{header_row}").value = label
        try:
            sheet.range(f"G{header_row}:J{header_row}").api.Font.Bold = True
            sheet.range(
                f"G{header_row}:J{header_row}"
            ).api.Interior.Color = _COLOR_BAND_LIGHT
        except Exception:
            pass

        for idx, e in enumerate(entries, start=DriversReportBuilder.TABLE_DATA_ROW):
            sheet.range(f"G{idx}").value = e.input_name
            sheet.range(f"H{idx}").value = e.correlation
            sheet.range(f"I{idx}").value = abs(e.correlation)
            sheet.range(f"J{idx}").value = _variance_share(e.correlation)
            try:
                sheet.range(f"H{idx}").api.NumberFormat = "0.000"
                sheet.range(f"I{idx}").api.NumberFormat = "0.000"
                sheet.range(f"J{idx}").api.NumberFormat = "0.0%"
                # Color the |r| cell by strength tier.
                col_strength = _driver_strength_color(abs(e.correlation))
                sheet.range(f"I{idx}").api.Font.Color = col_strength
                sheet.range(f"I{idx}").api.Font.Bold = True
            except Exception:
                pass

    @staticmethod
    def _write_chart_explanation(sheet: Any) -> None:
        row = DriversReportBuilder.EXPLAIN_HEADER_ROW
        sheet.range(f"A{row}").value = "HOW TO READ THIS CHART"
        try:
            sheet.range(f"A{row}").api.Font.Bold = True
            sheet.range(f"A{row}").api.Font.Size = 12
            sheet.range(f"A{row}").api.Font.Color = _rgb(100, 100, 100)
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
            sheet.range(f"A{i}").value = paragraph
            try:
                sheet.range(f"A{i}:J{i}").merge()
                sheet.range(f"A{i}").api.WrapText = True
                sheet.range(f"A{i}").row_height = 32
                sheet.range(f"A{i}").api.VerticalAlignment = -4160  # xlTop
            except Exception:
                pass

    @staticmethod
    def _write_recommendations(
        sheet: Any, recommendations: dict[str, list[str]],
    ) -> None:
        row = DriversReportBuilder.RECOMMEND_HEADER_ROW
        sheet.range(f"A{row}").value = "RECOMMENDED ACTIONS"
        try:
            sheet.range(f"A{row}").api.Font.Bold = True
            sheet.range(f"A{row}").api.Font.Size = 12
            sheet.range(f"A{row}").api.Font.Color = _rgb(100, 100, 100)
        except Exception:
            pass
        tier_rows = [
            ("Focus mitigation on:", "focus", _COLOR_DRIVER_STRONG),
            ("Monitor:", "monitor", _COLOR_DRIVER_MEDIUM),
            ("Can be deprioritised:", "deprioritise", _COLOR_DRIVER_WEAK),
        ]
        for i, (label, key, color) in enumerate(tier_rows, start=row + 1):
            sheet.range(f"A{i}").value = label
            inputs = recommendations.get(key, [])
            value = ", ".join(inputs) if inputs else "(none)"
            sheet.range(f"B{i}").value = value
            try:
                sheet.range(f"A{i}").api.Font.Bold = True
                sheet.range(f"B{i}").api.Font.Color = color
                sheet.range(f"B{i}").api.Font.Bold = True
                sheet.range(f"B{i}:J{i}").merge()
                sheet.range(f"B{i}").api.WrapText = True
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
