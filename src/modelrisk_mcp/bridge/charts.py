"""Chart writers — render simulation analysis as native Excel charts.

Currently:
- `TornadoChartWriter` — sorted bar chart of input sensitivity for one
  output. The standard Monte Carlo "what moves my output the most"
  visualization.
- `DistributionChartWriter` — histogram (frequency + cumulative
  overlay) or ascending-cumulative (CDF) chart of one output's
  simulation sample distribution. The Results-Viewer "what does my
  output's distribution look like" visualization, persisted as a native
  Excel chart on its own sheet.

Architecture:
- All chart logic stays in this module so `ExcelBridge` doesn't sprawl.
- The writer takes an opened workbook handle (via xlwings) and the
  precomputed data — it doesn't know anything about MRService.dll or
  the simulation pipeline.
- Chart creation drops to the raw Excel COM API for things xlwings
  doesn't expose directly (axis inversion, point-by-point bar
  colouring). Each COM call is wrapped in best-effort try/except so
  that even if formatting fails, the underlying data table is still
  written and visible.
- The distribution charts reuse the binning + chart-building helpers in
  `reports.py` (`_nice_bins`, `_add_histogram_chart`, the
  bind-with-verification workaround for bug #18) rather than
  reimplementing the fragile COM, so the standalone tool and the
  executive report render identical, style-guide-compliant charts.

Future: RiskProfileChartWriter (density overlay), ScenarioComparisonWriter.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from modelrisk_mcp.bridge.reports import (
    _COLOR_CHART_LINE,
    _XL_LINE_STYLE_NONE,
    _XL_TICK_MARK_NONE,
    _add_histogram_chart,
    _axis_scale_format,
    _bind_chart_to_range,
    _nice_bins,
    _percentile,
    _style_chart_frame,
    _write_histogram_data,
)

if TYPE_CHECKING:
    from modelrisk_mcp.schemas.results import SensitivityEntry


# Excel chart-type constants. Picked here so the bridge layer can stay
# numeric-literal-free.
_XL_BAR_CLUSTERED = 57
_XL_LINE = 4
_XL_COLUMNS_PLOT_BY = 2  # SetSourceData PlotBy: series are columns

# Layout — top-left corner of chart in points, plus dimensions.
_CHART_LEFT = 300
_CHART_TOP = 10
_CHART_WIDTH = 480
_CHART_HEIGHT = 360

# Layout for the distribution charts (each on its own sheet). The 3-col
# data table sits in A:C (~150pt); the chart starts past it.
_DIST_CHART_LEFT = 230
_DIST_CHART_TOP = 10
_DIST_CHART_WIDTH = 560
_DIST_CHART_HEIGHT = 360


@dataclass(frozen=True)
class TornadoChartResult:
    """What was created. Returned to the MCP layer so the response
    can carry a useful summary."""

    sheet_name: str
    chart_name: str
    output_name: str
    input_count: int
    top_input: str | None
    top_correlation: float | None


class TornadoChartWriter:
    """Renders a sensitivity ranking as a tornado chart on a new sheet.

    The sheet contains:
    - A1: "Input"
    - B1: "Spearman correlation"
    - C1: "|correlation|"     (sort key, hidden)
    - D1: "Regression coefficient"
    - A2:Dn — rows sorted by |correlation| descending

    The chart references A1:Bn so the bars show signed correlation
    (negative inputs render to the left of the axis). Sort order +
    inverted Y axis put the largest-magnitude input at the top, which
    is the tornado convention.
    """

    @staticmethod
    def write(
        book: Any,
        output_name: str,
        entries: list[SensitivityEntry],
        *,
        sheet_name: str | None = None,
    ) -> TornadoChartResult:
        """Create the chart on `book`. Returns metadata describing what
        was created. If `sheet_name` is taken it's overwritten — the
        whole point of this tool is to be re-runnable as the model
        evolves."""
        target_sheet = sheet_name or _default_sheet_name(output_name)

        # Replace existing sheet of the same name, then add fresh.
        TornadoChartWriter._remove_existing_sheet(book, target_sheet)
        sheet = book.sheets.add(target_sheet, after=book.sheets[-1])

        TornadoChartWriter._write_data_table(sheet, output_name, entries)

        # Hide the absolute-value sort-key column so the user sees only
        # the meaningful columns.
        try:
            sheet.range("C:C").api.EntireColumn.Hidden = True
        except Exception:
            pass

        chart_name = TornadoChartWriter._add_chart(
            sheet, output_name, len(entries)
        )

        top = entries[0] if entries else None
        return TornadoChartResult(
            sheet_name=target_sheet,
            chart_name=chart_name,
            output_name=output_name,
            input_count=len(entries),
            top_input=top.input_name if top else None,
            top_correlation=top.correlation if top else None,
        )

    # ----- internal -------------------------------------------------------

    @staticmethod
    def _remove_existing_sheet(book: Any, name: str) -> None:
        for sheet in list(book.sheets):
            if sheet.name == name:
                try:
                    sheet.delete()
                except Exception:
                    pass
                return

    @staticmethod
    def _write_data_table(
        sheet: Any,
        output_name: str,
        entries: list[SensitivityEntry],
    ) -> None:
        # Header row.
        sheet.range("A1").value = "Input"
        sheet.range("B1").value = "Spearman correlation"
        sheet.range("C1").value = "|correlation|"
        sheet.range("D1").value = "Regression coefficient"

        # Title cell off to the side so it's visible in the printed
        # sheet but doesn't interfere with the chart's source range.
        sheet.range("F1").value = f"Tornado — {output_name}"

        # Data rows. Already sorted by abs correlation in the ranking,
        # but re-sort defensively so the chart's category-axis-inverted
        # rendering matches.
        sorted_entries = sorted(
            entries, key=lambda e: abs(e.correlation), reverse=True
        )
        for i, e in enumerate(sorted_entries, start=2):
            sheet.range(f"A{i}").value = e.input_name
            sheet.range(f"B{i}").value = e.correlation
            sheet.range(f"C{i}").value = abs(e.correlation)
            sheet.range(f"D{i}").value = e.regression_coefficient

        # Light formatting that helps readability without being
        # obnoxious. Bold header, freeze top row.
        try:
            sheet.range("A1:D1").api.Font.Bold = True
            sheet.range("B2:B100").api.NumberFormat = "0.000"
            sheet.range("D2:D100").api.NumberFormat = "0.000"
        except Exception:
            pass

    @staticmethod
    def _add_chart(sheet: Any, output_name: str, n_rows: int) -> str:
        if n_rows == 0:
            return ""
        chart = sheet.charts.add(
            left=_CHART_LEFT, top=_CHART_TOP,
            width=_CHART_WIDTH, height=_CHART_HEIGHT,
        )
        chart.set_source_data(sheet.range(f"A1:B{n_rows + 1}"))

        # Drop to the raw Chart COM object for tornado-specific
        # configuration. `chart.api` returns (ChartObject, Chart) on
        # Windows xlwings; we want the inner Chart.
        try:
            chart_api = chart.api[1] if isinstance(chart.api, tuple) else chart.api
            # BarClustered (the tornado convention is a horizontal bar
            # chart, not column).
            chart_api.ChartType = _XL_BAR_CLUSTERED
            # Title.
            chart_api.HasTitle = True
            chart_api.ChartTitle.Text = f"Tornado — {output_name}"
            # Invert Y-axis so largest-|correlation| is at the top.
            # Category axis = xlCategory = 1 on Chart.Axes(type, group).
            category_axis = chart_api.Axes(1)
            category_axis.ReversePlotOrder = True
        except Exception:
            pass

        # Name the chart so the response can reference it.
        try:
            chart.name = f"Tornado_{output_name}"[:31]
        except Exception:
            pass
        return str(chart.name)


def _default_sheet_name(output_name: str) -> str:
    """Excel limits sheet names to 31 chars; truncate as needed."""
    base = f"Tornado_{output_name}"
    return base[:31]


# ---------------------------------------------------------------------------
# Distribution charts — histogram + cumulative (the Results-Viewer view)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class DistributionChartResult:
    """What was created for a distribution chart. Returned to the MCP
    layer so the response can carry the headline percentiles alongside
    the sheet/chart identifiers."""

    sheet_name: str
    chart_name: str
    output_name: str
    chart_kind: str  # "histogram" | "cdf"
    sample_count: int
    bin_count: int
    mean: float
    p10: float
    p50: float
    p90: float


class DistributionChartWriter:
    """Renders one output's simulation-sample distribution as a native
    Excel chart on its own sheet.

    Two kinds:
    - ``histogram`` — frequency columns with the cumulative-probability
      line overlaid on a secondary % axis and the central-80% (P10-P90)
      band highlighted. This is the full Results-Viewer histogram.
    - ``cdf`` — the ascending cumulative-probability curve on its own
      (the "what's the chance the output is below X" view).

    The sheet holds a 3-column data table (bin centre / frequency /
    cumulative %) plus the chart. Idempotent: a target sheet of the same
    name is replaced, so the tool is re-runnable as the model evolves.

    The binning and chart COM are reused from ``reports.py`` so the
    output is byte-for-byte the same style the executive report uses."""

    @staticmethod
    def write(
        book: Any,
        output_name: str,
        samples: list[float],
        *,
        chart_kind: str = "histogram",
        sheet_name: str | None = None,
    ) -> DistributionChartResult:
        kind = chart_kind.lower().strip()
        if kind not in ("histogram", "cdf"):
            raise ValueError(
                f"chart_kind must be 'histogram' or 'cdf', got {chart_kind!r}"
            )

        target_sheet = sheet_name or _distribution_sheet_name(output_name, kind)
        DistributionChartWriter._remove_existing_sheet(book, target_sheet)
        sheet = book.sheets.add(target_sheet, after=book.sheets[-1])

        bins = _nice_bins(samples)
        # Reuse the report's 3-column writer; the data table lives on the
        # visible chart sheet itself (A1:C) so the numbers travel with the
        # chart. The chart is placed past column C so they don't overlap.
        _write_histogram_data(sheet, bins, anchor_col="A", anchor_row=1)
        DistributionChartWriter._style_data_table(sheet, bins.n)

        p10 = _percentile(samples, 0.10)
        p50 = _percentile(samples, 0.50)
        p90 = _percentile(samples, 0.90)
        mean = sum(samples) / len(samples) if samples else 0.0

        chart_name = ""
        if bins.n:
            if kind == "histogram":
                title = f"Histogram — {output_name}"
                ok = _add_histogram_chart(
                    sheet, helper=sheet, helper_anchor_col="A", bins=bins,
                    p10=p10, p90=p90, title=title,
                    left=_DIST_CHART_LEFT, top=_DIST_CHART_TOP,
                    width=_DIST_CHART_WIDTH, height=_DIST_CHART_HEIGHT,
                )
                if ok:
                    chart_name = DistributionChartWriter._last_chart_name(sheet)
            else:
                title = f"Cumulative probability — {output_name}"
                chart_name = _add_cdf_chart(
                    sheet, helper=sheet, helper_anchor_col="A", bins=bins,
                    title=title,
                    left=_DIST_CHART_LEFT, top=_DIST_CHART_TOP,
                    width=_DIST_CHART_WIDTH, height=_DIST_CHART_HEIGHT,
                )

        return DistributionChartResult(
            sheet_name=target_sheet,
            chart_name=chart_name,
            output_name=output_name,
            chart_kind=kind,
            sample_count=len(samples),
            bin_count=bins.n,
            mean=mean,
            p10=p10,
            p50=p50,
            p90=p90,
        )

    # ----- internal -------------------------------------------------------

    @staticmethod
    def _remove_existing_sheet(book: Any, name: str) -> None:
        for sheet in list(book.sheets):
            if sheet.name == name:
                try:
                    sheet.delete()
                except Exception:
                    pass
                return

    @staticmethod
    def _style_data_table(sheet: Any, n: int) -> None:
        try:
            sheet.range("A1:C1").api.Font.Bold = True
        except Exception:
            pass
        if n:
            try:
                sheet.range(f"C2:C{n + 1}").api.NumberFormat = "0.0%"
            except Exception:
                pass

    @staticmethod
    def _last_chart_name(sheet: Any) -> str:
        """`_add_histogram_chart` names the chart itself but returns an
        int flag, not the name. Read the most-recently-added chart's
        name back off the sheet for the response."""
        try:
            charts = list(sheet.charts)
            if charts:
                return str(charts[-1].name)
        except Exception:
            pass
        return ""


def _add_cdf_chart(
    sheet: Any,
    *,
    helper: Any,
    helper_anchor_col: str,
    bins: Any,
    title: str,
    left: int, top: int, width: int, height: int,
) -> str:
    """Add an ascending-cumulative (CDF) line chart on `sheet`, bound to
    the cumulative-% column of the 3-column block at
    `helper_anchor_col`. Returns the chart name on success, "" if the
    chart couldn't be created or the COM bind failed.

    Layout assumption mirrors `_write_histogram_data`: col[0] = bin
    centres, col[1] = counts, col[2] = cumulative %, header at row 1.
    Binds the cumulative column as the single series, then assigns the
    bin centres as its X (category) values so the curve reads against
    round-number output values, not the 1..N row index."""
    if bins.n == 0:
        return ""
    try:
        chart = sheet.charts.add(left=left, top=top, width=width, height=height)
    except Exception:
        return ""
    bin_col = helper_anchor_col
    cum_col = chr(ord(helper_anchor_col) + 2)
    data_range = helper.range(f"{cum_col}1:{cum_col}{1 + bins.n}")
    try:
        chart_api = chart.api[1] if isinstance(chart.api, tuple) else chart.api
        chart_api.ChartType = _XL_LINE
    except Exception:
        pass
    bound = _bind_chart_to_range(chart, data_range, plot_by=_XL_COLUMNS_PLOT_BY)
    if not bound:
        return ""
    x_range = helper.range(f"{bin_col}2:{bin_col}{1 + bins.n}")
    try:
        chart_api = chart.api[1] if isinstance(chart.api, tuple) else chart.api
        chart_api.HasTitle = True
        chart_api.ChartTitle.Text = title
        try:
            series = chart_api.SeriesCollection(1)
            series.XValues = x_range.api
            series.Format.Line.ForeColor.RGB = _COLOR_CHART_LINE
            series.Format.Line.Weight = 2.25
            try:
                series.MarkerStyle = _XL_LINE_STYLE_NONE
            except Exception:
                pass
        except Exception:
            pass
        _style_chart_frame(chart_api)
        # X axis: round-number labels, thinned, no tick marks.
        try:
            xa = chart_api.Axes(1)
            xa.TickLabels.NumberFormat = _axis_scale_format(bins.centres)
            xa.TickLabels.Font.Size = 9
            xa.TickLabelSpacing = bins.label_every
            try:
                xa.TickMarkSpacing = bins.label_every
            except Exception:
                pass
            xa.MajorTickMark = _XL_TICK_MARK_NONE
            xa.MinorTickMark = _XL_TICK_MARK_NONE
        except Exception:
            pass
        # Y axis: cumulative probability, hard-capped 0..100%.
        try:
            ya = chart_api.Axes(2)
            ya.MinimumScale = 0
            ya.MaximumScale = 1.0
            ya.MajorUnit = 0.2
            ya.TickLabels.NumberFormat = "0%"
            ya.TickLabels.Font.Size = 9
        except Exception:
            pass
        chart.name = f"CDF_{title[:20]}"[:31]
    except Exception:
        pass
    return str(chart.name)


def _distribution_sheet_name(output_name: str, kind: str) -> str:
    """Excel limits sheet names to 31 chars; truncate as needed."""
    prefix = "Histogram_" if kind == "histogram" else "CDF_"
    return f"{prefix}{output_name}"[:31]


__all__ = [
    "DistributionChartResult",
    "DistributionChartWriter",
    "TornadoChartResult",
    "TornadoChartWriter",
]
