"""Chart writers — render simulation analysis as native Excel charts.

Currently:
- `TornadoChartWriter` — sorted bar chart of input sensitivity for one
  output. The standard Monte Carlo "what moves my output the most"
  visualization.

Architecture:
- All chart logic stays in this module so `ExcelBridge` doesn't sprawl.
- The writer takes an opened workbook handle (via xlwings) and the
  precomputed sensitivity data — it doesn't know anything about
  MRService.dll or the simulation pipeline.
- Chart creation drops to the raw Excel COM API for things xlwings
  doesn't expose directly (axis inversion, point-by-point bar
  colouring). Each COM call is wrapped in best-effort try/except so
  that even if formatting fails, the underlying data table is still
  written and visible.

Future: RiskProfileChartWriter (cumulative + density), HistogramWriter,
ScenarioComparisonWriter.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from modelrisk_mcp.schemas.results import SensitivityEntry


# Excel chart-type constants. Picked here so the bridge layer can stay
# numeric-literal-free.
_XL_BAR_CLUSTERED = 57

# Layout — top-left corner of chart in points, plus dimensions.
_CHART_LEFT = 300
_CHART_TOP = 10
_CHART_WIDTH = 480
_CHART_HEIGHT = 360


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


__all__ = ["TornadoChartResult", "TornadoChartWriter"]
