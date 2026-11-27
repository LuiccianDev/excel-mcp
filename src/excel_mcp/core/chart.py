"""Native Excel chart creation with cross-sheet and quoted-sheet support."""

from __future__ import annotations

from typing import Any, Final, cast

from openpyxl.chart import (
    AreaChart,
    BarChart,
    LineChart,
    PieChart,
    Reference,
    ScatterChart,
    Series,
)
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.core._io import atomic_save, workbook_session
from excel_mcp.utils.cell_utils import parse_cell_range

_CHART_CLASSES: Final[dict[str, type[Any]]] = {
    "line": LineChart,
    "bar": BarChart,
    "pie": PieChart,
    "scatter": ScatterChart,
    "area": AreaChart,
}


def _unquote_sheet_name(raw_sheet: str) -> str:
    stripped = raw_sheet.strip()
    if len(stripped) >= 2 and stripped.startswith("'") and stripped.endswith("'"):
        return stripped[1:-1].replace("''", "'")
    return stripped


def _parse_chart_data_range(
    wb: Any, default_ws: Worksheet, data_range: str
) -> tuple[Worksheet, int, int, int, int]:
    if "!" in data_range:
        raw_sheet, cell_range = data_range.rsplit("!", 1)
        range_sheet_name = _unquote_sheet_name(raw_sheet)
        if range_sheet_name not in wb.sheetnames:
            raise LookupError(
                f"Sheet '{range_sheet_name}' referenced in data range not found"
            )
        data_ws = cast(Worksheet, wb[range_sheet_name])
    else:
        data_ws = default_ws
        cell_range = data_range

    start_row, start_col, end_row, end_col = parse_cell_range(
        cell_range.replace("$", "")
    )
    if end_row is None or end_col is None:
        raise ValueError("Data range must specify both start and end cells")
    return data_ws, start_row, start_col, end_row, end_col


def _populate_chart_series(
    chart: Any,
    chart_type_lower: str,
    data_ws: Worksheet,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> None:
    if chart_type_lower == "scatter":
        for col in range(start_col + 1, end_col + 1):
            x_values = Reference(
                data_ws, min_row=start_row + 1, max_row=end_row, min_col=start_col
            )
            y_values = Reference(
                data_ws, min_row=start_row + 1, max_row=end_row, min_col=col
            )
            chart.series.append(Series(y_values, x_values))
        return

    data = Reference(
        data_ws,
        min_row=start_row,
        max_row=end_row,
        min_col=start_col + 1,
        max_col=end_col,
    )
    cats = Reference(data_ws, min_row=start_row + 1, max_row=end_row, min_col=start_col)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)


def _apply_chart_style(chart: Any, style: dict[str, Any] | None) -> None:
    if not style:
        return
    if style.get("show_legend", True):
        chart.legend = Legend()
        chart.legend.position = style.get("legend_position", "r")
    else:
        chart.legend = None
    if style.get("show_data_labels", False):
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
    if style.get("grid_lines", False):
        if hasattr(chart, "x_axis"):
            chart.x_axis.majorGridlines = ChartLines()
        if hasattr(chart, "y_axis"):
            chart.y_axis.majorGridlines = ChartLines()


def create_chart_in_sheet(
    filename: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = "",
    style: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Create a chart on ``sheet_name`` referencing ``data_range``."""
    try:
        chart_type_lower = chart_type.lower()
        chart_class = _CHART_CLASSES.get(chart_type_lower)
        if not chart_class:
            return {
                "status": "error",
                "message": (
                    f"Unsupported chart type: {chart_type}. "
                    f"Supported types: {', '.join(_CHART_CLASSES.keys())}"
                ),
            }

        with workbook_session(filename, mode="edit", save=False) as wb:
            if sheet_name not in wb.sheetnames:
                return {"status": "error", "message": f"Sheet '{sheet_name}' not found"}
            target_ws = cast(Worksheet, wb[sheet_name])

            try:
                data_ws, start_row, start_col, end_row, end_col = (
                    _parse_chart_data_range(wb, target_ws, data_range)
                )
            except LookupError as exc:
                return {"status": "error", "message": str(exc)}
            except ValueError as exc:
                return {
                    "status": "error",
                    "message": f"Invalid data range format: {exc}",
                }

            chart = chart_class()
            chart.title = title
            if hasattr(chart, "x_axis"):
                chart.x_axis.title = x_axis
            if hasattr(chart, "y_axis"):
                chart.y_axis.title = y_axis

            _populate_chart_series(
                chart, chart_type_lower, data_ws, start_row, start_col, end_row, end_col
            )
            _apply_chart_style(chart, style)
            chart.width = 15
            chart.height = 7.5

            try:
                target_ws.add_chart(chart, target_cell)
            except ValueError as exc:
                return {"status": "error", "message": f"Invalid target cell: {exc}"}

            atomic_save(wb, filename)

        return {
            "status": "success",
            "message": f"{chart_type.capitalize()} chart created successfully",
            "details": {
                "type": chart_type,
                "location": target_cell,
                "data_range": data_range,
            },
        }
    except Exception as exc:
        return {
            "status": "error",
            "message": f"Unexpected error creating chart: {exc}",
        }
