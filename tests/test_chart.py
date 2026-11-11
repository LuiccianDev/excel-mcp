"""Persistence tests for excel_mcp.core.chart."""

from __future__ import annotations

import openpyxl
from openpyxl.chart import BarChart, LineChart, ScatterChart

from excel_mcp.core.chart import create_chart_in_sheet


def _workbook_with_data(tmp_path):
    path = tmp_path / "chart.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Month", "Revenue"])
    ws.append(["Jan", 100])
    ws.append(["Feb", 150])
    ws.append(["Mar", 120])
    wb.save(path)
    return path


def test_bar_chart_is_persisted(tmp_path):
    path = _workbook_with_data(tmp_path)
    result = create_chart_in_sheet(
        filename=str(path),
        sheet_name="Data",
        data_range="A1:B4",
        chart_type="bar",
        target_cell="D2",
        title="Revenue",
    )
    assert result["status"] == "success"

    wb = openpyxl.load_workbook(path)
    charts = wb["Data"]._charts
    assert len(charts) == 1
    assert isinstance(charts[0], BarChart)
    assert charts[0].title is not None


def test_line_chart_with_style_options(tmp_path):
    path = _workbook_with_data(tmp_path)
    result = create_chart_in_sheet(
        filename=str(path),
        sheet_name="Data",
        data_range="A1:B4",
        chart_type="line",
        target_cell="D2",
        style={"show_legend": True, "legend_position": "b", "grid_lines": True},
    )
    assert result["status"] == "success"

    wb = openpyxl.load_workbook(path)
    chart = wb["Data"]._charts[0]
    assert isinstance(chart, LineChart)
    assert chart.legend.position == "b"


def test_scatter_chart_is_persisted(tmp_path):
    path = _workbook_with_data(tmp_path)
    result = create_chart_in_sheet(
        filename=str(path),
        sheet_name="Data",
        data_range="A1:B4",
        chart_type="scatter",
        target_cell="D2",
    )
    assert result["status"] == "success"

    wb = openpyxl.load_workbook(path)
    assert isinstance(wb["Data"]._charts[0], ScatterChart)


def test_unsupported_chart_type_is_rejected(tmp_path):
    path = _workbook_with_data(tmp_path)
    result = create_chart_in_sheet(
        filename=str(path),
        sheet_name="Data",
        data_range="A1:B4",
        chart_type="radar",
        target_cell="D2",
    )
    assert result["status"] == "error"
    assert "Unsupported chart type" in result["message"]


def test_missing_sheet_is_rejected(tmp_path):
    path = _workbook_with_data(tmp_path)
    result = create_chart_in_sheet(
        filename=str(path),
        sheet_name="Missing",
        data_range="A1:B4",
        chart_type="bar",
        target_cell="D2",
    )
    assert result["status"] == "error"
    assert "not found" in result["message"]


def test_invalid_data_range_is_rejected(tmp_path):
    path = _workbook_with_data(tmp_path)
    result = create_chart_in_sheet(
        filename=str(path),
        sheet_name="Data",
        data_range="not-a-range",
        chart_type="bar",
        target_cell="D2",
    )
    assert result["status"] == "error"
