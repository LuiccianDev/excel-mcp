"""Persistence tests for excel_mcp.core.formatting."""

from __future__ import annotations

import openpyxl
import pytest

from excel_mcp.core.formatting import format_range


@pytest.fixture
def workbook_path(tmp_path):
    path = tmp_path / "format.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Name"
    ws["B1"] = 10
    wb.save(path)
    return path


def test_font_fill_and_number_format_persist(workbook_path):
    result = format_range(
        filename=str(workbook_path),
        sheet_name="Sheet1",
        start_cell="B1",
        bold=True,
        font_size=14,
        font_color="FF0000",
        bg_color="FFF2CC",
        number_format="#,##0.00",
        alignment="center",
        wrap_text=True,
    )
    assert result["status"] == "success"

    cell = openpyxl.load_workbook(workbook_path)["Sheet1"]["B1"]
    assert cell.font.bold is True
    assert cell.font.size == 14
    assert str(cell.font.color.rgb)[-6:] == "FF0000"
    assert str(cell.fill.fgColor.rgb)[-6:] == "FFF2CC"
    assert cell.number_format == "#,##0.00"
    assert cell.alignment.horizontal == "center"
    assert cell.alignment.wrap_text is True


def test_border_persists(workbook_path):
    result = format_range(
        filename=str(workbook_path),
        sheet_name="Sheet1",
        start_cell="A1",
        end_cell="B1",
        border_style="thin",
        border_color="808080",
    )
    assert result["status"] == "success"

    cell = openpyxl.load_workbook(workbook_path)["Sheet1"]["A1"]
    assert cell.border.left.style == "thin"
    assert cell.border.top.color.rgb == "FF808080"


def test_merge_persists(workbook_path):
    result = format_range(
        filename=str(workbook_path),
        sheet_name="Sheet1",
        start_cell="A1",
        end_cell="B1",
        merge_cells=True,
    )
    assert result["status"] == "success"

    ws = openpyxl.load_workbook(workbook_path)["Sheet1"]
    assert "A1:B1" in [str(rng) for rng in ws.merged_cells.ranges]


def test_conditional_format_persists(workbook_path):
    result = format_range(
        filename=str(workbook_path),
        sheet_name="Sheet1",
        start_cell="B1",
        end_cell="B1",
        conditional_format={
            "type": "cell_is",
            "params": {
                "operator": "greaterThan",
                "formula": ["5"],
                "fill": {"fgColor": "FFC7CE"},
            },
        },
    )
    assert result["status"] == "success"

    ws = openpyxl.load_workbook(workbook_path)["Sheet1"]
    rules = [rule for rng in ws.conditional_formatting for rule in rng.rules]
    assert len(rules) == 1


def test_invalid_start_cell_is_rejected(workbook_path):
    result = format_range(
        filename=str(workbook_path),
        sheet_name="Sheet1",
        start_cell="not-a-cell",
    )
    assert result["status"] == "error"
    assert "Invalid start cell" in result["message"]


def test_missing_sheet_is_rejected(workbook_path):
    result = format_range(
        filename=str(workbook_path),
        sheet_name="Missing",
        start_cell="A1",
    )
    assert result["status"] == "error"
    assert "not found" in result["message"]


def test_invalid_conditional_format_type_is_rejected(workbook_path):
    result = format_range(
        filename=str(workbook_path),
        sheet_name="Sheet1",
        start_cell="A1",
        conditional_format={"type": "unknown", "params": {}},
    )
    assert result["status"] == "error"
    assert "Invalid conditional format type" in result["message"]
