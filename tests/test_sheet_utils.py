"""Persistence tests for excel_mcp.utils.sheet_utils operations."""

from __future__ import annotations

import openpyxl
import pytest

from excel_mcp.utils.sheet_utils import (
    copy_range,
    copy_range_operation,
    copy_sheet,
    delete_range_operation,
    delete_sheet,
    format_range_string,
    merge_range,
    rename_sheet,
    unmerge_range,
    delete_range,
)


@pytest.fixture
def workbook_path(tmp_path):
    path = tmp_path / "sheets.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Header"
    ws["A2"] = 10
    ws["B2"] = 20
    wb.create_sheet("Other")
    wb.save(path)
    return path


def test_format_range_string():
    assert format_range_string(1, 1, 5, 3) == "A1:C5"


def test_copy_range_copies_values_and_styles(workbook_path):
    workbook = openpyxl.load_workbook(workbook_path)
    source = workbook["Data"]
    target = workbook["Other"]
    source["A1"].font = openpyxl.styles.Font(bold=True, color="FF0000")
    source["A1"].fill = openpyxl.styles.PatternFill("solid", fgColor="00FF00")
    source["A1"].number_format = "0.00"
    source["A1"].alignment = openpyxl.styles.Alignment(horizontal="center")
    copy_range(source, target, "A1:B2", "C3")

    assert target["C3"].value == "Header"
    assert target["D4"].value == 20
    assert target["C3"].font.bold is True
    assert target["C3"].number_format == "0.00"


def test_delete_range_clears_values_and_styles(workbook_path):
    workbook = openpyxl.load_workbook(workbook_path)
    sheet = workbook["Data"]
    sheet["A1"].font = openpyxl.styles.Font(bold=True)
    delete_range(sheet, "A1:B2")

    assert sheet["A1"].value is None
    assert sheet["B2"].value is None
    assert sheet["A1"].font.bold is False


def test_copy_sheet_persists(workbook_path):
    result = copy_sheet(str(workbook_path), "Data", "DataCopy")
    assert "message" in result
    assert "DataCopy" in openpyxl.load_workbook(workbook_path).sheetnames


def test_copy_sheet_rejects_existing_target(workbook_path):
    result = copy_sheet(str(workbook_path), "Data", "Other")
    assert "error" in result


def test_rename_sheet_persists(workbook_path):
    result = rename_sheet(str(workbook_path), "Other", "Renamed")
    assert "message" in result
    assert "Renamed" in openpyxl.load_workbook(workbook_path).sheetnames


def test_delete_sheet_persists(workbook_path):
    result = delete_sheet(str(workbook_path), "Other")
    assert "message" in result
    assert openpyxl.load_workbook(workbook_path).sheetnames == ["Data"]


def test_delete_last_sheet_is_rejected(workbook_path):
    delete_sheet(str(workbook_path), "Other")
    result = delete_sheet(str(workbook_path), "Data")
    assert "error" in result


def test_merge_and_unmerge_persist(workbook_path):
    merged = merge_range(str(workbook_path), "Data", "A1", "B1")
    assert "message" in merged
    ws = openpyxl.load_workbook(workbook_path)["Data"]
    assert "A1:B1" in [str(rng) for rng in ws.merged_cells.ranges]

    unmerged = unmerge_range(str(workbook_path), "Data", "A1", "B1")
    assert "message" in unmerged
    ws = openpyxl.load_workbook(workbook_path)["Data"]
    assert [str(rng) for rng in ws.merged_cells.ranges] == []


def test_unmerge_unmerged_range_is_rejected(workbook_path):
    result = unmerge_range(str(workbook_path), "Data", "A1", "B1")
    assert "error" in result


def test_copy_range_operation_persists(workbook_path):
    result = copy_range_operation(
        filename=str(workbook_path),
        sheet_name="Data",
        source_start="A1",
        source_end="B2",
        target_start="D1",
    )
    assert "error" not in result
    ws = openpyxl.load_workbook(workbook_path)["Data"]
    assert ws["D1"].value == "Header"
    assert ws["D2"].value == 10
    assert ws["E2"].value == 20


def test_delete_range_operation_shifts_up(workbook_path):
    result = delete_range_operation(
        filename=str(workbook_path),
        sheet_name="Data",
        start_cell="A1",
        end_cell="A1",
        shift_direction="up",
    )
    assert "error" not in result
    ws = openpyxl.load_workbook(workbook_path)["Data"]
    assert ws["A1"].value == 10


def test_delete_range_operation_rejects_bad_direction(workbook_path):
    result = delete_range_operation(
        filename=str(workbook_path),
        sheet_name="Data",
        start_cell="A1",
        shift_direction="diagonal",
    )
    assert "error" in result


def test_missing_sheet_is_reported(workbook_path):
    result = merge_range(str(workbook_path), "Missing", "A1", "B1")
    assert "error" in result
