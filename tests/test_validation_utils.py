"""Direct branch coverage for shared workbook validation helpers."""

from __future__ import annotations

import openpyxl

from excel_mcp.utils.validation_utils import (
    validate_formula,
    validate_formula_in_cell_operation,
    validate_range_bounds,
    validate_range_in_sheet_operation,
)


def workbook_path(tmp_path):
    path = tmp_path / "validation-utils.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = "Header"
    sheet["B2"] = 42
    workbook.save(path)
    return path


def test_validate_formula_rejects_missing_prefix_and_parentheses() -> None:
    assert validate_formula("SUM(A1)")[0] is False
    assert validate_formula("=SUM(A1")[0] is False
    assert validate_formula("=SUM(A1))")[0] is False


def test_validate_formula_rejects_unsafe_functions() -> None:
    valid, message = validate_formula("=INDIRECT(A1)")
    assert valid is False
    assert "Unsafe function" in message


def test_validate_formula_accepts_balanced_formula() -> None:
    assert validate_formula("=SUM(A1:A2)") == (True, "Formula is valid")


def test_validate_formula_operation_reports_invalid_inputs(tmp_path) -> None:
    path = workbook_path(tmp_path)
    assert (
        validate_formula_in_cell_operation(str(path), "Data", "0A", "=SUM(A1)")[
            "status"
        ]
        == "error"
    )
    assert (
        validate_formula_in_cell_operation(str(path), "Missing", "A1", "=SUM(A1)")[
            "status"
        ]
        == "error"
    )
    assert (
        validate_formula_in_cell_operation(str(path), "Data", "A1", "SUM(A1)")["status"]
        == "error"
    )


def test_validate_formula_operation_accepts_valid_formula(tmp_path) -> None:
    result = validate_formula_in_cell_operation(
        str(workbook_path(tmp_path)), "Data", "A1", "=SUM(A1:B2)"
    )
    assert result["status"] == "success"
    assert result["valid"] is True


def test_validate_range_operation_reports_bounds_and_success(tmp_path) -> None:
    path = workbook_path(tmp_path)
    valid = validate_range_in_sheet_operation(str(path), "Data", "A1", "B2")
    assert valid["status"] == "success"
    assert valid["extends_beyond_data"] is False
    assert (
        validate_range_in_sheet_operation(str(path), "Data", "Z1")["status"] == "error"
    )
    assert (
        validate_range_in_sheet_operation(str(path), "Missing", "A1")["status"]
        == "error"
    )


def test_validate_range_bounds_checks_each_dimension(tmp_path) -> None:
    sheet = openpyxl.Workbook().active
    sheet["A1"] = "x"
    assert validate_range_bounds(sheet, 1, 1, 1, 1)[0] is True
    assert validate_range_bounds(sheet, 0, 1)[0] is False
    assert validate_range_bounds(sheet, 1, 0)[0] is False
    assert validate_range_bounds(sheet, 2, 1)[0] is False
    assert validate_range_bounds(sheet, 1, 2)[0] is False
    assert validate_range_bounds(sheet, 1, 1, 0, 1)[0] is False
    assert validate_range_bounds(sheet, 1, 1, 1, 0)[0] is False
