"""Tests for excel_add_data_validation."""

from __future__ import annotations

import openpyxl
import pytest

from excel_mcp.tools.validation_tools import add_data_validation


@pytest.fixture(autouse=True)
def allow_tmp_path(monkeypatch, tmp_path):
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    monkeypatch.delenv("DIRECTORY", raising=False)


@pytest.fixture
def workbook_path(tmp_path):
    path = tmp_path / "validation.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Status"
    wb.save(path)
    return path


def test_list_validation_persists(workbook_path):
    result = add_data_validation(
        filename=str(workbook_path),
        cell_range="A2:A10",
        validation_type="list",
        formula1='"Open,Closed"',
        sheet_name="Data",
    )
    assert result["status"] == "success"

    ws = openpyxl.load_workbook(workbook_path)["Data"]
    validations = ws.data_validations.dataValidation
    assert len(validations) == 1
    assert validations[0].type == "list"
    assert validations[0].formula1 == '"Open,Closed"'
    assert str(validations[0].sqref) == "A2:A10"


def test_whole_number_between_persists(workbook_path):
    result = add_data_validation(
        filename=str(workbook_path),
        cell_range="B2:B10",
        validation_type="whole",
        operator="between",
        formula1="1",
        formula2="100",
        prompt_title="Enter a number",
        prompt="Between 1 and 100",
        sheet_name="Data",
    )
    assert result["status"] == "success"

    validation = openpyxl.load_workbook(workbook_path)[
        "Data"
    ].data_validations.dataValidation[0]
    assert validation.type == "whole"
    assert validation.operator == "between"
    assert validation.formula1 == "1"
    assert validation.formula2 == "100"
    assert validation.promptTitle == "Enter a number"


def test_unknown_validation_type_is_rejected(workbook_path):
    result = add_data_validation(
        filename=str(workbook_path),
        cell_range="A2:A10",
        validation_type="dropdown",
        formula1='"a,b"',
    )
    assert result["status"] == "error"
    assert "Unsupported validation type" in result["message"]


def test_missing_formula_is_rejected(workbook_path):
    result = add_data_validation(
        filename=str(workbook_path),
        cell_range="A2:A10",
        validation_type="list",
    )
    assert result["status"] == "error"
    assert "formula1" in result["message"]


def test_missing_sheet_is_rejected(workbook_path):
    result = add_data_validation(
        filename=str(workbook_path),
        cell_range="A2:A10",
        validation_type="list",
        formula1='"a,b"',
        sheet_name="Missing",
    )
    assert result["status"] == "error"
    assert "not found" in result["message"]


def test_missing_filename_is_rejected():
    result = add_data_validation(
        cell_range="A2:A10",
        validation_type="list",
        formula1='"a,b"',
    )
    assert result["status"] == "error"
