"""Tests for excel_protect_sheet."""

from __future__ import annotations

import openpyxl
import pytest

from excel_mcp.tools.protection_tools import protect_sheet


@pytest.fixture(autouse=True)
def allow_tmp_path(monkeypatch, tmp_path):
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    monkeypatch.delenv("DIRECTORY", raising=False)


@pytest.fixture
def workbook_path(tmp_path):
    path = tmp_path / "protection.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Value"
    wb.save(path)
    return path


def test_protection_with_password_persists(workbook_path):
    result = protect_sheet(
        filename=str(workbook_path),
        password="s3cret",
        sheet_name="Data",
    )
    assert result["status"] == "success"

    protection = openpyxl.load_workbook(workbook_path)["Data"].protection
    assert protection.sheet is True
    # openpyxl stores a hash, never the plaintext password.
    assert protection.password is not None
    assert protection.password != "s3cret"


def test_allow_sort_clears_the_lock(workbook_path):
    result = protect_sheet(
        filename=str(workbook_path),
        allow_sort=True,
        sheet_name="Data",
    )
    assert result["status"] == "success"

    protection = openpyxl.load_workbook(workbook_path)["Data"].protection
    assert protection.sheet is True
    assert protection.sort is False
    # Untouched permissions stay locked by default.
    assert protection.insertRows is True


def test_enabled_false_removes_protection(workbook_path):
    protect_sheet(filename=str(workbook_path), sheet_name="Data")
    assert openpyxl.load_workbook(workbook_path)["Data"].protection.sheet is True

    result = protect_sheet(
        filename=str(workbook_path),
        enabled=False,
        sheet_name="Data",
    )
    assert result["status"] == "success"
    assert openpyxl.load_workbook(workbook_path)["Data"].protection.sheet is False


def test_missing_sheet_is_rejected(workbook_path):
    result = protect_sheet(filename=str(workbook_path), sheet_name="Missing")
    assert result["status"] == "error"
    assert "not found" in result["message"]


def test_missing_filename_is_rejected():
    result = protect_sheet(sheet_name="Data")
    assert result["status"] == "error"


def test_missing_workbook_is_rejected(tmp_path):
    missing = tmp_path / "nope.xlsx"
    result = protect_sheet(filename=str(missing))
    assert result["status"] == "error"
