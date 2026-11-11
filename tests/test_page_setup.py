"""Tests for excel_configure_page_setup."""

from __future__ import annotations

import openpyxl
import pytest

from excel_mcp.tools.page_tools import configure_page_setup


@pytest.fixture(autouse=True)
def allow_tmp_path(monkeypatch, tmp_path):
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    monkeypatch.delenv("DIRECTORY", raising=False)


@pytest.fixture
def workbook_path(tmp_path):
    path = tmp_path / "print.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    ws["A1"] = "Quarterly"
    wb.save(path)
    return path


def test_landscape_and_fit_to_width_persist(workbook_path):
    result = configure_page_setup(
        filename=str(workbook_path),
        orientation="landscape",
        fit_to_width=1,
        sheet_name="Report",
    )
    assert result["status"] == "success"

    ws = openpyxl.load_workbook(workbook_path)["Report"]
    assert ws.page_setup.orientation == "landscape"
    assert ws.sheet_properties.pageSetUpPr.fitToPage is True
    assert ws.page_setup.fitToWidth == 1


def test_paper_size_maps_to_openpyxl_constant(workbook_path):
    result = configure_page_setup(
        filename=str(workbook_path),
        paper_size="A3",
        sheet_name="Report",
    )
    assert result["status"] == "success"
    assert (
        str(openpyxl.load_workbook(workbook_path)["Report"].page_setup.paperSize) == "8"
    )


def test_print_area_persists(workbook_path):
    result = configure_page_setup(
        filename=str(workbook_path),
        print_area="A1:H50",
        sheet_name="Report",
    )
    assert result["status"] == "success"

    ws = openpyxl.load_workbook(workbook_path)["Report"]
    assert "$A$1:$H$50" in str(ws.print_area)


def test_header_and_footer_persist(workbook_path):
    result = configure_page_setup(
        filename=str(workbook_path),
        header="Quarterly Report",
        footer="Confidential",
        sheet_name="Report",
    )
    assert result["status"] == "success"

    ws = openpyxl.load_workbook(workbook_path)["Report"]
    assert ws.oddHeader.center.text == "Quarterly Report"
    assert ws.oddFooter.center.text == "Confidential"


def test_center_horizontally_persists(workbook_path):
    result = configure_page_setup(
        filename=str(workbook_path),
        center_horizontally=True,
        sheet_name="Report",
    )
    assert result["status"] == "success"
    assert (
        openpyxl.load_workbook(workbook_path)["Report"].print_options.horizontalCentered
        is True
    )


def test_invalid_orientation_is_rejected(workbook_path):
    result = configure_page_setup(
        filename=str(workbook_path),
        orientation="diagonal",
        sheet_name="Report",
    )
    assert result["status"] == "error"
    assert "orientation" in result["message"].lower()


def test_invalid_paper_size_is_rejected(workbook_path):
    result = configure_page_setup(
        filename=str(workbook_path),
        paper_size="Tabloid",
        sheet_name="Report",
    )
    assert result["status"] == "error"
    assert "paper size" in result["message"].lower()


def test_missing_sheet_is_rejected(workbook_path):
    result = configure_page_setup(filename=str(workbook_path), sheet_name="Missing")
    assert result["status"] == "error"
    assert "not found" in result["message"]


def test_missing_filename_is_rejected():
    result = configure_page_setup(sheet_name="Report")
    assert result["status"] == "error"
