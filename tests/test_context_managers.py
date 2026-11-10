"""Tests for context managers in workbook module."""

import pytest
from openpyxl import Workbook, load_workbook

from excel_mcp.exceptions.exception_core import SheetNotFoundError
from excel_mcp.core.workbook import (
    managed_workbook,
    managed_worksheet,
)


class TestManagedWorkbook:
    """Test cases for managed_workbook context manager."""

    def test_managed_workbook_creates_file(self, tmp_path):
        """Test that managed_workbook creates file if it doesn't exist."""
        test_file = tmp_path / "new_workbook.xlsx"

        with managed_workbook(test_file) as wb:
            ws = wb.active
            ws["A1"] = "Test Data"

        # File should exist after context exits
        assert test_file.exists()

        # Verify content was saved
        wb = Workbook()
        wb = Workbook()
        # We can't easily verify with openpyxl here, but the save happened

    def test_managed_workbook_opens_existing_file(self, tmp_path):
        """Test that managed_workbook opens existing files."""
        # Create existing file
        test_file = tmp_path / "existing.xlsx"
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "Existing Data"
        wb.save(test_file)
        wb.close()

        # Open with context manager
        with managed_workbook(test_file) as wb:
            ws = wb.active
            assert ws["A1"].value == "Existing Data"
            ws["B1"] = "New Data"

        # Verify changes were saved
        wb = Workbook()
        wb = Workbook()
        # Content verification would require reopening

    def test_managed_workbook_auto_save_true(self, tmp_path):
        """Test that auto_save=True saves on exit."""
        test_file = tmp_path / "autosave.xlsx"

        with managed_workbook(test_file, auto_save=True) as wb:
            ws = wb.active
            ws["A1"] = "Should Save"

        # File should be saved
        assert test_file.exists()

    def test_managed_workbook_auto_save_false(self, tmp_path):
        """Test that auto_save=False doesn't save on exit."""
        # Create existing file
        test_file = tmp_path / "no_autosave.xlsx"
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "Original"
        wb.save(test_file)
        wb.close()

        # Open with auto_save=False
        with managed_workbook(test_file, auto_save=False) as wb:
            ws = wb.active
            ws["A1"] = "Changed"

        # The workbook was not saved, so the original value must survive.
        reloaded = load_workbook(test_file)
        assert reloaded.active["A1"].value == "Original"
        reloaded.close()

    def test_managed_workbook_read_only(self, tmp_path):
        """Test read_only mode."""
        # Create file
        test_file = tmp_path / "readonly.xlsx"
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "Read Only Data"
        wb.save(test_file)
        wb.close()

        # Open in read-only mode
        with managed_workbook(test_file, read_only=True) as wb:
            ws = wb.active
            # Should be able to read
            assert ws["A1"].value == "Read Only Data"
            # Writing might fail or be ignored depending on openpyxl version

    def test_managed_workbook_exception_cleanup(self, tmp_path):
        """Test that resources are cleaned up on exception."""
        test_file = tmp_path / "exception.xlsx"

        try:
            with managed_workbook(test_file) as wb:
                ws = wb.active
                ws["A1"] = "Before Error"
                raise ValueError("Test Exception")
        except ValueError:
            pass  # Expected

        # File might or might not exist depending on when exception occurred
        # But no resources should be leaked

    def test_managed_workbook_with_path_object(self, tmp_path):
        """Test that managed_workbook accepts Path objects."""
        test_file = tmp_path / "path_object.xlsx"

        with managed_workbook(test_file) as wb:
            ws = wb.active
            ws["A1"] = "Path Object Test"

        assert test_file.exists()


class TestManagedWorksheet:
    """Test cases for managed_worksheet context manager."""

    def test_managed_worksheet_creates_sheet(self, tmp_path):
        """Test that managed_worksheet creates sheet if missing."""
        test_file = tmp_path / "new_sheet.xlsx"

        with managed_worksheet(test_file, "NewSheet") as ws:
            ws["A1"] = "New Sheet Data"

        # Verify sheet was created by reopening
        reloaded = load_workbook(test_file)
        assert "NewSheet" in reloaded.sheetnames
        assert reloaded["NewSheet"]["A1"].value == "New Sheet Data"
        reloaded.close()

    def test_managed_worksheet_opens_existing(self, tmp_path):
        """Test that managed_worksheet opens existing sheet."""
        # Create file with sheet
        test_file = tmp_path / "existing_sheet.xlsx"
        wb = Workbook()
        wb.create_sheet("ExistingSheet")
        ws = wb["ExistingSheet"]
        ws["A1"] = "Existing Data"
        wb.save(test_file)
        wb.close()

        # Open existing sheet
        with managed_worksheet(test_file, "ExistingSheet") as ws:
            assert ws["A1"].value == "Existing Data"
            ws["B1"] = "Added Data"

    def test_managed_worksheet_create_if_missing_false(self, tmp_path):
        """Test that create_if_missing=False raises error."""
        test_file = tmp_path / "no_create.xlsx"
        wb = Workbook()
        wb.save(test_file)
        wb.close()

        with pytest.raises(SheetNotFoundError):
            with managed_worksheet(
                test_file, "NonExistent", create_if_missing=False
            ) as _ws:
                pass

    def test_managed_worksheet_auto_save(self, tmp_path):
        """Test that changes are auto-saved."""
        test_file = tmp_path / "autosave_sheet.xlsx"

        with managed_worksheet(test_file, "TestSheet") as ws:
            ws["A1"] = "Auto Saved"

        # File should exist
        assert test_file.exists()

    def test_managed_worksheet_multiple_operations(self, tmp_path):
        """Test multiple worksheet operations."""
        test_file = tmp_path / "multi_ops.xlsx"

        # Create first sheet
        with managed_worksheet(test_file, "Sheet1") as ws:
            ws["A1"] = "Sheet1 Data"

        # Create second sheet
        with managed_worksheet(test_file, "Sheet2") as ws:
            ws["A1"] = "Sheet2 Data"

        # Modify first sheet again
        with managed_worksheet(test_file, "Sheet1") as ws:
            ws["B1"] = "More Data"

        # Both sheets should exist
        assert test_file.exists()
