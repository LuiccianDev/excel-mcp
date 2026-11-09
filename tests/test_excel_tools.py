from pathlib import Path
from unittest.mock import patch

import pytest

from excel_mcp.exceptions.exception_tools import ValidationError, WorkbookError
from excel_mcp.tools import excel_tools
from excel_mcp.config import ConfigurationManager


TEST_SHEET = "TestSheet"


@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_excel_workbook_success(tmp_path) -> None:
    """Test successful workbook creation."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_filename = str(tmp_path / "test_file.xlsx")

    with patch("excel_mcp.tools.excel_tools.create_workbook") as mock_create:
        mock_create.return_value = {"message": "created workbook"}

        result = await excel_tools.create_excel_workbook(test_filename)

        assert "created workbook" in result["message"]
        mock_create.assert_called_once()


@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_excel_workbook_workbook_error(tmp_path) -> None:
    """Test workbook creation with WorkbookError."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_filename = str(tmp_path / "test_file.xlsx")

    with patch("excel_mcp.tools.excel_tools.create_workbook") as mock_create:
        mock_create.side_effect = WorkbookError("Permission denied")

        result = await excel_tools.create_excel_workbook(test_filename)

        assert "error" in result


@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_excel_worksheet_success(tmp_path) -> None:
    """Test successful worksheet creation."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = str(tmp_path / "test_file.xlsx")

    with (
        patch("excel_mcp.tools.excel_tools.create_sheet") as mock_create_sheet,
        patch("excel_mcp.utils.file_utils.validate_file_access", lambda x: lambda f: f),
    ):
        mock_create_sheet.return_value = {
            "status": "success",
            "sheet_name": "TestSheet",
        }

        result = await excel_tools.create_excel_worksheet(test_file, TEST_SHEET)

        assert "sheet_name" in result
        mock_create_sheet.assert_called_once_with(test_file, TEST_SHEET)


@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_excel_worksheet_validation_error(tmp_path) -> None:
    """Test worksheet creation with ValidationError."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = str(tmp_path / "test_file.xlsx")

    with (
        patch("excel_mcp.tools.excel_tools.create_sheet") as mock_create_sheet,
        patch("excel_mcp.utils.file_utils.validate_file_access", lambda x: lambda f: f),
    ):
        mock_create_sheet.side_effect = ValidationError("Invalid sheet name")

        result = await excel_tools.create_excel_worksheet(test_file, "Invalid/Sheet")

        assert result["status"] == "error"
        assert "Invalid sheet name" in result["message"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_list_excel_documents_success() -> None:
    """Test successful listing of Excel documents."""
    test_files = [
        {"name": "test1.xlsx", "size": 1024, "modified": "2023-01-01"},
        {"name": "test2.xlsx", "size": 2048, "modified": "2023-01-02"},
    ]

    with patch(
        "excel_mcp.tools.excel_tools.list_excel_files_in_directory"
    ) as mock_list:
        mock_list.return_value = test_files

        result = await excel_tools.list_excel_documents()

        assert result["status"] == "success"
        assert result["count"] == 2
        assert len(result["files"]) == 2


@pytest.mark.asyncio  # type: ignore[misc]
async def test_list_excel_documents_error() -> None:
    """Test error handling when listing documents."""
    with patch(
        "excel_mcp.tools.excel_tools.list_excel_files_in_directory"
    ) as mock_list:
        mock_list.side_effect = Exception("Access denied")

        result = await excel_tools.list_excel_documents()

        assert result["status"] == "error"
