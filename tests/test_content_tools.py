from pathlib import Path
from unittest.mock import patch

import pytest

from excel_mcp.exceptions.exception_tools import ValidationError
from excel_mcp.tools import content_tools
from excel_mcp.config import ConfigurationManager


# Test data
TEST_SHEET = "Sheet1"
TEST_DATA = [["Name", "Age"], ["Alice", 30], ["Bob", 25]]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_read_data_from_excel_success(tmp_path) -> None:
    """Test successful read operation from Excel."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    # Create test file
    test_file = tmp_path / "test_file.xlsx"

    with patch("excel_mcp.tools.content_tools.read_excel_range") as mock_read:
        # Mock actual function
        mock_read.return_value = ["Row1", "Row2"]

        result = await content_tools.read_data_from_excel(str(test_file), TEST_SHEET)

    assert result["status"] == "success"
    assert result["data"] == ["Row1", "Row2"]
    mock_read.assert_called_once()


@pytest.mark.asyncio  # type: ignore[misc]
async def test_read_data_from_excel_no_data(tmp_path) -> None:
    """Test read operation when no data is found."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    # Create test file
    test_file = tmp_path / "test_file.xlsx"

    with patch("excel_mcp.tools.content_tools.read_excel_range") as mock_read:
        # Mock actual function to return empty data
        mock_read.return_value = []

        result = await content_tools.read_data_from_excel(str(test_file), TEST_SHEET)

    assert result["status"] == "error"
    assert "No data found" in result["message"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_read_data_from_excel_error(tmp_path) -> None:
    """Test error handling during read operation."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    # Create test file
    test_file = tmp_path / "test_file.xlsx"

    with patch("excel_mcp.tools.content_tools.read_excel_range") as mock_read:
        mock_read.side_effect = Exception("Read error")

        result = await content_tools.read_data_from_excel(str(test_file), TEST_SHEET)

    assert result["status"] == "error"
    assert "Failed to read Excel data" in result["message"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_write_data_to_excel_success(tmp_path) -> None:
    """Test successful write operation to Excel."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    # Create test file
    test_file = tmp_path / "test_file.xlsx"

    with patch("excel_mcp.core.data.write_data") as mock_write:
        mock_write.return_value = {"message": "Data written successfully"}

        result = await content_tools.write_data_to_excel(
            str(test_file), TEST_SHEET, TEST_DATA
        )

        assert "successfully" in result
        args, _ = mock_write.call_args
        assert str(args[0]).endswith("test_file.xlsx")
        assert args[1] == TEST_SHEET
        assert args[2] == TEST_DATA


@pytest.mark.asyncio  # type: ignore[misc]
async def test_write_data_to_excel_validation_error(tmp_path) -> None:
    """Test validation error during write operation."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    # Create test file
    test_file = tmp_path / "test_file.xlsx"

    with patch("excel_mcp.core.data.write_data") as mock_write:
        mock_write.side_effect = ValidationError("Invalid data")

        result = await content_tools.write_data_to_excel(
            str(test_file), TEST_SHEET, TEST_DATA
        )

        assert "Error: Invalid data" in result


@pytest.mark.asyncio  # type: ignore[misc]
async def test_write_data_to_excel_general_error(tmp_path) -> None:
    """Test general error during write operation."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    # Create test file
    test_file = tmp_path / "test_file.xlsx"

    with patch("excel_mcp.core.data.write_data") as mock_write:
        mock_write.side_effect = Exception("Write error")

        result = await content_tools.write_data_to_excel(
            str(test_file), TEST_SHEET, TEST_DATA
        )

        assert "Error: Failed to write data: Write error" in result
