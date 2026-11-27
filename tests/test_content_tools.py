from unittest.mock import patch

from excel_mcp.config import ConfigurationManager
from excel_mcp.exceptions.exception_tools import ValidationError
from excel_mcp.tools import content_tools


TEST_SHEET = "Sheet1"
TEST_DATA = [["Name", "Age"], ["Alice", 30], ["Bob", 25]]


def test_read_data_from_excel_success(tmp_path) -> None:
    """Test successful read operation from Excel."""
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = tmp_path / "test_file.xlsx"

    with patch("excel_mcp.tools.content_tools.read_excel_range") as mock_read:
        mock_read.return_value = ["Row1", "Row2"]

        result = content_tools.read_data_from_excel(str(test_file), TEST_SHEET)

    assert result["status"] == "success"
    assert result["data"] == ["Row1", "Row2"]
    mock_read.assert_called_once()


def test_read_data_from_excel_no_data(tmp_path) -> None:
    """Test read operation when no data is found."""
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = tmp_path / "test_file.xlsx"

    with patch("excel_mcp.tools.content_tools.read_excel_range") as mock_read:
        mock_read.return_value = []

        result = content_tools.read_data_from_excel(str(test_file), TEST_SHEET)

    assert result["status"] == "error"
    assert "No data found" in result["message"]


def test_read_data_from_excel_error(tmp_path) -> None:
    """Test error handling during read operation."""
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = tmp_path / "test_file.xlsx"

    with patch("excel_mcp.tools.content_tools.read_excel_range") as mock_read:
        mock_read.side_effect = Exception("Read error")

        result = content_tools.read_data_from_excel(str(test_file), TEST_SHEET)

    assert result["status"] == "error"
    assert "Read error" in result["message"]


def test_write_data_to_excel_success(tmp_path) -> None:
    """Test successful write operation to Excel."""
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = tmp_path / "test_file.xlsx"

    with patch("excel_mcp.tools.content_tools.write_data") as mock_write:
        mock_write.return_value = {
            "status": "success",
            "message": "Data written successfully",
        }

        result = content_tools.write_data_to_excel(
            str(test_file), TEST_SHEET, TEST_DATA
        )

        assert result["status"] == "success"
        assert "successfully" in result["message"]
        args, _ = mock_write.call_args
        assert str(args[0]).endswith("test_file.xlsx")
        assert args[1] == TEST_SHEET
        assert args[2] == TEST_DATA


def test_write_data_to_excel_validation_error(tmp_path) -> None:
    """Test validation error during write operation."""
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = tmp_path / "test_file.xlsx"

    with patch("excel_mcp.tools.content_tools.write_data") as mock_write:
        mock_write.side_effect = ValidationError("Invalid data")

        result = content_tools.write_data_to_excel(
            str(test_file), TEST_SHEET, TEST_DATA
        )

        assert result["status"] == "error"
        assert result["error_type"] == "ValidationError"
        assert "Invalid data" in result["message"]


def test_write_data_to_excel_general_error(tmp_path) -> None:
    """Test general error during write operation."""
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = tmp_path / "test_file.xlsx"

    with patch("excel_mcp.tools.content_tools.write_data") as mock_write:
        mock_write.side_effect = Exception("Write error")

        result = content_tools.write_data_to_excel(
            str(test_file), TEST_SHEET, TEST_DATA
        )

        assert result["status"] == "error"
        assert "Write error" in result["message"]
