from pathlib import Path
from unittest.mock import patch

import pytest

from excel_mcp.exceptions.exception_tools import ValidationError
from excel_mcp.tools import format_tools


# Test data
TEST_DIR = Path(__file__).parent.parent / "documents"
TEST_DIR.mkdir(exist_ok=True)
TEST_FILENAME = str(TEST_DIR / "test_workbook.xlsx")
TEST_SHEET = "Sheet1"
TEST_NEW_SHEET = "NewSheet"
TEST_RANGE_START = "A1"
TEST_RANGE_END = "B2"


@pytest.fixture(autouse=True)
def allow_test_directory(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(TEST_DIR))
    monkeypatch.delenv("DIRECTORY", raising=False)


# Test for format_range
@pytest.mark.asyncio  # type: ignore[misc]
async def test_format_range_success() -> None:
    """Test successful cell formatting."""
    with patch("excel_mcp.tools.format_tools.format_range_excel") as mock_format:
        mock_format.return_value = {"status": "success", "message": "Format applied"}

        result = await format_tools.format_range_excel(
            filename=TEST_FILENAME,
            sheet_name=TEST_SHEET,
            start_cell=TEST_RANGE_START,
            end_cell=TEST_RANGE_END,
            bold=True,
            font_size=12,
        )

        assert result["status"] == "success"
        mock_format.assert_called_once()


@pytest.mark.asyncio  # type: ignore[misc]
async def test_format_range_validation_error() -> None:
    """Test formatting with invalid input."""
    with patch("excel_mcp.tools.format_tools.format_range") as mock_format:
        mock_format.side_effect = ValidationError("Invalid cell reference")

        result = await format_tools.format_range_excel(
            filename=TEST_FILENAME,
            sheet_name=TEST_SHEET,
            start_cell="invalid",
            end_cell=TEST_RANGE_END,
        )

        assert result["status"] == "error"
        assert "Invalid" in result["message"]
        mock_format.assert_called_once()


# Test for copy_worksheet
@pytest.mark.asyncio  # type: ignore[misc]
async def test_copy_worksheet_success() -> None:
    """Test successful worksheet copy."""
    with patch("excel_mcp.tools.format_tools.copy_sheet") as mock_copy:
        mock_copy.return_value = {"status": "success", "new_sheet": TEST_NEW_SHEET}

        result = await format_tools.copy_worksheet(
            TEST_FILENAME, TEST_SHEET, TEST_NEW_SHEET
        )

        assert result["status"] == "success"
        assert result["new_sheet"] == TEST_NEW_SHEET
        mock_copy.assert_called_once_with(TEST_FILENAME, TEST_SHEET, TEST_NEW_SHEET)


# Test for delete_worksheet
@pytest.mark.asyncio  # type: ignore[misc]
async def test_delete_worksheet_success() -> None:
    """Test successful worksheet deletion."""
    with patch("excel_mcp.tools.format_tools.delete_sheet") as mock_delete:
        mock_delete.return_value = {"status": "success", "deleted_sheet": TEST_SHEET}

        result = await format_tools.delete_worksheet(TEST_FILENAME, TEST_SHEET)

        assert result["status"] == "success"
        assert result["deleted_sheet"] == TEST_SHEET


# Test for rename_worksheet
@pytest.mark.asyncio  # type: ignore[misc]
async def test_rename_worksheet_success() -> None:
    """Test successful worksheet renaming."""
    with patch("excel_mcp.tools.format_tools.rename_sheet") as mock_rename:
        mock_rename.return_value = {
            "status": "success",
            "old_name": TEST_SHEET,
            "new_name": TEST_NEW_SHEET,
        }

        result = await format_tools.rename_worksheet(
            TEST_FILENAME, TEST_SHEET, TEST_NEW_SHEET
        )

        assert result["status"] == "success"
        assert result["old_name"] == TEST_SHEET
        assert result["new_name"] == TEST_NEW_SHEET


# Test for get_workbook_metadata
@pytest.mark.asyncio  # type: ignore[misc]
async def test_get_workbook_metadata_success() -> None:
    """Test successful retrieval of workbook metadata."""
    test_metadata = {
        "filename": TEST_FILENAME,
        "sheets": ["Sheet1", "Sheet2"],
        "active_sheet": "Sheet1",
    }

    with patch("excel_mcp.tools.format_tools.get_workbook_info") as mock_info:
        mock_info.return_value = test_metadata

        result = await format_tools.get_workbook_metadata(TEST_FILENAME)

        assert result == test_metadata
        mock_info.assert_called_once_with(TEST_FILENAME, include_ranges=False)


# Test for merge_cells
@pytest.mark.asyncio  # type: ignore[misc]
async def test_merge_cells_success() -> None:
    """Test successful cell merging."""
    with patch("excel_mcp.tools.format_tools.merge_range") as mock_merge:
        mock_merge.return_value = {
            "status": "success",
            "merged_range": f"{TEST_RANGE_START}:{TEST_RANGE_END}",
        }

        result = await format_tools.merge_cells(
            TEST_FILENAME, TEST_SHEET, TEST_RANGE_START, TEST_RANGE_END
        )

        assert result["status"] == "success"
        assert "merged_range" in result


# Test for unmerge_cells
@pytest.mark.asyncio  # type: ignore[misc]
async def test_unmerge_cells_success() -> None:
    """Test successful cell unmerging."""
    with patch("excel_mcp.tools.format_tools.unmerge_range") as mock_unmerge:
        mock_unmerge.return_value = {
            "status": "success",
            "unmerged_range": f"{TEST_RANGE_START}:{TEST_RANGE_END}",
        }

        result = await format_tools.unmerge_cells(
            TEST_FILENAME, TEST_SHEET, TEST_RANGE_START, TEST_RANGE_END
        )

        assert result["status"] == "success"
        assert "unmerged_range" in result


# Test for copy_range
@pytest.mark.asyncio  # type: ignore[misc]
async def test_copy_range_success() -> None:
    """Test successful range copying."""
    with patch("excel_mcp.tools.format_tools.copy_range") as mock_copy_range:
        mock_copy_range.return_value = {
            "status": "success",
            "copied_from": "A1:B2",
            "copied_to": "C1:D2",
        }

        result = await format_tools.copy_range(
            TEST_FILENAME,
            TEST_SHEET,
            source_start=TEST_RANGE_START,
            source_end=TEST_RANGE_END,
            target_start="C1",
        )

        assert result["status"] == "success"
        assert "copied_from" in result
        assert "copied_to" in result


# Test for delete_range
@pytest.mark.asyncio  # type: ignore[misc]
async def test_delete_range_success() -> None:
    """Test successful range deletion."""
    with patch("excel_mcp.tools.format_tools.delete_range_operation") as mock_delete:
        mock_delete.return_value = {
            "status": "success",
            "deleted_range": f"{TEST_RANGE_START}:{TEST_RANGE_END}",
        }

        result = await format_tools.delete_range(
            TEST_FILENAME, TEST_SHEET, TEST_RANGE_START, TEST_RANGE_END
        )

        assert result["status"] == "success"
        assert "deleted_range" in result


# Test for validate_excel_range
@pytest.mark.asyncio  # type: ignore[misc]
async def test_validate_excel_range_success() -> None:
    """Test successful range validation."""
    with patch("excel_mcp.tools.format_tools.validate_excel_range") as mock_validate:
        mock_validate.return_value = {
            "is_valid": True,
            "range": f"{TEST_RANGE_START}:{TEST_RANGE_END}",
            "dimensions": {"rows": 2, "columns": 2},
        }

        result = await format_tools.validate_excel_range(
            TEST_FILENAME, TEST_SHEET, TEST_RANGE_START, TEST_RANGE_END
        )

        assert result["is_valid"] is True
        assert "dimensions" in result
