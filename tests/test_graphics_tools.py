from pathlib import Path
from unittest.mock import patch

import pytest

from excel_mcp.exceptions.exception_tools import ChartError, PivotError, ValidationError
from excel_mcp.tools import graphics_tools
from excel_mcp.config import ConfigurationManager


# Test data
TEST_SHEET = "Sheet1"
TEST_DATA_RANGE = "A1:C10"
TEST_TARGET_CELL = "E5"


# Test for create_chart
@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_chart_success(tmp_path) -> None:
    """Test successful chart creation."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = str(tmp_path / "test_workbook.xlsx")

    with patch("excel_mcp.tools.graphics_tools.create_chart_impl") as mock_create_chart:
        # Mock successful chart creation
        mock_create_chart.return_value = {
            "status": "success",
            "chart_type": "bar",
            "location": f"{TEST_SHEET}!{TEST_TARGET_CELL}",
        }

        result = await graphics_tools.create_chart(
            filename=test_file,
            sheet_name=TEST_SHEET,
            data_range=TEST_DATA_RANGE,
            chart_type="bar",
            target_cell=TEST_TARGET_CELL,
            title="Test Chart",
            x_axis="X Axis",
            y_axis="Y Axis",
        )

        assert result["status"] == "success"
        assert result["chart_type"] == "bar"
        mock_create_chart.assert_called_once_with(
            filename=test_file,
            sheet_name=TEST_SHEET,
            data_range=TEST_DATA_RANGE,
            chart_type="bar",
            target_cell=TEST_TARGET_CELL,
            title="Test Chart",
            x_axis="X Axis",
            y_axis="Y Axis",
        )


@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_chart_validation_error(tmp_path) -> None:
    """Test chart creation with validation error."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = str(tmp_path / "test_workbook.xlsx")

    with patch("excel_mcp.tools.graphics_tools.create_chart_impl") as mock_create_chart:
        # Mock validation error
        mock_create_chart.side_effect = ValidationError("Invalid chart type")

        result = await graphics_tools.create_chart(
            filename=test_file,
            sheet_name=TEST_SHEET,
            data_range=TEST_DATA_RANGE,
            chart_type="invalid_type",
            target_cell=TEST_TARGET_CELL,
        )

        assert "error" in result
        assert "Invalid chart type" in result["error"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_chart_chart_error(tmp_path) -> None:
    """Test chart creation with chart-specific error."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = str(tmp_path / "test_workbook.xlsx")

    with patch("excel_mcp.tools.graphics_tools.create_chart_impl") as mock_create_chart:
        # Mock chart error
        mock_create_chart.side_effect = ChartError("Data range is empty")

        result = await graphics_tools.create_chart(
            filename=test_file,
            sheet_name=TEST_SHEET,
            data_range="A1:A1",  # Empty range
            chart_type="bar",
            target_cell=TEST_TARGET_CELL,
        )

        assert "error" in result
        assert "Data range is empty" in result["error"]


# Test for create_pivot_table
@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_pivot_table_success(tmp_path) -> None:
    """Test successful pivot table creation."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = str(tmp_path / "test_workbook.xlsx")

    with patch(
        "excel_mcp.tools.graphics_tools.create_pivot_table_impl"
    ) as mock_create_pivot:
        # Mock successful pivot table creation
        mock_create_pivot.return_value = {
            "status": "success",
            "pivot_range": f"{TEST_SHEET}!A1:D10",
        }

        result = await graphics_tools.create_pivot_table(
            filename=test_file,
            sheet_name=TEST_SHEET,
            data_range=TEST_DATA_RANGE,
            rows=["Category"],
            values=["Sales"],
            columns=["Region"],
            agg_func="sum",
        )

        assert result["status"] == "success"
        assert "pivot_range" in result
        mock_create_pivot.assert_called_once_with(
            filename=test_file,
            sheet_name=TEST_SHEET,
            data_range=TEST_DATA_RANGE,
            rows=["Category"],
            values=["Sales"],
            columns=["Region"],
            agg_func="sum",
        )


@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_pivot_table_validation_error(tmp_path) -> None:
    """Test pivot table creation with validation error."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = str(tmp_path / "test_workbook.xlsx")

    with patch(
        "excel_mcp.tools.graphics_tools.create_pivot_table_impl"
    ) as mock_create_pivot:
        # Mock validation error
        mock_create_pivot.side_effect = ValidationError("Invalid data range")

        result = await graphics_tools.create_pivot_table(
            filename=test_file,
            sheet_name=TEST_SHEET,
            data_range="invalid_range",
            rows=["Category"],
            values=["Sales"],
        )

        assert "error" in result
        assert "Invalid data range" in result["error"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_pivot_table_pivot_error(tmp_path) -> None:
    """Test pivot table creation with pivot-specific error."""
    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    test_file = str(tmp_path / "test_workbook.xlsx")

    with patch(
        "excel_mcp.tools.graphics_tools.create_pivot_table_impl"
    ) as mock_create_pivot:
        # Mock pivot error
        mock_create_pivot.side_effect = PivotError("No numeric data in specified range")

        result = await graphics_tools.create_pivot_table(
            filename=test_file,
            sheet_name=TEST_SHEET,
            data_range=TEST_DATA_RANGE,
            rows=["Category"],
            values=["NonNumericColumn"],
        )

        assert "error" in result
        assert "No numeric data" in result["error"]
