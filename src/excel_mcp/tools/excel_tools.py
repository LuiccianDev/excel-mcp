"""Workbook and document management tools for Excel MCP Server."""

from __future__ import annotations

from typing import Any

from excel_mcp.core.workbook import create_sheet, create_workbook
from excel_mcp.utils.file_utils import list_excel_files_in_directory
from excel_mcp.utils.guards import tool_guard


@tool_guard(path_param="filename")
def create_excel_workbook(filename: str) -> dict[str, Any]:
    """
    Create an empty Excel workbook (.xlsx) in a secure, validated path.

    Note: For new reports or formatted tables, prefer `excel_build_workbook` to create sheets, data, styles, and charts in one call.

    Context for LLM/AI: Use this tool when you need to programmatically generate a new Excel workbook file as the starting point for a reporting, data collection, or automation workflow. This is typically used to initialize new datasets or reports, ensuring that the file is created only within authorized directories for security and compliance. The tool will automatically add the `.xlsx` extension if missing and will validate the path to prevent unauthorized file creation.

    Args:
        filename (str): Name or path of the workbook to create. Example: "data/reporte_diario"

    Returns:
        dict[str, Any]: A dictionary with status ("success" or "error"), message, active_sheet, and filepath.
    """
    result: dict[str, Any] = create_workbook(filename)
    return result


@tool_guard(path_param="filename")
def create_excel_worksheet(filename: str, sheet_name: str) -> dict[str, Any]:
    """
    Add a new worksheet to an existing Excel workbook.

    Context for LLM/AI: Use this tool to organize or extend data in an existing Excel file by adding a new worksheet. This is ideal when you want to segment data by category, time period, or any logical grouping.

    Args:
        filename (str): Path to the existing Excel workbook. Example: "reports/monthly_report.xlsx"
        sheet_name (str): Name of the new worksheet to add. Must be unique and valid per Excel rules.

    Returns:
        dict[str, Any]: Dictionary indicating success or describing the error.
    """
    result: dict[str, Any] = create_sheet(filename, sheet_name)
    return result


@tool_guard()
def list_excel_documents() -> dict[str, Any]:
    """
    List all .xlsx files in the allowed directories.

    Context for LLM/AI: Use this tool to discover and enumerate all Excel files in permitted directories, which is helpful for inventory, audit, batch processing, or automated data discovery scenarios.

    Returns:
        dict[str, Any]: A dictionary with status ("success" or "error"), count, and files metadata list.
    """
    files = list_excel_files_in_directory()
    return {
        "status": "success",
        "count": len(files),
        "files": files,
    }
