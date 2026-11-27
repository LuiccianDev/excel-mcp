"""Workbook and document management tools for Excel MCP Server."""

from __future__ import annotations

from typing import Any

from excel_mcp.core.workbook import create_sheet, create_workbook
from excel_mcp.utils.file_utils import list_excel_files_in_directory
from excel_mcp.utils.guards import tool_guard


@tool_guard(path_param="filename")
def create_excel_workbook(filename: str) -> dict[str, Any]:
    """Create an empty Excel workbook (.xlsx) in a secure, validated path.

    Note: For new reports or formatted tables, prefer `excel_build_workbook` to create sheets, data, styles, and charts in one call.

    Args:
        filename (str): Name or path of the workbook to create (e.g., "data/reporte_diario").

    Returns:
        dict[str, Any]: A dictionary with status ("success" or "error"), message, active_sheet, and filepath.
    """
    result: dict[str, Any] = create_workbook(filename)
    return result


@tool_guard(path_param="filename")
def create_excel_worksheet(filename: str, sheet_name: str) -> dict[str, Any]:
    """Add a new worksheet to an existing Excel workbook.

    Args:
        filename (str): Path to the existing Excel workbook (e.g., "reports/monthly_report.xlsx").
        sheet_name (str): Name of the new worksheet to add. Must be unique and valid per Excel rules.

    Returns:
        dict[str, Any]: Dictionary indicating success or describing the error.
    """
    result: dict[str, Any] = create_sheet(filename, sheet_name)
    return result


@tool_guard()
def list_excel_documents() -> dict[str, Any]:
    """List all .xlsx files in the allowed directories.

    Returns:
        dict[str, Any]: A dictionary with status ("success" or "error"), count, and files metadata list.
    """
    files = list_excel_files_in_directory()
    return {
        "status": "success",
        "count": len(files),
        "files": files,
    }
