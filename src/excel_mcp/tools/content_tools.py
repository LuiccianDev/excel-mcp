"""Content reading and writing tools for Excel MCP Server."""

from __future__ import annotations

from typing import Any

from excel_mcp.core.data import read_excel_range, write_data
from excel_mcp.utils.guards import tool_guard


@tool_guard(path_param="filename", read_only=True)
def read_data_from_excel(
    filename: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str | None = None,
    preview_only: bool = False,
) -> dict[str, Any]:
    """Read a specific range of tabular data from an Excel worksheet and return it in a structured format.

    Note: Call `excel_get_workbook_outline` first on unfamiliar or large workbooks to inspect dimensions and headers without loading full cell matrices.

    Args:
        filename (str): Absolute or relative path to the Excel workbook (.xlsx). The extension is enforced automatically.
        sheet_name (str): Name of the worksheet to read from.
        start_cell (str, optional): Top-left cell reference of the range to read. Defaults to "A1".
        end_cell (str | None, optional): Bottom-right cell reference. If None, reads until the first empty row/column. Defaults to None.
        preview_only (bool, optional): If True, returns only a bounded preview subset for quick inspection. Defaults to False.

    Returns:
        dict[str, Any]: A dictionary containing status ("success" or "error"), data (2-D list of cell values), and message.
    """
    data = read_excel_range(
        filename=filename,
        sheet_name=sheet_name,
        start_cell=start_cell,
        end_cell=end_cell,
        preview_only=preview_only,
    )

    if not data:
        return {
            "status": "error",
            "message": "No data found in the specified range",
            "data": [],
        }

    return {
        "status": "success",
        "data": data,
        "message": f"Successfully read {len(data)} rows from {sheet_name}",
    }


@tool_guard(path_param="filename")
def write_data_to_excel(
    filename: str,
    sheet_name: str,
    data: list[list[Any]],
    start_cell: str = "A1",
) -> dict[str, Any]:
    """Write data to an existing Excel worksheet beginning at the specified cell.

    Note: Prefer `excel_build_workbook` for new workbooks or `excel_batch_write_data` for multi-range updates in a single pass.

    Args:
        filename (str): Path to the workbook. The .xlsx extension is enforced.
        sheet_name (str): Target worksheet name. Must already exist.
        data (list[list[Any]]): 2-D list where each sub-list represents a row to write.
        start_cell (str, optional): Top-left cell where writing begins (e.g., "A1"). Defaults to "A1".

    Returns:
        dict[str, Any]: A status dictionary with status ("success" or "error"), message, active_sheet, and filepath.
    """
    return write_data(filename, sheet_name, data, start_cell)
