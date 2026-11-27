from typing import Any

from excel_mcp.core.data import (
    validate_range_in_sheet_operation as validate_range_impl,
)
from excel_mcp.core.formatting import BorderStyle, format_range
from excel_mcp.core.ranges import (
    copy_range_operation,
    copy_sheet,
    delete_range_operation,
    delete_sheet,
    merge_range,
    rename_sheet,
    unmerge_range,
)
from excel_mcp.core.workbook import get_workbook_info
from excel_mcp.utils.guards import tool_guard


@tool_guard(path_param="filename")
def format_range_excel(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str | None = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: int | None = None,
    font_color: str | None = None,
    bg_color: str | None = None,
    border_style: BorderStyle = None,
    border_color: str | None = None,
    number_format: str | None = None,
    alignment: str | None = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: dict[str, Any] | None = None,
    conditional_format: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Apply visual and data formatting styles to a specified cell range in an existing Excel worksheet.

    Note: For new workbooks, prefer `excel_build_workbook` (supports presets, headers, zebra striping, number formats, and charts in one call).

    Args:
        filename (str): Path to the Excel workbook.
        sheet_name (str): The name of the worksheet to format.
        start_cell (str): The top-left cell of the target range (e.g., "A1").
        end_cell (str | None): The bottom-right cell of the target range (e.g., "B2").
        bold (bool, optional): Apply bold font style. Defaults to False.
        italic (bool, optional): Apply italic font style. Defaults to False.
        underline (bool, optional): Apply underline. Defaults to False.
        font_size (int | None, optional): Set the font size. Defaults to None.
        font_color (str | None, optional): Font color as a hex code (e.g., "#FF0000"). Defaults to None.
        bg_color (str | None, optional): Cell background color as a hex code. Defaults to None.
        border_style (str | None, optional): Border style (e.g., "thin", "thick"). Defaults to None.
        border_color (str | None, optional): Border color as a hex code. Defaults to None.
        number_format (str | None, optional): Number format code (e.g., "0.00%", "YYYY-MM-DD"). Defaults to None.
        alignment (str | None, optional): Horizontal alignment ("left", "center", "right"). Defaults to None.
        wrap_text (bool, optional): Enable text wrapping within cells. Defaults to False.
        merge_cells (bool, optional): Merge the entire specified range into a single cell. Defaults to False.
        protection (dict[str, Any] | None, optional): Cell protection settings (e.g., `{"locked": True}`). Defaults to None.
        conditional_format (dict[str, Any] | None, optional): Rules for conditional formatting. Defaults to None.

    Returns:
        dict[str, Any]: A dictionary indicating the status ("success" or "error") and a descriptive message.
    """
    return format_range(
        filename=filename,
        sheet_name=sheet_name,
        start_cell=start_cell,
        end_cell=end_cell,
        bold=bold,
        italic=italic,
        underline=underline,
        font_size=font_size,
        font_color=font_color,
        bg_color=bg_color,
        border_style=border_style,
        border_color=border_color,
        number_format=number_format,
        alignment=alignment,
        wrap_text=wrap_text,
        merge_cells=merge_cells,
        protection=protection,
        conditional_format=conditional_format,
    )


@tool_guard(path_param="filename")
def copy_worksheet(
    filename: str, source_sheet: str, target_sheet: str
) -> dict[str, Any]:
    """Create a duplicate of an existing worksheet within the same Excel workbook."""
    return copy_sheet(filename, source_sheet, target_sheet)


@tool_guard(path_param="filename")
def delete_worksheet(filename: str, sheet_name: str) -> dict[str, Any]:
    """Permanently delete a worksheet from a workbook."""
    return delete_sheet(filename, sheet_name)


@tool_guard(path_param="filename")
def rename_worksheet(filename: str, old_name: str, new_name: str) -> dict[str, Any]:
    """Change the name of an existing worksheet within a workbook."""
    return rename_sheet(filename, old_name, new_name)


@tool_guard(path_param="filename", read_only=True)
def get_workbook_metadata(
    filename: str, include_ranges: bool = False
) -> dict[str, Any]:
    """Retrieve metadata about an Excel workbook, including sheets and optional used ranges."""
    return {
        "status": "success",
        **get_workbook_info(filename, include_ranges=include_ranges),
    }


@tool_guard(path_param="filename")
def merge_cells(
    filename: str, sheet_name: str, start_cell: str, end_cell: str
) -> dict[str, Any]:
    """Merge a rectangular range of cells into a single, larger cell."""
    return merge_range(filename, sheet_name, start_cell, end_cell)


@tool_guard(path_param="filename")
def unmerge_cells(
    filename: str, sheet_name: str, start_cell: str, end_cell: str
) -> dict[str, Any]:
    """Unmerge a previously merged cell range, reverting it to individual cells."""
    return unmerge_range(filename, sheet_name, start_cell, end_cell)


@tool_guard(path_param="filename")
def copy_range(
    filename: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str | None = None,
) -> dict[str, Any]:
    """Copy a range of cells, including their values and formatting, to a new location."""
    return copy_range_operation(
        filename, sheet_name, source_start, source_end, target_start, target_sheet
    )


@tool_guard(path_param="filename")
def delete_range(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str | None = None,
    shift_direction: str = "up",
) -> dict[str, Any]:
    """Delete a range of cells and optionally shift surrounding cells to fill the gap."""
    return delete_range_operation(
        filename, sheet_name, start_cell, end_cell, shift_direction
    )


@tool_guard(path_param="filename", read_only=True)
def validate_excel_range(
    filename: str, sheet_name: str, start_cell: str, end_cell: str | None = None
) -> dict[str, Any]:
    """Validate that a given cell range is valid within a specific worksheet."""
    return validate_range_impl(filename, sheet_name, start_cell, end_cell)
