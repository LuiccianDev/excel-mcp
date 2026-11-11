from typing import Any

from excel_mcp.core.formatting import BorderStyle, format_range
from excel_mcp.core.workbook import get_workbook_info

# Import exceptions
from excel_mcp.exceptions.exception_tools import (
    FormattingError,
    SheetError,
    ValidationError,
    WorkbookError,
)
from excel_mcp.utils.file_utils import ensure_xlsx_extension, validate_file_access
from excel_mcp.utils.sheet_utils import (
    copy_range_operation,
    copy_sheet,
    delete_range_operation,
    delete_sheet,
    merge_range,
    rename_sheet,
    unmerge_range,
)


@validate_file_access("filename")
async def format_range_excel(
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
    """Apply a wide range of visual and data formatting styles to a specified cell range in an Excel worksheet.

    Context for AI/LLM:
        Use this comprehensive tool to programmatically style reports, highlight data, or enforce a consistent visual layout in an Excel sheet. It is ideal for automating the final presentation layer of data generation workflows.

    Typical use cases:
        1. Styling table headers with bold text, background colors, and borders.
        2. Applying currency or date formats to numerical data columns.
        3. Setting up conditional formatting to highlight values above a certain threshold.
        4. Merging cells to create report titles or grouped headers.

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

    filename = ensure_xlsx_extension(filename)
    try:
        result: dict[str, Any] = format_range(
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
        return result
    except (ValidationError, FormattingError) as e:
        return {"status": "error", "message": f"Error: {str(e)}"}
    except Exception as e:
        return {"status": "error", "message": f"Failed to format range: {str(e)}"}


@validate_file_access("filename")
async def copy_worksheet(
    filename: str, source_sheet: str, target_sheet: str
) -> dict[str, Any]:
    """Create a duplicate of an existing worksheet within the same Excel workbook.

    Context for AI/LLM:
        Use this tool when you need to create a new worksheet based on an existing template or data sheet. It preserves all content, formatting, and formulas from the original.

    Args:
        filename (str): Path to the Excel workbook.
        source_sheet (str): The name of the worksheet to be copied.
        target_sheet (str): The name for the newly created worksheet. Must not already exist.

    Returns:
        dict[str, Any]: A status dictionary with a confirmation message or error details.
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result: dict[str, Any] = copy_sheet(filename, source_sheet, target_sheet)
        return result
    except (ValidationError, SheetError) as e:
        return {"status": "error", "message": f"Error: {str(e)}"}
    except Exception as e:
        return {"status": "error", "message": f"Failed to copy worksheet: {str(e)}"}


@validate_file_access("filename")
async def delete_worksheet(filename: str, sheet_name: str) -> dict[str, Any]:
    """Permanently delete a worksheet from a workbook.

    Context for AI/LLM:
        This is a destructive operation. Use it to remove temporary, outdated, or unnecessary worksheets from an Excel file as part of a cleanup or automation workflow. Ensure the agent confirms this action if the sheet contains data.

    Args:
        filename (str): Path to the Excel workbook.
        sheet_name (str): The name of the worksheet to be deleted.

    Returns:
        dict[str, Any]: A status dictionary indicating success or failure with a message.
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result: dict[str, Any] = delete_sheet(filename, sheet_name)
        return result
    except (ValidationError, SheetError) as e:
        return {"status": "error", "message": f"Error: {str(e)}"}
    except Exception as e:
        return {"status": "error", "message": f"Failed to delete worksheet: {str(e)}"}


@validate_file_access("filename")
async def rename_worksheet(
    filename: str, old_name: str, new_name: str
) -> dict[str, Any]:
    """Change the name of an existing worksheet within a workbook.

    Context for AI/LLM:
        Use this tool to update worksheet names to be more descriptive or to follow a required naming convention as part of an automated process.

    Args:
        filename (str): Path to the Excel workbook.
        old_name (str): The current name of the worksheet to be renamed.
        new_name (str): The new name for the worksheet. Must be unique and follow Excel's naming rules.

    Returns:
        dict[str, Any]: A status dictionary with a confirmation or error message.
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result: dict[str, Any] = rename_sheet(filename, old_name, new_name)
        return result
    except (ValidationError, SheetError) as e:
        return {"status": "error", "message": f"Error: {str(e)}"}
    except Exception as e:
        return {"status": "error", "message": f"Failed to rename worksheet: {str(e)}"}


@validate_file_access("filename")
async def get_workbook_metadata(
    filename: str, include_ranges: bool = False
) -> dict[str, Any]:
    """Retrieve metadata about an Excel workbook, including a list of all worksheets and optional details about named ranges.

    Context for AI/LLM:
        Use this tool for discovery and inspection. It allows an agent to understand the structure of a workbook before performing read, write, or format operations. It's a crucial first step in many interactive or complex workflows.

    Args:
        filename (str): Path to the Excel workbook to inspect.
        include_ranges (bool, optional): If True, the metadata will include details about named ranges. Defaults to False.

    Returns:
        dict[str, Any]: A dictionary containing workbook metadata, such as a list of sheet names and properties.
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result: dict[str, Any] = get_workbook_info(
            filename, include_ranges=include_ranges
        )
        return result
    except WorkbookError as e:
        return {"status": "error", "message": f"Error: {str(e)}"}
    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to get workbook metadata: {str(e)}",
        }


@validate_file_access("filename")
async def merge_cells(
    filename: str, sheet_name: str, start_cell: str, end_cell: str
) -> dict[str, Any]:
    """Merge a rectangular range of cells into a single, larger cell.

    Context for AI/LLM:
        Use this tool to create titles, headers, or grouped labels that span multiple columns or rows. This is a common formatting step in report generation.

    Args:
        filename (str): Path to the Excel workbook.
        sheet_name (str): The name of the worksheet where the merge will occur.
        start_cell (str): The top-left cell of the range to merge.
        end_cell (str): The bottom-right cell of the range to merge.

    Returns:
        dict[str, Any]: A status dictionary indicating success or failure.
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result: dict[str, Any] = merge_range(filename, sheet_name, start_cell, end_cell)
        return result
    except (ValidationError, SheetError) as e:
        return {"status": "error", "message": f"Error: {str(e)}"}
    except Exception as e:
        return {"status": "error", "message": f"Failed to merge cells: {str(e)}"}


@validate_file_access("filename")
async def unmerge_cells(
    filename: str, sheet_name: str, start_cell: str, end_cell: str
) -> dict[str, Any]:
    """Unmerge a previously merged cell range, reverting it to individual cells.

    Context for AI/LLM:
        Use this tool to reverse a merge operation, typically as part of a re-formatting or data extraction workflow where individual cell access is required.

    Args:
        filename (str): Path to the Excel workbook.
        sheet_name (str): The name of the worksheet containing the merged cells.
        start_cell (str): The top-left cell of the range to unmerge.
        end_cell (str): The bottom-right cell of the range to unmerge.

    Returns:
        dict[str, Any]: A status dictionary indicating success or failure.
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result: dict[str, Any] = unmerge_range(
            filename, sheet_name, start_cell, end_cell
        )
        return result
    except (ValidationError, SheetError) as e:
        return {"status": "error", "message": f"Error: {str(e)}"}
    except Exception as e:
        return {"status": "error", "message": f"Failed to unmerge cells: {str(e)}"}


@validate_file_access("filename")
async def copy_range(
    filename: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str | None = None,
) -> dict[str, Any]:
    """Copy a range of cells, including their values and formatting, to a new location, potentially in a different worksheet.

    Context for AI/LLM:
        Use this tool to duplicate data or templates within a workbook. It's useful for creating summaries, staging data for charts, or replicating formatted tables.

    Args:
        filename (str): Path to the Excel workbook.
        sheet_name (str): The name of the source worksheet.
        source_start (str): The top-left cell of the source range to copy.
        source_end (str): The bottom-right cell of the source range.
        target_start (str): The top-left cell of the destination.
        target_sheet (str | None, optional): The name of the destination worksheet. If None, the same sheet is used. Defaults to None.

    Returns:
        dict[str, Any]: A status dictionary indicating success or failure.
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result: dict[str, Any] = copy_range_operation(
            filename, sheet_name, source_start, source_end, target_start, target_sheet
        )
        return result
    except (ValidationError, SheetError) as e:
        return {"status": "error", "message": f"Error: {str(e)}"}
    except Exception as e:
        return {"status": "error", "message": f"Failed to copy range: {str(e)}"}


@validate_file_access("filename")
async def delete_range(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up",
) -> dict[str, Any]:
    """Delete a range of cells and optionally shift the surrounding cells to fill the gap.

    Context for AI/LLM:
        This is a destructive operation used to remove rows or columns of data programmatically. It's useful for cleaning up datasets by removing invalid or unnecessary records.

    Args:
        filename (str): Path to the Excel workbook.
        sheet_name (str): The name of the worksheet.
        start_cell (str): The top-left cell of the range to delete.
        end_cell (str): The bottom-right cell of the range to delete.
        shift_direction (str, optional): Direction to shift cells after deletion ('up' or 'left'). Defaults to "up".

    Returns:
        dict[str, Any]: A status dictionary indicating success or failure.
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result: dict[str, Any] = delete_range_operation(
            filename, sheet_name, start_cell, end_cell, shift_direction
        )
        return result
    except (ValidationError, SheetError) as e:
        return {"status": "error", "message": f"Error: {str(e)}"}
    except Exception as e:
        return {"status": "error", "message": f"Failed to delete range: {str(e)}"}


@validate_file_access("filename")
async def validate_excel_range(
    filename: str, sheet_name: str, start_cell: str, end_cell: str | None = None
) -> dict[str, Any]:
    """Validate that a given cell range is valid within a specific worksheet.

    Context for AI/LLM:
        Use this tool as a precondition check before attempting to read from or write to a range. This helps prevent errors by ensuring the target sheet and cell references are valid.

    Args:
        filename (str): Path to the Excel workbook.
        sheet_name (str): The name of the worksheet to check.
        start_cell (str): The top-left cell of the range.
        end_cell (str | None, optional): The bottom-right cell of the range. Defaults to None.

    Returns:
        dict[str, Any]: A dictionary with validation status ("success" or "error") and a descriptive message.
    """
    filename = ensure_xlsx_extension(filename)

    try:
        range_str = start_cell if not end_cell else f"{start_cell}:{end_cell}"
        from excel_mcp.utils.validation_utils import (
            validate_range_in_sheet_operation as validate_range_impl,
        )

        result: dict[str, Any] = validate_range_impl(filename, sheet_name, range_str)
        return result
    except ValidationError as e:
        return {"status": "error", "message": f"Error: {str(e)}"}
    except Exception as e:
        return {"status": "error", "message": f"Failed to validate range: {str(e)}"}
