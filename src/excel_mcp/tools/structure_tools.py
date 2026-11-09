"""Worksheet structure and layout tools for Excel MCP Server.

Provides tools for row/column modification, appending rows, configuring
worksheet views, creating native tables, and sorting cell ranges.
"""

from __future__ import annotations

from typing import Any, Literal

from excel_mcp.core.structure import (
    append_rows_to_sheet,
    create_excel_table,
    delete_worksheet_dimension,
    format_worksheet_properties,
    insert_worksheet_dimension,
    sort_worksheet_range,
)
from excel_mcp.utils.guards import tool_guard
from excel_mcp.utils.security import resolve_and_check_path


@tool_guard()
def append_rows(
    filename: str | None = None,
    file_path: str | None = None,
    rows: list[list[Any]] | None = None,
    sheet_name: str | None = None,
    start_column: int = 1,
) -> dict[str, Any]:
    """Append batch rows of data below the existing content in a worksheet in an atomic operation.

    Detects the first unoccupied row automatically, validates formulas securely,
    and commits changes atomically.

    Args:
        filename: Path to the target .xlsx file.
        file_path: Alternative alias for filename.
        rows: 2D array of row values to append (e.g. [["val1", 10], ["val2", 20]]).
        sheet_name: Target sheet name (defaults to active sheet).
        start_column: Starting 1-based column index (defaults to 1 = column A).

    Returns:
        Structured result with sheet name and appended row boundaries.
    """
    target, err = resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    if rows is None:
        return {"status": "error", "message": "Missing required parameter 'rows'."}

    res = append_rows_to_sheet(
        filename=target,
        rows=rows,
        sheet_name=sheet_name,
        start_column=start_column,
    )
    return {"status": "success", "data": res}


@tool_guard()
def insert_rows_columns(
    filename: str | None = None,
    file_path: str | None = None,
    dimension: Literal["row", "column"] = "row",
    index: int = 1,
    amount: int = 1,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Insert empty rows or columns at the specified 1-based index in a worksheet.

    Args:
        filename: Path to the target .xlsx file.
        file_path: Alternative alias for filename.
        dimension: 'row' to insert rows, or 'column' to insert columns.
        index: 1-based index where rows or columns should be inserted.
        amount: Number of empty rows or columns to insert (default 1).
        sheet_name: Target sheet name (defaults to active sheet).

    Returns:
        Summary of the dimension inserted.
    """
    target, err = resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    res = insert_worksheet_dimension(
        filename=target,
        dimension=dimension,
        idx=index,
        amount=amount,
        sheet_name=sheet_name,
    )
    return {"status": "success", "data": res}


@tool_guard()
def delete_rows_columns(
    filename: str | None = None,
    file_path: str | None = None,
    dimension: Literal["row", "column"] = "row",
    index: int = 1,
    amount: int = 1,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Delete rows or columns starting at the specified 1-based index in a worksheet.

    Args:
        filename: Path to the target .xlsx file.
        file_path: Alternative alias for filename.
        dimension: 'row' to delete rows, or 'column' to delete columns.
        index: 1-based starting index of rows or columns to delete.
        amount: Number of rows or columns to delete (default 1).
        sheet_name: Target sheet name (defaults to active sheet).

    Returns:
        Summary of the dimension deleted.
    """
    target, err = resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    res = delete_worksheet_dimension(
        filename=target,
        dimension=dimension,
        idx=index,
        amount=amount,
        sheet_name=sheet_name,
    )
    return {"status": "success", "data": res}


@tool_guard()
def format_sheet(
    filename: str | None = None,
    file_path: str | None = None,
    sheet_name: str | None = None,
    freeze_panes: str | None = None,
    auto_filter: bool | str | None = None,
    show_grid_lines: bool | None = None,
    column_widths: dict[str, float] | None = None,
) -> dict[str, Any]:
    """Configure worksheet-level display settings and layout properties.

    Enables setting freeze panes (e.g. 'A2' to freeze top row), enabling autofilter
    on the data range, toggling gridline visibility, and specifying custom column widths.

    Args:
        filename: Path to the target .xlsx file.
        file_path: Alternative alias for filename.
        sheet_name: Target sheet name (defaults to active sheet).
        freeze_panes: Cell reference to freeze panes below/right (e.g. 'A2', 'B2', or empty to clear).
        auto_filter: True to enable on used range, range string (e.g. 'A1:E50'), or False to clear.
        show_grid_lines: True to display gridlines, False to hide.
        column_widths: Mapping of column letters to widths (e.g. {'A': 18.0, 'B': 25.0}).

    Returns:
        Summary of sheet properties updated.
    """
    target, err = resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    res = format_worksheet_properties(
        filename=target,
        sheet_name=sheet_name,
        freeze_panes=freeze_panes,
        auto_filter=auto_filter,
        show_grid_lines=show_grid_lines,
        column_widths=column_widths,
    )
    return {"status": "success", "data": res}


@tool_guard()
def create_table(
    filename: str | None = None,
    file_path: str | None = None,
    range_string: str | None = None,
    table_name: str | None = None,
    style_name: str = "TableStyleMedium9",
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Convert a range of cells into a native, styled Excel Table.

    Native Excel tables provide automatic filtering drop-downs, structured references,
    and banded row formatting.

    Args:
        filename: Path to the target .xlsx file.
        file_path: Alternative alias for filename.
        range_string: Cell range coordinate (e.g. 'A1:D15').
        table_name: Optional unique table identifier (alphanumeric). Auto-generated if omitted.
        style_name: Built-in Excel table style name (default 'TableStyleMedium9').
        sheet_name: Target sheet name (defaults to active sheet).

    Returns:
        Summary with table name, range, and applied style.
    """
    target, err = resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    if not range_string:
        return {"status": "error", "message": "Missing required parameter 'range_string'."}

    res = create_excel_table(
        filename=target,
        range_string=range_string,
        table_name=table_name,
        style_name=style_name,
        sheet_name=sheet_name,
    )
    return {"status": "success", "data": res}


@tool_guard()
def sort_range(
    filename: str | None = None,
    file_path: str | None = None,
    range_string: str | None = None,
    key_column: int | str | None = None,
    ascending: bool = True,
    has_header: bool = True,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Sort a 2D range of cells in-place by a specific key column.

    Supports numeric, date, and text ordering with type-safe comparator logic.

    Args:
        filename: Path to the target .xlsx file.
        file_path: Alternative alias for filename.
        range_string: Range of cells to sort (e.g. 'A1:E50').
        key_column: Column index (1-based), column letter (e.g. 'B'), or header name.
        ascending: True for ascending order, False for descending.
        has_header: True if the first row of range contains column headers that should stay in place.
        sheet_name: Target sheet name (defaults to active sheet).

    Returns:
        Summary with sorted row count and key column used.
    """
    target, err = resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    if not range_string:
        return {"status": "error", "message": "Missing required parameter 'range_string'."}
    if key_column is None:
        return {"status": "error", "message": "Missing required parameter 'key_column'."}

    res = sort_worksheet_range(
        filename=target,
        range_string=range_string,
        key_column=key_column,
        ascending=ascending,
        has_header=has_header,
        sheet_name=sheet_name,
    )
    return {"status": "success", "data": res}
