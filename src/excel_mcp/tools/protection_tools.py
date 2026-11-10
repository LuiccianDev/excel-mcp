"""Worksheet protection tools for Excel MCP Server.

Locks a worksheet, optionally with a password, while granting explicit
exceptions for the actions the recipient still needs.
"""

from __future__ import annotations

from typing import Any

from excel_mcp.core.protection import protect_sheet as _protect_sheet
from excel_mcp.utils.guards import tool_guard
from excel_mcp.utils.security import resolve_and_check_path


@tool_guard()
def protect_sheet(
    filename: str | None = None,
    file_path: str | None = None,
    enabled: bool = True,
    password: str | None = None,
    sheet_name: str | None = None,
    allow_select_locked_cells: bool = True,
    allow_select_unlocked_cells: bool = True,
    allow_format_cells: bool = False,
    allow_format_columns: bool = False,
    allow_format_rows: bool = False,
    allow_insert_columns: bool = False,
    allow_insert_rows: bool = False,
    allow_insert_hyperlinks: bool = False,
    allow_delete_columns: bool = False,
    allow_delete_rows: bool = False,
    allow_sort: bool = False,
    allow_auto_filter: bool = False,
    allow_pivot_tables: bool = False,
    allow_edit_objects: bool = False,
    allow_edit_scenarios: bool = False,
) -> dict[str, Any]:
    """Protect or unprotect a worksheet, with optional per-action exceptions.

    Use this when delivering a workbook template: lock the sheet so structure
    and formulas survive, then open only the cells or actions the recipient
    needs (for example, allow sorting and autofilter but not row deletion).

    Args:
        filename: Path to the target .xlsx file.
        file_path: Alternative alias for filename.
        enabled: True to protect the sheet, False to remove existing protection.
        password: Optional password required to unprotect the sheet.
        sheet_name: Target sheet name (defaults to active sheet).
        allow_select_locked_cells: Allow selecting locked cells.
        allow_select_unlocked_cells: Allow selecting unlocked cells.
        allow_format_cells: Allow formatting cells.
        allow_format_columns: Allow formatting columns.
        allow_format_rows: Allow formatting rows.
        allow_insert_columns: Allow inserting columns.
        allow_insert_rows: Allow inserting rows.
        allow_insert_hyperlinks: Allow inserting hyperlinks.
        allow_delete_columns: Allow deleting columns.
        allow_delete_rows: Allow deleting rows.
        allow_sort: Allow sorting.
        allow_auto_filter: Allow using autofilters.
        allow_pivot_tables: Allow creating or refreshing pivot tables.
        allow_edit_objects: Allow editing embedded objects.
        allow_edit_scenarios: Allow editing scenarios.

    Returns:
        Structured result with the sheet name and resulting protection state.
    """
    target, err = resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    res = _protect_sheet(
        filename=target,
        enabled=enabled,
        password=password,
        sheet_name=sheet_name,
        allow_select_locked_cells=allow_select_locked_cells,
        allow_select_unlocked_cells=allow_select_unlocked_cells,
        allow_format_cells=allow_format_cells,
        allow_format_columns=allow_format_columns,
        allow_format_rows=allow_format_rows,
        allow_insert_columns=allow_insert_columns,
        allow_insert_rows=allow_insert_rows,
        allow_insert_hyperlinks=allow_insert_hyperlinks,
        allow_delete_columns=allow_delete_columns,
        allow_delete_rows=allow_delete_rows,
        allow_sort=allow_sort,
        allow_auto_filter=allow_auto_filter,
        allow_pivot_tables=allow_pivot_tables,
        allow_edit_objects=allow_edit_objects,
        allow_edit_scenarios=allow_edit_scenarios,
    )
    return {"status": "success", "data": res}
