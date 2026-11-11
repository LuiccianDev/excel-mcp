"""Worksheet protection for Excel MCP Server.

Turns the "deliver a template, not a scratch file" use case into a single
atomic operation: lock a worksheet, optionally with a password, and grant
explicit exceptions for the actions the recipient still needs.
"""

from __future__ import annotations

import os
from typing import Any

import openpyxl

from excel_mcp.core._io import atomic_save
from excel_mcp.exceptions.exception_core import SheetNotFoundError


def _get_sheet(wb: openpyxl.Workbook, sheet_name: str | None = None) -> Any:
    """Return the target worksheet or the active sheet when no name is given."""
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise SheetNotFoundError(f"Worksheet '{sheet_name}' not found in workbook.")
        return wb[sheet_name]
    return wb.active


def protect_sheet(
    filename: str,
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
    """Enable or remove worksheet protection.

    openpyxl's :class:`~openpyxl.worksheet.protection.SheetProtection` flags mean
    "this action is locked", so every ``allow_*`` parameter is stored as the
    negation of its attribute. Passwords are hashed by openpyxl; the plaintext is
    never written to the file.

    Args:
        filename: Path to the target .xlsx file.
        enabled: True to protect the sheet, False to remove existing protection.
        password: Optional password required to unprotect the sheet.
        sheet_name: Target sheet name (defaults to the active sheet).
        allow_select_locked_cells: Allow selecting locked cells while protected.
        allow_select_unlocked_cells: Allow selecting unlocked cells while protected.
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
        Summary dictionary describing the resulting protection state.

    Raises:
        FileNotFoundError: If the workbook does not exist.
        SheetNotFoundError: If ``sheet_name`` is not present.
    """
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Workbook not found: {filename}")

    wb = openpyxl.load_workbook(filename)
    try:
        ws = _get_sheet(wb, sheet_name)
        protection = ws.protection

        if not enabled:
            # ``set_password`` re-enables protection as a side effect, so the
            # password is cleared first and the sheet disabled afterwards.
            protection.set_password(None, already_hashed=True)
            protection.disable()
        else:
            protection.sheet = True
            if password:
                protection.password = password

            # openpyxl attributes are "locked" flags: allowed == not locked.
            protection.formatCells = not allow_format_cells
            protection.formatColumns = not allow_format_columns
            protection.formatRows = not allow_format_rows
            protection.insertColumns = not allow_insert_columns
            protection.insertRows = not allow_insert_rows
            protection.insertHyperlinks = not allow_insert_hyperlinks
            protection.deleteColumns = not allow_delete_columns
            protection.deleteRows = not allow_delete_rows
            protection.sort = not allow_sort
            protection.autoFilter = not allow_auto_filter
            protection.pivotTables = not allow_pivot_tables
            protection.objects = not allow_edit_objects
            protection.scenarios = not allow_edit_scenarios
            protection.selectLockedCells = not allow_select_locked_cells
            protection.selectUnlockedCells = not allow_select_unlocked_cells

        atomic_save(wb, filename)
    finally:
        wb.close()

    return {
        "sheet": ws.title,
        "protected": enabled,
        "password_protected": bool(password) and enabled,
    }
