"""Worksheet structure and layout operations for Excel MCP Server.

Provides robust functions for appending rows, modifying dimensions,
creating native tables, formatting sheet view properties, and sorting ranges.
"""

from __future__ import annotations

import logging
import os
import re
from typing import Any, Literal

import openpyxl
from openpyxl.utils.cell import (
    column_index_from_string,
    get_column_letter,
    range_boundaries,
)
from openpyxl.worksheet.table import Table, TableStyleInfo

from excel_mcp.core._io import atomic_save
from excel_mcp.core.calculations import FormulaSecurityError, validate_formula_secure
from excel_mcp.exceptions.exception_core import SheetNotFoundError
from excel_mcp.exceptions.exception_tools import ValidationError

logger = logging.getLogger(__name__)

VALID_TABLE_NAME_PATTERN = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")


def _get_sheet(wb: openpyxl.Workbook, sheet_name: str | None = None) -> Any:
    """Retrieve target worksheet or active sheet if not specified."""
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise SheetNotFoundError(f"Worksheet '{sheet_name}' not found in workbook.")
        return wb[sheet_name]
    return wb.active


def append_rows_to_sheet(
    filename: str,
    rows: list[list[Any]],
    sheet_name: str | None = None,
    start_column: int = 1,
) -> dict[str, Any]:
    """Append rows of data below existing content in a worksheet.

    Args:
        filename: Path to the target .xlsx file.
        rows: 2D list of row values to append.
        sheet_name: Target sheet name (defaults to active sheet).
        start_column: Starting 1-based column index (defaults to 1 = A).

    Returns:
        Summary dictionary with rows appended and cell range bounds.
    """
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Workbook not found: {filename}")
    if not rows:
        return {"sheet_name": sheet_name or "active", "rows_appended": 0, "message": "No rows to append."}

    wb = openpyxl.load_workbook(filename)
    try:
        ws = _get_sheet(wb, sheet_name)

        # Detect the actual last row containing non-empty data
        max_r = ws.max_row or 0
        actual_last_row = 0
        for r in range(max_r, 0, -1):
            if any(ws.cell(row=r, column=c).value is not None for c in range(1, (ws.max_column or 1) + 1)):
                actual_last_row = r
                break

        start_row = actual_last_row + 1

        for r_offset, row in enumerate(rows):
            curr_row = start_row + r_offset
            for c_offset, val in enumerate(row):
                curr_col = start_column + c_offset
                if isinstance(val, str) and val.startswith("="):
                    is_valid, msg = validate_formula_secure(val)
                    if not is_valid:
                        raise FormulaSecurityError(
                            f"Formula security validation failed at row {curr_row}, col {curr_col}: {msg}"
                        )
                ws.cell(row=curr_row, column=curr_col, value=val)

        end_row = start_row + len(rows) - 1
        atomic_save(wb, filename)

        return {
            "sheet_name": ws.title,
            "rows_appended": len(rows),
            "start_row": start_row,
            "end_row": end_row,
        }
    finally:
        wb.close()


def insert_worksheet_dimension(
    filename: str,
    dimension: Literal["row", "column"],
    idx: int,
    amount: int = 1,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Insert empty rows or columns at the specified index."""
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Workbook not found: {filename}")
    if idx < 1:
        raise ValidationError(f"Index must be >= 1, got {idx}.")
    if amount < 1:
        raise ValidationError(f"Amount must be >= 1, got {amount}.")

    dim_lower = dimension.strip().lower()
    if dim_lower not in ("row", "column"):
        raise ValidationError(f"Dimension must be 'row' or 'column', got '{dimension}'.")

    wb = openpyxl.load_workbook(filename)
    try:
        ws = _get_sheet(wb, sheet_name)
        if dim_lower == "row":
            ws.insert_rows(idx, amount)
        else:
            ws.insert_cols(idx, amount)

        atomic_save(wb, filename)
        return {
            "sheet_name": ws.title,
            "dimension": dim_lower,
            "index": idx,
            "amount": amount,
            "action": "insert",
        }
    finally:
        wb.close()


def delete_worksheet_dimension(
    filename: str,
    dimension: Literal["row", "column"],
    idx: int,
    amount: int = 1,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Delete rows or columns starting at the specified index."""
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Workbook not found: {filename}")
    if idx < 1:
        raise ValidationError(f"Index must be >= 1, got {idx}.")
    if amount < 1:
        raise ValidationError(f"Amount must be >= 1, got {amount}.")

    dim_lower = dimension.strip().lower()
    if dim_lower not in ("row", "column"):
        raise ValidationError(f"Dimension must be 'row' or 'column', got '{dimension}'.")

    wb = openpyxl.load_workbook(filename)
    try:
        ws = _get_sheet(wb, sheet_name)
        if dim_lower == "row":
            ws.delete_rows(idx, amount)
        else:
            ws.delete_cols(idx, amount)

        atomic_save(wb, filename)
        return {
            "sheet_name": ws.title,
            "dimension": dim_lower,
            "index": idx,
            "amount": amount,
            "action": "delete",
        }
    finally:
        wb.close()


def format_worksheet_properties(
    filename: str,
    sheet_name: str | None = None,
    freeze_panes: str | None = None,
    auto_filter: bool | str | None = None,
    show_grid_lines: bool | None = None,
    column_widths: dict[str, float] | None = None,
) -> dict[str, Any]:
    """Configure worksheet display properties: freeze panes, autofilter, gridlines, column widths."""
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Workbook not found: {filename}")

    wb = openpyxl.load_workbook(filename)
    try:
        ws = _get_sheet(wb, sheet_name)
        applied: dict[str, Any] = {}

        if freeze_panes is not None:
            # e.g., "A2" or "B2", or None/empty to remove
            ws.freeze_panes = freeze_panes.strip().upper() if freeze_panes else None
            applied["freeze_panes"] = ws.freeze_panes

        if auto_filter is not None:
            if isinstance(auto_filter, bool):
                if auto_filter:
                    dim = getattr(ws, "dimensions", None) or f"A1:{get_column_letter(ws.max_column or 1)}{ws.max_row or 1}"
                    ws.auto_filter.ref = dim
                    applied["auto_filter"] = dim
                else:
                    ws.auto_filter.ref = None
                    applied["auto_filter"] = None
            elif isinstance(auto_filter, str):
                ws.auto_filter.ref = auto_filter.strip().upper()
                applied["auto_filter"] = ws.auto_filter.ref

        if show_grid_lines is not None:
            if hasattr(ws, "views") and getattr(ws.views, "sheetView", None):
                ws.views.sheetView[0].showGridLines = bool(show_grid_lines)
                applied["show_grid_lines"] = bool(show_grid_lines)

        if column_widths:
            widths_applied = {}
            for col_key, width in column_widths.items():
                col_letter = col_key.strip().upper()
                ws.column_dimensions[col_letter].width = float(width)
                widths_applied[col_letter] = float(width)
            applied["column_widths"] = widths_applied

        atomic_save(wb, filename)
        return {
            "sheet_name": ws.title,
            "properties_updated": applied,
        }
    finally:
        wb.close()


def create_excel_table(
    filename: str,
    range_string: str,
    table_name: str | None = None,
    style_name: str = "TableStyleMedium9",
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Create a native formatted Excel Table over a given range."""
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Workbook not found: {filename}")

    wb = openpyxl.load_workbook(filename)
    try:
        ws = _get_sheet(wb, sheet_name)

        # Validate range syntax
        try:
            min_col, min_row, max_col, max_row = range_boundaries(range_string)
        except Exception as e:
            raise ValidationError(f"Invalid range specification '{range_string}': {e}") from e

        if min_row >= max_row:
            raise ValidationError(f"Table range '{range_string}' must include at least one header row and one data row.")

        # Determine table name
        existing_names = set(ws.tables.keys())
        for sheet in wb.worksheets:
            existing_names.update(sheet.tables.keys())

        if not table_name:
            t_idx = 1
            while f"Table{t_idx}" in existing_names:
                t_idx += 1
            table_name = f"Table{t_idx}"
        else:
            table_name = table_name.strip()
            if not VALID_TABLE_NAME_PATTERN.match(table_name):
                raise ValidationError(
                    f"Invalid table name '{table_name}'. Names must start with a letter or underscore and contain only alphanumeric characters or underscores."
                )
            if table_name in existing_names:
                raise ValidationError(f"Table name '{table_name}' already exists in workbook.")

        norm_range = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
        table = Table(displayName=table_name, ref=norm_range)
        table.tableStyleInfo = TableStyleInfo(
            name=style_name,
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(table)

        atomic_save(wb, filename)
        return {
            "sheet_name": ws.title,
            "table_name": table_name,
            "range": norm_range,
            "style": style_name,
        }
    finally:
        wb.close()


def sort_worksheet_range(
    filename: str,
    range_string: str,
    key_column: int | str,
    ascending: bool = True,
    has_header: bool = True,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Sort a 2D range of cells in-place by a specific key column."""
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Workbook not found: {filename}")

    wb = openpyxl.load_workbook(filename)
    try:
        ws = _get_sheet(wb, sheet_name)

        try:
            min_col, min_row, max_col, max_row = range_boundaries(range_string)
        except Exception as e:
            raise ValidationError(f"Invalid range specification '{range_string}': {e}") from e

        start_data_row = min_row + 1 if has_header else min_row
        if start_data_row >= max_row:
            return {
                "sheet_name": ws.title,
                "range": range_string,
                "sorted_rows": 0,
                "message": "Range contains no data rows to sort.",
            }

        # Resolve relative column index (0-based)
        if isinstance(key_column, int):
            if min_col <= key_column <= max_col:
                rel_idx = key_column - min_col
            elif 1 <= key_column <= (max_col - min_col + 1):
                rel_idx = key_column - 1
            else:
                raise ValidationError(
                    f"Key column index {key_column} is outside range bounds (columns {min_col} to {max_col})."
                )
        else:
            col_str = str(key_column).strip()
            # Try column letter first
            try:
                abs_col = column_index_from_string(col_str.upper())
                if min_col <= abs_col <= max_col:
                    rel_idx = abs_col - min_col
                else:
                    rel_idx = -1
            except ValueError:
                rel_idx = -1

            # If not a valid column letter within bounds, search headers
            if rel_idx < 0 and has_header:
                headers = [str(ws.cell(row=min_row, column=c).value or "").strip().lower() for c in range(min_col, max_col + 1)]
                target = col_str.lower()
                if target in headers:
                    rel_idx = headers.index(target)

            if rel_idx < 0:
                raise ValidationError(f"Key column '{key_column}' cannot be resolved within range '{range_string}'.")

        # Read rows
        rows_data = []
        for r in range(start_data_row, max_row + 1):
            row_vals = [ws.cell(row=r, column=c).value for c in range(min_col, max_col + 1)]
            rows_data.append(row_vals)

        # Sort with type-safe ordering (numbers, then strings, then None)
        def _sort_key(row: list[Any]) -> tuple[int, Any]:
            val = row[rel_idx] if rel_idx < len(row) else None
            if val is None:
                return (2, "")
            if isinstance(val, (int, float)):
                return (0, val)
            return (1, str(val).lower())

        rows_data.sort(key=_sort_key, reverse=not ascending)

        # Write back sorted values
        for r_offset, row_vals in enumerate(rows_data):
            row_num = start_data_row + r_offset
            for c_offset, val in enumerate(row_vals):
                ws.cell(row=row_num, column=min_col + c_offset, value=val)

        atomic_save(wb, filename)
        return {
            "sheet_name": ws.title,
            "range": range_string,
            "sorted_rows": len(rows_data),
            "key_column": key_column,
            "ascending": ascending,
        }
    finally:
        wb.close()
