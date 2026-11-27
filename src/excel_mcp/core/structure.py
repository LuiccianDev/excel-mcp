"""Worksheet structure and layout operations for Excel MCP Server.

Provides robust functions for appending rows, modifying dimensions,
creating native tables, formatting sheet view properties, and sorting ranges.
"""

from __future__ import annotations

import logging
from pathlib import Path
import re
from typing import Any, Final, Literal

from openpyxl.utils.cell import (
    column_index_from_string,
    get_column_letter,
    range_boundaries,
)
from openpyxl.worksheet.table import Table, TableStyleInfo

from excel_mcp.core._io import atomic_save, resolve_sheet, workbook_session
from excel_mcp.core.calculations import assert_safe_cell_value
from excel_mcp.exceptions.exception_core import ValidationError
from excel_mcp.utils.cell_utils import assert_range_cell_limit

logger = logging.getLogger(__name__)

VALID_TABLE_NAME_PATTERN = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
MAX_SORT_CELLS: Final[int] = 100_000
MAX_TABLE_CELLS: Final[int] = 1_000_000
MAX_APPEND_CELLS: Final[int] = 100_000
MAX_DIMENSION_AMOUNT: Final[int] = 10_000
MAX_EXCEL_ROWS: Final[int] = 1_048_576
MAX_EXCEL_COLS: Final[int] = 16_384


def _last_non_empty_row(ws: Any) -> int:
    return max(
        (
            row_idx
            for (row_idx, _col_idx), cell in ws._cells.items()
            if cell.value is not None
        ),
        default=0,
    )


def _validate_dimension_args(dimension: str, idx: int, amount: int) -> str:
    if isinstance(idx, bool) or idx < 1:
        raise ValidationError(f"Index must be >= 1, got {idx}.")
    if isinstance(amount, bool) or amount < 1:
        raise ValidationError(f"Amount must be >= 1, got {amount}.")

    dim_lower = dimension.strip().lower()
    if dim_lower not in ("row", "column"):
        raise ValidationError(
            f"Dimension must be 'row' or 'column', got '{dimension}'."
        )
    if amount > MAX_DIMENSION_AMOUNT:
        raise ValidationError(
            f"Amount {amount} exceeds maximum allowed ({MAX_DIMENSION_AMOUNT})."
        )
    max_idx = MAX_EXCEL_ROWS if dim_lower == "row" else MAX_EXCEL_COLS
    if idx > max_idx:
        raise ValidationError(
            f"Index {idx} exceeds Excel {dim_lower} limit ({max_idx})."
        )
    return dim_lower


def _mutate_worksheet_dimension(
    filename: str,
    dimension: Literal["row", "column"],
    idx: int,
    amount: int,
    sheet_name: str | None,
    action: Literal["insert", "delete"],
) -> dict[str, Any]:
    dim_lower = _validate_dimension_args(dimension, idx, amount)
    with workbook_session(filename, mode="edit") as wb:
        ws = resolve_sheet(wb, sheet_name)
        if action == "insert":
            ws.insert_rows(idx, amount) if dim_lower == "row" else ws.insert_cols(
                idx, amount
            )
        else:
            ws.delete_rows(idx, amount) if dim_lower == "row" else ws.delete_cols(
                idx, amount
            )
        title = ws.title

    return {
        "sheet_name": title,
        "dimension": dim_lower,
        "index": idx,
        "amount": amount,
        "action": action,
    }


def _parse_bounded_range(
    range_string: str, max_cells: int, label: str
) -> tuple[int, int, int, int]:
    try:
        min_col, min_row, max_col, max_row = range_boundaries(range_string)
    except Exception as e:
        raise ValidationError(
            f"Invalid range specification '{range_string}': {e}"
        ) from e

    if min_col is None or min_row is None or max_col is None or max_row is None:
        raise ValidationError(f"Range '{range_string}' must specify concrete bounds.")

    assert_range_cell_limit(
        min_row,
        min_col,
        max_row,
        max_col,
        max_cells,
        ValidationError,
        label=label,
    )
    return min_col, min_row, max_col, max_row


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
    if isinstance(start_column, bool) or start_column < 1:
        raise ValidationError(f"start_column must be >= 1, got {start_column}.")
    if start_column > MAX_EXCEL_COLS:
        raise ValidationError(
            f"start_column {start_column} exceeds Excel column limit ({MAX_EXCEL_COLS})."
        )

    if not rows:
        path = Path(filename).resolve()
        if not path.exists():
            raise FileNotFoundError(f"Workbook not found: {path}")
        return {
            "sheet_name": sheet_name or "active",
            "rows_appended": 0,
            "message": "No rows to append.",
        }

    if len(rows) > MAX_EXCEL_ROWS:
        raise ValidationError(
            f"Append row count ({len(rows)}) exceeds Excel row limit ({MAX_EXCEL_ROWS})."
        )

    total_cells = sum(len(row) for row in rows)
    if total_cells > MAX_APPEND_CELLS:
        raise ValidationError(
            f"Append payload ({total_cells} cells) exceeds maximum allowed ({MAX_APPEND_CELLS})."
        )

    max_end_col = max(
        (start_column + len(row) - 1 for row in rows if row),
        default=start_column,
    )
    if max_end_col > MAX_EXCEL_COLS:
        raise ValidationError(
            f"Target end column ({max_end_col}) exceeds Excel column limit ({MAX_EXCEL_COLS})."
        )

    with workbook_session(filename, mode="edit") as wb:
        ws = resolve_sheet(wb, sheet_name)
        start_row = _last_non_empty_row(ws) + 1
        end_row = start_row + len(rows) - 1
        if end_row > MAX_EXCEL_ROWS:
            raise ValidationError(
                f"Target end row ({end_row}) exceeds Excel row limit ({MAX_EXCEL_ROWS})."
            )

        for r_offset, row in enumerate(rows):
            curr_row = start_row + r_offset
            for c_offset, val in enumerate(row):
                curr_col = start_column + c_offset
                assert_safe_cell_value(val, f"at row {curr_row}, col {curr_col}")
                ws.cell(row=curr_row, column=curr_col, value=val)

        title = ws.title

    return {
        "sheet_name": title,
        "rows_appended": len(rows),
        "start_row": start_row,
        "end_row": end_row,
    }


def insert_worksheet_dimension(
    filename: str,
    dimension: Literal["row", "column"],
    idx: int,
    amount: int = 1,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Insert empty rows or columns at the specified index."""
    return _mutate_worksheet_dimension(
        filename, dimension, idx, amount, sheet_name, action="insert"
    )


def delete_worksheet_dimension(
    filename: str,
    dimension: Literal["row", "column"],
    idx: int,
    amount: int = 1,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Delete rows or columns starting at the specified index."""
    return _mutate_worksheet_dimension(
        filename, dimension, idx, amount, sheet_name, action="delete"
    )


def format_worksheet_properties(
    filename: str,
    sheet_name: str | None = None,
    freeze_panes: str | None = None,
    auto_filter: bool | str | None = None,
    show_grid_lines: bool | None = None,
    column_widths: dict[str, float] | None = None,
) -> dict[str, Any]:
    """Configure worksheet display properties: freeze panes, autofilter, gridlines, column widths."""
    with workbook_session(filename, mode="edit") as wb:
        ws = resolve_sheet(wb, sheet_name)
        applied: dict[str, Any] = {}

        if freeze_panes is not None:
            ws.freeze_panes = freeze_panes.strip().upper() if freeze_panes else None
            applied["freeze_panes"] = ws.freeze_panes

        if auto_filter is not None:
            if isinstance(auto_filter, bool):
                if auto_filter:
                    dim = (
                        getattr(ws, "dimensions", None)
                        or f"A1:{get_column_letter(ws.max_column or 1)}{ws.max_row or 1}"
                    )
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

        title = ws.title

    return {
        "sheet_name": title,
        "properties_updated": applied,
    }


def create_excel_table(
    filename: str,
    range_string: str,
    table_name: str | None = None,
    style_name: str = "TableStyleMedium9",
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Create a native formatted Excel Table over a given range."""
    min_col, min_row, max_col, max_row = _parse_bounded_range(
        range_string, MAX_TABLE_CELLS, "table creation"
    )

    if min_row >= max_row:
        raise ValidationError(
            f"Table range '{range_string}' must include at least one header row and one data row."
        )

    with workbook_session(filename, mode="edit") as wb:
        ws = resolve_sheet(wb, sheet_name)

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
                raise ValidationError(
                    f"Table name '{table_name}' already exists in workbook."
                )

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
        title = ws.title

    return {
        "sheet_name": title,
        "table_name": table_name,
        "range": norm_range,
        "style": style_name,
    }


def sort_worksheet_range(
    filename: str,
    range_string: str,
    key_column: int | str,
    ascending: bool = True,
    has_header: bool = True,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Sort a 2D range of cells in-place by a specific key column."""
    min_col, min_row, max_col, max_row = _parse_bounded_range(
        range_string, MAX_SORT_CELLS, "sorting"
    )

    with workbook_session(filename, mode="edit", save=False) as wb:
        ws = resolve_sheet(wb, sheet_name)

        start_data_row = min_row + 1 if has_header else min_row
        if start_data_row >= max_row:
            return {
                "sheet_name": ws.title,
                "range": range_string,
                "sorted_rows": 0,
                "message": "Range contains no data rows to sort.",
            }

        if isinstance(key_column, bool):
            raise ValidationError("Key column cannot be a boolean value.")
        if isinstance(key_column, int):
            width = max_col - min_col + 1
            if 1 <= key_column <= width:
                rel_idx = key_column - 1
            else:
                raise ValidationError(
                    f"Key column index {key_column} is outside range bounds (1 to {width})."
                )
        else:
            col_str = str(key_column).strip()
            try:
                abs_col = column_index_from_string(col_str.upper())
                if min_col <= abs_col <= max_col:
                    rel_idx = abs_col - min_col
                else:
                    rel_idx = -1
            except ValueError:
                rel_idx = -1

            if rel_idx < 0 and has_header:
                headers = [
                    str(ws.cell(row=min_row, column=c).value or "").strip().lower()
                    for c in range(min_col, max_col + 1)
                ]
                target = col_str.lower()
                if target in headers:
                    rel_idx = headers.index(target)

            if rel_idx < 0:
                raise ValidationError(
                    f"Key column '{key_column}' cannot be resolved within range '{range_string}'."
                )

        rows_data = []
        for r in range(start_data_row, max_row + 1):
            row_vals = [
                ws.cell(row=r, column=c).value for c in range(min_col, max_col + 1)
            ]
            rows_data.append(row_vals)

        def _sort_key(row: list[Any]) -> tuple[int, Any]:
            val = row[rel_idx] if rel_idx < len(row) else None
            if val is None:
                return (2, "")
            if isinstance(val, int | float):
                return (0, val)
            return (1, str(val).lower())

        rows_data.sort(key=_sort_key, reverse=not ascending)

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
