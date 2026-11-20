import re

from openpyxl.utils import column_index_from_string


def parse_cell_range(
    cell_ref: str, end_ref: str | None = None
) -> tuple[int, int, int | None, int | None]:
    """Parse Excel cell reference into row and column indices."""
    if end_ref:
        start_cell = cell_ref
        end_cell = end_ref
    elif ":" in cell_ref:
        # Handle range format like "A1:B2"
        start_cell, end_cell = cell_ref.split(":", 1)
    else:
        start_cell = cell_ref
        end_cell = None

    match = re.fullmatch(r"([A-Z]+)([0-9]+)", start_cell.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {start_cell}")
    col_str, row_str = match.groups()
    start_row = int(row_str)
    if start_row < 1:
        raise ValueError(f"Invalid row number: {start_row}")
    start_col = column_index_from_string(col_str)

    if end_cell:
        match = re.fullmatch(r"([A-Z]+)([0-9]+)", end_cell.upper())
        if not match:
            raise ValueError(f"Invalid cell reference: {end_cell}")
        col_str, row_str = match.groups()
        end_row = int(row_str)
        if end_row < 1:
            raise ValueError(f"Invalid row number: {end_row}")
        end_col = column_index_from_string(col_str)
        if (end_row, end_col) < (start_row, start_col):
            raise ValueError("Range end must not precede its start")
    else:
        end_row = None
        end_col = None

    return start_row, start_col, end_row, end_col


def validate_cell_reference(cell_ref: str) -> bool:
    """Validate Excel cell reference format (e.g., 'A1', 'BC123')"""
    if not cell_ref:
        return False

    # Split into column and row parts
    col = row = ""
    for c in cell_ref:
        if c.isalpha():
            if row:  # Letters after numbers not allowed
                return False
            col += c
        elif c.isdigit():
            row += c
        else:
            return False

    return bool(col and row)


def assert_range_cell_limit(
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    max_cells: int,
    error_cls: type[Exception] = ValueError,
    label: str = "operation",
) -> int:
    if end_row < start_row or end_col < start_col:
        return 0
    cell_count = (end_row - start_row + 1) * (end_col - start_col + 1)
    if cell_count > max_cells:
        raise error_cls(
            f"Range too large ({cell_count} cells); maximum allowed for {label} is {max_cells} cells."
        )
    return cell_count
