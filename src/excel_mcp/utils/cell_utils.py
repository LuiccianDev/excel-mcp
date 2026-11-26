"""Pure Excel cell coordinate, column-letter, and range-bounds utilities."""

from __future__ import annotations

import re
from typing import Final

MAX_EXCEL_ROWS: Final[int] = 1_048_576
MAX_EXCEL_COLS: Final[int] = 16_384
_CELL_REF_RE: Final[re.Pattern[str]] = re.compile(r"([A-Z]+)([0-9]+)")


def column_index_from_string(col: str) -> int:
    """Convert an Excel column letter sequence (e.g. 'A', 'XFD') to a 1-based index."""
    if not col or not col.isascii() or not col.isalpha():
        raise ValueError(f"Invalid column letter: {col!r}")
    idx = 0
    for ch in col.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    if idx < 1 or idx > MAX_EXCEL_COLS:
        raise ValueError(
            f"Column '{col}' index {idx} exceeds Excel bounds (1-{MAX_EXCEL_COLS})."
        )
    return idx


def get_column_letter(col_idx: int) -> str:
    """Convert a 1-based Excel column index (1..16384) to column letters."""
    if isinstance(col_idx, bool) or not isinstance(col_idx, int):
        raise ValueError(f"Column index must be an integer, got {col_idx!r}.")
    if col_idx < 1 or col_idx > MAX_EXCEL_COLS:
        raise ValueError(
            f"Column index {col_idx} out of Excel bounds (1-{MAX_EXCEL_COLS})."
        )
    letters: list[str] = []
    remaining = col_idx
    while remaining > 0:
        remaining, rem = divmod(remaining - 1, 26)
        letters.append(chr(ord("A") + rem))
    return "".join(reversed(letters))


def format_range_string(
    start_row: int, start_col: int, end_row: int, end_col: int
) -> str:
    """Format 1-based coordinates into an A1:B2 range string."""
    return f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"


def _parse_single_cell(cell_ref: str) -> tuple[int, int]:
    match = _CELL_REF_RE.fullmatch(cell_ref.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    col_str, row_str = match.groups()
    row = int(row_str)
    if row < 1 or row > MAX_EXCEL_ROWS:
        raise ValueError(f"Invalid row number: {row}")
    col = column_index_from_string(col_str)
    return row, col


def parse_cell_range(
    cell_ref: str, end_ref: str | None = None
) -> tuple[int, int, int | None, int | None]:
    """Parse Excel cell reference into 1-based row and column indices."""
    if end_ref:
        start_cell = cell_ref
        end_cell = end_ref
    elif ":" in cell_ref:
        start_cell, end_cell = cell_ref.split(":", 1)
    else:
        start_cell = cell_ref
        end_cell = None

    start_row, start_col = _parse_single_cell(start_cell)

    if end_cell:
        end_row, end_col = _parse_single_cell(end_cell)
        if end_row < start_row or end_col < start_col:
            raise ValueError("Range end must not precede its start")
    else:
        end_row = None
        end_col = None

    return start_row, start_col, end_row, end_col


def parse_bounded_range(
    range_string: str,
    max_cells: int,
    error_cls: type[Exception] = ValueError,
    label: str = "operation",
    *,
    end_ref: str | None = None,
    require_range: bool = False,
) -> tuple[int, int, int, int]:
    """Parse a cell or range string and enforce a maximum cell-area cap."""
    try:
        start_row, start_col, end_row, end_col = parse_cell_range(range_string, end_ref)
    except ValueError as exc:
        raise error_cls(f"Invalid range specification '{range_string}': {exc}") from exc
    if require_range and (end_row is None or end_col is None):
        raise error_cls("Both start and end cells must be specified")
    eff_end_row = end_row if end_row is not None else start_row
    eff_end_col = end_col if end_col is not None else start_col
    assert_range_cell_limit(
        start_row,
        start_col,
        eff_end_row,
        eff_end_col,
        max_cells,
        error_cls,
        label=label,
    )
    return start_row, start_col, eff_end_row, eff_end_col


def validate_cell_reference(cell_ref: str) -> bool:
    """Validate Excel cell reference format and bounds (e.g., 'A1', 'XFD1048576')."""
    if not isinstance(cell_ref, str) or not cell_ref:
        return False
    try:
        _parse_single_cell(cell_ref)
        return True
    except ValueError:
        return False


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
