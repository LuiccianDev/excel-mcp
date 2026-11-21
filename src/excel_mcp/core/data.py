"""Core functionality for reading and writing Excel data.

This module provides functions to interact with Excel files, including reading ranges,
writing data, and detecting headers, following the Model Context Protocol (MCP) standards.
"""

from pathlib import Path
from typing import Any, Final, TypedDict

from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook as OpenpyxlWorkbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.core._io import resolve_sheet, workbook_session
from excel_mcp.core.calculations import FormulaSecurityError, assert_safe_cell_value
from excel_mcp.exceptions.exception_core import (
    DataError,
    InvalidCellReferenceError,
    InvalidDataError,
    RangeError,
    SheetError,
    SheetNotFoundError,
    WorkbookError,
)
from excel_mcp.utils.cell_utils import assert_range_cell_limit, parse_cell_range


# Type aliases for better code readability
CellValue = Any
RowData = list[CellValue]
SheetData = list[RowData]
CellRange = str

MAX_READ_CELLS: Final[int] = 250_000
MAX_USED_RANGE_SCAN_ROWS: Final[int] = 10_000
MAX_USED_RANGE_SCAN_COLS: Final[int] = 200
PREVIEW_MAX_ROWS: Final[int] = 100


class RangeCoordinates(TypedDict):
    """Type definition for cell range coordinates."""

    start_row: int
    start_col: int
    end_row: int
    end_col: int


def _get_worksheet(workbook: OpenpyxlWorkbook, sheet_name: str) -> Worksheet:
    """Get a worksheet by name, raising appropriate exceptions.

    Args:
        workbook: Openpyxl workbook object.
        sheet_name: Name of the worksheet to retrieve.

    Returns:
        The requested worksheet.

    Raises:
        SheetNotFoundError: If the specified sheet doesn't exist.
    """
    if sheet_name not in workbook.sheetnames:
        raise SheetNotFoundError(f"Sheet '{sheet_name}' not found")
    return workbook[sheet_name]


def _parse_cell_reference(cell_ref: str) -> tuple[int, int]:
    """Parse cell reference into (row, column) coordinates.

    Args:
        cell_ref: Cell reference (e.g., 'A1', 'B2').

    Returns:
        Tuple of (row, column) coordinates (1-based).

    Raises:
        InvalidCellReferenceError: If the cell reference is invalid.
    """
    try:
        coords = parse_cell_range(f"{cell_ref}:{cell_ref}")
        if not coords or None in coords[:2]:
            raise InvalidCellReferenceError(f"Invalid cell reference: {cell_ref}")
        return coords[0], coords[1]  # row, column
    except ValueError as e:
        raise InvalidCellReferenceError(f"Invalid cell format: {str(e)}") from e


def _get_used_range(
    worksheet: Worksheet,
    start_row: int,
    start_col: int,
    *,
    max_rows: int | None = None,
) -> tuple[int, int]:
    """Find the used range in a worksheet starting from given coordinates.

    Args:
        worksheet: The worksheet to search.
        start_row: Starting row (1-based).
        start_col: Starting column (1-based).
        max_rows: Optional explicit row window (e.g. preview mode).

    Returns:
        Tuple of (end_row, end_col) coordinates.
    """
    if worksheet.max_row < start_row or worksheet.max_column < start_col:
        return start_row - 1, start_col - 1

    if max_rows is None and (
        worksheet.max_row - start_row + 1 > MAX_USED_RANGE_SCAN_ROWS
        or worksheet.max_column - start_col + 1 > MAX_USED_RANGE_SCAN_COLS
    ):
        raise RangeError(
            f"Worksheet dimensions ({worksheet.max_row}x{worksheet.max_column}) exceed "
            f"auto-scan limit ({MAX_USED_RANGE_SCAN_ROWS}x{MAX_USED_RANGE_SCAN_COLS}); specify an explicit range."
        )

    row_limit = max_rows if max_rows is not None else MAX_USED_RANGE_SCAN_ROWS
    max_scan_row = min(worksheet.max_row, start_row + row_limit - 1)
    max_scan_col = min(worksheet.max_column, start_col + MAX_USED_RANGE_SCAN_COLS - 1)

    last_row = start_row - 1
    last_col = start_col - 1
    for r_idx, row in enumerate(
        worksheet.iter_rows(
            min_row=start_row,
            max_row=max_scan_row,
            min_col=start_col,
            max_col=max_scan_col,
            values_only=True,
        ),
        start=start_row,
    ):
        for c_idx, val in enumerate(row, start=start_col):
            if val is not None:
                if r_idx > last_row:
                    last_row = r_idx
                if c_idx > last_col:
                    last_col = c_idx

    return last_row, last_col


# * Read a range of data from an Excel worksheet
def read_excel_range(
    filename: Path | str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str | None = None,
    preview_only: bool = False,
) -> SheetData:
    """Read a range of data from an Excel worksheet.

    Args:
        filename: Path to the Excel file.
        sheet_name: Name of the worksheet to read from.
        start_cell: Starting cell reference (e.g., 'A1').
        end_cell: Optional ending cell reference. If not provided, will detect used range.
        preview_only: If True, only returns a preview of the data.

    Returns:
        A 2D list containing the cell values in the specified range.

    Raises:
        FileNotFoundError: If the specified file doesn't exist.
        SheetNotFoundError: If the specified sheet doesn't exist.
        InvalidCellReferenceError: If cell references are invalid.
        DataError: For errors related to data processing.
    """
    if not sheet_name:
        raise SheetNotFoundError(f"Sheet '{sheet_name}' not found")

    path = Path(filename).resolve()

    try:
        with workbook_session(path, mode="read", data_only=True) as wb:
            ws = resolve_sheet(wb, sheet_name)

            # Handle range string (e.g., "A1:B2")
            if ":" in start_cell:
                start_cell, end_cell = start_cell.split(":", 1)

            # Parse start cell coordinates
            start_row, start_col = _parse_cell_reference(start_cell)

            # Parse or detect end cell coordinates
            if end_cell:
                end_row, end_col = _parse_cell_reference(end_cell)
            else:
                end_row, end_col = _get_used_range(
                    ws,
                    start_row,
                    start_col,
                    max_rows=PREVIEW_MAX_ROWS if preview_only else None,
                )

            # Validate range is within worksheet bounds
            if start_row > ws.max_row or start_col > ws.max_column:
                max_cell = f"{get_column_letter(ws.max_column)}{ws.max_row}"
                raise RangeError(
                    f"Start cell out of bounds. Sheet dimensions are A1:{max_cell}"
                )

            if preview_only:
                end_row = min(end_row, start_row + PREVIEW_MAX_ROWS - 1)

            effective_end_row = min(end_row, ws.max_row)
            effective_end_col = min(end_col, ws.max_column)

            data: SheetData = []
            if effective_end_row >= start_row and effective_end_col >= start_col:
                assert_range_cell_limit(
                    start_row,
                    start_col,
                    effective_end_row,
                    effective_end_col,
                    MAX_READ_CELLS,
                    RangeError,
                    label="reading",
                )
                for row in ws.iter_rows(
                    min_row=start_row,
                    max_row=effective_end_row,
                    min_col=start_col,
                    max_col=effective_end_col,
                    values_only=True,
                ):
                    row_data = list(row)
                    if any(v is not None for v in row_data):
                        data.append(row_data)

            return data

    except Exception as e:
        if not isinstance(
            e,
            FileNotFoundError
            | SheetNotFoundError
            | InvalidCellReferenceError
            | DataError,
        ):
            raise DataError(f"Failed to read Excel range: {str(e)}") from e
        raise


def _write_data_to_worksheet(
    worksheet: Worksheet, data: SheetData, start_row: int, start_col: int
) -> None:
    """Write data to a worksheet starting at the specified coordinates.

    Args:
        worksheet: The worksheet to write to.
        data: 2D list of data to write.
        start_row: Starting row (1-based).
        start_col: Starting column (1-based).

    Note:
        None values are skipped to preserve existing cell formatting.
    """
    for row_idx, row_data in enumerate(data):
        for col_idx, value in enumerate(row_data):
            if value is not None:
                assert_safe_cell_value(
                    value, f"at row offset {row_idx}, col offset {col_idx}"
                )
                worksheet.cell(
                    row=start_row + row_idx, column=start_col + col_idx, value=value
                )


# * Write data to an Excel worksheet
def write_data(
    filename: str | Path,
    sheet_name: str | None,
    data: SheetData | None,
    start_cell: str = "A1",
) -> dict[str, str]:
    """Write data to an Excel worksheet.

    Args:
        filename: Path to the Excel file.
        sheet_name: Name of the worksheet to write to. If None, uses active sheet.
        data: 2D list of data to write.
        start_cell: Starting cell reference (e.g., 'A1').

    Returns:
        Dictionary with operation status and details.

    Raises:
        InvalidDataError: If input data is invalid.
        InvalidCellReferenceError: If cell reference is invalid.
        SheetError: For sheet-related errors.
        WorkbookError: For other workbook-related errors.
    """
    if not data:
        raise InvalidDataError("No data provided to write")

    path = Path(filename).resolve()
    try:
        # Parse start cell before mutating or creating files
        try:
            start_row, start_col = _parse_cell_reference(start_cell)
        except InvalidCellReferenceError as e:
            raise InvalidCellReferenceError(
                f"Invalid start cell format: {str(e)}"
            ) from e

        with workbook_session(
            path,
            mode="upsert",
            initial_sheet=sheet_name or "Sheet1",
        ) as wb:
            try:
                ws = resolve_sheet(wb, sheet_name or None, create_if_missing=True)
            except Exception as e:
                raise SheetError(f"Failed to access worksheet: {str(e)}") from e

            _write_data_to_worksheet(ws, data, start_row, start_col)
            title = ws.title

        return {
            "status": "success",
            "message": f"Data written to {title}",
            "active_sheet": title,
            "filepath": str(path),
        }

    except Exception as e:
        if not isinstance(
            e,
            InvalidDataError
            | InvalidCellReferenceError
            | SheetError
            | WorkbookError
            | FormulaSecurityError,
        ):
            raise WorkbookError(f"Failed to write data: {str(e)}") from e
        raise
