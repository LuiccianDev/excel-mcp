"""Core functionality for reading, writing, and validating Excel data ranges."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Final

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
    ValidationError,
    WorkbookError,
)
from excel_mcp.models.response_models import BatchWriteResult
from excel_mcp.utils.cell_utils import (
    MAX_EXCEL_COLS,
    assert_range_cell_limit,
    get_column_letter,
    parse_cell_range,
    validate_cell_reference,
)

CellValue = Any
RowData = list[CellValue]
SheetData = list[RowData]
CellRange = str

MAX_READ_CELLS: Final[int] = 250_000
MAX_USED_RANGE_SCAN_ROWS: Final[int] = 10_000
MAX_USED_RANGE_SCAN_COLS: Final[int] = 200
PREVIEW_MAX_ROWS: Final[int] = 100


def _parse_cell_reference(cell_ref: str) -> tuple[int, int]:
    """Parse cell reference into (row, column) coordinates (1-based)."""
    try:
        start_row, start_col, end_row, end_col = parse_cell_range(cell_ref)
        if end_row is not None or end_col is not None:
            raise InvalidCellReferenceError(
                f"Expected single cell reference: {cell_ref}"
            )
        return start_row, start_col
    except ValueError as e:
        raise InvalidCellReferenceError(f"Invalid cell format: {e}") from e


def _get_used_range(
    worksheet: Worksheet,
    start_row: int,
    start_col: int,
    *,
    max_rows: int | None = None,
) -> tuple[int, int]:
    """Find the populated end coordinates starting from (start_row, start_col)."""
    ws_max_row = worksheet.max_row
    ws_max_col = worksheet.max_column
    if ws_max_row is not None and ws_max_col is not None:
        if ws_max_row < start_row or ws_max_col < start_col:
            return start_row - 1, start_col - 1

        if max_rows is None and (
            ws_max_row - start_row + 1 > MAX_USED_RANGE_SCAN_ROWS
            or ws_max_col - start_col + 1 > MAX_USED_RANGE_SCAN_COLS
        ):
            raise RangeError(
                f"Worksheet dimensions ({ws_max_row}x{ws_max_col}) exceed "
                f"auto-scan limit ({MAX_USED_RANGE_SCAN_ROWS}x{MAX_USED_RANGE_SCAN_COLS}); specify an explicit range."
            )

    row_limit = max_rows if max_rows is not None else MAX_USED_RANGE_SCAN_ROWS
    max_scan_row = (
        min(ws_max_row, start_row + row_limit - 1)
        if ws_max_row is not None
        else start_row + row_limit - 1
    )
    max_scan_col = (
        min(ws_max_col, start_col + MAX_USED_RANGE_SCAN_COLS - 1)
        if ws_max_col is not None
        else start_col + MAX_USED_RANGE_SCAN_COLS - 1
    )

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


def read_excel_range(
    filename: Path | str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str | None = None,
    preview_only: bool = False,
) -> SheetData:
    """Read a range of data from an Excel worksheet."""
    if not sheet_name:
        raise SheetNotFoundError(f"Sheet '{sheet_name}' not found")

    path = Path(filename).resolve()

    try:
        with workbook_session(path, mode="read", data_only=True) as wb:
            ws = resolve_sheet(wb, sheet_name)

            if ":" in start_cell and end_cell is None:
                start_cell, end_cell = start_cell.split(":", 1)

            start_row, start_col = _parse_cell_reference(start_cell)

            if end_cell:
                end_row, end_col = _parse_cell_reference(end_cell)
                if end_row < start_row or end_col < start_col:
                    raise RangeError("Range end must not precede its start")
            else:
                end_row, end_col = _get_used_range(
                    ws,
                    start_row,
                    start_col,
                    max_rows=PREVIEW_MAX_ROWS if preview_only else None,
                )

            ws_max_row = ws.max_row
            ws_max_col = ws.max_column
            if (ws_max_row is not None and start_row > ws_max_row) or (
                ws_max_col is not None and start_col > ws_max_col
            ):
                max_cell = f"{get_column_letter(ws_max_col or 1)}{ws_max_row or 1}"
                raise RangeError(
                    f"Start cell out of bounds. Sheet dimensions are A1:{max_cell}"
                )

            if preview_only:
                end_row = min(end_row, start_row + PREVIEW_MAX_ROWS - 1)

            effective_end_row = (
                min(end_row, ws_max_row) if ws_max_row is not None else end_row
            )
            effective_end_col = (
                min(end_col, ws_max_col) if ws_max_col is not None else end_col
            )

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
            | DataError
            | RangeError
            | ValidationError,
        ):
            raise DataError(f"Failed to read Excel range: {e}") from e
        raise


def _write_data_to_worksheet(
    worksheet: Worksheet,
    data: SheetData,
    start_row: int,
    start_col: int,
    *,
    skip_none: bool = True,
) -> int:
    """Write a 2D data matrix into a worksheet starting at (start_row, start_col)."""
    cells_written = 0
    for row_idx, row_data in enumerate(data):
        for col_idx, value in enumerate(row_data):
            if skip_none and value is None:
                continue
            assert_safe_cell_value(
                value, f"at row offset {row_idx}, col offset {col_idx}"
            )
            worksheet.cell(
                row=start_row + row_idx, column=start_col + col_idx, value=value
            )
            cells_written += 1
    return cells_written


def batch_write_matrix(
    filename: str,
    sheet_name: str | None,
    start_cell: str,
    data_matrix: list[list[Any]],
) -> BatchWriteResult:
    """Write a 2D matrix of values into a sheet starting from a given cell in one atomic pass."""
    start_row, start_col = _parse_cell_reference(start_cell)
    rows_affected = len(data_matrix)
    cols_affected = max((len(r) for r in data_matrix), default=0)

    with workbook_session(filename, mode="edit") as wb:
        ws = resolve_sheet(wb, sheet_name)
        total_cells = _write_data_to_worksheet(
            ws, data_matrix, start_row, start_col, skip_none=False
        )
        title = ws.title

    end_cell = (
        f"{get_column_letter(start_col + max(cols_affected - 1, 0))}"
        f"{start_row + max(rows_affected - 1, 0)}"
    )
    return BatchWriteResult(
        filename=filename,
        sheet_name=title,
        range_written=f"{start_cell}:{end_cell}",
        rows_affected=rows_affected,
        columns_affected=cols_affected,
        cells_updated=total_cells,
    )


def write_data(
    filename: str | Path,
    sheet_name: str | None,
    data: SheetData | None,
    start_cell: str = "A1",
) -> dict[str, str]:
    """Write a 2D data matrix to an Excel worksheet."""
    if not data:
        raise InvalidDataError("No data provided to write")

    path = Path(filename).resolve()
    try:
        start_row, start_col = _parse_cell_reference(start_cell)

        with workbook_session(
            path,
            mode="upsert",
            initial_sheet=sheet_name or "Sheet1",
        ) as wb:
            try:
                ws = resolve_sheet(wb, sheet_name or None, create_if_missing=True)
            except Exception as e:
                raise SheetError(f"Failed to access worksheet: {e}") from e

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
            raise WorkbookError(f"Failed to write data: {e}") from e
        raise


def validate_range_bounds(
    worksheet: Worksheet,
    start_row: int,
    start_col: int,
    end_row: int | None = None,
    end_col: int | None = None,
) -> tuple[bool, str]:
    """Check whether cell coordinates fall inside populated worksheet bounds."""
    max_row = worksheet.max_row or 0
    max_col = worksheet.max_column or 0
    if start_row < 1 or start_row > max_row:
        return False, f"Start row {start_row} out of bounds (1-{max_row})"
    if start_col < 1 or start_col > max_col:
        start_col_label = (
            get_column_letter(start_col)
            if 1 <= start_col <= MAX_EXCEL_COLS
            else str(start_col)
        )
        max_col_letter = get_column_letter(max(max_col, 1))
        return False, (
            f"Start column {start_col_label} out of bounds (A-{max_col_letter})"
        )
    if end_row is not None and end_col is not None:
        if end_row < start_row:
            return False, "End row cannot be before start row"
        if end_col < start_col:
            return False, "End column cannot be before start column"
        if end_row > max_row:
            return False, f"End row {end_row} out of bounds (1-{max_row})"
        if end_col > max_col:
            end_col_label = (
                get_column_letter(end_col)
                if 1 <= end_col <= MAX_EXCEL_COLS
                else str(end_col)
            )
            max_col_letter = get_column_letter(max(max_col, 1))
            return False, (
                f"End column {end_col_label} out of bounds (A-{max_col_letter})"
            )
    return True, "Range is valid"


def validate_range_in_sheet_operation(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str | None = None,
) -> dict[str, Any]:
    """Validate that a cell range falls within a worksheet's populated dimensions."""
    if not validate_cell_reference(start_cell):
        return {
            "status": "error",
            "message": f"Invalid start cell reference: {start_cell}",
        }
    if end_cell and not validate_cell_reference(end_cell):
        return {"status": "error", "message": f"Invalid end cell reference: {end_cell}"}

    try:
        with workbook_session(filepath, mode="read") as wb:
            worksheet = resolve_sheet(wb, sheet_name)
            data_max_row = worksheet.max_row or 0
            data_max_col = worksheet.max_column or 0
            try:
                start_row, start_col, end_row, end_col = parse_cell_range(
                    start_cell, end_cell
                )
            except ValueError as e:
                return {"status": "error", "message": f"Invalid range: {e}"}
            if end_row is None:
                end_row = start_row
            if end_col is None:
                end_col = start_col
            is_valid, message = validate_range_bounds(
                worksheet, start_row, start_col, end_row, end_col
            )
            if not is_valid:
                return {"status": "error", "message": message}
            range_str = (
                f"{start_cell}" if end_cell is None else f"{start_cell}:{end_cell}"
            )
            max_col_letter = get_column_letter(max(data_max_col, 1))
            data_range_str = f"A1:{max_col_letter}{data_max_row}"
            extends_beyond_data = end_row > data_max_row or end_col > data_max_col
            return {
                "status": "success",
                "message": (
                    f"Range '{range_str}' is valid. "
                    f"Sheet contains data in range '{data_range_str}'"
                ),
                "valid": True,
                "range": range_str,
                "data_range": data_range_str,
                "extends_beyond_data": extends_beyond_data,
                "data_dimensions": {
                    "max_row": data_max_row,
                    "max_col": data_max_col,
                    "max_col_letter": max_col_letter,
                },
            }
    except SheetNotFoundError:
        return {"status": "error", "message": f"Sheet '{sheet_name}' not found"}
    except Exception as e:
        return {"status": "error", "message": str(e)}
