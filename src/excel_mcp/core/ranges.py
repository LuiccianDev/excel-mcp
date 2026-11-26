"""Core worksheet and cell-range manipulation operations."""

from __future__ import annotations

from copy import copy
from typing import Any, Final

from openpyxl.styles import Alignment, Border, Font, GradientFill, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.core._io import resolve_sheet, validate_sheet_name, workbook_session
from excel_mcp.exceptions.exception_core import (
    RangeError,
    SheetExistsError,
    ValidationError,
)
from excel_mcp.utils.cell_utils import (
    MAX_EXCEL_COLS,
    MAX_EXCEL_ROWS,
    format_range_string,
    parse_bounded_range,
    parse_cell_range,
)

MAX_RANGE_CELLS: Final[int] = 250_000
MAX_MERGE_CELLS: Final[int] = 100_000

_CellSnapshot = tuple[
    Any,
    Font | None,
    Border | None,
    PatternFill | GradientFill | None,
    Alignment | None,
    str | None,
]


def _snapshot_cell(cell: Any) -> _CellSnapshot:
    if not cell.has_style:
        return (cell.value, None, None, None, None, None)
    return (
        cell.value,
        copy(cell.font) if cell.font else None,
        copy(cell.border) if cell.border else None,
        copy(cell.fill) if cell.fill else None,
        copy(cell.alignment) if cell.alignment else None,
        cell.number_format,
    )


def _apply_cell_snapshot(target_cell: Any, snapshot: _CellSnapshot) -> None:
    value, font, border, fill, alignment, number_format = snapshot
    target_cell.value = value
    target_cell.font = font if font is not None else Font()
    target_cell.border = border if border is not None else Border()
    target_cell.fill = fill if fill is not None else PatternFill()
    target_cell.alignment = alignment if alignment is not None else Alignment()
    target_cell.number_format = (
        number_format if number_format is not None else "General"
    )


def _clear_cell(cell: Any) -> None:
    cell.value = None
    cell.font = Font()
    cell.border = Border()
    cell.fill = PatternFill()
    cell.number_format = "General"
    cell.alignment = Alignment()


def copy_sheet(filename: str, source_sheet: str, target_sheet: str) -> dict[str, Any]:
    """Duplicate an existing worksheet within the same workbook."""
    validate_sheet_name(target_sheet)
    with workbook_session(filename, mode="edit") as wb:
        source = resolve_sheet(wb, source_sheet)
        if any(name.lower() == target_sheet.lower() for name in wb.sheetnames):
            raise SheetExistsError(f"Target sheet '{target_sheet}' already exists")
        target = wb.copy_worksheet(source)
        target.title = target_sheet
    return {
        "status": "success",
        "message": f"Sheet '{source_sheet}' copied to '{target_sheet}'",
    }


def delete_sheet(filename: str, sheet_name: str) -> dict[str, Any]:
    """Delete a worksheet while preventing removal of the sole remaining sheet."""
    with workbook_session(filename, mode="edit") as wb:
        resolve_sheet(wb, sheet_name)
        if len(wb.sheetnames) == 1:
            raise ValidationError("Cannot delete the only sheet in workbook")
        del wb[sheet_name]
    return {"status": "success", "message": f"Sheet '{sheet_name}' deleted"}


def rename_sheet(filename: str, old_name: str, new_name: str) -> dict[str, Any]:
    """Rename an existing worksheet."""
    validate_sheet_name(new_name)
    with workbook_session(filename, mode="edit") as wb:
        sheet = resolve_sheet(wb, old_name)
        if any(
            name.lower() == new_name.lower() and name != sheet.title
            for name in wb.sheetnames
        ):
            raise SheetExistsError(f"Sheet '{new_name}' already exists")
        sheet.title = new_name
    return {
        "status": "success",
        "message": f"Sheet renamed from '{old_name}' to '{new_name}'",
    }


def copy_range(
    source_ws: Worksheet,
    target_ws: Worksheet,
    source_range: str,
    target_start: str | None = None,
) -> None:
    """Copy a cell range and its styles losslessly, including overlapping self-copies."""
    src_start_row, src_start_col, eff_end_row, eff_end_col = parse_bounded_range(
        source_range, MAX_RANGE_CELLS, RangeError, label="copying"
    )

    dest_ref = target_start if target_start is not None else source_range.split(":")[0]
    try:
        tgt_start_row, tgt_start_col, tgt_end_row, tgt_end_col = parse_cell_range(
            dest_ref
        )
    except ValueError as exc:
        raise ValidationError(f"Invalid target cell '{dest_ref}': {exc}") from exc
    if tgt_end_row is not None or tgt_end_col is not None:
        raise ValidationError(f"Target start must be a single cell, got '{dest_ref}'.")

    row_count = eff_end_row - src_start_row + 1
    col_count = eff_end_col - src_start_col + 1
    if tgt_start_row + row_count - 1 > MAX_EXCEL_ROWS:
        raise RangeError(
            f"Destination range exceeds Excel row limit ({MAX_EXCEL_ROWS})."
        )
    if tgt_start_col + col_count - 1 > MAX_EXCEL_COLS:
        raise RangeError(
            f"Destination range exceeds Excel column limit ({MAX_EXCEL_COLS})."
        )

    snapshots = [
        [
            _snapshot_cell(source_ws.cell(row=row, column=col))
            for col in range(src_start_col, eff_end_col + 1)
        ]
        for row in range(src_start_row, eff_end_row + 1)
    ]

    for r_offset, row_snapshots in enumerate(snapshots):
        for c_offset, snapshot in enumerate(row_snapshots):
            target_cell = target_ws.cell(
                row=tgt_start_row + r_offset, column=tgt_start_col + c_offset
            )
            _apply_cell_snapshot(target_cell, snapshot)


def delete_range(
    worksheet: Worksheet, start_cell: str, end_cell: str | None = None
) -> None:
    """Clear cell values and reset formatting within a worksheet range."""
    start_row, start_col, eff_end_row, eff_end_col = parse_bounded_range(
        start_cell, MAX_RANGE_CELLS, RangeError, label="deletion", end_ref=end_cell
    )
    for row in range(start_row, eff_end_row + 1):
        for col in range(start_col, eff_end_col + 1):
            _clear_cell(worksheet.cell(row=row, column=col))


def merge_range(
    filename: str, sheet_name: str, start_cell: str, end_cell: str
) -> dict[str, Any]:
    """Merge a bounded rectangular range of cells."""
    start_row, start_col, end_row, end_col = parse_bounded_range(
        start_cell,
        MAX_MERGE_CELLS,
        RangeError,
        label="merging",
        end_ref=end_cell,
        require_range=True,
    )
    range_string = format_range_string(start_row, start_col, end_row, end_col)
    with workbook_session(filename, mode="edit") as wb:
        worksheet = resolve_sheet(wb, sheet_name)
        worksheet.merge_cells(range_string)
    return {
        "status": "success",
        "message": f"Range '{range_string}' merged in sheet '{sheet_name}'",
    }


def unmerge_range(
    filename: str, sheet_name: str, start_cell: str, end_cell: str
) -> dict[str, Any]:
    """Unmerge a previously merged cell range."""
    start_row, start_col, end_row, end_col = parse_bounded_range(
        start_cell,
        MAX_MERGE_CELLS,
        ValidationError,
        label="unmerging",
        end_ref=end_cell,
        require_range=True,
    )
    range_string = format_range_string(start_row, start_col, end_row, end_col)
    target_range = range_string.upper()
    with workbook_session(filename, mode="edit") as wb:
        worksheet = resolve_sheet(wb, sheet_name)
        merged_ranges = worksheet.merged_cells.ranges
        if not any(str(rng).upper() == target_range for rng in merged_ranges):
            raise RangeError(f"Range '{range_string}' is not merged")
        worksheet.unmerge_cells(range_string)
    return {
        "status": "success",
        "message": f"Range '{range_string}' unmerged successfully",
    }


def copy_range_operation(
    filename: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str | None = None,
) -> dict[str, Any]:
    """Copy a cell range and styles within or across worksheets."""
    parse_bounded_range(
        source_start,
        MAX_RANGE_CELLS,
        RangeError,
        label="copying",
        end_ref=source_end,
        require_range=True,
    )
    with workbook_session(filename, mode="edit") as wb:
        source_ws = resolve_sheet(wb, sheet_name)
        target_ws = resolve_sheet(wb, target_sheet) if target_sheet else source_ws
        copy_range(source_ws, target_ws, f"{source_start}:{source_end}", target_start)
    return {"status": "success", "message": "Range copied successfully"}


def _shift_subrange(
    worksheet: Worksheet,
    start_row: int,
    start_col: int,
    eff_end_row: int,
    eff_end_col: int,
    max_row: int,
    max_col: int,
    shift_direction: str,
) -> None:
    if shift_direction == "up":
        span = eff_end_row - start_row + 1
        for col in range(start_col, eff_end_col + 1):
            for r in range(start_row, max_row - span + 1):
                _apply_cell_snapshot(
                    worksheet.cell(row=r, column=col),
                    _snapshot_cell(worksheet.cell(row=r + span, column=col)),
                )
            for r in range(max(start_row, max_row - span + 1), max_row + 1):
                _clear_cell(worksheet.cell(row=r, column=col))
        return

    span = eff_end_col - start_col + 1
    for row in range(start_row, eff_end_row + 1):
        for c in range(start_col, max_col - span + 1):
            _apply_cell_snapshot(
                worksheet.cell(row=row, column=c),
                _snapshot_cell(worksheet.cell(row=row, column=c + span)),
            )
        for c in range(max(start_col, max_col - span + 1), max_col + 1):
            _clear_cell(worksheet.cell(row=row, column=c))


def delete_range_operation(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str | None = None,
    shift_direction: str = "up",
) -> dict[str, Any]:
    """Delete a sub-range and shift only the affected column or row slice."""
    if shift_direction not in ("up", "left"):
        raise ValidationError(
            f"Invalid shift direction: {shift_direction}. Must be 'up' or 'left'"
        )
    start_row, start_col, eff_end_row, eff_end_col = parse_bounded_range(
        start_cell, MAX_RANGE_CELLS, RangeError, label="deletion", end_ref=end_cell
    )

    with workbook_session(filename, mode="edit") as wb:
        worksheet = resolve_sheet(wb, sheet_name)
        max_row = worksheet.max_row or 0
        max_col = worksheet.max_column or 0
        if eff_end_row > max_row:
            raise RangeError(f"End row {eff_end_row} out of bounds (1-{max_row})")
        if eff_end_col > max_col:
            raise RangeError(f"End column {eff_end_col} out of bounds (1-{max_col})")
        _shift_subrange(
            worksheet,
            start_row,
            start_col,
            eff_end_row,
            eff_end_col,
            max_row,
            max_col,
            shift_direction,
        )

    range_string = format_range_string(start_row, start_col, eff_end_row, eff_end_col)
    return {
        "status": "success",
        "message": f"Range {range_string} deleted successfully",
    }
