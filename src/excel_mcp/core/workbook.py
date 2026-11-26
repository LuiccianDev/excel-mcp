"""Core functionality for Excel workbook operations."""

from __future__ import annotations

from collections.abc import Generator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.core._io import resolve_sheet, validate_sheet_name, workbook_session
from excel_mcp.exceptions.exception_core import (
    SheetExistsError,
    ValidationError,
    WorkbookError,
    WorksheetError,
)


def create_workbook(
    filename: str | Path, sheet_name: str = "Sheet1", data_only: bool = False
) -> dict[str, Any]:
    """Create a new Excel workbook with an optional custom sheet name."""
    path = Path(filename).resolve()

    try:
        validate_sheet_name(sheet_name)
        with workbook_session(path, mode="create", initial_sheet=sheet_name):
            pass

        return {
            "status": "success",
            "message": f"Created workbook: {path}",
            "active_sheet": sheet_name,
            "filepath": str(path),
        }

    except PermissionError as e:
        raise PermissionError(f"Cannot write to {path}: {e}") from e
    except Exception as e:
        raise WorkbookError(f"Failed to create workbook: {e}") from e


def _create_new_sheet(workbook: Workbook, sheet_name: str) -> None:
    """Create a new sheet in the specified workbook."""
    if not isinstance(workbook, Workbook):
        raise TypeError(f"Expected Workbook instance, got {type(workbook).__name__}")

    if sheet_name in workbook.sheetnames:
        raise SheetExistsError(f"Sheet '{sheet_name}' already exists")

    try:
        workbook.create_sheet(sheet_name)
    except ValueError as e:
        raise ValidationError(f"Invalid sheet name '{sheet_name}': {e}") from e
    except Exception as e:
        raise WorksheetError(f"Failed to create sheet '{sheet_name}': {e}") from e


def create_sheet(filename: str | Path, sheet_name: str) -> dict[str, Any]:
    """Create a new sheet in an existing workbook, or create the workbook if missing."""
    path = Path(filename).resolve()

    if not path.exists():
        return create_workbook(path, sheet_name)

    try:
        validate_sheet_name(sheet_name)
        with workbook_session(path, mode="edit") as wb:
            _create_new_sheet(wb, sheet_name)

        return {
            "status": "success",
            "message": f"Created sheet '{sheet_name}' in {path}",
            "sheet_name": sheet_name,
            "filepath": str(path),
        }

    except (SheetExistsError, ValidationError, PermissionError):
        raise
    except Exception as e:
        raise WorkbookError(f"Failed to create sheet '{sheet_name}': {e}") from e


def get_workbook_info(
    filename: str | Path, include_ranges: bool = False
) -> dict[str, Any]:
    """Get metadata about an Excel workbook including sheets, file info, and ranges."""
    path = Path(filename).resolve()

    try:
        with workbook_session(path, mode="read", data_only=True) as wb:
            info: dict[str, Any] = {
                "filename": path.name,
                "sheets": wb.sheetnames,
                "size": path.stat().st_size,
                "modified": path.stat().st_mtime,
                "used_ranges": None,
            }

            if include_ranges:
                ranges: dict[str, str] = {}
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    if range_str := _get_worksheet_range(ws):
                        ranges[sheet_name] = range_str

                info["used_ranges"] = ranges or None

            return info

    except FileNotFoundError:
        raise
    except PermissionError as e:
        raise PermissionError(f"Cannot access {path}: {e}") from e
    except Exception as e:
        raise WorkbookError(f"Failed to get workbook info: {e}") from e


def _get_worksheet_range(worksheet: Any) -> str | None:
    """Get the used range of a worksheet as a string (e.g., 'A1:D10')."""
    max_row = worksheet.max_row or 0
    max_col = worksheet.max_column or 0
    if max_row > 0 and max_col > 0:
        return f"A1:{get_column_letter(max_col)}{max_row}"
    return None


@contextmanager
def managed_workbook(
    filename: str | Path, read_only: bool = False, auto_save: bool = True
) -> Generator[Workbook, None, None]:
    """Context manager for workbooks with automatic resource management."""
    with workbook_session(
        filename,
        mode="read" if read_only else "upsert",
        save=False if read_only else auto_save,
    ) as wb:
        yield wb


@contextmanager
def managed_worksheet(
    filename: str | Path, sheet_name: str, create_if_missing: bool = True
) -> Generator[Worksheet, None, None]:
    """Context manager for specific worksheets."""
    with workbook_session(filename, mode="upsert") as wb:
        yield resolve_sheet(wb, sheet_name, create_if_missing=create_if_missing)
