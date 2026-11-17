"""Atomic and safe file I/O operations for Excel workbooks."""

from __future__ import annotations

from contextlib import contextmanager
import os
from pathlib import Path
import uuid
from collections.abc import Iterator
from typing import TYPE_CHECKING, Literal

import openpyxl

from excel_mcp.exceptions import SheetNotFoundError

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet

SessionMode = Literal["read", "edit", "create", "upsert"]


@contextmanager
def workbook_session(
    filename: str | Path,
    *,
    mode: SessionMode = "edit",
    data_only: bool = False,
    initial_sheet: str | None = "Sheet1",
    save: bool | None = None,
) -> Iterator[openpyxl.Workbook]:
    """Open or initialize an openpyxl Workbook, atomically save on clean write exit, and always close."""
    path = Path(filename).resolve()
    should_save = (mode != "read") if save is None else save

    if mode == "read":
        if not path.exists():
            raise FileNotFoundError(f"Workbook not found: {path}")
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=data_only)
    elif mode == "edit":
        if not path.exists():
            raise FileNotFoundError(f"Workbook not found: {path}")
        wb = openpyxl.load_workbook(str(path), data_only=data_only)
    elif mode == "create" or (mode == "upsert" and not path.exists()):
        path.parent.mkdir(parents=True, exist_ok=True)
        wb = openpyxl.Workbook()
        if initial_sheet is not None:
            active = wb.active
            if active is not None:
                active.title = initial_sheet
            else:
                wb.create_sheet(initial_sheet)
    elif mode == "upsert":
        wb = openpyxl.load_workbook(str(path), data_only=data_only)
    else:
        raise ValueError(f"Invalid workbook session mode: {mode!r}")

    try:
        yield wb
        if should_save:
            atomic_save(wb, path)
    finally:
        wb.close()


def resolve_sheet(
    wb: openpyxl.Workbook,
    sheet_name: str | None = None,
    *,
    create_if_missing: bool = False,
) -> Worksheet:
    """Resolve sheet_name (or active sheet), optionally creating it, or raise SheetNotFoundError."""
    if sheet_name is not None:
        if sheet_name in wb.sheetnames:
            return wb[sheet_name]
        if create_if_missing:
            return wb.create_sheet(sheet_name)
        raise SheetNotFoundError(f"Worksheet '{sheet_name}' not found in workbook.")
    if wb.active is None:
        raise SheetNotFoundError("Workbook has no active worksheet.")
    return wb.active


def atomic_save(wb: openpyxl.Workbook, target_path: str | Path) -> Path:
    """Save an openpyxl Workbook to target_path atomically.

    Writes to a temporary file in the same directory (guaranteeing the same filesystem)
    and then renames it via os.replace to prevent partial or corrupted writes if an
    interruption, process kill, or crash occurs during serialization.

    Args:
        wb: The openpyxl Workbook object to serialize.
        target_path: Destination path for the .xlsx file.

    Returns:
        Path: Resolved destination Path.
    """
    dest = Path(target_path).resolve()
    parent_dir = dest.parent
    parent_dir.mkdir(parents=True, exist_ok=True)

    # Use hidden temp file in the same parent directory for atomic same-device rename
    temp_name = f".tmp_{uuid.uuid4().hex}_{dest.name}"
    temp_path = parent_dir / temp_name

    try:
        wb.save(str(temp_path))
        os.replace(temp_path, dest)
        return dest
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass
