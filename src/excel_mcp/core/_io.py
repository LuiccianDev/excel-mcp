"""Atomic and safe file I/O operations for Excel workbooks."""

from __future__ import annotations

import os
from pathlib import Path
import uuid
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    import openpyxl


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
