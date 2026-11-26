"""Atomic and safe file I/O operations for Excel workbooks."""

from __future__ import annotations

from contextlib import contextmanager
import os
from pathlib import Path
import re
import uuid
import zipfile
from collections.abc import Iterator
from typing import TYPE_CHECKING, Final, Literal

import openpyxl

from excel_mcp.exceptions import SheetNotFoundError, ValidationError
from excel_mcp.utils.cell_utils import parse_cell_range
from excel_mcp.utils.metrics import measure_phase
from excel_mcp.utils.paths import DEFAULT_MAX_FILE_SIZE

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet

SessionMode = Literal["read", "edit", "create", "upsert"]
_INVALID_SHEET_CHARS = frozenset(r"[]:*?/\\")
MAX_XLSX_ENTRIES: Final[int] = 10_000
MAX_XLSX_UNCOMPRESSED_BYTES: Final[int] = 250 * 1024 * 1024
MAX_XLSX_COMPRESSION_RATIO: Final[float] = 100.0
_RATIO_CHECK_MIN_BYTES: Final[int] = 1024 * 1024
MAX_XLSX_MERGE_CELLS: Final[int] = 10_000
MAX_XLSX_MERGE_AREA: Final[int] = 100_000
MAX_XLSX_HYPERLINKS: Final[int] = 10_000
MAX_XLSX_HYPERLINK_AREA: Final[int] = 100_000
MAX_WORKSHEET_XML_BYTES: Final[int] = 50 * 1024 * 1024
_MERGE_CELL_REF_RE: Final[re.Pattern[bytes]] = re.compile(
    rb"<(?:\w+:)?mergeCell\b[^<>]*\bref\s*=\s*[\"']([^\"']+)[\"']"
)
_HYPERLINK_TAG_RE: Final[re.Pattern[bytes]] = re.compile(
    rb"<(?:\w+:)?hyperlink\b([^<>]*)/?>"
)
_REF_ATTR_RE: Final[re.Pattern[bytes]] = re.compile(rb"\bref\s*=\s*[\"']([^\"']+)[\"']")


def _range_area_from_ref(ref_str: str) -> int:
    try:
        start_row, start_col, end_row, end_col = parse_cell_range(
            ref_str.replace("$", "")
        )
    except ValueError:
        return 0
    eff_end_row = end_row if end_row is not None else start_row
    eff_end_col = end_col if end_col is not None else start_col
    return (eff_end_row - start_row + 1) * (eff_end_col - start_col + 1)


def _check_xml_tag_bounds(
    data: bytes,
    tag_pattern: re.Pattern[bytes],
    max_count: int,
    max_area: int,
    label: str,
    member_name: str,
    archive_name: str,
    *,
    extract_attr_ref: bool = False,
) -> None:
    count = 0
    total_area = 0
    for match in tag_pattern.finditer(data):
        count += 1
        if count > max_count:
            raise ValidationError(
                f"Worksheet '{member_name}' in '{archive_name}' exceeds {label} limit ({max_count})."
            )
        if extract_attr_ref:
            ref_match = _REF_ATTR_RE.search(match.group(1))
            raw_ref = ref_match.group(1) if ref_match is not None else b""
        else:
            raw_ref = match.group(1)
        if raw_ref:
            ref_str = raw_ref.decode("ascii", errors="ignore")
            total_area += _range_area_from_ref(ref_str)
            if total_area > max_area:
                raise ValidationError(
                    f"Worksheet '{member_name}' in '{archive_name}' exceeds {label} area limit ({max_area} cells)."
                )


def _read_bounded_zip_member(
    zf: zipfile.ZipFile, info: zipfile.ZipInfo, max_bytes: int, archive_name: str
) -> bytes:
    with zf.open(info, "r") as member:
        data = member.read(max_bytes + 1)
    if len(data) > max_bytes:
        raise ValidationError(
            f"Archive member '{info.filename}' in '{archive_name}' exceeds decompressed byte limit ({max_bytes} bytes)."
        )
    return data


def _validate_worksheet_xml_bounds(
    zf: zipfile.ZipFile, infos: list[zipfile.ZipInfo], archive_name: str
) -> None:
    for info in infos:
        lower_name = info.filename.lower()
        if not (lower_name.startswith("xl/") and lower_name.endswith(".xml")):
            continue
        if lower_name in (
            "xl/sharedstrings.xml",
            "xl/styles.xml",
            "xl/calcchain.xml",
            "xl/workbook.xml",
        ):
            continue
        data = _read_bounded_zip_member(zf, info, MAX_WORKSHEET_XML_BYTES, archive_name)
        _check_xml_tag_bounds(
            data,
            _MERGE_CELL_REF_RE,
            MAX_XLSX_MERGE_CELLS,
            MAX_XLSX_MERGE_AREA,
            "mergeCell",
            info.filename,
            archive_name,
        )
        _check_xml_tag_bounds(
            data,
            _HYPERLINK_TAG_RE,
            MAX_XLSX_HYPERLINKS,
            MAX_XLSX_HYPERLINK_AREA,
            "hyperlink",
            info.filename,
            archive_name,
            extract_attr_ref=True,
        )


def _validate_xlsx_archive(path: Path) -> None:
    if not zipfile.is_zipfile(path):
        return
    with zipfile.ZipFile(path, "r") as zf:
        infos = zf.infolist()
        if len(infos) > MAX_XLSX_ENTRIES:
            raise ValidationError(
                f"Archive '{path.name}' has {len(infos)} entries, exceeding limit of {MAX_XLSX_ENTRIES}."
            )
        total_uncompressed = 0
        total_compressed = 0
        for info in infos:
            total_uncompressed += info.file_size
            total_compressed += info.compress_size
            if total_uncompressed > MAX_XLSX_UNCOMPRESSED_BYTES:
                raise ValidationError(
                    f"Archive '{path.name}' decompressed size exceeds {MAX_XLSX_UNCOMPRESSED_BYTES} bytes."
                )
            if info.file_size > _RATIO_CHECK_MIN_BYTES and (
                info.compress_size == 0
                or info.file_size / info.compress_size > MAX_XLSX_COMPRESSION_RATIO
            ):
                raise ValidationError(
                    f"Archive member '{info.filename}' in '{path.name}' exceeds compression ratio limit."
                )
        if total_uncompressed > _RATIO_CHECK_MIN_BYTES and (
            total_compressed == 0
            or total_uncompressed / total_compressed > MAX_XLSX_COMPRESSION_RATIO
        ):
            raise ValidationError(
                f"Archive '{path.name}' exceeds aggregate compression ratio limit."
            )
        _validate_worksheet_xml_bounds(zf, infos, path.name)


def _validate_new_sheet_name(name: str) -> None:
    if not name or not name.strip():
        raise ValidationError("Sheet name cannot be empty.")
    if len(name) > 31:
        raise ValidationError(
            f"Sheet name '{name}' exceeds Excel's 31-character limit."
        )
    if any(ch in _INVALID_SHEET_CHARS for ch in name):
        raise ValidationError(f"Sheet name '{name}' contains invalid Excel characters.")


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

    if mode == "read" and should_save:
        raise ValueError("Cannot save a workbook opened in 'read' mode.")
    if data_only and should_save:
        raise ValueError(
            "Refusing to save a workbook loaded with data_only=True; pass save=False."
        )

    is_new = mode == "create" or (mode == "upsert" and not path.exists())
    if mode == "read" or mode == "edit" or (mode == "upsert" and not is_new):
        if not path.exists():
            raise FileNotFoundError(f"Workbook not found: {path}")
        _validate_xlsx_archive(path)
        with measure_phase("io_read_ms"):
            wb = openpyxl.load_workbook(
                str(path), read_only=(mode == "read"), data_only=data_only
            )
    elif is_new:
        if initial_sheet is not None:
            _validate_new_sheet_name(initial_sheet)
        wb = openpyxl.Workbook()
    else:
        raise ValueError(f"Invalid workbook session mode: {mode!r}")

    try:
        if is_new and initial_sheet is not None:
            active = wb.active
            if active is not None:
                active.title = initial_sheet
            else:
                wb.create_sheet(initial_sheet)
        yield wb
        if should_save:
            with measure_phase("io_save_ms"):
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
    if sheet_name:
        if sheet_name in wb.sheetnames:
            return wb[sheet_name]
        if create_if_missing:
            if getattr(wb, "read_only", False):
                raise ValueError("Cannot create a worksheet on a read-only workbook.")
            _validate_new_sheet_name(sheet_name)
            return wb.create_sheet(sheet_name)
        raise SheetNotFoundError(f"Worksheet '{sheet_name}' not found in workbook.")
    if create_if_missing and sheet_name == "":
        raise ValidationError("Sheet name cannot be empty.")
    if wb.active is None:
        raise SheetNotFoundError("Workbook has no active worksheet.")
    return wb.active


def atomic_save(
    wb: openpyxl.Workbook,
    target_path: str | Path,
    *,
    max_file_size: int = DEFAULT_MAX_FILE_SIZE,
) -> Path:
    """Save an openpyxl Workbook to target_path atomically.

    Writes to a temporary file in the same directory (guaranteeing the same filesystem)
    and then renames it via os.replace to prevent partial or corrupted writes if an
    interruption, process kill, or crash occurs during serialization.

    Args:
        wb: The openpyxl Workbook object to serialize.
        target_path: Destination path for the .xlsx file.
        max_file_size: Maximum allowed serialized file size in bytes.

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
        size = temp_path.stat().st_size
        if size > max_file_size:
            raise ValidationError(
                f"Saved workbook ({size} bytes) exceeds maximum allowed size ({max_file_size} bytes)."
            )
        os.replace(temp_path, dest)
        return dest
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass
