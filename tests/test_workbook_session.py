"""Tests for workbook_session and resolve_sheet lifecycle helpers."""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from excel_mcp.core._io import resolve_sheet, workbook_session
from excel_mcp.exceptions import SheetNotFoundError


def test_workbook_session_create_and_read_roundtrip(tmp_path: Path) -> None:
    target = tmp_path / "nested" / "sample.xlsx"
    with workbook_session(target, mode="create", initial_sheet="Metrics") as wb:
        ws = resolve_sheet(wb, "Metrics")
        ws["A1"] = "Revenue"
        ws["B1"] = 42

    assert target.exists()
    with workbook_session(target, mode="read", data_only=True) as wb:
        ws = resolve_sheet(wb, "Metrics")
        assert ws["A1"].value == "Revenue"
        assert ws["B1"].value == 42


def test_workbook_session_edit_aborts_save_on_exception(tmp_path: Path) -> None:
    target = tmp_path / "abort.xlsx"
    with workbook_session(target, mode="create", initial_sheet="Data") as wb:
        wb["Data"]["A1"] = "original"

    with pytest.raises(RuntimeError, match="boom"):
        with workbook_session(target, mode="edit") as wb:
            wb["Data"]["A1"] = "corrupted"
            raise RuntimeError("boom")

    with workbook_session(target, mode="read") as wb:
        assert wb["Data"]["A1"].value == "original"


def test_workbook_session_missing_file_raises(tmp_path: Path) -> None:
    missing = tmp_path / "missing.xlsx"
    with pytest.raises(FileNotFoundError):
        with workbook_session(missing, mode="read"):
            pass
    with pytest.raises(FileNotFoundError):
        with workbook_session(missing, mode="edit"):
            pass


def test_resolve_sheet_missing_and_create_if_missing() -> None:
    wb = openpyxl.Workbook()
    try:
        with pytest.raises(SheetNotFoundError):
            resolve_sheet(wb, "Missing")
        created = resolve_sheet(wb, "Created", create_if_missing=True)
        assert created.title == "Created"
        assert resolve_sheet(wb) is wb.active
    finally:
        wb.close()
