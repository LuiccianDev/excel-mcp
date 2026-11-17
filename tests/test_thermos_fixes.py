"""Regression tests for security, contract, and exception fixes identified by Thermos."""

from __future__ import annotations

import asyncio
from pathlib import Path

import openpyxl
import pytest

from excel_mcp.exceptions import exception_core, exception_tools
from excel_mcp.tools.excel_tools import create_excel_workbook
from excel_mcp.tools.format_tools import validate_excel_range
from excel_mcp.tools.formulas_excel_tools import validate_formula_syntax
from excel_mcp.utils.file_utils import validate_file_access
from excel_mcp.utils.paths import validate_file_constraints
from excel_mcp.utils.validation_utils import validate_formula


def test_exception_hierarchies_are_unified() -> None:
    """Tool exception classes must catch corresponding core exceptions."""
    assert issubclass(exception_core.ValidationError, exception_tools.ValidationError)
    assert issubclass(exception_core.WorkbookError, exception_tools.WorkbookError)
    assert issubclass(exception_core.SheetError, exception_tools.SheetError)
    assert issubclass(exception_core.DataError, exception_tools.DataError)
    assert issubclass(exception_core.PivotError, exception_tools.PivotError)
    assert issubclass(exception_core.WorkbookNotFoundError, exception_tools.WorkbookError)


def test_validate_file_access_does_not_create_ghost_file(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """validate_file_access must not create a 0-byte file on disk when checking access."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))

    @validate_file_access("filename")
    def write_tool(filename: str) -> dict[str, str]:
        return {"status": "ok", "filename": filename}

    @validate_file_access("filename", read_only=True)
    def read_tool(filename: str) -> dict[str, str]:
        return {"status": "ok", "filename": filename}

    write_res = write_tool("ghost_write.xlsx")
    read_res = read_tool("ghost_read.xlsx")

    assert write_res["status"] == "ok"
    assert read_res["status"] == "ok"
    assert not (docs_dir / "ghost_write.xlsx").exists()
    assert not (docs_dir / "ghost_read.xlsx").exists()


def test_validate_formula_syntax_blocks_traversal(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """validate_formula_syntax must enforce sandboxed path validation."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))

    result = asyncio.run(
        validate_formula_syntax(
            "../../outside.xlsx", "Sheet1", "A1", "=SUM(B1:B5)"
        )
    )
    assert result["status"] == "error"
    assert "Access denied" in result["message"]


def test_validate_file_constraints_rejects_dotfiles_and_extensionless(
    tmp_path: Path,
) -> None:
    """Dotfiles like .env and extensionless filenames must be rejected."""
    assert validate_file_constraints(tmp_path / ".env") is not None
    assert validate_file_constraints(tmp_path / "report") is not None
    assert validate_file_constraints(tmp_path / "report.xlsx") is None


def test_create_excel_workbook_enforces_xlsx_extension(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """create_excel_workbook must reject non-.xlsx extensions and dotfiles."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))

    for bad_name in ("payload.exe", "macro.xlsm", ".env"):
        res = asyncio.run(create_excel_workbook(bad_name))
        assert "error" in res
        assert not (docs_dir / bad_name).exists()

    ok_res = asyncio.run(create_excel_workbook("daily_report"))
    assert ok_res.get("status") == "success"
    assert (docs_dir / "daily_report.xlsx").exists()


def test_validate_formula_blocks_lowercase_unsafe_functions() -> None:
    """Unsafe Excel functions must be blocked regardless of case."""
    for formula in (
        '=hyperlink("https://example.com")',
        '=webservice("https://example.com")',
        '=indirect("A1")',
    ):
        is_valid, msg = validate_formula(formula)
        assert not is_valid
        assert "Unsafe function" in msg


def test_validate_excel_range_accepts_end_cell(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """validate_excel_range must pass start_cell and end_cell separately."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))

    wb_path = docs_dir / "range_test.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Data"
    ws["A1"] = "Header"
    ws["B2"] = 42
    wb.save(wb_path)
    wb.close()

    res = asyncio.run(validate_excel_range("range_test.xlsx", "Data", "A1", "B2"))
    assert res["status"] == "success"
    assert res["valid"] is True
    assert res["range"] == "A1:B2"
