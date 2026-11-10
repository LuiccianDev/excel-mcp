"""Tests for the unified allowed-directory resolver."""

from __future__ import annotations

import asyncio
import os

import openpyxl

from excel_mcp.utils.paths import is_path_allowed, resolve_allowed_directories


def test_mcp_allowed_directories_takes_precedence(monkeypatch, tmp_path):
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    monkeypatch.setenv("DIRECTORY", str(tmp_path / "legacy"))
    assert resolve_allowed_directories() == [str(tmp_path.resolve())]


def test_directory_env_is_used_as_fallback(monkeypatch, tmp_path):
    monkeypatch.delenv("MCP_ALLOWED_DIRECTORIES", raising=False)
    monkeypatch.setenv("DIRECTORY", str(tmp_path))
    assert resolve_allowed_directories() == [str(tmp_path.resolve())]


def test_multiple_roots_and_trailing_separator(monkeypatch, tmp_path):
    first, second = tmp_path / "a", tmp_path / "b"
    first.mkdir()
    second.mkdir()
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", f"{first}{os.sep}, {second}")
    assert resolve_allowed_directories() == [
        str(first.resolve()),
        str(second.resolve()),
    ]


def test_traversal_outside_root_is_rejected(monkeypatch, tmp_path):
    root = tmp_path / "root"
    root.mkdir()
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(root))
    allowed, _ = is_path_allowed(str(root / ".." / "escape.xlsx"))
    assert allowed is False


def test_list_documents_finds_files_under_mcp_allowed_directories(
    monkeypatch, tmp_path
):
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    monkeypatch.delenv("DIRECTORY", raising=False)
    workbook = tmp_path / "book.xlsx"
    openpyxl.Workbook().save(workbook)

    from excel_mcp.tools.excel_tools import list_excel_documents

    result = asyncio.run(list_excel_documents())
    assert result["status"] == "success"
    assert result["count"] == 1
    assert result["files"][0]["filename"] == "book.xlsx"


def test_read_tool_accepts_path_under_mcp_allowed_directories(monkeypatch, tmp_path):
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    monkeypatch.delenv("DIRECTORY", raising=False)
    workbook = tmp_path / "book.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.save(workbook)

    from excel_mcp.tools.content_tools import read_data_from_excel

    result = asyncio.run(
        read_data_from_excel(filename=str(workbook), sheet_name="Sheet")
    )
    assert result["status"] == "success"
