"""Adversarial path and input validation regression tests."""

from __future__ import annotations

import asyncio
import os
from pathlib import Path

import openpyxl
import pytest

from excel_mcp.core.data_validation import add_data_validation
from excel_mcp.tools.format_tools import format_range_excel
from excel_mcp.utils.paths import is_path_allowed, resolve_allowed_directories
from excel_mcp.utils.security import resolve_and_check_path


def test_multiple_roots_and_trailing_separators_are_normalized(
    tmp_path: Path, monkeypatch
) -> None:
    first = tmp_path / "first"
    second = tmp_path / "second"
    first.mkdir()
    second.mkdir()
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", f"{first}, {second}{os.sep}")
    monkeypatch.delenv("DIRECTORY", raising=False)

    roots = resolve_allowed_directories()

    assert roots == [str(first.resolve()), str(second.resolve())]
    assert is_path_allowed(second / "book.xlsx")[0] is True


def test_relative_traversal_is_rejected_without_creating_directories(
    tmp_path: Path, monkeypatch
) -> None:
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(allowed))
    outside = tmp_path / "outside"

    target, error = resolve_and_check_path("../outside/book.xlsx")

    assert target is None
    assert error and error["status"] == "error"
    assert not outside.exists()


def test_absolute_external_path_is_rejected(tmp_path: Path, monkeypatch) -> None:
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    external = tmp_path / "external.xlsx"
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(allowed))

    target, error = resolve_and_check_path(str(external))

    assert target is None
    assert error and error["status"] == "error"
    assert not external.exists()


def test_symlink_escape_is_rejected(tmp_path: Path, monkeypatch) -> None:
    allowed = tmp_path / "allowed"
    outside = tmp_path / "outside"
    allowed.mkdir()
    outside.mkdir()
    link = allowed / "link"
    try:
        link.symlink_to(outside, target_is_directory=True)
    except (OSError, NotImplementedError) as exc:
        pytest.skip(f"symlink creation unavailable: {exc}")
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(allowed))

    target, error = resolve_and_check_path(str(link / "book.xlsx"))

    assert target is None
    assert error and error["status"] == "error"


def test_reversed_range_is_rejected_before_workbook_mutation(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsx"
    openpyxl.Workbook().save(workbook)

    with pytest.raises(Exception, match="precede"):
        add_data_validation(str(workbook), "B5:A1", "list", formula1='"a,b"')


def test_invalid_color_returns_structured_error(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsx"
    openpyxl.Workbook().save(workbook)

    result = asyncio.run(
        format_range_excel(
            filename=str(workbook),
            sheet_name="Sheet",
            start_cell="A1",
            bg_color="not-a-color",
        )
    )

    assert result["status"] == "error"


def test_invalid_border_enum_returns_structured_error(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsx"
    openpyxl.Workbook().save(workbook)

    result = asyncio.run(
        format_range_excel(
            filename=str(workbook),
            sheet_name="Sheet",
            start_cell="A1",
            border_style="not-a-border",  # type: ignore[arg-type]
        )
    )

    assert result["status"] == "error"
