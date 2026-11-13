"""Tests for default Documents directory fallback and environment overrides."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_mcp.config import MCPExcelConfig, _default_documents_directory
from excel_mcp.utils.paths import is_path_allowed, resolve_target_path
from excel_mcp.utils.security import resolve_and_check_path


def test_default_directory_uses_user_documents(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """When ~/Documents is writable and env vars are unset, default to ~/Documents."""
    fake_home = tmp_path / "Users" / "Tester"
    (fake_home / "Documents").mkdir(parents=True)

    monkeypatch.delenv("DIRECTORY", raising=False)
    monkeypatch.delenv("MCP_ALLOWED_DIRECTORIES", raising=False)
    monkeypatch.setattr(Path, "home", staticmethod(lambda: fake_home))

    expected = str(fake_home / "Documents")
    assert _default_documents_directory() == expected
    assert MCPExcelConfig().directory == expected


def test_directory_env_var_overrides_default(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Explicit DIRECTORY env var still overrides the ~/Documents default."""
    custom_dir = tmp_path / "custom_sheets"
    custom_dir.mkdir(parents=True)

    monkeypatch.setenv("DIRECTORY", str(custom_dir))
    assert MCPExcelConfig().directory == str(custom_dir)


def test_relative_filename_resolves_inside_allowed_root(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Relative filenames anchor to the primary allowed directory instead of cwd."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    other_cwd = tmp_path / "System32"
    other_cwd.mkdir(parents=True)

    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))
    monkeypatch.setattr(Path, "cwd", staticmethod(lambda: other_cwd))

    resolved, err = resolve_and_check_path("reporte.xlsx")
    assert err is None
    assert resolved == str((docs_dir / "reporte.xlsx").resolve())
    assert (
        resolve_target_path("sub/enero.xlsx")
        == (docs_dir / "sub" / "enero.xlsx").resolve()
    )


def test_relative_path_traversal_is_blocked(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Relative path traversal escaping the allowed root must still be rejected."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))

    allowed, msg = is_path_allowed("../../Windows/System32/evil.xlsx")
    assert not allowed
    assert msg is not None and "not in allowed directories" in msg

    resolved, err = resolve_and_check_path("../../Windows/System32/evil.xlsx")
    assert resolved is None
    assert err is not None and err["status"] == "error"
