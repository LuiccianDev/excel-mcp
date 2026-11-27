"""Tests for side-effect-free FileConfig validation and path security helpers."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_mcp.config import (
    ConfigurationError,
    ConfigurationManager,
    FileConfig,
    MCPExcelConfig,
    _default_documents_directory,
    get_config,
    get_directory,
    reload_configuration,
)
from excel_mcp.utils.paths import resolve_and_check_path


def test_file_config_validation_does_not_create_directory(tmp_path: Path) -> None:
    non_existent = tmp_path / "should_not_exist_yet"
    cfg = FileConfig(directory=str(non_existent))
    assert cfg.directory == str(non_existent.resolve())
    assert not non_existent.exists()


def test_file_config_rejects_file_as_directory(tmp_path: Path) -> None:
    existing_file = tmp_path / "regular_file.txt"
    existing_file.write_text("hello", encoding="utf-8")
    with pytest.raises(ValueError, match="Path is not a directory"):
        FileConfig(directory=str(existing_file))


def test_file_config_rejects_empty_dir_and_bad_extensions(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="Directory cannot be empty"):
        FileConfig(directory="   ")
    with pytest.raises(ValueError, match="File extension must start with"):
        FileConfig(directory=str(tmp_path), allowed_extensions=["xlsx"])


def test_mcp_config_placeholder_and_user_config_resolution(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir()
    monkeypatch.setattr(Path, "home", staticmethod(lambda: tmp_path))
    monkeypatch.setenv("DIRECTORY", "${workspaceFolder}")

    cfg = MCPExcelConfig()
    assert cfg.directory == str(docs_dir)

    custom_dir = tmp_path / "custom"
    custom_dir.mkdir()
    monkeypatch.setenv("MY_DIR", str(custom_dir))
    resolved_cfg = MCPExcelConfig(directory="${user_config.my_dir}")
    assert resolved_cfg.directory == str(custom_dir)

    fallback_cwd = tmp_path / "cwd_fallback"
    fallback_cwd.mkdir()
    monkeypatch.chdir(fallback_cwd)
    monkeypatch.delenv("DIRECTORY", raising=False)
    cwd_cfg = MCPExcelConfig(directory="${user_config.directory}")
    assert cwd_cfg.directory == str(fallback_cwd)

    default_log_cfg = MCPExcelConfig(log_level="${user_config.missing_level}")
    assert default_log_cfg.log_level == "INFO"


def test_default_documents_directory_fallbacks(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    missing_home = tmp_path / "no_home"
    monkeypatch.setattr(Path, "home", staticmethod(lambda: missing_home))
    monkeypatch.delenv("WINDIR", raising=False)
    assert _default_documents_directory() == "./documents"

    win_dir = tmp_path / "Windows"
    win_dir.mkdir()
    monkeypatch.chdir(win_dir)
    monkeypatch.setenv("WINDIR", str(win_dir))
    assert _default_documents_directory() == str(missing_home / "Documents")


def test_configuration_manager_helpers_and_error_translation(tmp_path: Path) -> None:
    manager = ConfigurationManager()
    original_config = manager._config
    try:
        reload_configuration(directory=str(tmp_path), log_level="DEBUG")
        assert get_config().directory == str(tmp_path)
        assert get_directory() == str(tmp_path)
        assert manager.get_log_level() == "DEBUG"
        summary = manager.get_configuration_summary()
        assert str(tmp_path) in summary
        assert "DEBUG" in summary

        with pytest.raises(ConfigurationError, match="Configuration loading failed"):
            reload_configuration(directory="")
    finally:
        manager._config = original_config


def test_resolve_and_check_path_missing_target_returns_error() -> None:
    resolved, err = resolve_and_check_path(None, None)
    assert resolved is None
    assert err is not None
    assert err["status"] == "error"
