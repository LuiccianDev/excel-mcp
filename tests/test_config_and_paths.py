"""Tests for side-effect-free FileConfig validation and path security helpers."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_mcp.config import FileConfig
from excel_mcp.utils.security import resolve_and_check_path


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


def test_resolve_and_check_path_missing_target_returns_error() -> None:
    resolved, err = resolve_and_check_path(None, None)
    assert resolved is None
    assert err is not None
    assert err["status"] == "error"
