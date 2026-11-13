"""Security hardening tests for unified path resolution and sandbox enforcement."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_mcp.utils.file_utils import resolve_safe_path, validate_file_access
from excel_mcp.utils.paths import is_path_allowed, validate_file_constraints
from excel_mcp.utils.security import resolve_and_check_path


def test_validate_file_access_anchors_relative_path_to_allowed_dir(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """validate_file_access must resolve relative filenames against allowed root, not cwd."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    other_cwd = tmp_path / "System32"
    other_cwd.mkdir(parents=True)

    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))
    monkeypatch.chdir(other_cwd)

    @validate_file_access("filename")
    def dummy_tool(filename: str) -> dict[str, str]:
        return {"status": "ok", "filename": filename}

    result = dummy_tool("report.xlsx")
    assert result["status"] == "ok"
    assert Path(result["filename"]).resolve() == (docs_dir / "report.xlsx").resolve()


def test_validate_file_access_blocks_relative_traversal(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """validate_file_access must reject relative path traversal escaping the sandbox."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))

    @validate_file_access("filename")
    def dummy_tool(filename: str) -> dict[str, str]:
        return {"status": "ok", "filename": filename}

    result = dummy_tool("../../Windows/System32/evil.xlsx")
    assert result["status"] == "error"
    assert "Access denied" in result["message"]


def test_resolve_safe_path_blocks_traversal(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """resolve_safe_path must raise PermissionError on relative traversal."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))

    with pytest.raises(PermissionError):
        resolve_safe_path("../outside.xlsx")


def test_resolve_and_check_path_rejects_empty_and_whitespace(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """resolve_and_check_path must reject empty or whitespace-only filenames."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))

    for bad in ("", "   "):
        resolved, err = resolve_and_check_path(bad)
        assert resolved is None
        assert err is not None and err["status"] == "error"


@pytest.mark.parametrize(
    "hazardous_path",
    [
        "report\x00.xlsx",
        r"\\server\share\report.xlsx",
        "//server/share/report.xlsx",
        "report.xlsx:hidden",
        "CON.xlsx",
        "nul.xlsx",
        "PRN",
        "AUX",
        "COM1.xlsx",
        "LPT9.xlsx",
    ],
)
def test_os_path_hazards_are_blocked(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, hazardous_path: str
) -> None:
    """Null bytes, UNC paths, Windows ADS, and reserved device names must be rejected."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))

    allowed, msg = is_path_allowed(hazardous_path)
    assert not allowed
    assert msg is not None

    resolved, err = resolve_and_check_path(hazardous_path)
    assert resolved is None
    assert err is not None and err["status"] == "error"


def test_disallowed_extensions_are_blocked(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Non-.xlsx extensions such as .xlsm or .exe must be rejected."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))

    for bad_ext in ("macro.xlsm", "payload.exe", "script.py"):
        resolved, err = resolve_and_check_path(bad_ext)
        assert resolved is None
        assert err is not None and err["status"] == "error"

    @validate_file_access("filename")
    def dummy_tool(filename: str) -> dict[str, str]:
        return {"status": "ok", "filename": filename}

    result = dummy_tool("macro.xlsm")
    assert result["status"] == "error"


def test_max_file_size_is_enforced(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Existing files larger than max_file_size must be rejected before opening."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    big_file = docs_dir / "huge.xlsx"
    big_file.write_bytes(b"x" * 2048)

    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))

    err_msg = validate_file_constraints(big_file, max_file_size=1024)
    assert err_msg is not None and "exceeds maximum allowed size" in err_msg
