"""Single source of truth for filesystem sandboxing.

Every tool that reads or writes a workbook resolves its permitted roots here,
so the documented ``MCP_ALLOWED_DIRECTORIES`` variable can never be honoured by
one module and ignored by another.
"""

from __future__ import annotations

import os
from pathlib import Path

from excel_mcp.config import get_directory


def _split(value: str) -> list[str]:
    """Split a comma-separated list of directories, dropping empty entries."""
    return [item.strip() for item in value.split(",") if item.strip()]


def resolve_allowed_directories() -> list[str]:
    """Return the resolved, de-duplicated list of allowed directories.

    Precedence is ``MCP_ALLOWED_DIRECTORIES``, then the legacy ``DIRECTORY``
    variable, then the configured default directory. ``Path.resolve()``
    normalizes separators, resolves symlinks and removes trailing separators.
    """
    candidates = _split(os.environ.get("MCP_ALLOWED_DIRECTORIES", "").strip())
    if not candidates:
        candidates = _split(os.environ.get("DIRECTORY", "").strip())
    if not candidates:
        candidates = [get_directory()]

    resolved: list[str] = []
    for candidate in candidates:
        path = str(Path(candidate).resolve())
        if path not in resolved:
            resolved.append(path)
    return resolved


def is_path_allowed(path: str | Path) -> tuple[bool, str | None]:
    """Return whether ``path`` lives inside an allowed root, plus an error message."""
    target = Path(path).resolve()
    allowed = resolve_allowed_directories()
    for root in allowed:
        root_path = Path(root)
        if target == root_path or target.is_relative_to(root_path):
            return True, None
    return False, f"Path '{path}' is not in allowed directories: {', '.join(allowed)}"
