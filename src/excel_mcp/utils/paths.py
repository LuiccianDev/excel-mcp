"""Single source of truth for filesystem sandboxing.

Every tool that reads or writes a workbook resolves its permitted roots here,
so the documented ``MCP_ALLOWED_DIRECTORIES`` variable can never be honoured by
one module and ignored by another.
"""

from __future__ import annotations

import os
import re
from pathlib import Path

from excel_mcp.config import get_directory

DEFAULT_MAX_FILE_SIZE = 100 * 1024 * 1024  # 100 MB
_WINDOWS_RESERVED_NAMES = frozenset(
    {"CON", "PRN", "AUX", "NUL"}
    | {f"COM{i}" for i in range(1, 10)}
    | {f"LPT{i}" for i in range(1, 10)}
)
_DRIVE_PREFIX_RE = re.compile(r"^[A-Za-z]:")


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


def validate_path_safety(path: str | Path) -> str | None:
    """Return an error message if ``path`` contains OS-level path hazards, else None."""
    raw = str(path)
    if not raw.strip():
        return "Path cannot be empty."
    if "\x00" in raw:
        return "Path contains null byte."
    if raw.startswith("\\\\") or raw.startswith("//"):
        return f"UNC network paths are not permitted: '{path}'"

    remainder = _DRIVE_PREFIX_RE.sub("", raw, count=1)
    if ":" in remainder:
        return f"Alternate data streams or invalid ':' in path are not permitted: '{path}'"

    for part in Path(raw).parts:
        stem = part.rstrip(".").split(".", 1)[0].upper()
        if stem in _WINDOWS_RESERVED_NAMES:
            return f"Reserved Windows device name '{part}' is not permitted."
    return None


def validate_file_constraints(
    target: Path, max_file_size: int = DEFAULT_MAX_FILE_SIZE
) -> str | None:
    """Enforce allowed .xlsx extension and maximum file size on existing files."""
    if target.suffix and target.suffix.lower() != ".xlsx":
        return f"Disallowed file extension '{target.suffix}': only '.xlsx' files are permitted."
    try:
        if target.exists() and target.is_file():
            size = target.stat().st_size
            if size > max_file_size:
                return (
                    f"File '{target.name}' ({size} bytes) exceeds maximum allowed size "
                    f"({max_file_size} bytes)."
                )
    except OSError as exc:
        return f"Cannot stat file '{target}': {exc}"
    return None


def resolve_target_path(path: str | Path) -> Path:
    """Resolve an input path, anchoring relative paths to the primary allowed root."""
    candidate = Path(path)
    if candidate.is_absolute():
        return candidate.resolve()
    allowed = resolve_allowed_directories()
    return (Path(allowed[0]) / candidate).resolve()


def is_path_allowed(path: str | Path) -> tuple[bool, str | None]:
    """Return whether ``path`` is hazard-free and inside an allowed root."""
    hazard_error = validate_path_safety(path)
    if hazard_error:
        return False, hazard_error

    target = resolve_target_path(path)
    allowed = resolve_allowed_directories()
    for root in allowed:
        root_path = Path(root)
        if target == root_path or target.is_relative_to(root_path):
            return True, None
    return False, f"Path '{path}' is not in allowed directories: {', '.join(allowed)}"
