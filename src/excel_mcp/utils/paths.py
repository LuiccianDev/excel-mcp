"""Single source of truth for filesystem sandboxing.

Every tool that reads or writes a workbook resolves its permitted roots here,
so the documented ``EXCEL_MCP_DIRS`` variable can never be honoured by
one module and ignored by another.
"""

from __future__ import annotations

import logging
import os
import re
from pathlib import Path

from excel_mcp.config import get_directory

logger = logging.getLogger(__name__)

DEFAULT_MAX_FILE_SIZE = 100 * 1024 * 1024
_WINDOWS_RESERVED_NAMES = frozenset(
    {"CON", "PRN", "AUX", "NUL"}
    | {f"COM{i}" for i in range(1, 10)}
    | {f"LPT{i}" for i in range(1, 10)}
)
_DRIVE_PREFIX_RE = re.compile(r"^[A-Za-z]:")
_UNEXPANDED_PLACEHOLDER_RE = re.compile(r"\$\{[^}]+\}")


def _parse_dir_entries(value: str) -> list[str]:
    items: list[str] = []
    for item in value.split(","):
        stripped = item.strip()
        if not stripped:
            continue
        if _UNEXPANDED_PLACEHOLDER_RE.search(stripped):
            logger.warning("Ignoring unexpanded directory placeholder: %r", stripped)
            continue
        items.append(stripped)
    return items


def resolve_allowed_directories() -> list[str]:
    """Return the resolved, de-duplicated list of allowed directories.

    Precedence is ``EXCEL_MCP_DIRS``, then the legacy ``DIRECTORY``
    variable, then the configured default directory. ``Path.resolve()``
    normalizes separators, resolves symlinks and removes trailing separators.
    """
    candidates = _parse_dir_entries(os.environ.get("EXCEL_MCP_DIRS", "").strip())
    if not candidates:
        candidates = _parse_dir_entries(os.environ.get("DIRECTORY", "").strip())
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
        return (
            f"Alternate data streams or invalid ':' in path are not permitted: '{path}'"
        )

    for part in Path(raw).parts:
        stem = part.split(".", 1)[0].rstrip(". ").upper()
        if stem in _WINDOWS_RESERVED_NAMES:
            return f"Reserved Windows device name '{part}' is not permitted."
    return None


def validate_file_constraints(
    target: Path, max_file_size: int = DEFAULT_MAX_FILE_SIZE
) -> str | None:
    """Enforce allowed .xlsx extension and maximum file size on existing files."""
    if target.suffix.lower() != ".xlsx":
        return f"Disallowed file extension '{target.suffix}': only '.xlsx' files are permitted."
    try:
        if target.exists():
            if target.is_dir():
                return f"Path '{target.name}' is a directory."
            if target.is_file():
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


def normalize_xlsx_candidate(path: str | Path) -> str:
    """Append .xlsx when path has no suffix and is not blank or a hidden dotfile."""
    raw = str(path)
    if not raw.strip():
        return raw
    candidate = Path(raw)
    if not candidate.suffix and not candidate.name.startswith("."):
        return f"{raw}.xlsx"
    return raw


def validate_and_resolve_path(
    path: str | Path,
    *,
    auto_extension: bool = False,
) -> tuple[Path | None, str | None]:
    """Validate sandbox access and file constraints, returning (resolved_path, error_msg)."""
    raw_str = normalize_xlsx_candidate(path) if auto_extension else str(path)
    is_allowed, dir_error = is_path_allowed(raw_str)
    if not is_allowed:
        return None, dir_error
    resolved = resolve_target_path(raw_str)
    constraint_error = validate_file_constraints(resolved)
    if constraint_error:
        return None, constraint_error
    return resolved, None
