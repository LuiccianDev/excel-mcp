"""File utility functions for MCP Excel Office Server."""

from __future__ import annotations

import os
from pathlib import Path

from excel_mcp.config import ConfigurationError
from excel_mcp.utils import paths


def resolve_safe_path(filename: str | Path) -> Path:
    """Resolve file path to an allowed directory.

    Raises:
        PermissionError: If no allowed directories are available or path escapes sandbox.
    """
    allowed_dirs = paths.resolve_allowed_directories()
    if not allowed_dirs:
        raise PermissionError("No allowed directories available for file operations")

    allowed, error = paths.is_path_allowed(filename)
    if not allowed:
        raise PermissionError(
            error or f"Path '{filename}' is not in allowed directories"
        )
    return paths.resolve_target_path(filename)


def _check_file_writeable(filename: str) -> tuple[bool, str | None]:
    """Check if a file can be written to, including directory permissions."""
    try:
        is_allowed, error = paths.is_path_allowed(filename)
        if not is_allowed:
            return False, error

        resolved_path = paths.resolve_target_path(filename)
        parent_dir = resolved_path.parent

        if resolved_path.exists():
            if resolved_path.is_dir():
                return False, f"Path is a directory: {resolved_path}"
            if not os.access(resolved_path, os.W_OK):
                return False, f"Permission denied: {resolved_path}"
        else:
            if not parent_dir.exists():
                return False, f"Directory does not exist: {parent_dir}"
            if not os.access(parent_dir, os.W_OK):
                return False, f"Directory not writeable: {parent_dir}"

        return True, ""

    except Exception as e:
        return False, f"Error checking file write permissions: {str(e)}"


def ensure_xlsx_extension(filename: str) -> str:
    """Ensure filename has .xlsx extension."""
    if not filename.lower().endswith(".xlsx"):
        return filename + ".xlsx"
    return filename


def list_excel_files_in_directory() -> list[dict]:
    """Retrieve metadata for all Excel (.xlsx) files in allowed directories."""
    try:
        directories = paths.resolve_allowed_directories()
    except Exception as e:
        raise ConfigurationError(f"Failed to get configured directory: {e}") from e

    existing = [directory for directory in directories if os.path.isdir(directory)]
    if not existing:
        raise FileNotFoundError(
            f"Configured directory not found: {', '.join(directories)}"
        )

    excel_files: list[dict] = []
    seen: set[str] = set()

    try:
        for directory in existing:
            with os.scandir(directory) as entries:
                for entry in entries:
                    try:
                        if entry.is_symlink():
                            continue
                        if entry.is_file(
                            follow_symlinks=False
                        ) and entry.name.lower().endswith(".xlsx"):
                            if entry.path in seen:
                                continue
                            seen.add(entry.path)
                            stat = entry.stat(follow_symlinks=False)
                            excel_files.append(
                                {
                                    "filename": entry.name,
                                    "size_kb": round(stat.st_size / 1024, 2),
                                    "modified": stat.st_mtime,
                                    "path": entry.path,
                                }
                            )
                    except (OSError, PermissionError):
                        continue

        return sorted(excel_files, key=lambda x: x["filename"].lower())

    except OSError as e:
        raise OSError(
            f"Error accessing configured directory '{', '.join(existing)}': {str(e)}"
        ) from e
