"""Security utilities for file path and directory validation.

Provides functions for validating file paths against permitted directories
configured via MCP_ALLOWED_DIRECTORIES or DIRECTORY environment variables.
Delegates the actual resolution to :mod:`excel_mcp.utils.paths`.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from excel_mcp.utils import paths


def get_allowed_directories() -> list[str]:
    """Get the list of allowed directories from the environment.

    Returns:
        List of absolute paths to directories where spreadsheets can be created/accessed.
        Defaults to ['./documents'] if MCP_ALLOWED_DIRECTORIES or DIRECTORY is not set.
    """
    return paths.resolve_allowed_directories()


def is_path_in_allowed_directories(file_path: str | Path) -> tuple[bool, str | None]:
    """Check if the given file path is within allowed directories.

    Args:
        file_path: The file path to validate.

    Returns:
        Tuple of (is_allowed, error_message) where is_allowed is a boolean
        indicating if the path is allowed, and error_message provides details
        if the path is not allowed.
    """
    return paths.is_path_allowed(file_path)


def resolve_and_check_path(
    filename: str | None,
    file_path: str | None = None,
) -> tuple[str | None, dict[str, Any] | None]:
    """Resolve target path, enforce allowed directory security, and ensure parent directory exists."""
    target = filename or file_path
    if not target:
        return None, {
            "status": "error",
            "message": "Missing required parameter 'filename' (or 'file_path').",
        }

    is_allowed, error_msg = is_path_in_allowed_directories(target)
    if not is_allowed:
        return None, {
            "status": "error",
            "message": f"Security restriction: {error_msg}",
        }

    directory = os.path.dirname(os.path.abspath(target))
    if directory and not os.path.exists(directory):
        try:
            os.makedirs(directory, exist_ok=True)
        except OSError as e:
            return None, {
                "status": "error",
                "message": f"Cannot create directory '{directory}': {e}",
            }

    return target, None
