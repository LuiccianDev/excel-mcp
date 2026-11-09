"""Security utilities for file path and directory validation.

Provides functions for validating file paths against permitted directories
configured via MCP_ALLOWED_DIRECTORIES or DIRECTORY environment variables.
"""

from __future__ import annotations

import os
from pathlib import Path


def get_allowed_directories() -> list[str]:
    """Get the list of allowed directories from environment variables.

    Returns:
        List of absolute paths to directories where spreadsheets can be created/accessed.
        Defaults to ['./documents'] if MCP_ALLOWED_DIRECTORIES or DIRECTORY is not set.
    """
    allowed_dirs_str = os.environ.get(
        "MCP_ALLOWED_DIRECTORIES",
        os.environ.get("DIRECTORY", "./documents")
    )
    allowed_dirs = [d.strip() for d in allowed_dirs_str.split(",") if d.strip()]
    return [str(Path(d).resolve()) for d in allowed_dirs]


def is_path_in_allowed_directories(file_path: str | Path) -> tuple[bool, str | None]:
    """Check if the given file path is within allowed directories.

    Args:
        file_path: The file path to validate.

    Returns:
        Tuple of (is_allowed, error_message) where is_allowed is a boolean
        indicating if the path is allowed, and error_message provides details
        if the path is not allowed.
    """
    allowed_dirs = get_allowed_directories()
    resolved_path = Path(file_path).resolve()

    for allowed_dir in allowed_dirs:
        allowed_path = Path(allowed_dir).resolve()
        try:
            if resolved_path == allowed_path or resolved_path.is_relative_to(allowed_path):
                return True, None
        except (ValueError, TypeError):
            continue

    error_msg = (
        f"Path '{file_path}' is not in allowed directories: {', '.join(allowed_dirs)}"
    )
    return False, error_msg
