"""
File utility functions for MCP Excel Office Server.

This module provides secure file path validation, directory access control,
and file operation utilities for Excel file manipulation. It ensures that
all file operations are performed within authorized directories and with
proper security validation.
"""

import functools
import inspect
import os
from collections.abc import Callable
from pathlib import Path
from typing import Any, TypeVar, cast

from excel_mcp.config import ConfigurationError
from excel_mcp.utils import paths


F = TypeVar("F", bound=Callable[..., Any])


def resolve_safe_path(filename: str | Path) -> Path:
    """Resolve file path to an allowed directory.

    Args:
        filename: File name or path to resolve.

    Returns:
        Path object pointing to an allowed directory.

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


# Check if a file can be written to, including directory permissions
def _check_file_writeable(filename: str) -> tuple[bool, str | None]:
    """
    Check if a file can be written to, including directory permissions.
    This function handles several scenarios:
    - File doesn't exist: checks if parent directory is writeable
    - File exists: checks if file is writeable
    - Handles permission and I/O errors gracefully
    - Uses Path.resolve() to properly handle symlinks

    Args:
        filename: Absolute or relative path to the file

    Returns:
        tuple[bool, str]: (success, error_message)
            - success: True if file can be written, False otherwise
            - error_message: Empty string if successful, otherwise contains error details
    """
    try:
        is_allowed, error = paths.is_path_allowed(filename)
        if not is_allowed:
            return False, error

        resolved_path = paths.resolve_target_path(filename)
        parent_dir = resolved_path.parent

        # If file exists, check write permissions
        if resolved_path.exists():
            if resolved_path.is_dir():
                return False, f"Path is a directory: {resolved_path}"
            if not os.access(resolved_path, os.W_OK):
                return False, f"Permission denied: {resolved_path}"
        # If file doesn't exist, check parent directory
        else:
            if not parent_dir.exists():
                return False, f"Directory does not exist: {parent_dir}"
            if not os.access(parent_dir, os.W_OK):
                return False, f"Directory not writeable: {parent_dir}"

        return True, ""

    except Exception as e:
        return False, f"Error checking file write permissions: {str(e)}"


#! Ensure filename has .xlsx extension
def ensure_xlsx_extension(filename: str) -> str:
    """
    Ensure filename has .xlsx extension.

    Args:
        filename: The filename to check

    Returns:
        Filename with .xlsx extension
    """
    if not filename.lower().endswith(".xlsx"):
        return filename + ".xlsx"
    return filename


# * Retrieve metadata for all Excel (.xlsx) files in the specified directory
def list_excel_files_in_directory() -> list[dict]:
    """
    Retrieve metadata for all Excel (.xlsx) files in the allowed directories.

    Every root returned by :func:`excel_mcp.utils.paths.resolve_allowed_directories`
    is scanned, so ``EXCEL_MCP_DIRS`` governs discovery exactly as it
    governs access.

    Returns:
        list[dict]: List of dictionaries containing file metadata with keys:
            - filename (str): Name of the file
            - size_kb (float): File size in kilobytes (2 decimal places)
            - modified (float): Last modification timestamp
            - path (str): Full absolute path to the file

    Raises:
        FileNotFoundError: If no allowed directory exists
        OSError: For other filesystem-related errors
        ConfigurationError: If directory configuration is invalid
    """
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
            # Use scandir with absolute paths for better error handling
            with os.scandir(directory) as entries:
                for entry in entries:
                    try:
                        if entry.is_file() and entry.name.lower().endswith(".xlsx"):
                            if entry.path in seen:
                                continue
                            seen.add(entry.path)
                            stat = entry.stat()
                            excel_files.append(
                                {
                                    "filename": entry.name,
                                    "size_kb": round(stat.st_size / 1024, 2),
                                    "modified": stat.st_mtime,
                                    "path": entry.path,
                                }
                            )
                    except (OSError, PermissionError):
                        # Skip files we can't access but continue with others
                        continue

        return sorted(excel_files, key=lambda x: x["filename"].lower())

    except OSError as e:
        # Re-raise with more context
        raise OSError(
            f"Error accessing configured directory '{', '.join(existing)}': {str(e)}"
        ) from e


# * Decorator to validate file access
def validate_file_access(
    param: str = "filename", *, read_only: bool = False
) -> Callable[[F], F]:
    """
    Decorator to validate file access before executing the decorated function.
    - Verifies that the file is in an allowed directory.
    - Verifies read or write permissions without mutating files on disk.

    Args:
        param: Name of the parameter that contains the file path.
        read_only: When True, checks read access instead of write access.

    Returns:
        A decorator function that can be applied to sync or async functions.
    """

    def decorator(func: F) -> F:
        sig = inspect.signature(func)

        def _validate_file(
            args: tuple[Any, ...], kwargs: dict[str, Any]
        ) -> tuple[tuple[Any, ...], dict[str, Any], str] | dict[str, str]:
            try:
                bound = sig.bind(*args, **kwargs)
                bound.apply_defaults()

                if param not in bound.arguments:
                    return {
                        "status": "error",
                        "message": f"'{param}' parameter not found in function arguments",
                    }

                raw_value = bound.arguments[param]
                if not isinstance(raw_value, str | Path) or not str(raw_value).strip():
                    return {
                        "status": "error",
                        "message": f"Access denied: '{param}' must be a non-empty file path",
                    }

                resolved_obj, validation_error = paths.validate_and_resolve_path(
                    raw_value, auto_extension=True
                )
                if validation_error or resolved_obj is None:
                    return {
                        "status": "error",
                        "message": f"Access denied: {validation_error}",
                        "path": paths.normalize_xlsx_candidate(raw_value),
                    }
                file_path: str = str(resolved_obj)

                if read_only:
                    if resolved_obj.exists():
                        if resolved_obj.is_dir():
                            return {
                                "status": "error",
                                "message": f"Access denied: Path is a directory: {file_path}",
                                "path": file_path,
                            }
                        if not os.access(resolved_obj, os.R_OK):
                            return {
                                "status": "error",
                                "message": f"Read access denied: Permission denied: {file_path}",
                                "path": file_path,
                            }
                else:
                    # Validate write permissions
                    is_writable, write_error = _check_file_writeable(file_path)
                    if not is_writable:
                        return {
                            "status": "error",
                            "message": f"Write access denied: {write_error}",
                            "path": file_path,
                        }

                bound.arguments[param] = file_path
                return bound.args, bound.kwargs, file_path

            except Exception as e:
                return {
                    "status": "error",
                    "message": f"Error during file validation: {str(e)}",
                }

        if inspect.iscoroutinefunction(func):

            @functools.wraps(func)
            async def async_wrapper(*args: Any, **kwargs: Any) -> Any:
                result = _validate_file(args, kwargs)
                if isinstance(result, dict):
                    return result
                bound_args, bound_kwargs, _resolved_path = result
                return await func(*bound_args, **bound_kwargs)

            return cast(F, async_wrapper)

        else:

            @functools.wraps(func)
            def sync_wrapper(*args: Any, **kwargs: Any) -> Any:
                result = _validate_file(args, kwargs)
                if isinstance(result, dict):
                    return result
                bound_args, bound_kwargs, _resolved_path = result
                return func(*bound_args, **bound_kwargs)

            return cast(F, sync_wrapper)

    return decorator
