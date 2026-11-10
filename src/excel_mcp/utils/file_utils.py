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
import shutil
from collections.abc import Callable
from pathlib import Path
from typing import Any, TypeVar, cast

from excel_mcp.config import ConfigurationError
from excel_mcp.utils import paths


F = TypeVar("F", bound=Callable[..., Any])


# Get allowed directories from the shared resolver
def _get_allowed_directories() -> list[str]:
    """Get the list of allowed directories from the shared resolver.

    Delegates to :mod:`excel_mcp.utils.paths` so this module can never honour a
    different sandbox than :mod:`excel_mcp.utils.security`.
    """
    return paths.resolve_allowed_directories()


# Check if the given file path is within allowed directories
def _is_path_in_allowed_directories(file_path: str) -> tuple[bool, str | None]:
    """Check if the given file path is within allowed directories.

    Uses Path.resolve() to properly resolve symlinks before validation,
    preventing symlink-based path traversal attacks.
    """
    return paths.is_path_allowed(file_path)


def resolve_safe_path(filename: str | Path) -> Path:
    """Resolve file path to an allowed directory.
    Args:
        filename: File name or path to resolve.
    Returns:
        Path object pointing to an allowed directory.
    Raises:
        PermissionError: If no allowed directories are available.
    """
    path = Path(filename)
    allowed_dirs = _get_allowed_directories()

    if not allowed_dirs:
        raise PermissionError("No allowed directories available for file operations")

    # If absolute path, validate it's in an allowed directory
    if path.is_absolute():
        for allowed_dir in allowed_dirs:
            try:
                # Check if the path is within an allowed directory
                path.resolve().relative_to(Path(allowed_dir).resolve())
                return path  # Path is allowed
            except ValueError:
                continue  # Path is not within this allowed directory

        # If not in any allowed directory, use filename in first allowed dir
        return Path(allowed_dirs[0]) / path.name

    # If relative path, use first allowed directory
    return Path(allowed_dirs[0]) / filename


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
        # Normalize and get resolved path (resolves symlinks)
        resolved_path = Path(filename).resolve()
        parent_dir = resolved_path.parent

        # Check if path is in allowed directories
        is_allowed, error = _is_path_in_allowed_directories(str(resolved_path))
        if not is_allowed:
            return False, error

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

        # Test actual write operation
        try:
            with open(resolved_path, "a"):
                pass
            return True, ""
        except OSError as e:
            return False, f"Write test failed: {str(e)}"

    except Exception as e:
        return False, f"Error checking file write permissions: {str(e)}"


# * Create a copy of a document
def create_document_copy(
    source_path: str, dest_path: str | None = None
) -> tuple[bool, str, str | None]:
    """
    Create a copy of a document.

    Args:
        source_path: Path to the source document
        dest_path: Optional path for the new document. If not provided, will use source_path + '_copy.docx'

    Returns:
        Tuple of (success, message, new_filepath)
    """
    if not os.path.exists(source_path):
        return False, f"Source document {source_path} does not exist", None

    if not dest_path:
        # Generate a new filename if not provided
        base, ext = os.path.splitext(source_path)
        dest_path = f"{base}_copy{ext}"

    try:
        # Simple file copy
        shutil.copy2(source_path, dest_path)
        return True, f"Document copied to {dest_path}", dest_path
    except Exception as e:
        return False, f"Failed to copy document: {str(e)}", None


#! Ensure filename has .xlsx extension
def ensure_xlsx_extension(filename: str) -> str:
    """
    Ensure filename has .xlsx extension.

    Args:
        filename: The filename to check

    Returns:
        Filename with .xlsx extension
    """
    if not filename.endswith(".xlsx"):
        return filename + ".xlsx"
    return filename


# * Retrieve metadata for all Excel (.xlsx) files in the specified directory
def list_excel_files_in_directory() -> list[dict]:
    """
    Retrieve metadata for all Excel (.xlsx) files in the allowed directories.

    Every root returned by :func:`excel_mcp.utils.paths.resolve_allowed_directories`
    is scanned, so ``MCP_ALLOWED_DIRECTORIES`` governs discovery exactly as it
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
def validate_file_access(param: str = "filename") -> Callable[[F], F]:
    """
    Decorator to validate file access before executing the decorated function.
    - Verifies that the file is in an allowed directory.
    - Verifies that it has write permissions.

    Args:
        param: Name of the parameter that contains the file path.

    Returns:
        A decorator function that can be applied to sync or async functions.
    """

    def decorator(func: F) -> F:
        def _validate_file(
            args: tuple[Any, ...], kwargs: dict[str, Any]
        ) -> str | dict[str, str]:
            try:
                sig = inspect.signature(func)
                bound = sig.bind(*args, **kwargs)
                bound.apply_defaults()

                if param not in bound.arguments:
                    return {
                        "status": "error",
                        "message": f"'{param}' parameter not found in function arguments",
                    }

                # Use resolve() instead of abspath() to resolve symlinks and normalize path
                file_path: str = str(Path(bound.arguments[param]).resolve())

                # Validate allowed directory
                is_allowed, dir_error = _is_path_in_allowed_directories(file_path)
                if not is_allowed:
                    return {
                        "status": "error",
                        "message": f"Access denied: {dir_error}",
                        "path": file_path,
                    }

                # Validate write permissions
                is_writable, write_error = _check_file_writeable(file_path)
                if not is_writable:
                    return {
                        "status": "error",
                        "message": f"Write access denied: {write_error}",
                        "path": file_path,
                    }

                return file_path

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
                try:
                    return await func(*args, **kwargs)
                except Exception as e:
                    return {
                        "status": "error",
                        "message": f"Error in async function: {str(e)}",
                        "path": result,
                    }

            return cast(F, async_wrapper)

        else:

            @functools.wraps(func)
            def sync_wrapper(*args: Any, **kwargs: Any) -> Any:
                result = _validate_file(args, kwargs)
                if isinstance(result, dict):
                    return result
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    return {
                        "status": "error",
                        "message": f"Error in function: {str(e)}",
                        "path": result,
                    }

            return cast(F, sync_wrapper)

    return decorator
