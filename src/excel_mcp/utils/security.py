"""Security utilities for file path and directory validation.

Provides functions for validating file paths against permitted directories
configured via EXCEL_MCP_DIRS or DIRECTORY environment variables.
Delegates the actual resolution to :mod:`excel_mcp.utils.paths`.
"""

from __future__ import annotations

from collections.abc import Callable
import functools
import inspect
import os
from pathlib import Path
from typing import Any, TypeVar

from excel_mcp.utils import paths

F = TypeVar("F", bound=Callable[..., Any])


def get_allowed_directories() -> list[str]:
    """Get the list of allowed directories from the environment.

    Returns:
        List of absolute paths to directories where spreadsheets can be created/accessed.
        Defaults to ['./documents'] if EXCEL_MCP_DIRS or DIRECTORY is not set.
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
    *,
    create_dirs: bool = False,
) -> tuple[str | None, dict[str, Any] | None]:
    """Resolve target path and enforce allowed directory and file-type security."""
    target = filename or file_path
    if not target or not target.strip():
        return None, {
            "status": "error",
            "message": "Missing required parameter 'filename' (or 'file_path').",
        }

    resolved_path, error_msg = paths.validate_and_resolve_path(target)
    if error_msg or resolved_path is None:
        return None, {
            "status": "error",
            "message": f"Security restriction: {error_msg}",
        }

    resolved_target = str(resolved_path)
    if create_dirs:
        directory = os.path.dirname(resolved_target)
        if directory and not os.path.exists(directory):
            try:
                os.makedirs(directory, exist_ok=True)
            except OSError as e:
                return None, {
                    "status": "error",
                    "message": f"Cannot create directory '{directory}': {e}",
                }

    return resolved_target, None


def require_valid_path(*, create_dirs: bool = False) -> Callable[[F], F]:
    """Decorator that resolves (filename, file_path) via resolve_and_check_path and injects resolved filename."""

    def decorator(fn: F) -> F:
        sig = inspect.signature(fn)

        def _prepare(args: tuple[Any, ...], kwargs: dict[str, Any]) -> tuple[Any, Any]:
            bound = sig.bind_partial(*args, **kwargs)
            bound.apply_defaults()
            target, err = resolve_and_check_path(
                bound.arguments.get("filename"),
                bound.arguments.get("file_path"),
                create_dirs=create_dirs,
            )
            if err is not None:
                return None, err
            bound.arguments["filename"] = target
            return bound, None

        if inspect.iscoroutinefunction(fn):

            @functools.wraps(fn)
            async def async_wrapper(*args: Any, **kwargs: Any) -> Any:
                bound, err = _prepare(args, kwargs)
                if err is not None:
                    return err
                return await fn(*bound.args, **bound.kwargs)

            return async_wrapper  # type: ignore[return-value]

        @functools.wraps(fn)
        def sync_wrapper(*args: Any, **kwargs: Any) -> Any:
            bound, err = _prepare(args, kwargs)
            if err is not None:
                return err
            return fn(*bound.args, **bound.kwargs)

        return sync_wrapper  # type: ignore[return-value]

    return decorator
