"""Unified execution guard decorator for Excel MCP tools.

Provides standardized exception catching, logging, and error response structuring
across all tools.
"""

from __future__ import annotations

import functools
import inspect
from typing import Any, Callable, TypeVar

from excel_mcp.exceptions.exception_tools import ExceptionTool

F = TypeVar("F", bound=Callable[..., Any])


def tool_guard(
    tool_name: str | None = None,
) -> Callable[[F], F]:
    """Decorator to standardize tool execution and error handling.

    Catches unhandled exceptions, logs them with tool context, and returns a consistent
    {"status": "error", "message": ..., "operation": ..., "error_type": ...} dict.
    Supports both synchronous and asynchronous tool functions.
    """

    def decorator(fn: F) -> F:
        name = tool_name or fn.__name__

        if inspect.iscoroutinefunction(fn):

            @functools.wraps(fn)
            async def async_wrapper(*args: Any, **kwargs: Any) -> Any:
                target_file = kwargs.get("filename") or kwargs.get("file_path")
                try:
                    return await fn(*args, **kwargs)
                except Exception as e:
                    return ExceptionTool.handle_error(
                        e, filename=target_file, operation=name
                    )

            return async_wrapper  # type: ignore

        @functools.wraps(fn)
        def sync_wrapper(*args: Any, **kwargs: Any) -> Any:
            target_file = kwargs.get("filename") or kwargs.get("file_path")
            try:
                return fn(*args, **kwargs)
            except Exception as e:
                return ExceptionTool.handle_error(
                    e, filename=target_file, operation=name
                )

        return sync_wrapper  # type: ignore

    return decorator
