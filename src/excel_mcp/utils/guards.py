"""Unified execution guard decorator for Excel MCP tools.

Provides standardized exception catching, logging, and error response structuring
across all tools.
"""

from __future__ import annotations

from collections.abc import Callable
import functools
import inspect
from typing import Any, TypeVar

from excel_mcp.exceptions.exception_tools import ExceptionTool
from excel_mcp.utils.metrics import is_metrics_enabled, track_tool_metrics

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
        sig = inspect.signature(fn)

        def _extract_call_context(
            args: tuple[Any, ...], kwargs: dict[str, Any]
        ) -> tuple[Any, Any]:
            target = kwargs.get("filename") or kwargs.get("file_path")
            action = kwargs.get("action")
            if target is not None and (
                action is not None or "action" not in sig.parameters
            ):
                return target, action
            try:
                bound = sig.bind_partial(*args, **kwargs).arguments
                if target is None:
                    target = bound.get("filename") or bound.get("file_path")
                if action is None:
                    action = bound.get("action")
            except TypeError:
                pass
            return target, action

        if inspect.iscoroutinefunction(fn):

            @functools.wraps(fn)
            async def async_wrapper(*args: Any, **kwargs: Any) -> Any:
                target_file, action = (
                    _extract_call_context(args, kwargs)
                    if is_metrics_enabled()
                    else (None, None)
                )
                with track_tool_metrics(name, action) as tracker:
                    try:
                        result = await fn(*args, **kwargs)
                        tracker.observe_result(result)
                        return result
                    except Exception as e:
                        if target_file is None:
                            target_file, _ = _extract_call_context(args, kwargs)
                        tracker.observe_exception(e)
                        return ExceptionTool.handle_error(
                            e, filename=target_file, operation=name
                        )

            return async_wrapper  # type: ignore

        @functools.wraps(fn)
        def sync_wrapper(*args: Any, **kwargs: Any) -> Any:
            target_file, action = (
                _extract_call_context(args, kwargs)
                if is_metrics_enabled()
                else (None, None)
            )
            with track_tool_metrics(name, action) as tracker:
                try:
                    result = fn(*args, **kwargs)
                    tracker.observe_result(result)
                    return result
                except Exception as e:
                    if target_file is None:
                        target_file, _ = _extract_call_context(args, kwargs)
                    tracker.observe_exception(e)
                    return ExceptionTool.handle_error(
                        e, filename=target_file, operation=name
                    )

        return sync_wrapper  # type: ignore

    return decorator
