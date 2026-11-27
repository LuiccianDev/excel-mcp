"""Unified execution and path-security guard decorator for Excel MCP tools."""

from __future__ import annotations

from collections.abc import Callable
import functools
import inspect
import os
from pathlib import Path
from typing import Any, TypeVar, cast

from excel_mcp.exceptions.exception_tools import ExceptionTool
from excel_mcp.utils import paths
from excel_mcp.utils.metrics import is_metrics_enabled, track_tool_metrics

F = TypeVar("F", bound=Callable[..., Any])


def _extract_raw_path(arguments: dict[str, Any], path_param: str) -> Any:
    primary = arguments.get(path_param)
    if path_param == "filename" and "file_path" in arguments:
        alias = arguments.get("file_path")
        if primary is not None and alias is not None and str(primary) != str(alias):
            raise PermissionError(
                "Access denied: Conflicting 'filename' and 'file_path' values provided."
            )
        return primary if primary is not None else alias
    return primary


def _check_resolved_access(
    resolved_obj: Path, *, read_only: bool, create_dirs: bool
) -> None:
    if read_only:
        if resolved_obj.exists():
            if resolved_obj.is_dir():
                raise PermissionError(
                    f"Access denied: Path is a directory: {resolved_obj}"
                )
            if not os.access(resolved_obj, os.R_OK):
                raise PermissionError(
                    f"Read access denied: Permission denied: {resolved_obj}"
                )
        return

    if resolved_obj.exists():
        if resolved_obj.is_dir():
            raise PermissionError(f"Access denied: Path is a directory: {resolved_obj}")
        if not os.access(resolved_obj, os.W_OK):
            raise PermissionError(
                f"Write access denied: Permission denied: {resolved_obj}"
            )
        return

    ancestor = resolved_obj.parent
    while not ancestor.exists() and ancestor != ancestor.parent:
        ancestor = ancestor.parent
    if (
        not ancestor.exists()
        or not ancestor.is_dir()
        or not os.access(ancestor, os.W_OK)
    ):
        raise PermissionError(
            f"Write access denied: Directory not writeable: {ancestor}"
        )
    if create_dirs and not resolved_obj.parent.exists():
        resolved_obj.parent.mkdir(parents=True, exist_ok=True)


def _bind_and_validate_path(
    sig: inspect.Signature,
    args: tuple[Any, ...],
    kwargs: dict[str, Any],
    *,
    path_param: str,
    read_only: bool,
    auto_extension: bool,
    create_dirs: bool,
) -> tuple[tuple[Any, ...], dict[str, Any], str]:
    bound = sig.bind(*args, **kwargs)
    bound.apply_defaults()
    if path_param not in bound.arguments:
        raise PermissionError(
            f"Access denied: '{path_param}' parameter not found in function arguments."
        )

    raw_value = _extract_raw_path(bound.arguments, path_param)
    if not isinstance(raw_value, str | Path) or not str(raw_value).strip():
        raise PermissionError(
            f"Access denied: '{path_param}' must be a non-empty file path."
        )

    resolved_obj, validation_error = paths.validate_and_resolve_path(
        raw_value, auto_extension=auto_extension
    )
    if validation_error or resolved_obj is None:
        raise PermissionError(f"Access denied: {validation_error}")

    _check_resolved_access(resolved_obj, read_only=read_only, create_dirs=create_dirs)
    resolved_str = str(resolved_obj)
    bound.arguments[path_param] = resolved_str
    return bound.args, bound.kwargs, resolved_str


def tool_guard(
    tool_name: str | None = None,
    *,
    path_param: str | None = None,
    read_only: bool = False,
    auto_extension: bool = True,
    create_dirs: bool = False,
) -> Callable[[F], F]:
    """Standardize tool path enforcement, execution metrics, and error payloads."""

    def decorator(fn: F) -> F:
        name = tool_name or fn.__name__
        sig = inspect.signature(fn)

        def _extract_call_context(
            args: tuple[Any, ...], kwargs: dict[str, Any]
        ) -> tuple[Any, Any]:
            target_key = path_param or "filename"
            target = (
                kwargs.get(target_key)
                or kwargs.get("filename")
                or kwargs.get("file_path")
            )
            action = kwargs.get("action")
            if target is not None and (
                action is not None or "action" not in sig.parameters
            ):
                return target, action
            try:
                bound = sig.bind_partial(*args, **kwargs).arguments
                if target is None:
                    target = (
                        bound.get(target_key)
                        or bound.get("filename")
                        or bound.get("file_path")
                    )
                if action is None:
                    action = bound.get("action")
            except TypeError:
                pass
            return target, action

        def _prepare(
            args: tuple[Any, ...], kwargs: dict[str, Any]
        ) -> tuple[tuple[Any, ...], dict[str, Any], Any]:
            if path_param is None:
                target = (
                    _extract_call_context(args, kwargs)[0]
                    if is_metrics_enabled()
                    else None
                )
                return args, kwargs, target
            return _bind_and_validate_path(
                sig,
                args,
                kwargs,
                path_param=path_param,
                read_only=read_only,
                auto_extension=auto_extension,
                create_dirs=create_dirs,
            )

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
                        call_args, call_kwargs, resolved_target = _prepare(args, kwargs)
                        if resolved_target is not None:
                            target_file = resolved_target
                        result = await fn(*call_args, **call_kwargs)
                        tracker.observe_result(result)
                        return result
                    except Exception as e:
                        if target_file is None:
                            target_file, _ = _extract_call_context(args, kwargs)
                        tracker.observe_exception(e)
                        return ExceptionTool.handle_error(
                            e, filename=target_file, operation=name
                        )

            return cast(F, async_wrapper)

        @functools.wraps(fn)
        def sync_wrapper(*args: Any, **kwargs: Any) -> Any:
            target_file, action = (
                _extract_call_context(args, kwargs)
                if is_metrics_enabled()
                else (None, None)
            )
            with track_tool_metrics(name, action) as tracker:
                try:
                    call_args, call_kwargs, resolved_target = _prepare(args, kwargs)
                    if resolved_target is not None:
                        target_file = resolved_target
                    result = fn(*call_args, **call_kwargs)
                    tracker.observe_result(result)
                    return result
                except Exception as e:
                    if target_file is None:
                        target_file, _ = _extract_call_context(args, kwargs)
                    tracker.observe_exception(e)
                    return ExceptionTool.handle_error(
                        e, filename=target_file, operation=name
                    )

        return cast(F, sync_wrapper)

    return decorator
