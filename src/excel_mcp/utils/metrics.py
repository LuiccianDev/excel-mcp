"""Optional development metrics logger for MCP tool executions."""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass
from datetime import UTC, datetime
import json
import os
from pathlib import Path
import time
from typing import Any, Final

DEFAULT_MAX_METRICS_LOG_BYTES: Final[int] = 10 * 1024 * 1024

_current_phases: ContextVar[dict[str, float] | None] = ContextVar(
    "_current_phases", default=None
)


def _get_metrics_path() -> str | None:
    """Return stripped EXCEL_MCP_METRICS_LOG path or None when disabled."""
    raw = os.environ.get("EXCEL_MCP_METRICS_LOG", "").strip()
    return raw or None


def is_metrics_enabled() -> bool:
    """Return True when EXCEL_MCP_METRICS_LOG is configured."""
    return _get_metrics_path() is not None


def add_metric_phase(name: str, duration_ms: float) -> None:
    """Accumulate duration_ms for a named internal phase in the active scope."""
    phases = _current_phases.get()
    if phases is None:
        return
    phases[name] = phases.get(name, 0.0) + duration_ms


@contextmanager
def measure_phase(name: str) -> Iterator[None]:
    """Measure elapsed time for a named internal phase if a metrics scope is active."""
    if _current_phases.get() is None:
        yield
        return
    start = time.perf_counter()
    try:
        yield
    finally:
        add_metric_phase(name, (time.perf_counter() - start) * 1000)


@dataclass
class MetricTracker:
    """Mutable outcome holder for a single tracked tool invocation."""

    status: str = "success"
    error_type: str | None = None

    def observe_result(self, result: Any) -> None:
        """Extract status and error_type from a dict or FastMCP CallToolResult."""
        is_err = bool(
            getattr(result, "isError", False) or getattr(result, "is_error", False)
        )
        payload = (
            result
            if isinstance(result, dict)
            else getattr(result, "structured_content", None)
        )
        if isinstance(payload, dict):
            self.status = payload.get("status", "error" if is_err else "success")
            self.error_type = payload.get("error_type") or (
                "ToolError" if is_err else None
            )
            return
        if is_err:
            self.status = "error"
            self.error_type = "ToolError"

    def observe_exception(self, exc: Exception) -> None:
        """Record an unhandled exception as an error outcome."""
        self.status = "error"
        self.error_type = type(exc).__name__


@contextmanager
def track_tool_metrics(tool_name: str, action: Any = None) -> Iterator[MetricTracker]:
    """Time a tool call and emit a JSONL record on exit if no outer scope is active."""
    tracker = MetricTracker()
    if not is_metrics_enabled() or _current_phases.get() is not None:
        yield tracker
        return

    token = _current_phases.set({})
    start = time.perf_counter()
    try:
        yield tracker
    finally:
        elapsed_ms = (time.perf_counter() - start) * 1000
        phases = _current_phases.get() or {}
        _current_phases.reset(token)
        record_tool_metric(
            tool_name,
            action,
            elapsed_ms,
            tracker.status,
            error_type=tracker.error_type,
            phases=phases,
        )


def record_tool_metric(
    tool_name: str,
    action: Any,
    duration_ms: float,
    status: str,
    error_type: str | None = None,
    phases: dict[str, float] | None = None,
) -> None:
    """Append a single JSONL metric record when EXCEL_MCP_METRICS_LOG is set."""
    metrics_path = _get_metrics_path()
    if not metrics_path:
        return
    record: dict[str, Any] = {
        "timestamp": datetime.now(UTC).isoformat(),
        "tool": tool_name,
        "duration_ms": round(duration_ms, 2),
        "status": status,
    }
    if action is not None:
        record["action"] = str(action)
    if error_type is not None:
        record["error_type"] = error_type
    if phases:
        record["phases"] = {k: round(v, 2) for k, v in phases.items()}
    try:
        path = Path(metrics_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists() and path.stat().st_size >= DEFAULT_MAX_METRICS_LOG_BYTES:
            path.replace(path.with_suffix(path.suffix + ".1"))
        with path.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(record) + "\n")
    except OSError:
        pass
