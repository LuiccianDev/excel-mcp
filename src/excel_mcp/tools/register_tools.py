"""Tool registration module for MCP Excel Server.

Coordinates registration of all Excel manipulation, batch synthesis,
and inspection tools with the FastMCP server.
"""

from __future__ import annotations

from collections.abc import Sequence
from enum import StrEnum
import logging
import os
from typing import Any

try:
    from mcp.types import ToolAnnotations
except ImportError:

    class ToolAnnotations:  # type: ignore
        def __init__(self, **kwargs: Any) -> None:
            self.kwargs = kwargs


from excel_mcp.tools import batch_tools, compose_tools

logger = logging.getLogger(__name__)


class ToolMode(StrEnum):
    """Supported MCP tool surface modes."""

    COMPOSE = "compose"
    FULL = "full"


def normalize_tool_mode(mode: str | None) -> ToolMode:
    """Normalize and validate an explicit tool mode string."""
    if mode is None:
        return get_tool_mode()
    normalized = mode.strip().lower()
    try:
        return ToolMode(normalized)
    except ValueError as exc:
        raise ValueError(
            f"Invalid tool mode '{mode}': expected 'compose' or 'full'."
        ) from exc


def should_register_legacy_aliases() -> bool:
    """Check if legacy unprefixed and redundant aliases should be registered."""
    val = os.environ.get("EXCEL_MCP_LEGACY_ALIASES", "").strip().lower()
    return val in ("1", "true", "yes")


def get_tool_mode() -> ToolMode:
    """Return the configured MCP tool surface mode ('compose' default or 'full')."""
    raw = os.environ.get("EXCEL_MCP_MODE", "").strip().lower()
    if raw in (ToolMode.FULL, ToolMode.COMPOSE):
        return ToolMode(raw)
    if should_register_legacy_aliases():
        if raw:
            logger.warning(
                "Unrecognized EXCEL_MCP_MODE=%r; using 'full' because EXCEL_MCP_LEGACY_ALIASES is enabled.",
                raw,
            )
        return ToolMode.FULL
    if raw:
        logger.warning("Unrecognized EXCEL_MCP_MODE=%r; defaulting to 'compose'.", raw)
    return ToolMode.COMPOSE


def register_all_tools(mcp: Any, mode: str | None = None) -> None:
    """Register Excel MCP tools with the server in 'compose' or 'full' mode."""
    read_only = ToolAnnotations(read_only_hint=True)
    write = ToolAnnotations(read_only_hint=False)
    destructive = ToolAnnotations(read_only_hint=False, destructive_hint=True)

    resolved_mode = normalize_tool_mode(mode)
    include_aliases = should_register_legacy_aliases()

    def _reg(
        canonical_name: str,
        fn: Any,
        annotations: ToolAnnotations,
        legacy_aliases: Sequence[str] | None = None,
    ) -> None:
        # Register canonical tool
        try:
            mcp.tool(name=canonical_name, annotations=annotations)(fn)
        except TypeError:
            mcp.tool(name=canonical_name)(fn)

        # Register legacy aliases if enabled via EXCEL_MCP_LEGACY_ALIASES=1
        if include_aliases and legacy_aliases and resolved_mode != "compose":
            for alias in legacy_aliases:
                if alias == canonical_name:
                    continue
                try:
                    mcp.tool(name=alias, annotations=annotations)(fn)
                except TypeError:
                    mcp.tool(name=alias)(fn)

    # 1. BATCH & SINGLE-PASS TOOLS (always registered in both modes)
    _reg("excel_build_workbook", batch_tools.build_workbook, write, ["build_workbook"])
    _reg(
        "excel_get_workbook_outline",
        batch_tools.get_workbook_outline,
        read_only,
        ["get_workbook_outline"],
    )
    _reg(
        "excel_batch_write_data",
        batch_tools.batch_write_data,
        write,
        ["batch_write_data"],
    )

    if resolved_mode == "compose":
        _reg("excel_workbook", compose_tools.excel_workbook, write)
        _reg("excel_data", compose_tools.excel_data, write)
        _reg("excel_format", compose_tools.excel_format, write)
        _reg("excel_formula", compose_tools.excel_formula, write)
        _reg("excel_analytics", compose_tools.excel_analytics, write)
        logger.info("Excel MCP tools registered in compact compose mode (8 tools).")
        return

    for group_actions in compose_tools.COMPOSE_REGISTRY.values():
        for spec in group_actions.values():
            annotations = (
                destructive
                if spec.destructive
                else (read_only if spec.read_only else write)
            )
            _reg(
                spec.full_name,
                spec.handler,
                annotations,
                spec.legacy_aliases,
            )

    logger.info(
        f"Excel MCP tools registered in full mode (legacy_aliases_enabled={include_aliases})."
    )
