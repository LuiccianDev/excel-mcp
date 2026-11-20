"""Tool registration module for MCP Excel Server.

Coordinates registration of all Excel manipulation, batch synthesis,
and inspection tools with the FastMCP server.
"""

from __future__ import annotations

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


from excel_mcp.tools import batch_tools, compound_tools

logger = logging.getLogger(__name__)


class ToolMode(StrEnum):
    """Supported MCP tool surface modes."""

    COMPOUND = "compound"
    FULL = "full"


def normalize_tool_mode(mode: ToolMode | str | None) -> ToolMode:
    """Normalize and validate an explicit tool mode string."""
    if mode is None:
        return get_tool_mode()
    if isinstance(mode, ToolMode):
        return mode
    normalized = mode.strip().lower()
    try:
        return ToolMode(normalized)
    except ValueError as exc:
        raise ValueError(
            f"Invalid tool mode '{mode}': expected 'compound' or 'full'."
        ) from exc


def get_tool_mode() -> ToolMode:
    """Return the configured MCP tool surface mode ('compound' default or 'full')."""
    raw = os.environ.get("EXCEL_MCP_TOOLS", "").strip().lower()
    if raw in (ToolMode.FULL, ToolMode.COMPOUND):
        return ToolMode(raw)
    if raw:
        logger.warning(
            "Unrecognized EXCEL_MCP_TOOLS=%r; defaulting to 'compound'.", raw
        )
    return ToolMode.COMPOUND


def register_all_tools(mcp: Any, mode: str | None = None) -> None:
    """Register Excel MCP tools with the server in 'compound' or 'full' mode."""
    read_only = ToolAnnotations(read_only_hint=True)
    write = ToolAnnotations(read_only_hint=False)
    destructive = ToolAnnotations(read_only_hint=False, destructive_hint=True)

    resolved_mode = normalize_tool_mode(mode)

    def _reg(
        canonical_name: str,
        fn: Any,
        annotations: ToolAnnotations,
    ) -> None:
        try:
            mcp.tool(name=canonical_name, annotations=annotations)(fn)
        except TypeError:
            logger.warning(
                "Failed to attach ToolAnnotations for '%s'; registering without annotations.",
                canonical_name,
            )
            mcp.tool(name=canonical_name)(fn)

    # 1. BATCH & SINGLE-PASS TOOLS (always registered in both modes)
    _reg("excel_build_workbook", batch_tools.build_workbook, destructive)
    _reg("excel_get_workbook_outline", batch_tools.get_workbook_outline, read_only)
    _reg("excel_batch_write_data", batch_tools.batch_write_data, write)

    if resolved_mode == ToolMode.COMPOUND:
        for group, actions in compound_tools.COMPOUND_REGISTRY.items():
            if any(spec.destructive for spec in actions.values()):
                group_annotations = destructive
            elif all(spec.read_only for spec in actions.values()):
                group_annotations = read_only
            else:
                group_annotations = write
            _reg(
                f"excel_{group}",
                getattr(compound_tools, f"excel_{group}"),
                group_annotations,
            )
        logger.info("Excel MCP tools registered in compact compound mode (8 tools).")
        return

    for group_actions in compound_tools.COMPOUND_REGISTRY.values():
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
            )

    logger.info("Excel MCP tools registered in full mode (33 canonical tools).")
