"""Tool registration module for MCP Excel Server.

Coordinates registration of all Excel manipulation, batch synthesis,
and inspection tools with the FastMCP server.
"""

from __future__ import annotations

import logging
from typing import Any

try:
    from fastmcp import FastMCP
except ImportError:
    try:
        from mcp.server.fastmcp import FastMCP  # type: ignore
    except ImportError:
        FastMCP = Any  # type: ignore

try:
    from mcp.types import ToolAnnotations
except ImportError:
    class ToolAnnotations:  # type: ignore
        def __init__(self, **kwargs: Any) -> None:
            self.kwargs = kwargs

from excel_mcp.tools import (
    batch_tools,
    content_tools,
    excel_tools,
    format_tools,
    formulas_excel_tools,
    graphics_tools,
)

logger = logging.getLogger(__name__)


def register_all_tools(mcp: Any) -> None:
    """Register all tools with the MCP server, providing both standard and prefixed aliases."""
    read_only = ToolAnnotations(readOnlyHint=True)
    write = ToolAnnotations(readOnlyHint=False)
    destructive = ToolAnnotations(readOnlyHint=False, destructiveHint=True)

    def _register(tool_fn: object, annotations: ToolAnnotations) -> None:
        fn_name = getattr(tool_fn, "__name__", "")
        # Register base tool name (e.g. build_workbook, read_data_from_excel)
        try:
            mcp.tool(annotations=annotations)(tool_fn)
        except TypeError:
            mcp.tool()(tool_fn)

        # Register prefixed alias (e.g. excel_build_workbook) if not already prefixed
        if fn_name and not fn_name.startswith("excel_"):
            try:
                mcp.tool(name=f"excel_{fn_name}", annotations=annotations)(tool_fn)
            except TypeError:
                mcp.tool(name=f"excel_{fn_name}")(tool_fn)

    def register_read_only(tool_fn: object) -> None:
        _register(tool_fn, read_only)

    def register_write(tool_fn: object) -> None:
        _register(tool_fn, write)

    def register_destructive(tool_fn: object) -> None:
        _register(tool_fn, destructive)

    # -------------------------------------------------------------------------
    # 1. BATCH & SINGLE-PASS TOOLS (High-efficiency, token-saving tools)
    # -------------------------------------------------------------------------
    register_write(batch_tools.build_workbook)
    register_read_only(batch_tools.get_workbook_outline)
    register_write(batch_tools.batch_write_data)

    # -------------------------------------------------------------------------
    # 2. WORKBOOK & WORKSHEET TOOLS
    # -------------------------------------------------------------------------
    register_write(excel_tools.create_excel_workbook)
    register_write(excel_tools.create_excel_worksheet)
    register_read_only(excel_tools.list_excel_documents)

    # -------------------------------------------------------------------------
    # 3. CONTENT TOOLS
    # -------------------------------------------------------------------------
    register_read_only(content_tools.read_data_from_excel)
    register_write(content_tools.write_data_to_excel)

    # -------------------------------------------------------------------------
    # 4. FORMATTING & SHEET MANIPULATION TOOLS
    # -------------------------------------------------------------------------
    register_write(format_tools.format_range_excel)
    register_write(format_tools.copy_worksheet)
    register_destructive(format_tools.delete_worksheet)
    register_write(format_tools.rename_worksheet)
    register_read_only(format_tools.get_workbook_metadata)
    register_write(format_tools.merge_cells)
    register_write(format_tools.unmerge_cells)
    register_write(format_tools.copy_range)
    register_destructive(format_tools.delete_range)
    register_read_only(format_tools.validate_excel_range)

    # -------------------------------------------------------------------------
    # 5. FORMULA TOOLS
    # -------------------------------------------------------------------------
    register_read_only(formulas_excel_tools.validate_formula_syntax)
    register_write(formulas_excel_tools.apply_formula_excel)

    # -------------------------------------------------------------------------
    # 6. GRAPHICS & ANALYTICS TOOLS
    # -------------------------------------------------------------------------
    register_write(graphics_tools.create_chart)
    register_write(graphics_tools.create_pivot_table)

    logger.info("All Excel MCP tools registered successfully.")
