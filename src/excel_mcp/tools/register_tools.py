"""Tool registration module for MCP Excel Server.

Coordinates registration of all Excel manipulation, batch synthesis,
and inspection tools with the FastMCP server.
"""

from __future__ import annotations

import logging
import os
from typing import Any

try:
    from mcp.types import ToolAnnotations
except ImportError:

    class ToolAnnotations:  # type: ignore
        def __init__(self, **kwargs: Any) -> None:
            self.kwargs = kwargs


from excel_mcp.tools import (
    analysis_tools,
    batch_tools,
    content_tools,
    excel_tools,
    format_tools,
    formulas_excel_tools,
    graphics_tools,
    structure_tools,
)

logger = logging.getLogger(__name__)


def should_register_legacy_aliases() -> bool:
    """Check if legacy unprefixed and redundant aliases should be registered."""
    val = os.environ.get("EXCEL_MCP_LEGACY_ALIASES", "").strip().lower()
    return val in ("1", "true", "yes")


def register_all_tools(mcp: Any) -> None:
    """Register canonical Excel MCP tools with the server, with optional legacy aliases."""
    read_only = ToolAnnotations(read_only_hint=True)
    write = ToolAnnotations(read_only_hint=False)
    destructive = ToolAnnotations(read_only_hint=False, destructive_hint=True)

    include_aliases = should_register_legacy_aliases()

    def _reg(
        canonical_name: str,
        fn: Any,
        annotations: ToolAnnotations,
        legacy_aliases: list[str] | None = None,
    ) -> None:
        # Register canonical tool
        try:
            mcp.tool(name=canonical_name, annotations=annotations)(fn)
        except TypeError:
            mcp.tool(name=canonical_name)(fn)

        # Register legacy aliases if enabled via EXCEL_MCP_LEGACY_ALIASES=1
        if include_aliases and legacy_aliases:
            for alias in legacy_aliases:
                if alias == canonical_name:
                    continue
                try:
                    mcp.tool(name=alias, annotations=annotations)(fn)
                except TypeError:
                    mcp.tool(name=alias)(fn)

    # 1. BATCH & SINGLE-PASS TOOLS
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
    _reg("excel_append_rows", structure_tools.append_rows, write, ["append_rows"])

    # 2. WORKBOOK & WORKSHEET TOOLS
    _reg(
        "excel_create_workbook",
        excel_tools.create_excel_workbook,
        write,
        ["create_excel_workbook", "excel_create_excel_workbook"],
    )
    _reg(
        "excel_create_worksheet",
        excel_tools.create_excel_worksheet,
        write,
        ["create_excel_worksheet", "excel_create_excel_worksheet"],
    )
    _reg(
        "excel_list_documents",
        excel_tools.list_excel_documents,
        read_only,
        ["list_excel_documents", "excel_list_excel_documents"],
    )

    # 3. CONTENT TOOLS
    _reg(
        "excel_read_data",
        content_tools.read_data_from_excel,
        read_only,
        ["read_data_from_excel", "excel_read_data_from_excel"],
    )
    _reg(
        "excel_write_data",
        content_tools.write_data_to_excel,
        write,
        ["write_data_to_excel", "excel_write_data_to_excel"],
    )

    # 4. FORMATTING & SHEET MANIPULATION TOOLS
    _reg(
        "excel_format_range",
        format_tools.format_range_excel,
        write,
        ["format_range_excel", "excel_format_range_excel"],
    )
    _reg("excel_format_sheet", structure_tools.format_sheet, write, ["format_sheet"])
    _reg("excel_create_table", structure_tools.create_table, write, ["create_table"])
    _reg("excel_sort_range", structure_tools.sort_range, write, ["sort_range"])
    _reg(
        "excel_insert_rows_columns",
        structure_tools.insert_rows_columns,
        write,
        ["insert_rows_columns"],
    )
    _reg(
        "excel_delete_rows_columns",
        structure_tools.delete_rows_columns,
        destructive,
        ["delete_rows_columns"],
    )
    _reg("excel_copy_worksheet", format_tools.copy_worksheet, write, ["copy_worksheet"])
    _reg(
        "excel_delete_worksheet",
        format_tools.delete_worksheet,
        destructive,
        ["delete_worksheet"],
    )
    _reg(
        "excel_rename_worksheet",
        format_tools.rename_worksheet,
        write,
        ["rename_worksheet"],
    )
    _reg(
        "excel_get_workbook_metadata",
        format_tools.get_workbook_metadata,
        read_only,
        ["get_workbook_metadata"],
    )
    _reg("excel_merge_cells", format_tools.merge_cells, write, ["merge_cells"])
    _reg("excel_unmerge_cells", format_tools.unmerge_cells, write, ["unmerge_cells"])
    _reg("excel_copy_range", format_tools.copy_range, write, ["copy_range"])
    _reg("excel_delete_range", format_tools.delete_range, destructive, ["delete_range"])
    _reg(
        "excel_validate_range",
        format_tools.validate_excel_range,
        read_only,
        ["validate_excel_range", "excel_validate_excel_range"],
    )

    # 5. FORMULA TOOLS
    _reg(
        "excel_validate_formula_syntax",
        formulas_excel_tools.validate_formula_syntax,
        read_only,
        ["validate_formula_syntax"],
    )
    _reg(
        "excel_apply_formula",
        formulas_excel_tools.apply_formula_excel,
        write,
        ["apply_formula_excel", "excel_apply_formula_excel"],
    )

    # 6. GRAPHICS & ANALYTICS TOOLS
    _reg("excel_create_chart", graphics_tools.create_chart, write, ["create_chart"])
    _reg(
        "excel_create_pivot_table",
        graphics_tools.create_pivot_table,
        write,
        ["create_pivot_table"],
    )

    # 7. SEARCH & PROFILING TOOLS
    _reg("excel_find_cells", analysis_tools.find_cells, read_only, ["find_cells"])
    _reg(
        "excel_profile_sheet",
        analysis_tools.profile_sheet,
        read_only,
        ["profile_sheet"],
    )

    logger.info(
        f"Excel MCP tools registered successfully (legacy_aliases_enabled={include_aliases})."
    )
