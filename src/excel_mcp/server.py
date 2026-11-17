"""Main entry point for the Excel MCP Server.

Provides the FastMCP server instance for Excel manipulation and single-pass creation.
"""

from __future__ import annotations

import logging
import os
import sys
import warnings
from typing import Any

# Disable FastMCP PyPI update checks and ASCII banner on stdio startup by default
# to prevent cold-start latency and client initialization timeouts.
os.environ.setdefault("FASTMCP_CHECK_FOR_UPDATES", "off")
os.environ.setdefault("FASTMCP_SHOW_SERVER_BANNER", "false")

try:
    from fastmcp import FastMCP
except ImportError:
    try:
        from mcp.server.fastmcp import FastMCP  # type: ignore
    except ImportError:
        FastMCP = Any  # type: ignore

from excel_mcp import __version__
from excel_mcp.tools.register_tools import get_tool_mode, register_all_tools
from excel_mcp.utils.security import get_allowed_directories

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger(__name__)


def build_server_instructions(mode: str, allowed_dirs: list[Any]) -> str:
    """Build mode-aware MCP server instructions matching the active tool surface."""
    dirs_list = "\n".join([f"- `{d}`" for d in allowed_dirs])
    if mode == "compose":
        surface_section = """### Active Tool Surface (`compose` mode, 8 tools)
- **`excel_workbook`**: Action-dispatched workbook & sheet management (`create`, `list`, `metadata`, `create_sheet`, `copy_sheet`, `delete_sheet`, `rename_sheet`).
- **`excel_data`**: Action-dispatched data, range, row/column, search & profiling (`read`, `write`, `append_rows`, `copy_range`, `delete_range`, `merge_cells`, `unmerge_cells`, `insert_rows_columns`, `delete_rows_columns`, `sort_range`, `validate_range`, `find_cells`, `profile_sheet`).
- **`excel_format`**: Action-dispatched styling, tables, validation, protection & print layout (`format_range`, `format_sheet`, `create_table`, `add_data_validation`, `protect_sheet`, `configure_page_setup`).
- **`excel_formula`**: Action-dispatched formula validation & assignment (`validate`, `apply`).
- **`excel_analytics`**: Action-dispatched charts & pivot tables (`create_chart`, `create_pivot_table`)."""
    else:
        surface_section = """### Active Tool Surface (`full` mode, 33 tools)
- **Workbook & Sheet Management**: `excel_create_workbook`, `excel_list_documents`, `excel_get_workbook_metadata`, `excel_create_worksheet`, `excel_copy_worksheet`, `excel_delete_worksheet`, `excel_rename_worksheet`.
- **Data, Structure & Search**: `excel_read_data`, `excel_write_data`, `excel_append_rows`, `excel_copy_range`, `excel_delete_range`, `excel_merge_cells`, `excel_unmerge_cells`, `excel_insert_rows_columns`, `excel_delete_rows_columns`, `excel_sort_range`, `excel_validate_range`, `excel_find_cells`, `excel_profile_sheet`.
- **Formatting, Formulas & Analytics**: `excel_format_range`, `excel_format_sheet`, `excel_create_table`, `excel_add_data_validation`, `excel_protect_sheet`, `excel_configure_page_setup`, `excel_validate_formula_syntax`, `excel_apply_formula`, `excel_create_chart`, `excel_create_pivot_table`."""

    return f"""# Excel MCP Server

High-performance Model Context Protocol (MCP) server for programmatic Excel (.xlsx) spreadsheet generation, analysis, and manipulation.

## Core Capabilities

### Single-Pass Document Construction (>90% Token Savings)
- **`excel_build_workbook`**: Generate complete multi-sheet workbooks with design presets (`corporate`, `financial`, `modern`, `compact`), table headers, zebra striping, currency/date formats, formulas, and charts in **one single MCP call**.
- **`excel_get_workbook_outline`**: Extract structural outlines (sheet names, dimensions, column headers, formula presence) without dumping full data grids into context.
- **`excel_batch_write_data`**: Insert 2D matrices of values and formulas into worksheets in a single atomic pass.

{surface_section}

## Single-Pass Construction Directive (Mandatory for AI Assistants)

- **Do NOT make micro-calls for new workbooks:**
  Avoid calling `create_workbook` -> `set_cell_value` 50 times -> `set_cell_format` 20 times.
  Always prefer `excel_build_workbook` to emit complete data, formulas, styles, and charts in a single operation.
- **When inspecting large workbooks:**
  Always run `excel_get_workbook_outline` first to inspect dimensions and column headers before reading specific ranges.

## Security & Path Restrictions

- **ALLOWED DIRECTORIES:**
  You are STRICTLY RESTRICTED to reading, creating, or modifying files ONLY within the following permitted directories:
{dirs_list}

  If the user asks to create or manipulate a spreadsheet without specifying an absolute path, you MUST default to one of the above allowed directories.

## Tool Naming

All tools use the `excel_` prefix for clarity (e.g., `excel_build_workbook`, `excel_workbook`).
Set `EXCEL_MCP_MODE=full` (or launch `server --tool-mode full`) for all 33 granular tools; legacy unprefixed aliases require `EXCEL_MCP_LEGACY_ALIASES=1`.
"""


def create_server(mode: str | None = None) -> FastMCP:
    """Create and configure the Excel MCP Server instance."""
    resolved_mode = (mode if mode is not None else get_tool_mode()).strip().lower()
    if resolved_mode not in ("compose", "full"):
        raise ValueError(f"Invalid tool mode '{mode}': expected 'compose' or 'full'.")

    allowed_dirs = get_allowed_directories()
    instructions = build_server_instructions(resolved_mode, allowed_dirs)

    mcp = FastMCP(
        name="excel_mcp",
        instructions=instructions,
        version=__version__,
        on_duplicate="error",
    )

    register_all_tools(mcp, mode=resolved_mode)
    return mcp


def configure_warning_filters() -> None:
    """Suppress noisy openpyxl UserWarnings that would otherwise pollute stdio stderr."""
    warnings.filterwarnings("ignore", category=UserWarning, module=r"openpyxl(\.|$)")


configure_warning_filters()


def run_server(mode: str | None = None) -> FastMCP:
    """Backward-compatible helper to instantiate and return configured server."""
    return create_server(mode=mode)


def run_stdio(server: Any) -> None:
    """Run the MCP server over stdio with banner output disabled when supported."""
    try:
        server.run(transport="stdio", show_banner=False)
    except TypeError:
        server.run(transport="stdio")


if __name__ == "__main__":
    try:
        server = create_server()
        run_stdio(server)
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Server startup failed: {e}")
        sys.exit(1)
