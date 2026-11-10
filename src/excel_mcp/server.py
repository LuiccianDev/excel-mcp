"""Main entry point for the Excel MCP Server.

Provides the FastMCP server instance for Excel manipulation and single-pass creation.
"""

from __future__ import annotations

import logging
import sys
from typing import Any

try:
    from fastmcp import FastMCP
except ImportError:
    try:
        from mcp.server.fastmcp import FastMCP  # type: ignore
    except ImportError:
        FastMCP = Any  # type: ignore

from excel_mcp import __version__
from excel_mcp.tools.register_tools import register_all_tools
from excel_mcp.utils.security import get_allowed_directories

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger(__name__)


def create_server() -> FastMCP:
    """Create and configure the Excel MCP Server instance."""
    allowed_dirs = get_allowed_directories()
    dirs_list = "\n".join([f"- `{d}`" for d in allowed_dirs])

    instructions = f"""# Excel MCP Server

High-performance Model Context Protocol (MCP) server for programmatic Excel (.xlsx) spreadsheet generation, analysis, and manipulation.

## Core Capabilities

### Single-Pass Document Construction (>90% Token Savings)
- **`excel_build_workbook`**: Generate complete multi-sheet workbooks with design presets (`corporate`, `financial`, `modern`, `compact`), table headers, zebra striping, currency/date formats, formulas, and charts in **one single MCP call**.
- **`excel_get_workbook_outline`**: Extract structural outlines (sheet names, dimensions, column headers, formula presence) without dumping full data grids into context.
- **`excel_batch_write_data`**: Insert 2D matrices of values and formulas into worksheets in a single atomic pass.

### Standard Spreadsheet Operations
- **Workbook management**: Create, copy, rename, delete worksheets, and retrieve metadata.
- **Data operations**: Read ranges, write cells, merge/unmerge cells, copy and delete cell ranges.
- **Formatting**: Fonts, borders, background fills, number formats (accounting, currency, percent).
- **Formulas**: Apply formulas (`SUM`, `AVERAGE`, `VLOOKUP`, etc.) and validate syntax.
- **Graphics & Analytics**: Native bar/line/pie charts and pivot tables.

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

All tools use the `excel_` prefix for clarity (e.g., `excel_build_workbook`, `excel_create_excel_workbook`).
Legacy unprefixed aliases are opt-in: set `EXCEL_MCP_LEGACY_ALIASES=1` to register them.
"""

    mcp = FastMCP(
        name="excel_mcp",
        instructions=instructions,
        version=__version__,
        on_duplicate="error",
    )

    register_all_tools(mcp)
    return mcp


def run_server() -> FastMCP:
    """Backward-compatible helper to instantiate and return configured server."""
    return create_server()


if __name__ == "__main__":
    try:
        server = create_server()
        server.run(transport="stdio")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Server startup failed: {e}")
        sys.exit(1)
