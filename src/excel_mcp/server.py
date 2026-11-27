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

from fastmcp import FastMCP

from excel_mcp import __version__
from excel_mcp.tools.compound_tools import COMPOUND_REGISTRY
from excel_mcp.tools.register_tools import (
    ToolMode,
    normalize_tool_mode,
    register_all_tools,
)
from excel_mcp.utils.metrics import is_metrics_enabled, track_tool_metrics
from excel_mcp.utils.paths import resolve_allowed_directories

logger = logging.getLogger(__name__)


def configure_logging(level: str | None = None) -> None:
    """Configure stderr logging once for CLI and server entrypoints."""
    resolved_level = (
        level or os.environ.get("LOG_LEVEL", "INFO").strip().upper() or "INFO"
    )
    numeric_level = logging.getLevelNamesMapping().get(resolved_level, logging.INFO)
    logging.basicConfig(
        level=numeric_level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        stream=sys.stderr,
    )


def build_server_instructions(mode: str, allowed_dirs: list[Any]) -> str:
    """Build mode-aware MCP server instructions derived from COMPOUND_REGISTRY."""
    resolved_mode = normalize_tool_mode(mode)
    dirs_list = "\n".join([f"- `{d}`" for d in allowed_dirs])
    single_pass_count = 3
    if resolved_mode == ToolMode.COMPOUND:
        total_tools = single_pass_count + len(COMPOUND_REGISTRY)
        lines = [
            f"- **`excel_{group}`**: Action-dispatched ({', '.join(f'`{action}`' for action in actions)})."
            for group, actions in COMPOUND_REGISTRY.items()
        ]
        surface_section = (
            f"### Active Tool Surface (`compound` mode, {total_tools} tools)\n"
            + "\n".join(lines)
        )
    else:
        total_tools = single_pass_count + sum(
            len(actions) for actions in COMPOUND_REGISTRY.values()
        )
        lines = [
            f"- **`{group}`**: {', '.join(f'`{spec.full_name}`' for spec in actions.values())}."
            for group, actions in COMPOUND_REGISTRY.items()
        ]
        surface_section = (
            f"### Active Tool Surface (`full` mode, {total_tools} tools)\n"
            + "\n".join(lines)
        )

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
Set `EXCEL_MCP_TOOLS=full` for all 33 granular tools.
"""


def _attach_metrics_middleware(mcp: FastMCP) -> None:
    """Attach FastMCP tool-call metrics middleware when EXCEL_MCP_METRICS_LOG is set."""
    if not is_metrics_enabled():
        return
    from fastmcp.server.middleware import Middleware

    class _MetricsMiddleware(Middleware):
        async def on_call_tool(self, context: Any, call_next: Any) -> Any:
            msg = getattr(context, "message", None)
            tool_name = getattr(msg, "name", None) or "unknown"
            args = getattr(msg, "arguments", None) or {}
            action = args.get("action") if isinstance(args, dict) else None
            with track_tool_metrics(tool_name, action) as tracker:
                try:
                    result = await call_next(context)
                    tracker.observe_result(result)
                    return result
                except Exception as exc:
                    tracker.observe_exception(exc)
                    raise

    mcp.add_middleware(_MetricsMiddleware())


def create_server(mode: str | None = None) -> FastMCP:
    """Create and configure the Excel MCP Server instance."""
    resolved_mode = normalize_tool_mode(mode)

    allowed_dirs = resolve_allowed_directories()
    instructions = build_server_instructions(resolved_mode, allowed_dirs)

    mcp = FastMCP(
        name="excel_mcp",
        instructions=instructions,
        version=__version__,
        on_duplicate="error",
    )

    _attach_metrics_middleware(mcp)
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
    configure_logging()
    try:
        server = create_server()
        run_stdio(server)
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Server startup failed: {e}")
        sys.exit(1)
