"""Tools package for Excel MCP Server.

Provides tools for workbook generation, content editing, formatting, formulas,
graphics, and token-optimized batch construction.
"""

from __future__ import annotations

import importlib
from typing import Any

__all__ = [
    "batch_tools",
    "content_tools",
    "excel_tools",
    "format_tools",
    "formulas_excel_tools",
    "graphics_tools",
    "register_all_tools",
]


def __getattr__(name: str) -> Any:
    if name == "register_all_tools":
        from excel_mcp.tools.register_tools import register_all_tools
        return register_all_tools
    if name in __all__:
        return importlib.import_module(f"excel_mcp.tools.{name}")
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
