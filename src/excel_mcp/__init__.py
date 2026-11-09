"""Excel MCP Server package.

High-performance Model Context Protocol (MCP) server for programmatic
Excel (.xlsx) document generation, single-pass synthesis, and inspection.
"""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("luiccian-excel-mcp")
except PackageNotFoundError:
    try:
        __version__ = version("excel_mcp")
    except PackageNotFoundError:
        __version__ = "2.0.0"

__title__ = "excel_mcp"
__description__ = "High-performance Model Context Protocol (MCP) server for Excel spreadsheets"
__author__ = "LuiccianDev"
