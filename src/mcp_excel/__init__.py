"""Backward compatibility shim for mcp_excel.

Redirects to excel_mcp to maintain compatibility with existing installations.
"""

from excel_mcp import (
    __author__,
    __description__,
    __title__,
    __version__,
)

__all__ = [
    "__version__",
    "__author__",
    "__description__",
    "__title__",
]
