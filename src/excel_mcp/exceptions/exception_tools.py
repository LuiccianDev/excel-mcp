"""Tools exceptions for the MCP Excel Office tools module.

These exceptions are used by the tools layer for handling
errors related to Excel operations at the tool/interface level.
"""

from __future__ import annotations

import logging
from typing import Any

from excel_mcp.exceptions.exception_core import (
    CoreError,
    DataError,
    PivotError,
    SheetError,
    ValidationError,
    WorkbookError,
)

logger = logging.getLogger(__name__)

# Unify tool and core exception hierarchies so legacy tool handlers catch core errors.
ExcelMCPError = CoreError


class FormattingError(ExcelMCPError):
    """Raised when formatting operations fail."""

    pass


class CalculationError(ExcelMCPError):
    """Raised when formula calculations fail."""

    pass


class ChartError(ExcelMCPError):
    """Raised when chart operations fail."""

    pass


__all__ = [
    "ExcelMCPError",
    "WorkbookError",
    "SheetError",
    "DataError",
    "ValidationError",
    "FormattingError",
    "CalculationError",
    "PivotError",
    "ChartError",
    "ExceptionTool",
]


class ExceptionTool:
    """Centralized error handling and standardized response generator for MCP tools."""

    @staticmethod
    def handle_error(
        error: Exception,
        filename: str | None = None,
        operation: str | None = None,
    ) -> dict[str, Any]:
        """Convert exceptions into clean, structured dictionaries for LLM consumption."""
        logger.error(
            f"Error during {operation or 'operation'} on {filename or 'file'}: {error}",
            exc_info=True,
        )
        return {
            "status": "error",
            "message": str(error),
            "operation": operation or "unknown",
            "filename": filename,
            "error_type": type(error).__name__,
        }
