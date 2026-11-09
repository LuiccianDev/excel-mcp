"""Tools exceptions for the MCP Excel Office tools module.

These exceptions are used by the tools layer for handling
errors related to Excel operations at the tool/interface level.
"""

from __future__ import annotations

import logging
from typing import Any

logger = logging.getLogger(__name__)


class ExcelMCPError(Exception):
    """Base exception for Excel MCP errors."""

    pass


class WorkbookError(ExcelMCPError):
    """Raised when workbook operations fail."""

    pass


class SheetError(ExcelMCPError):
    """Raised when sheet operations fail."""

    pass


class DataError(ExcelMCPError):
    """Raised when data operations fail."""

    pass


class ValidationError(ExcelMCPError):
    """Raised when validation fails."""

    pass


class FormattingError(ExcelMCPError):
    """Raised when formatting operations fail."""

    pass


class CalculationError(ExcelMCPError):
    """Raised when formula calculations fail."""

    pass


class PivotError(ExcelMCPError):
    """Raised when pivot table operations fail."""

    pass


class ChartError(ExcelMCPError):
    """Raised when chart operations fail."""

    pass


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
