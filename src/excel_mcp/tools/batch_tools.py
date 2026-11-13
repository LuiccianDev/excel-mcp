"""Batch and composite document building tools for Excel MCP Server.

Provides high-efficiency tools to build full workbooks, worksheets, or insert
multiple tabular rows in a single pass, saving >90% tokens.
"""

from __future__ import annotations

from typing import Any, cast

from excel_mcp.core.batch import (
    batch_write_matrix,
    build_single_pass_workbook,
    inspect_workbook_outline,
)
from excel_mcp.utils.guards import tool_guard
from excel_mcp.utils.security import resolve_and_check_path as _resolve_and_check_path


def _dump(model: Any) -> dict[str, Any]:
    """Dump Pydantic model supporting both v1 and v2."""
    if hasattr(model, "model_dump"):
        return cast(dict[str, Any], model.model_dump())
    if hasattr(model, "dict"):
        return cast(dict[str, Any], model.dict())
    if isinstance(model, dict):
        return cast(dict[str, Any], model)
    return {"data": model}


@tool_guard()
def build_workbook(
    filename: str | None = None,
    file_path: str | None = None,
    sheets: list[dict[str, Any]] | None = None,
    preset: str = "corporate",
    auto_fit: bool = True,
    font_name: str | None = None,
    font_size: float | None = None,
) -> dict[str, Any]:
    """Build a complete Excel workbook with typography, zebra striping, currency/date formats, formulas, and charts in a single pass.

    Args:
        filename: Path to the .xlsx file to create or overwrite.
        file_path: Alternative alias for filename.
        sheets: List of sheet definitions. Each sheet dict contains:
            - 'name': Sheet title (str).
            - 'title': Optional banner title header (str).
            - 'headers': List of column header names (list[str]).
            - 'rows': 2D array of data values (list[list[Any]]).
            - 'auto_totals': bool or list of 0-based column indices to calculate SUM.
            - 'currency_columns': list of column names or 0-based indices for currency formatting ($#,##0.00).
            - 'chart': Optional chart definition {'type': 'bar'|'line'|'pie', 'title': '...', 'data_column': 2, 'label_column': 1}.
            - 'freeze_panes': Optional cell coordinate to freeze panes (e.g. 'A2', 'A4').
            - 'auto_filter': Optional bool to enable autofilter on the table header range.
        preset: Visual styling preset ('corporate', 'financial', 'modern', 'compact').
        auto_fit: If True, calculates optimal column widths automatically.
        font_name: Optional uniform font family override across the entire workbook (e.g. 'Calibri', 'Arial').
        font_size: Optional uniform font size override in points (e.g. 11).

    Returns:
        Structured result summary including sheets created and total rows.
    """
    target, err = _resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    if not sheets:
        sheets = [
            {"name": "Sheet1", "headers": ["Column 1"], "rows": [["No data provided"]]}
        ]

    result = build_single_pass_workbook(
        filename=target,
        sheets=sheets,
        preset=preset,
        auto_fit=auto_fit,
        font_name=font_name,
        font_size=font_size,
    )
    return {"status": "success", "data": _dump(result)}


@tool_guard()
def get_workbook_outline(
    filename: str | None = None,
    file_path: str | None = None,
) -> dict[str, Any]:
    """Inspect structural outline of an Excel workbook without dumping massive cell grids into LLM context.

    Args:
        filename: Path to the target .xlsx file.
        file_path: Alternative alias for filename.

    Returns:
        Lightweight outline containing sheet names, dimensions (e.g. A1:H45),
        detected header titles, and formula presence.
    """
    target, err = _resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    outline = inspect_workbook_outline(target)
    return {"status": "success", "data": _dump(outline)}


@tool_guard()
def batch_write_data(
    filename: str | None = None,
    file_path: str | None = None,
    sheet_name: str | None = None,
    start_cell: str = "A1",
    data_matrix: list[list[Any]] | None = None,
) -> dict[str, Any]:
    """Write a 2D matrix of values or formulas into a worksheet starting from a specific cell in one atomic operation.

    Args:
        filename: Path to the target .xlsx file.
        file_path: Alternative alias for filename.
        sheet_name: Worksheet name (defaults to active sheet if omitted).
        start_cell: Top-left coordinate (e.g., 'A1', 'B5').
        data_matrix: 2D array of values or formula strings to write.

    Returns:
        Summary of updated cell range and row/column counts.
    """
    target, err = _resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    if not data_matrix:
        return {"status": "error", "message": "data_matrix cannot be empty."}

    result = batch_write_matrix(
        filename=target,
        sheet_name=sheet_name,
        start_cell=start_cell,
        data_matrix=data_matrix,
    )
    return {"status": "success", "data": _dump(result)}
