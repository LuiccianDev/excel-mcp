"""Pydantic response models for Excel MCP Server.

These models provide structured, typed outputs for MCP tools,
optimizing context consumption and enabling programmatic LLM processing.
"""

from __future__ import annotations

from pydantic import BaseModel, Field


class SheetOutline(BaseModel):
    """Structural outline of a single worksheet."""

    name: str = Field(description="Name of the worksheet")
    max_row: int = Field(description="Total number of populated rows")
    max_column: int = Field(description="Total number of populated columns")
    dimension: str = Field(description="Populated cell range (e.g., 'A1:H50')")
    headers: list[str] = Field(
        default_factory=list, description="Top row headers if present"
    )
    header_row: int = Field(
        default=1, description="Detected 1-indexed row number containing headers"
    )
    has_formulas: bool = Field(
        default=False, description="Whether the sheet contains formulas"
    )
    formula_count: int = Field(
        default=0, description="Number of cells containing formulas"
    )
    warnings: list[str] = Field(
        default_factory=list, description="Non-fatal warnings or anomalies detected"
    )


class WorkbookOutline(BaseModel):
    """Lightweight structural summary of an entire Excel workbook."""

    filename: str = Field(description="Path to the Excel workbook file")
    sheet_count: int = Field(description="Total number of worksheets in the workbook")
    sheets: list[SheetOutline] = Field(description="List of sheet outlines")
    total_cells_populated: int = Field(
        default=0, description="Estimated count of non-empty cells"
    )


class BuildWorkbookResult(BaseModel):
    """Result of single-pass workbook building."""

    filename: str = Field(description="Path to the created or updated workbook")
    sheets_created: list[str] = Field(description="Names of created worksheets")
    total_rows_inserted: int = Field(
        description="Total data rows inserted across all sheets"
    )
    preset_applied: str = Field(
        description="Visual design preset applied ('corporate', 'financial', etc.)"
    )
    charts_created: int = Field(default=0, description="Number of charts generated")
    warnings: list[str] = Field(
        default_factory=list,
        description="Non-fatal warnings (e.g. chart generation issues)",
    )
    message: str = Field(description="Status confirmation message")


class BatchWriteResult(BaseModel):
    """Result of batch data write operation."""

    filename: str = Field(description="Target workbook path")
    sheet_name: str = Field(description="Target worksheet name")
    range_written: str = Field(description="Cell range updated (e.g. 'B2:E20')")
    rows_affected: int = Field(description="Number of rows updated")
    columns_affected: int = Field(description="Number of columns updated")
    cells_updated: int = Field(description="Total cells updated")
