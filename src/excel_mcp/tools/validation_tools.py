"""Data validation tools for Excel MCP Server.

Exposes worksheet input constraints (dropdown lists, numeric ranges, dates,
text length and custom formulas) as a single MCP tool.
"""

from __future__ import annotations

from typing import Any, Literal

from excel_mcp.core.data_validation import add_data_validation as _add_data_validation
from excel_mcp.utils.guards import tool_guard
from excel_mcp.utils.security import resolve_and_check_path


@tool_guard()
def add_data_validation(
    filename: str | None = None,
    file_path: str | None = None,
    cell_range: str | None = None,
    validation_type: Literal[
        "list", "whole", "decimal", "date", "time", "textLength", "custom"
    ] = "list",
    operator: Literal[
        "between",
        "notBetween",
        "equal",
        "notEqual",
        "greaterThan",
        "lessThan",
        "greaterThanOrEqual",
        "lessThanOrEqual",
    ] = "between",
    formula1: str | None = None,
    formula2: str | None = None,
    allow_blank: bool = True,
    prompt_title: str | None = None,
    prompt: str | None = None,
    error_title: str | None = None,
    error_message: str | None = None,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Add a data validation rule to a worksheet range.

    Restricts what users (or generated templates) may enter in a range: a
    dropdown list, a whole/decimal number bound, a date or time window, a text
    length limit, or a custom formula.

    Args:
        filename: Path to the target .xlsx file.
        file_path: Alternative alias for filename.
        cell_range: Target range in A1 notation (e.g. "A2:A100").
        validation_type: Rule type: list, whole, decimal, date, time, textLength or custom.
        operator: Comparison operator for non-list rules.
        formula1: List source (e.g. '"Open,Closed"'), lower bound or custom formula.
        formula2: Upper bound for between/notBetween rules.
        allow_blank: Whether empty cells bypass the validation.
        prompt_title: Optional input-message title.
        prompt: Optional input message shown when the cell is selected.
        error_title: Optional error-dialog title.
        error_message: Optional error-dialog message.
        sheet_name: Target sheet name (defaults to active sheet).

    Returns:
        Structured result with the sheet, range and applied validation type.
    """
    target, err = resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    if not cell_range:
        return {
            "status": "error",
            "message": "Missing required parameter 'cell_range'.",
        }

    res = _add_data_validation(
        filename=target,
        cell_range=cell_range,
        validation_type=validation_type,
        operator=operator,
        formula1=formula1,
        formula2=formula2,
        allow_blank=allow_blank,
        prompt_title=prompt_title,
        prompt=prompt,
        error_title=error_title,
        error_message=error_message,
        sheet_name=sheet_name,
    )
    return {"status": "success", "data": res}
