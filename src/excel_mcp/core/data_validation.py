"""Worksheet data validation rules for Excel MCP Server.

Creates openpyxl-native :class:`~openpyxl.worksheet.datavalidation.DataValidation`
rules so generated workbooks can constrain user input (dropdown lists, numeric
bounds, dates, text length and custom formulas) instead of relying on post-hoc
checks.
"""

from __future__ import annotations

import os
from typing import Any

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

from excel_mcp.core._io import atomic_save
from excel_mcp.exceptions.exception_core import SheetNotFoundError
from excel_mcp.exceptions.exception_tools import ValidationError
from excel_mcp.utils.cell_utils import parse_cell_range

#: Validation types supported by openpyxl's DataValidation.
VALID_VALIDATION_TYPES = frozenset(
    {"list", "whole", "decimal", "date", "time", "textLength", "custom"}
)

#: Comparison operators supported by openpyxl's DataValidation.
VALID_OPERATORS = frozenset(
    {
        "between",
        "notBetween",
        "equal",
        "notEqual",
        "greaterThan",
        "lessThan",
        "greaterThanOrEqual",
        "lessThanOrEqual",
    }
)


def _get_sheet(wb: openpyxl.Workbook, sheet_name: str | None = None) -> Any:
    """Return the target worksheet or the active sheet when no name is given."""
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise SheetNotFoundError(f"Worksheet '{sheet_name}' not found in workbook.")
        return wb[sheet_name]
    return wb.active


def add_data_validation(
    filename: str,
    cell_range: str,
    validation_type: str,
    operator: str = "between",
    formula1: str | None = None,
    formula2: str | None = None,
    allow_blank: bool = True,
    prompt_title: str | None = None,
    prompt: str | None = None,
    error_title: str | None = None,
    error_message: str | None = None,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Attach a data validation rule to a worksheet range.

    Args:
        filename: Path to the target .xlsx file.
        cell_range: Target range in A1 notation (e.g. ``"A2:A100"``).
        validation_type: One of ``list``, ``whole``, ``decimal``, ``date``,
            ``time``, ``textLength`` or ``custom``.
        operator: Comparison operator for non-list validations.
        formula1: First operand (list source, lower bound or custom formula).
        formula2: Second operand for ``between``/``notBetween`` rules.
        allow_blank: Whether empty cells bypass the validation.
        prompt_title: Optional input-message title.
        prompt: Optional input message shown when the cell is selected.
        error_title: Optional error-dialog title.
        error_message: Optional error-dialog message.
        sheet_name: Target sheet name (defaults to the active sheet).

    Returns:
        Summary dictionary describing the applied rule.

    Raises:
        FileNotFoundError: If the workbook does not exist.
        SheetNotFoundError: If ``sheet_name`` is not present.
        ValidationError: If the type, operator, range or formulas are invalid.
    """
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Workbook not found: {filename}")

    if validation_type not in VALID_VALIDATION_TYPES:
        raise ValidationError(
            f"Unsupported validation type '{validation_type}'. "
            f"Expected one of: {', '.join(sorted(VALID_VALIDATION_TYPES))}."
        )

    if validation_type != "list" and operator not in VALID_OPERATORS:
        raise ValidationError(
            f"Unsupported operator '{operator}'. "
            f"Expected one of: {', '.join(sorted(VALID_OPERATORS))}."
        )

    if not formula1:
        raise ValidationError(
            "Missing required parameter 'formula1' for the selected validation type."
        )

    if (
        operator in ("between", "notBetween")
        and validation_type != "list"
        and not formula2
    ):
        raise ValidationError(
            f"Operator '{operator}' requires both 'formula1' and 'formula2'."
        )

    try:
        parse_cell_range(cell_range)
    except ValueError as exc:
        raise ValidationError(f"Invalid cell range '{cell_range}': {exc}") from exc

    wb = openpyxl.load_workbook(filename)
    try:
        ws = _get_sheet(wb, sheet_name)

        validation = DataValidation(
            type=validation_type,
            operator=None if validation_type == "list" else operator,
            formula1=formula1,
            formula2=formula2,
            allow_blank=allow_blank,
        )
        validation.add(cell_range)

        if prompt_title or prompt:
            validation.promptTitle = prompt_title
            validation.prompt = prompt
            validation.showInputMessage = True
        if error_title or error_message:
            validation.errorTitle = error_title
            validation.error = error_message
            validation.showErrorMessage = True

        ws.add_data_validation(validation)
        atomic_save(wb, filename)
    finally:
        wb.close()

    return {
        "sheet": ws.title,
        "range": cell_range,
        "validation_type": validation_type,
        "operator": None if validation_type == "list" else operator,
    }
