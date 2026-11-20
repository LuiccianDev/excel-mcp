"""Worksheet data validation rules for Excel MCP Server.

Creates openpyxl-native :class:`~openpyxl.worksheet.datavalidation.DataValidation`
rules so generated workbooks can constrain user input (dropdown lists, numeric
bounds, dates, text length and custom formulas) instead of relying on post-hoc
checks.
"""

from __future__ import annotations

from typing import Any

from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.core._io import resolve_sheet, workbook_session
from excel_mcp.core.calculations import (
    FormulaSecurityError,
    validate_cell_value_secure,
    validate_formula_secure,
)
from excel_mcp.exceptions.exception_core import ValidationError
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


def _validate_data_validation_params(
    cell_range: str,
    validation_type: str,
    operator: str,
    formula1: Any,
    formula2: Any,
) -> tuple[str, str | None]:
    """Validate and normalize data validation parameters into openpyxl strings."""
    if validation_type not in VALID_VALIDATION_TYPES:
        raise ValidationError(
            f"Unsupported validation type '{validation_type}'. "
            f"Expected one of: {', '.join(sorted(VALID_VALIDATION_TYPES))}."
        )

    if operator not in VALID_OPERATORS:
        raise ValidationError(
            f"Unsupported operator '{operator}'. "
            f"Expected one of: {', '.join(sorted(VALID_OPERATORS))}."
        )

    if formula1 is None or str(formula1) == "":
        raise ValidationError(
            "Missing required parameter 'formula1' for the selected validation type."
        )

    if (
        operator in ("between", "notBetween")
        and validation_type != "list"
        and (formula2 is None or str(formula2) == "")
    ):
        raise ValidationError(
            f"Operator '{operator}' requires both 'formula1' and 'formula2'."
        )

    try:
        parse_cell_range(cell_range)
    except ValueError as exc:
        raise ValidationError(f"Invalid cell range '{cell_range}': {exc}") from exc

    norm_f1 = str(formula1)
    norm_f2 = str(formula2) if formula2 is not None else None

    if validation_type == "custom":
        is_valid, msg = validate_formula_secure(norm_f1)
    else:
        is_valid, msg = validate_cell_value_secure(norm_f1)
    if not is_valid:
        raise FormulaSecurityError(f"Unsafe formula1 in data validation: {msg}")

    if norm_f2 is not None:
        is_valid_f2, msg_f2 = validate_cell_value_secure(norm_f2)
        if not is_valid_f2:
            raise FormulaSecurityError(f"Unsafe formula2 in data validation: {msg_f2}")

    return norm_f1, norm_f2


def apply_data_validation_to_sheet(
    ws: Worksheet,
    cell_range: str,
    validation_type: str = "list",
    operator: str = "between",
    formula1: Any = None,
    formula2: Any = None,
    allow_blank: bool = True,
    prompt_title: str | None = None,
    prompt: str | None = None,
    error_title: str | None = None,
    error_message: str | None = None,
) -> DataValidation:
    """Validate parameters and attach a DataValidation rule to an open worksheet."""
    norm_f1, norm_f2 = _validate_data_validation_params(
        cell_range=cell_range,
        validation_type=validation_type,
        operator=operator,
        formula1=formula1,
        formula2=formula2,
    )

    validation = DataValidation(
        type=validation_type,
        operator=None if validation_type == "list" else operator,
        formula1=norm_f1,
        formula2=norm_f2,
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
    return validation


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
    _validate_data_validation_params(
        cell_range=cell_range,
        validation_type=validation_type,
        operator=operator,
        formula1=formula1,
        formula2=formula2,
    )

    with workbook_session(filename, mode="edit") as wb:
        ws = resolve_sheet(wb, sheet_name)
        apply_data_validation_to_sheet(
            ws=ws,
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
        )
        title = ws.title

    return {
        "sheet": title,
        "range": cell_range,
        "validation_type": validation_type,
        "operator": None if validation_type == "list" else operator,
    }
