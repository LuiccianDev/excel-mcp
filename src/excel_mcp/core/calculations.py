from pathlib import Path
from typing import Any

from excel_mcp.core._io import workbook_session
from excel_mcp.exceptions.exception_core import FormulaError, ValidationError
from excel_mcp.utils.cell_utils import validate_cell_reference
from excel_mcp.utils.validation_utils import (
    DANGEROUS_FUNCTIONS,
    FORMULA_PREFIX,
    INJECTION_PATTERNS,
    MAX_FORMULA_LENGTH,
    validate_cell_value_secure,
    validate_formula_secure,
)

__all__ = [
    "DANGEROUS_FUNCTIONS",
    "FORMULA_PREFIX",
    "FormulaSecurityError",
    "INJECTION_PATTERNS",
    "MAX_FORMULA_LENGTH",
    "apply_formula",
    "apply_formula_secure",
    "assert_safe_cell_value",
    "validate_cell_value_secure",
    "validate_formula_secure",
]


class FormulaSecurityError(FormulaError):
    """Exception raised when formula fails security validation."""


def assert_safe_cell_value(value: Any, context: str) -> None:
    is_valid, msg = validate_cell_value_secure(value)
    if not is_valid:
        raise FormulaSecurityError(
            f"Formula security validation failed {context}: {msg}"
        )


def _write_formula_cell(
    filename: str | Path,
    sheet_name: str,
    cell: str,
    formula: str,
) -> dict[str, Any]:
    with workbook_session(filename, mode="upsert", initial_sheet=None) as workbook:
        if sheet_name not in workbook.sheetnames:
            raise ValidationError(f"Sheet '{sheet_name}' not found")
        worksheet = workbook[sheet_name]

        try:
            worksheet[cell].value = formula
        except Exception as e:
            raise FormulaError(f"Failed to apply formula to cell: {str(e)}") from e

    return {
        "status": "success",
        "message": f"Applied formula '{formula}' to cell {cell}",
        "cell": cell,
        "formula": formula,
    }


def apply_formula_secure(
    filename: str | Path,
    sheet_name: str,
    cell: str,
    formula: str,
) -> dict[str, Any]:
    """Apply formula with enhanced security validation.

    This is a secure wrapper around apply_formula() that performs
    comprehensive security checks before applying the formula.

    Args:
        filename: Path to the Excel file
        sheet_name: Name of the worksheet
        cell: Cell reference (e.g., 'A1')
        formula: Excel formula to apply

    Returns:
        Dict containing operation result details
    """
    is_valid, message = validate_formula_secure(formula)
    if not is_valid:
        return {
            "status": "error",
            "message": f"Security validation failed: {message}",
            "security_error": True,
            "cell": cell,
        }

    if not validate_cell_reference(cell):
        return {
            "status": "error",
            "message": f"Invalid cell reference: {cell}",
            "cell": cell,
        }

    normalized = formula if formula.startswith("=") else f"={formula}"
    try:
        return _write_formula_cell(filename, sheet_name, cell, normalized)
    except (FormulaError, ValidationError) as e:
        return {
            "status": "error",
            "message": str(e),
            "cell": cell,
        }


def apply_formula(
    filename: str | Path,
    sheet_name: str,
    cell: str,
    formula: str,
) -> dict[str, Any]:
    """Apply an Excel formula to a specific cell in a worksheet.

    Args:
        filename: Path to the Excel file.
        sheet_name: Name of the worksheet.
        cell: Cell reference (e.g., 'A1').
        formula: Excel formula to apply (with or without '=' prefix).

    Returns:
        Dict containing operation result details.

    Raises:
        ValidationError: If cell reference or sheet is invalid.
        FormulaError: If formula application or file save fails.
    """
    if not validate_cell_reference(cell):
        raise ValidationError(f"Invalid cell reference: {cell}")

    normalized = formula if formula.startswith("=") else f"={formula}"
    is_valid, message = validate_formula_secure(normalized)
    if not is_valid:
        raise FormulaSecurityError(f"Invalid formula: {message}")

    return _write_formula_cell(filename, sheet_name, cell, normalized)
