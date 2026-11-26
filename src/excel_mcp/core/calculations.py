from pathlib import Path
import re
from typing import Any

from excel_mcp.core._io import resolve_sheet, workbook_session
from excel_mcp.exceptions.exception_core import FormulaError, ValidationError
from excel_mcp.utils.cell_utils import validate_cell_reference
from excel_mcp.utils.validation_utils import (
    DANGEROUS_FUNCTIONS,
    FORMULA_PREFIX,
    INJECTION_PATTERNS,
    MAX_FORMULA_LENGTH,
    validate_cell_value_secure,
    validate_formula,
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
    "validate_formula_in_cell_operation",
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

    stripped = formula.strip()
    normalized = stripped if stripped.startswith("=") else f"={stripped}"
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

    stripped = formula.strip()
    normalized = stripped if stripped.startswith("=") else f"={stripped}"
    is_valid, message = validate_formula_secure(normalized)
    if not is_valid:
        raise FormulaSecurityError(f"Invalid formula: {message}")

    return _write_formula_cell(filename, sheet_name, cell, normalized)


def validate_formula_in_cell_operation(
    filepath: str, sheet_name: str, cell: str, formula: str
) -> dict[str, Any]:
    """Validate formula syntax, security, and referenced cells against an existing sheet."""
    if not validate_cell_reference(cell):
        return {"status": "error", "message": f"Invalid cell reference: {cell}"}

    try:
        with workbook_session(filepath, mode="read") as wb:
            if sheet_name not in wb.sheetnames:
                return {"status": "error", "message": f"Sheet '{sheet_name}' not found"}
            resolve_sheet(wb, sheet_name)
            is_valid, message = validate_formula(formula)
            if not is_valid:
                return {
                    "status": "error",
                    "message": f"Invalid formula syntax: {message}",
                }
            sec_valid, sec_message = validate_formula_secure(formula)
            if not sec_valid:
                return {
                    "status": "error",
                    "message": f"Formula security validation failed: {sec_message}",
                }
            cell_refs = re.findall(r"[A-Za-z]+[0-9]+(?::[A-Za-z]+[0-9]+)?", formula)
            for ref in cell_refs:
                if ":" in ref:
                    start, end = ref.split(":")
                    if not (
                        validate_cell_reference(start) and validate_cell_reference(end)
                    ):
                        return {
                            "status": "error",
                            "message": f"Invalid cell range reference in formula: {ref}",
                        }
                elif not validate_cell_reference(ref):
                    return {
                        "status": "error",
                        "message": f"Invalid cell reference in formula: {ref}",
                    }
            return {
                "status": "success",
                "message": f"Formula '{formula}' is valid and ready to be applied to cell {cell}",
                "valid": True,
                "cell": cell,
                "formula": formula,
            }
    except Exception as e:
        return {"status": "error", "message": str(e)}
