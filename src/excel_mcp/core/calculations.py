from pathlib import Path
import re
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
    "validate_formula",
    "validate_formula_in_cell_operation",
    "validate_formula_secure",
]

_SECURITY_ERROR_PREFIXES = ("Dangerous", "Unsafe function", "Suspicious pattern")


class FormulaSecurityError(FormulaError):
    """Exception raised when formula fails security validation."""


def assert_safe_cell_value(value: Any, context: str) -> None:
    is_valid, msg = validate_cell_value_secure(value)
    if not is_valid:
        raise FormulaSecurityError(
            f"Formula security validation failed {context}: {msg}"
        )


_STRING_LITERAL_RE = re.compile(r'"[^"]*"')
_QUOTED_SHEET_PREFIX_RE = re.compile(r"'(?:[^']|'')*'!")
_UNQUOTED_SHEET_PREFIX_RE = re.compile(r"[A-Za-z0-9_]+!")
_CELL_OR_RANGE_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_])\$?[A-Za-z]+\$?[0-9]+(?::\$?[A-Za-z]+\$?[0-9]+)?(?![A-Za-z0-9_(])"
)


def _validate_formula_body(cell: str, formula: str, *, require_equals: bool) -> str:
    if not validate_cell_reference(cell):
        raise ValidationError(f"Invalid cell reference: {cell}")

    stripped = formula.strip()
    if require_equals and not stripped.startswith("="):
        raise ValidationError("Invalid formula syntax: Formula must start with '='")

    normalized = stripped if stripped.startswith("=") else f"={stripped}"
    is_valid, message = validate_formula_secure(normalized)
    if not is_valid:
        if message.startswith(_SECURITY_ERROR_PREFIXES):
            raise FormulaSecurityError(f"Invalid formula: {message}")
        raise ValidationError(f"Invalid formula syntax: {message}")

    scrubbed = _STRING_LITERAL_RE.sub("", normalized)
    scrubbed = _QUOTED_SHEET_PREFIX_RE.sub("", scrubbed)
    scrubbed = _UNQUOTED_SHEET_PREFIX_RE.sub("", scrubbed)
    for ref in _CELL_OR_RANGE_REF_RE.findall(scrubbed):
        clean_ref = ref.replace("$", "")
        if ":" in clean_ref:
            start, end = clean_ref.split(":", 1)
            if not (validate_cell_reference(start) and validate_cell_reference(end)):
                raise ValidationError(f"Invalid cell range reference in formula: {ref}")
        elif not validate_cell_reference(clean_ref):
            raise ValidationError(f"Invalid cell reference in formula: {ref}")

    return normalized


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
    """Validate syntax and security rules before writing a formula cell."""
    try:
        normalized = _validate_formula_body(cell, formula, require_equals=False)
        return _write_formula_cell(filename, sheet_name, cell, normalized)
    except FormulaSecurityError as exc:
        return {
            "status": "error",
            "message": f"Security validation failed: {exc}",
            "security_error": True,
            "cell": cell,
        }
    except (FormulaError, ValidationError) as exc:
        return {
            "status": "error",
            "message": str(exc),
            "cell": cell,
        }


def apply_formula(
    filename: str | Path,
    sheet_name: str,
    cell: str,
    formula: str,
) -> dict[str, Any]:
    """Validate and apply an Excel formula in a single workbook session."""
    normalized = _validate_formula_body(cell, formula, require_equals=False)
    return _write_formula_cell(filename, sheet_name, cell, normalized)


def validate_formula_in_cell_operation(
    filepath: str, sheet_name: str, cell: str, formula: str
) -> dict[str, Any]:
    """Validate formula syntax, security, and referenced cells against an existing sheet."""
    try:
        _validate_formula_body(cell, formula, require_equals=True)
        with workbook_session(filepath, mode="read") as wb:
            if sheet_name not in wb.sheetnames:
                return {"status": "error", "message": f"Sheet '{sheet_name}' not found"}
        return {
            "status": "success",
            "message": f"Formula '{formula}' is valid and ready to be applied to cell {cell}",
            "valid": True,
            "cell": cell,
            "formula": formula,
        }
    except Exception as exc:
        return {"status": "error", "message": str(exc)}
