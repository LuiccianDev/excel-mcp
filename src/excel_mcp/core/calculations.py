"""Core functionality for Excel workbook operations.

This module provides functions to apply formulas to cells with enhanced
security validation.
"""

import re
from pathlib import Path
from typing import Any, Final

from excel_mcp.core.workbook import get_or_create_workbook
from excel_mcp.exceptions.exception_core import FormulaError, ValidationError
from excel_mcp.utils.cell_utils import validate_cell_reference
from excel_mcp.utils.validation_utils import validate_formula


# Constants
FORMULA_PREFIX: Final[str] = "="

# Security: Dangerous Excel functions that should be blocked
DANGEROUS_FUNCTIONS: Final[set[str]] = {
    "CALL",
    "EXEC",
    "RUN",
    "SYSTEM",
    "SHELL",
    "EVAL",
    "REGISTER.ID",
    "XLM.CALL",
    "XLM.EXEC",
    "XLM.RUN",
    "REGISTER",
    "CALL.LIBRARY",
    "VBA.CALL",
}

# Security: Patterns that indicate potential injection attacks
INJECTION_PATTERNS: Final[list[str]] = [
    r"\b(?:http|ftp|https)://",  # URLs in formulas (data exfiltration)
    r"<script",  # Script tags
    r"document\.write",  # JavaScript injection
    r"eval\s*\(",  # JavaScript eval
    r"__import__",  # Python imports
    r"os\.",  # OS module access
    r"subprocess",  # Subprocess calls
    r"sys\.exit",  # System exit calls
    r"open\s*\(",  # File operations
    r"file://",  # File protocol
]

# Maximum formula length (Excel limit is 8192)
MAX_FORMULA_LENGTH: Final[int] = 8192


class FormulaSecurityError(FormulaError):
    """Exception raised when formula fails security validation."""

    pass


def validate_formula_secure(formula: str) -> tuple[bool, str]:
    """Comprehensive security validation for Excel formulas.

    Performs multiple security checks to prevent formula injection
    attacks and malicious code execution.

    Args:
        formula: Excel formula to validate

    Returns:
        Tuple[bool, str]: (is_valid, message)

    Checks performed:
        1. Dangerous function detection (CALL, EXEC, etc.)
        2. Injection pattern detection (URLs, scripts)
        3. Parentheses balancing
        4. Formula length validation
        5. Syntax validation
    """
    # Normalize formula for analysis
    normalized = formula.strip()
    if not normalized:
        return False, "Formula cannot be empty"

    # Ensure formula starts with '='
    if not normalized.startswith("="):
        normalized = "=" + normalized

    normalized_upper = normalized.upper()

    # Check 1: Dangerous functions
    # Extract function names using regex (supports dots for XLM functions like REGISTER.ID)
    function_pattern = r"([A-Z][A-Z0-9_.]*)\s*\("
    functions_found = set(re.findall(function_pattern, normalized_upper))

    dangerous_found = functions_found.intersection(DANGEROUS_FUNCTIONS)
    if dangerous_found:
        return False, f"Dangerous functions detected: {dangerous_found}"

    # Check 2: Injection patterns
    formula_lower = normalized.lower()
    for pattern in INJECTION_PATTERNS:
        if re.search(pattern, formula_lower):
            return False, "Suspicious pattern detected (potential injection)"

    # Check 3: Parentheses balance
    open_count = normalized.count("(")
    close_count = normalized.count(")")
    if open_count != close_count:
        return False, f"Unbalanced parentheses: {open_count} open, {close_count} close"

    # Check 4: Formula length
    if len(normalized) > MAX_FORMULA_LENGTH:
        return (
            False,
            f"Formula exceeds maximum length of {MAX_FORMULA_LENGTH} characters",
        )

    # Check 5: Basic syntax validation (delegates to existing validator)
    is_valid, message = validate_formula(normalized)
    if not is_valid:
        return False, message

    return True, "Formula is valid and secure"


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

    Example:
        result = apply_formula_secure(
            "workbook.xlsx",
            "Sheet1",
            "B2",
            "=SUM(A1:A10)"
        )
    """
    # Security validation
    is_valid, message = validate_formula_secure(formula)
    if not is_valid:
        return {
            "status": "error",
            "message": f"Security validation failed: {message}",
            "security_error": True,
            "cell": cell,
        }

    # Apply formula using existing function
    try:
        return apply_formula(filename, sheet_name, cell, formula)
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
    """
    Apply an Excel formula to a specific cell in a worksheet.

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
    # Input validation
    if not validate_cell_reference(cell):
        raise ValidationError(f"Invalid cell reference: {cell}")

    # Load workbook and validate sheet
    # Use data_only=False to preserve existing formulas when modifying the file
    workbook = get_or_create_workbook(str(filename), data_only=False)
    if sheet_name not in workbook.sheetnames:
        raise ValidationError(f"Sheet '{sheet_name}' not found")
    worksheet = workbook[sheet_name]

    # Process and validate formula - ensure formula starts with '='
    formula = formula if formula.startswith("=") else f"={formula}"

    # Validate formula syntax
    is_valid, message = validate_formula(formula)
    if not is_valid:
        raise FormulaError(f"Invalid formula syntax: {message}")

    # Apply formula to the specified cell in the worksheet
    try:
        worksheet[cell].value = formula
    except Exception as e:
        raise FormulaError(f"Failed to apply formula to cell: {str(e)}") from e

    # Save the workbook to the specified file
    try:
        workbook.save(str(filename))
    except Exception as e:
        raise FormulaError(f"Failed to save workbook: {str(e)}") from e

    # Return success result
    result = {
        "status": "success",
        "message": f"Applied formula '{formula}' to cell {cell}",
        "cell": cell,
        "formula": formula,
    }
    return result
