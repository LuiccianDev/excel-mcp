# Import exceptions
from typing import Any

from excel_mcp.core.calculations import apply_formula
from excel_mcp.exceptions.exception_tools import CalculationError, ValidationError
from excel_mcp.utils.file_utils import ensure_xlsx_extension, validate_file_access

# Import core/tools/utils with new structure
from excel_mcp.utils.validation_utils import validate_formula_in_cell_operation


async def validate_formula_syntax(
    filename: str,
    sheet_name: str,
    cell: str,
    formula: str,
) -> dict[str, Any]:
    """Validate the syntax of an Excel formula without writing it to a cell.

    Context for AI/LLM:
        Use this as a crucial pre-flight check before applying a formula, especially if the formula is user-supplied or dynamically generated. This prevents writing invalid formulas to the sheet, which could corrupt calculations.

    Args:
        filename (str): Path to the Excel workbook.
        sheet_name (str): The worksheet context for the validation.
        cell (str): The cell reference to use as context for the formula.
        formula (str): The Excel formula string to validate (must start with '=').

    Returns:
        dict[str, Any]: A dictionary with validation status ("success" or "error") and a message.
    """
    filename = ensure_xlsx_extension(filename)
    try:
        result: dict[str, Any] = validate_formula_in_cell_operation(
            filename, sheet_name, cell, formula
        )
        return result
    except (ValidationError, CalculationError) as e:
        return {"status": "error", "message": f"Error: {str(e)}"}
    except Exception as e:
        return {"status": "error", "message": f"Failed to validate formula: {str(e)}"}


@validate_file_access("filename")
async def apply_formula_excel(
    filename: str,
    sheet_name: str,
    cell: str,
    formula: str,
) -> dict[str, Any]:
    """
    Apply a validated Excel formula to a specific cell in a worksheet.

    Context for AI/LLM:
        Use this tool to perform calculations within Excel by programmatically inserting formulas. It's ideal for automating summaries, financial models, or any task requiring dynamic calculations based on other cell values.

    Typical use cases:
        1. Adding a `=SUM(A1:A10)` formula to cell A11 to total a column.
        2. Placing a `=VLOOKUP(...)` formula to link data between sheets.
        3. Automating the creation of calculated columns in a data table.

    Args:
        filename (str): Path to the Excel workbook.
        sheet_name (str): The name of the worksheet where the formula will be applied.
        cell (str): The cell reference to write the formula to (e.g., "A1").
        formula (str): The Excel formula to apply (e.g., "=SUM(A1:A10)").

    Returns:
        dict[str, Any]: A status dictionary indicating success or failure, with a descriptive message.

    Notes:
        • The tool automatically validates the formula's syntax before attempting to apply it.
        • The cell's existing value or formula will be overwritten.
    """
    filename = ensure_xlsx_extension(filename)

    try:
        # First validate the formula
        validation: dict[str, Any] = validate_formula_in_cell_operation(
            filename, sheet_name, cell, formula
        )
        if isinstance(validation, dict) and validation.get("status") == "error":
            return validation

        result: dict[str, Any] = apply_formula(filename, sheet_name, cell, formula)
        return result
    except (ValidationError, CalculationError) as e:
        return {"status": "error", "message": f"Error: {str(e)}"}
    except Exception as e:
        return {"status": "error", "message": f"Failed to apply formula: {str(e)}"}
