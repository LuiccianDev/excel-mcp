from typing import Any

from excel_mcp.core.data import read_excel_range
from excel_mcp.exceptions.exception_tools import DataError, ValidationError
from excel_mcp.utils.file_utils import ensure_xlsx_extension, validate_file_access


@validate_file_access("filename")
async def read_data_from_excel(
    filename: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str | None = None,
    preview_only: bool = False,
) -> dict[str, Any]:
    """
    Read a specific range of tabular data from an Excel worksheet and return it in a structured format.

    Note: Call `excel_get_workbook_outline` first on unfamiliar or large workbooks to inspect dimensions and headers without loading full cell matrices.

    Context for AI/LLM:
        Use this tool to extract tabular data from a defined worksheet and cell range within an Excel workbook.
        Ideal for automated data ingestion, previewing spreadsheet contents, or validating required data before processing.

    Typical use cases:
        1. Loading a data range for analysis or transformation.
        2. Previewing the first N rows of a sheet to confirm layout.
        3. Validating that a required sheet exists and contains data.

    Args:
        filename (str): Absolute or relative path to the Excel workbook (.xlsx). The extension is enforced automatically.
        sheet_name (str): Name of the worksheet to read from.
        start_cell (str, optional): Top-left cell reference of the range to read. Defaults to "A1".
        end_cell (str | None, optional): Bottom-right cell reference. If None, reads until the first empty row/column. Defaults to None.
        preview_only (bool, optional): If True, returns only a small subset (e.g., first 100 rows) for quick inspection. Defaults to False.

    Returns:
        dict[str, Any]: A dictionary containing:
            - status (str): "success" or "error".
            - data (list[list[Any]] | None): 2-D list of cell values from the worksheet range when status is "success".
            - message (str): Human-readable message or error description.

    Raises:
        FileNotFoundError: If the workbook cannot be located.
        ValidationError | DataError: If the sheet name or range is invalid or missing.

    Notes:
        • The specified worksheet must already exist; the function does not create new sheets.
        • If the target range is empty, status will be "error" with a descriptive message.
    """
    try:
        # Ensure filename has .xlsx extension
        filename = ensure_xlsx_extension(filename)

        # Read data from Excel
        data = read_excel_range(
            filename=filename,
            sheet_name=sheet_name,
            start_cell=start_cell,
            end_cell=end_cell,
            preview_only=preview_only,
        )

        # Handle empty results
        if not data:
            return {
                "status": "error",
                "message": "No data found in the specified range",
                "data": [],
            }

        return {
            "status": "success",
            "data": data,
            "message": f"Successfully read {len(data)} rows from {sheet_name}",
        }

    except FileNotFoundError:
        return {
            "status": "error",
            "message": f"File not found: {filename}",
            "data": None,
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to read Excel data: {str(e)}",
            "data": None,
        }


@validate_file_access("filename")
async def write_data_to_excel(
    filename: str,
    sheet_name: str,
    data: list[list[Any]],
    start_cell: str = "A1",
) -> str:
    """
    Write data to an existing Excel worksheet beginning at the specified cell.

    Note: Prefer `excel_build_workbook` for new workbooks or `excel_batch_write_data` for multi-range updates in a single pass.

    Context for AI/LLM:
        Employ this tool to programmatically populate or update a worksheet with tabular data produced by prior computations or external systems. Suitable for automated report generation, data export, or incremental updates to existing files.

    Typical use cases:
        1. Dumping processed data frames into Excel for business users.
        2. Appending new monthly records to a reporting sheet.
        3. Overwriting a template region with fresh values for distribution.

    Args:
        filename (str): Path to the workbook. The .xlsx extension is enforced.
        sheet_name (str): Target worksheet name. Must already exist.
        data (list[list[Any]]): 2-D list where each sub-list represents a row to write.
        start_cell (str, optional): Top-left cell where writing begins (e.g., "A1"). Defaults to "A1".

    Returns:
        str: Success confirmation message, or an error description prefixed with "Error:".

    Raises:
        ValidationError | DataError: When validation of input parameters or write operation fails.
        Exception: For unexpected I/O or library errors.

    Notes:
        • Data will overwrite any existing content in the target range.
        • The function does not create new worksheets; the target sheet must already exist.
        • Input data is written as-is, without formula injection or type conversion.
    """
    filename = ensure_xlsx_extension(filename)
    try:
        from excel_mcp.core.data import write_data

        result = write_data(filename, sheet_name, data, start_cell)
        if isinstance(result, dict) and "error" in result:
            return f"Error: {result['error']}"
        return str(result.get("message", "Data written successfully"))
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Error: Failed to write data: {str(e)}"
