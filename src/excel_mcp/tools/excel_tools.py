from typing import Any

from excel_mcp.core.workbook import create_sheet, create_workbook
from excel_mcp.utils import paths
from excel_mcp.utils.file_utils import (
    list_excel_files_in_directory,
    validate_file_access,
)
from excel_mcp.utils.guards import tool_guard


@tool_guard()
async def create_excel_workbook(filename: str) -> dict[str, Any]:
    """
    Create an empty Excel workbook (.xlsx) in a secure, validated path.

    Note: For new reports or formatted tables, prefer `excel_build_workbook` to create sheets, data, styles, and charts in one call.

    Context for LLM/AI: Use this tool when you need to programmatically generate a new Excel workbook file as the starting point for a reporting, data collection, or automation workflow. This is typically used to initialize new datasets or reports, ensuring that the file is created only within authorized directories for security and compliance. The tool will automatically add the `.xlsx` extension if missing and will validate the path to prevent unauthorized file creation.

    When to use:
        - When starting a new workflow that requires a fresh Excel file.
        - When an automated agent needs to ensure the file is created securely.
        - When the filename or path may be user-supplied and must be validated for safety.

    Args:
        filename (str): Name or path of the workbook to create. Example: "data/reporte_diario"

    Returns:
        dict[str, Any]: A dictionary with status or error details.
            Example: {"status": "success", "filename": "/secure/path/reporte_diario.xlsx"}
        On failure, returns an error dictionary with a descriptive message.
    """
    resolved_path, constraint_error = paths.validate_and_resolve_path(
        filename, auto_extension=True
    )
    if constraint_error or resolved_path is None:
        return {"status": "error", "message": f"Access denied: {constraint_error}"}

    result: dict[str, Any] = create_workbook(str(resolved_path))
    return result


@validate_file_access("filename")
@tool_guard()
async def create_excel_worksheet(filename: str, sheet_name: str) -> dict[str, Any]:
    """
    Add a new worksheet to an existing Excel workbook.

    Context for LLM/AI: Use this tool to organize or extend data in an existing Excel file by adding a new worksheet. This is ideal when you want to segment data by category, time period, or any logical grouping (e.g., adding a new month to a financial report or a new department to a tracking file). The tool ensures the file exists, is accessible, and that the new sheet name does not conflict with existing sheets. It enforces Excel naming conventions and security constraints.

    When to use:
        - When augmenting an existing Excel file with additional data sections.
        - When automating workflows that require dynamic worksheet creation.
        - When you need to ensure that sheet names remain unique and valid.

    Args:
        filename (str): Path to the existing Excel workbook. Example: "reports/monthly_report.xlsx"
        sheet_name (str): Name of the new worksheet to add. Must be unique and valid per Excel rules.

    Returns:
        dict[str, Any]: Dictionary indicating success or describing the error.
            Example: {"status": "success", "sheet": "NewData"}
        On error, returns a dictionary with status "error" and a descriptive message.
    """
    result: dict[str, Any] = create_sheet(filename, sheet_name)
    return result


@tool_guard()
async def list_excel_documents() -> dict[str, Any]:
    """
    List all .xlsx files in the specified directory.

    Context for LLM/AI: Use this tool to discover and enumerate all Excel files in a specific directory, which is helpful for inventory, audit, batch processing, or automated data discovery scenarios. The tool validates the directory for security, ensuring only authorized locations are scanned. It returns detailed metadata for each file, enabling downstream processing, selection, or reporting by an automated agent.

    When to use:
        - When an AI needs to present the user with available Excel files for further action.
        - When preparing to process, analyze, or summarize multiple Excel documents in a folder.
        - When auditing or verifying the presence and properties of .xlsx files in a given path.

    Args:
        directory (str): Path to the folder to scan. Defaults to the current directory (".")

    Returns:
        dict[str, Any]: A dictionary with:
            - status (str): "success" or "error"
            - count (int): Number of .xlsx files found
            - directory (str): Absolute path to the searched directory
            - files (list[dict]): List of file metadata, e.g.:
                [{"name": "report.xlsx", "size": 2048, "modified": "2025-07-30T10:45:00"}]
        On error, returns a dictionary with status "error" and a descriptive message.
    """
    files = list_excel_files_in_directory()
    return {
        "status": "success",
        "count": len(files),
        "files": files,
    }
