"""Analysis and search tools for Excel MCP Server.

Provides tools for discovering cells matching text/regex queries and profiling
worksheet column data types, nulls, and statistics without context window bloat.
"""

from __future__ import annotations

from typing import Any

from excel_mcp.core.search import find_cells_in_workbook, profile_worksheet
from excel_mcp.utils.guards import tool_guard
from excel_mcp.utils.security import resolve_and_check_path


@tool_guard()
def find_cells(
    filename: str | None = None,
    file_path: str | None = None,
    query: str | None = None,
    sheet_name: str | None = None,
    match_case: bool = False,
    exact_match: bool = False,
    use_regex: bool = False,
    max_results: int = 100,
) -> dict[str, Any]:
    """Search for cells containing text or regular expression patterns across sheets.

    Token-efficient: returns only matching cell coordinates, sheet names, and values
    without dumping the surrounding grid into LLM context.

    Args:
        filename: Path to the target .xlsx file.
        file_path: Alternative alias for filename.
        query: String or regex pattern to search for.
        sheet_name: Specific sheet name to search, or None to search all sheets.
        match_case: If True, search is case-sensitive.
        exact_match: If True, cell value must match query exactly.
        use_regex: If True, interprets query as a regular expression.
        max_results: Maximum number of matches to return (default 100).

    Returns:
        Structured list of matching cells with coordinates, values, and formula flags.
    """
    target, err = resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    if query is None:
        return {"status": "error", "message": "Missing required parameter 'query'."}

    res = find_cells_in_workbook(
        filename=target,
        query=query,
        sheet_name=sheet_name,
        match_case=match_case,
        exact_match=exact_match,
        use_regex=use_regex,
        max_results=max_results,
    )
    return {"status": "success", "data": res}


@tool_guard()
def profile_sheet(
    filename: str | None = None,
    file_path: str | None = None,
    sheet_name: str | None = None,
    sample_rows: int | None = None,
) -> dict[str, Any]:
    """Generate lightweight column-level profiling and statistical summaries for a worksheet.

    Token-efficient: reports detected column data types (numeric, integer, float,
    string, datetime, boolean), null counts, min/max for numerics, and distinct samples
    without transmitting full row payloads.

    Args:
        filename: Path to the target .xlsx file.
        file_path: Alternative alias for filename.
        sheet_name: Specific sheet name to profile (defaults to active sheet).
        sample_rows: Maximum data rows to sample (default all rows).

    Returns:
        Structured summary with column data types, min/max, nulls, and sample values.
    """
    target, err = resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    res = profile_worksheet(
        filename=target,
        sheet_name=sheet_name,
        sample_rows=sample_rows,
    )
    return {"status": "success", "data": res}
