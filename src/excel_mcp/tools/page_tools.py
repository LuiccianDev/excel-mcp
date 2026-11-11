"""Print layout tools for Excel MCP Server.

Exposes worksheet page setup (orientation, paper size, scaling, print area,
repeating title rows, headers and footers) as a single MCP tool.
"""

from __future__ import annotations

from typing import Any, Literal

from excel_mcp.core.page_setup import configure_page_setup as _configure_page_setup
from excel_mcp.utils.guards import tool_guard
from excel_mcp.utils.security import resolve_and_check_path


@tool_guard()
def configure_page_setup(
    filename: str | None = None,
    file_path: str | None = None,
    orientation: Literal["portrait", "landscape"] = "portrait",
    paper_size: Literal["A4", "A3", "Letter", "Legal"] = "A4",
    fit_to_width: int | None = 1,
    fit_to_height: int | None = None,
    print_area: str | None = None,
    print_title_rows: str | None = None,
    header: str | None = None,
    footer: str | None = None,
    center_horizontally: bool = False,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Configure the print layout of a worksheet.

    Sets the page orientation, paper size, fit-to-page scaling, print area,
    repeating title rows, centered header and footer, and horizontal centering
    so the workbook prints as a publication-ready document.

    Args:
        filename: Path to the target .xlsx file.
        file_path: Alternative alias for filename.
        orientation: Page orientation: portrait or landscape.
        paper_size: Paper size: A4, A3, Letter or Legal.
        fit_to_width: Pages to fit horizontally (default 1).
        fit_to_height: Pages to fit vertically (default unlimited).
        print_area: Print area in A1 notation (e.g. "A1:H50").
        print_title_rows: Repeating title rows (e.g. "1:2").
        header: Text for the centered odd-page header.
        footer: Text for the centered odd-page footer.
        center_horizontally: Center printed content horizontally.
        sheet_name: Target sheet name (defaults to active sheet).

    Returns:
        Structured result with the applied print layout settings.
    """
    target, err = resolve_and_check_path(filename, file_path)
    if err:
        return err
    assert target is not None

    res = _configure_page_setup(
        filename=target,
        orientation=orientation,
        paper_size=paper_size,
        fit_to_width=fit_to_width,
        fit_to_height=fit_to_height,
        print_area=print_area,
        print_title_rows=print_title_rows,
        header=header,
        footer=footer,
        center_horizontally=center_horizontally,
        sheet_name=sheet_name,
    )
    return {"status": "success", "data": res}
