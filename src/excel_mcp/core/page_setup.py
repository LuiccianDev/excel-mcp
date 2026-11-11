"""Print layout configuration for Excel MCP Server.

Applies openpyxl-native page setup so generated workbooks print as intended:
orientation, paper size, fit-to-page scaling, print area, repeating title rows,
headers and footers, and horizontal centering.
"""

from __future__ import annotations

import os
from typing import Any

import openpyxl
from openpyxl.worksheet.properties import PageSetupProperties

from excel_mcp.core._io import atomic_save
from excel_mcp.exceptions.exception_core import SheetNotFoundError
from excel_mcp.exceptions.exception_tools import ValidationError
from excel_mcp.utils.cell_utils import parse_cell_range

#: Orientation values accepted by openpyxl's page setup.
VALID_ORIENTATIONS = frozenset({"portrait", "landscape"})

#: Mapping from human-readable paper names to openpyxl paper size codes.
PAPER_SIZES = {"A4": "9", "A3": "8", "Letter": "1", "Legal": "5"}


def _get_sheet(wb: openpyxl.Workbook, sheet_name: str | None = None) -> Any:
    """Return the target worksheet or the active sheet when no name is given."""
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise SheetNotFoundError(f"Worksheet '{sheet_name}' not found in workbook.")
        return wb[sheet_name]
    return wb.active


def configure_page_setup(
    filename: str,
    orientation: str = "portrait",
    paper_size: str = "A4",
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

    Args:
        filename: Path to the target .xlsx file.
        orientation: ``portrait`` or ``landscape``.
        paper_size: One of ``A4``, ``A3``, ``Letter`` or ``Legal``.
        fit_to_width: Number of pages the sheet is scaled to fit horizontally.
        fit_to_height: Number of pages the sheet is scaled to fit vertically.
        print_area: Print area in A1 notation (e.g. ``"A1:H50"``).
        print_title_rows: Repeating title rows (e.g. ``"1:2"``).
        header: Text placed in the centered odd-page header.
        footer: Text placed in the centered odd-page footer.
        center_horizontally: Center the printed content horizontally.
        sheet_name: Target sheet name (defaults to the active sheet).

    Returns:
        Summary dictionary describing the applied print layout.

    Raises:
        FileNotFoundError: If the workbook does not exist.
        SheetNotFoundError: If ``sheet_name`` is not present.
        ValidationError: If the orientation, paper size or print area is invalid.
    """
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Workbook not found: {filename}")

    if orientation not in VALID_ORIENTATIONS:
        raise ValidationError(
            f"Unsupported orientation '{orientation}'. "
            f"Expected one of: {', '.join(sorted(VALID_ORIENTATIONS))}."
        )

    if paper_size not in PAPER_SIZES:
        raise ValidationError(
            f"Unsupported paper size '{paper_size}'. "
            f"Expected one of: {', '.join(sorted(PAPER_SIZES))}."
        )

    if print_area:
        try:
            parse_cell_range(print_area)
        except ValueError as exc:
            raise ValidationError(f"Invalid print area '{print_area}': {exc}") from exc

    wb = openpyxl.load_workbook(filename)
    try:
        ws = _get_sheet(wb, sheet_name)

        ws.page_setup.orientation = orientation
        ws.page_setup.paperSize = PAPER_SIZES[paper_size]

        if fit_to_width is not None or fit_to_height is not None:
            if ws.sheet_properties.pageSetUpPr is None:
                ws.sheet_properties.pageSetUpPr = PageSetupProperties()
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            if fit_to_width is not None:
                ws.page_setup.fitToWidth = fit_to_width
            if fit_to_height is not None:
                ws.page_setup.fitToHeight = fit_to_height

        if print_area:
            ws.print_area = print_area
        if print_title_rows:
            ws.print_title_rows = print_title_rows
        if header:
            ws.oddHeader.center.text = header
        if footer:
            ws.oddFooter.center.text = footer
        if center_horizontally:
            ws.print_options.horizontalCentered = True

        atomic_save(wb, filename)
    finally:
        wb.close()

    return {
        "sheet": ws.title,
        "orientation": orientation,
        "paper_size": paper_size,
        "fit_to_page": fit_to_width is not None or fit_to_height is not None,
        "print_area": str(ws.print_area) if print_area else None,
        "print_title_rows": print_title_rows,
        "has_header": bool(header),
        "has_footer": bool(footer),
        "center_horizontally": center_horizontally,
    }
