"""Design presets and visual styling definitions for Excel spreadsheets.

Provides professional visual design presets (corporate, financial, modern, compact)
with standardized typography, header fills, zebra-striping, accounting borders,
and automatic column width calculation.
"""

from __future__ import annotations

from typing import Any
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


# -----------------------------------------------------------------------------
# Color Palettes & Preset Definitions
# -----------------------------------------------------------------------------

PRESETS: dict[str, dict[str, Any]] = {
    "corporate": {
        "name": "corporate",
        "font_name": "Calibri",
        "header_font_size": 11,
        "data_font_size": 11,
        "header_fill": "1F4E79",  # Deep Navy Blue
        "header_font_color": "FFFFFF",  # Pure White
        "zebra_fill": "EDF2F8",  # Soft Ice Blue
        "border_color": "D9D9D9",  # Light Gray
        "total_fill": "D9E1F2",  # Highlighted Soft Blue
        "number_format": "#,##0.00",
        "currency_format": "$#,##0.00",
        "date_format": "YYYY-MM-DD",
    },
    "financial": {
        "name": "financial",
        "font_name": "Calibri",
        "header_font_size": 11,
        "data_font_size": 11,
        "header_fill": "1E4620",  # Forest Green
        "header_font_color": "FFFFFF",  # Pure White
        "zebra_fill": "F0F6F1",  # Pale Mint Tint
        "border_color": "BFBFBF",  # Medium Gray
        "total_fill": "E2EFDA",  # Soft Green Accent
        "number_format": "#,##0.00",
        "currency_format": "$#,##0.00",
        "date_format": "YYYY-MM-DD",
    },
    "modern": {
        "name": "modern",
        "font_name": "Segoe UI",
        "header_font_size": 10.5,
        "data_font_size": 10.5,
        "header_fill": "2F3542",  # Charcoal / Slate
        "header_font_color": "FFFFFF",  # Pure White
        "zebra_fill": "F8F9FA",  # Off-white / Cool Gray
        "border_color": "E0E0E0",  # Minimalist Soft Gray
        "total_fill": "E9ECEF",  # Cool Gray Highlight
        "number_format": "#,##0.00",
        "currency_format": "$#,##0.00",
        "date_format": "YYYY-MM-DD",
    },
    "compact": {
        "name": "compact",
        "font_name": "Arial",
        "header_font_size": 9.5,
        "data_font_size": 9.5,
        "header_fill": "4A6572",  # Steel Gray
        "header_font_color": "FFFFFF",  # Pure White
        "zebra_fill": "F5F5F5",  # Very Light Gray
        "border_color": "CCCCCC",  # Standard Border
        "total_fill": "E0E0E0",  # Distinct Total Fill
        "number_format": "#,##0.00",
        "currency_format": "$#,##0.00",
        "date_format": "YYYY-MM-DD",
    },
}


def get_preset(preset_name: str | None) -> dict[str, Any]:
    """Retrieve preset dictionary with fallback to 'corporate'."""
    key = (preset_name or "corporate").lower().strip()
    return PRESETS.get(key, PRESETS["corporate"])


def apply_header_style(cell: Any, preset_name: str | None = None) -> None:
    """Apply standard high-contrast header formatting to a cell."""
    preset = get_preset(preset_name)
    cell.font = Font(
        name=preset["font_name"],
        size=preset["header_font_size"],
        bold=True,
        color=preset["header_font_color"],
    )
    cell.fill = PatternFill(
        start_color=preset["header_fill"],
        end_color=preset["header_fill"],
        fill_type="solid",
    )
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin = Side(border_style="thin", color=preset["border_color"])
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def apply_data_cell_style(
    cell: Any,
    row_idx: int,
    preset_name: str | None = None,
    is_total: bool = False,
    is_numeric: bool = False,
    is_currency: bool = False,
) -> None:
    """Apply data cell typography, fill (with zebra striping), and borders."""
    preset = get_preset(preset_name)
    font_name = preset["font_name"]
    font_size = preset["data_font_size"]
    border_color = preset["border_color"]

    if is_total:
        # Total / Summary row styling (Bold, total fill, accounting double bottom line)
        cell.font = Font(name=font_name, size=font_size, bold=True, color="000000")
        cell.fill = PatternFill(
            start_color=preset["total_fill"],
            end_color=preset["total_fill"],
            fill_type="solid",
        )
        thin = Side(border_style="thin", color=border_color)
        double = Side(border_style="double", color="000000")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=double)
    else:
        # Standard data cell
        cell.font = Font(name=font_name, size=font_size, bold=False, color="000000")
        # Apply zebra striping on even rows
        if row_idx % 2 == 0:
            cell.fill = PatternFill(
                start_color=preset["zebra_fill"],
                end_color=preset["zebra_fill"],
                fill_type="solid",
            )
        else:
            cell.fill = PatternFill(fill_type=None)

        thin = Side(border_style="thin", color=border_color)
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Alignment and number formats
    if is_numeric or is_currency:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        if is_currency:
            cell.number_format = preset["currency_format"]
        else:
            cell.number_format = preset["number_format"]
    else:
        cell.alignment = Alignment(horizontal="left", vertical="center")


def auto_fit_column_widths(
    ws: Worksheet, min_width: int = 12, max_width: int = 50
) -> None:
    """Automatically adjust column widths based on maximum cell content length."""
    for col_idx, col in enumerate(ws.columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for cell in col:
            # Skip MergedCell objects
            if type(cell).__name__ == "MergedCell":
                continue
            val = cell.value
            if val is not None:
                sval = str(val)
                if sval.startswith("="):
                    sval = "123,456.78"
                line_len = max(len(line) for line in sval.split("\n"))
                if line_len > max_len:
                    max_len = line_len

        adjusted_width = max(max_len + 4, min_width)
        ws.column_dimensions[col_letter].width = min(adjusted_width, max_width)
