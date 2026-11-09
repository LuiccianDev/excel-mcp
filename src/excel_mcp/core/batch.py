"""Single-pass batch workbook building and lightweight outline inspection.

Enables generation of entire multi-sheet Excel workbooks with typography,
formulas, zebra striping, currency formatting, and charts in a single MCP tool call,
saving >90% tokens and preventing chatty micro-iterations.
"""

from __future__ import annotations

import os
from typing import Any
import openpyxl
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils import get_column_letter

from excel_mcp.core.presets import (
    apply_data_cell_style,
    apply_header_style,
    auto_fit_column_widths,
    get_preset,
)
from excel_mcp.models.response_models import (
    BatchWriteResult,
    BuildWorkbookResult,
    SheetOutline,
    WorkbookOutline,
)


def _render_chart(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    chart_config: dict[str, Any],
    start_row: int,
    end_row: int,
    num_cols: int,
) -> None:
    """Render an openpyxl chart into the worksheet if configured."""
    ctype = chart_config.get("type", "bar").lower()
    title = chart_config.get("title", "Chart")
    chart_cell = chart_config.get("position", f"B{end_row + 3}")

    data_col = chart_config.get("data_column", 2)  # 1-indexed
    label_col = chart_config.get("label_column", 1)  # 1-indexed

    if ctype == "line":
        chart = LineChart()
    elif ctype == "pie":
        chart = PieChart()
    else:
        chart = BarChart()

    chart.title = title
    chart.style = 10
    chart.height = chart_config.get("height", 12)
    chart.width = chart_config.get("width", 18)

    data = Reference(ws, min_col=data_col, min_row=start_row, max_row=end_row)
    labels = Reference(ws, min_col=label_col, min_row=start_row + 1, max_row=end_row)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    ws.add_chart(chart, chart_cell)


def build_single_pass_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    sheet_data: dict[str, Any],
    preset_name: str = "corporate",
    auto_fit: bool = True,
) -> int:
    """Build a complete formatted sheet from structured data."""
    headers = sheet_data.get("headers", [])
    rows = sheet_data.get("rows", [])
    title_banner = sheet_data.get("title")
    auto_totals = sheet_data.get("auto_totals", False)
    currency_cols = set(sheet_data.get("currency_columns", []))
    numeric_cols = set(sheet_data.get("numeric_columns", []))
    chart_config = sheet_data.get("chart")

    current_row = 1
    num_cols = len(headers) if headers else (len(rows[0]) if rows else 1)

    # 1. Optional Title Banner
    if title_banner:
        ws.merge_cells(
            start_row=current_row,
            start_column=1,
            end_row=current_row,
            end_column=max(num_cols, 1),
        )
        banner_cell = ws.cell(row=current_row, column=1, value=title_banner)
        apply_header_style(banner_cell, preset_name)
        banner_cell.font = openpyxl.styles.Font(
            name=get_preset(preset_name)["font_name"],
            size=14,
            bold=True,
            color=get_preset(preset_name)["header_font_color"],
        )
        ws.row_dimensions[current_row].height = 28
        current_row += 2

    # 2. Table Headers
    header_start_row = current_row
    if headers:
        ws.row_dimensions[current_row].height = 24
        for col_idx, header_text in enumerate(headers, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=header_text)
            apply_header_style(cell, preset_name)
        current_row += 1

    data_start_row = current_row

    # 3. Data Rows
    for r_idx, row_values in enumerate(rows, start=1):
        ws.row_dimensions[current_row].height = 20
        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            is_curr = (col_idx - 1 in currency_cols) or (headers and col_idx - 1 < len(headers) and headers[col_idx - 1] in currency_cols)
            is_num = is_curr or (col_idx - 1 in numeric_cols) or isinstance(val, (int, float))
            apply_data_cell_style(
                cell=cell,
                row_idx=r_idx,
                preset_name=preset_name,
                is_total=False,
                is_numeric=is_num,
                is_currency=is_curr,
            )
        current_row += 1

    data_end_row = current_row - 1

    # 4. Optional Auto-Totals Row
    if auto_totals and rows:
        ws.row_dimensions[current_row].height = 22
        first_cell = ws.cell(row=current_row, column=1, value="Total")
        apply_data_cell_style(first_cell, row_idx=1, preset_name=preset_name, is_total=True)

        # Detect numeric columns to sum
        cols_to_sum = auto_totals if isinstance(auto_totals, list) else list(range(1, num_cols))
        for col_idx in range(2, num_cols + 1):
            cell = ws.cell(row=current_row, column=col_idx)
            is_curr = (col_idx - 1 in currency_cols) or (headers and col_idx - 1 < len(headers) and headers[col_idx - 1] in currency_cols)

            if (col_idx - 1 in cols_to_sum) or (headers and col_idx - 1 < len(headers) and headers[col_idx - 1] in cols_to_sum):
                col_letter = get_column_letter(col_idx)
                cell.value = f"=SUM({col_letter}{data_start_row}:{col_letter}{data_end_row})"
                apply_data_cell_style(
                    cell,
                    row_idx=1,
                    preset_name=preset_name,
                    is_total=True,
                    is_numeric=True,
                    is_currency=is_curr,
                )
            else:
                apply_data_cell_style(cell, row_idx=1, preset_name=preset_name, is_total=True)
        current_row += 1

    # 5. Optional Chart Generation
    if chart_config and rows:
        try:
            _render_chart(
                ws=ws,
                chart_config=chart_config,
                start_row=header_start_row,
                end_row=data_end_row,
                num_cols=num_cols,
            )
        except Exception as e:
            # Non-fatal chart rendering fallback
            ws.cell(row=current_row + 1, column=2, value=f"[Chart generation error: {e}]")

    # 6. Auto-fit column widths
    if auto_fit:
        auto_fit_column_widths(ws)

    return len(rows)


def build_single_pass_workbook(
    filename: str,
    sheets: list[dict[str, Any]],
    preset: str = "corporate",
    auto_fit: bool = True,
) -> BuildWorkbookResult:
    """Construct an entire Excel workbook in one atomic pass."""
    os.makedirs(os.path.dirname(os.path.abspath(filename)), exist_ok=True)

    wb = openpyxl.Workbook()
    default_sheet = wb.active
    sheets_created: list[str] = []
    total_rows = 0
    charts_count = 0

    for i, s_data in enumerate(sheets):
        sheet_name = s_data.get("name", f"Sheet{i+1}")
        if i == 0 and default_sheet is not None:
            ws = default_sheet
            ws.title = sheet_name
        else:
            ws = wb.create_sheet(title=sheet_name)

        sheets_created.append(sheet_name)
        rows_inserted = build_single_pass_sheet(
            ws=ws,
            sheet_data=s_data,
            preset_name=preset,
            auto_fit=auto_fit,
        )
        total_rows += rows_inserted
        if s_data.get("chart"):
            charts_count += 1

    wb.save(filename)

    return BuildWorkbookResult(
        filename=filename,
        sheets_created=sheets_created,
        total_rows_inserted=total_rows,
        preset_applied=preset,
        charts_created=charts_count,
        message=f"Successfully built workbook with {len(sheets_created)} sheets and {total_rows} rows.",
    )


def inspect_workbook_outline(filename: str) -> WorkbookOutline:
    """Inspect structural outline of workbook without loading full cell grids into context."""
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Workbook not found at {filename}")

    wb = openpyxl.load_workbook(filename, data_only=False)
    sheet_outlines: list[SheetOutline] = []
    total_cells = 0

    for name in wb.sheetnames:
        ws = wb[name]
        max_r = ws.max_row or 0
        max_c = ws.max_column or 0
        dim = ws.dimensions or f"A1:{get_column_letter(max(max_c, 1))}{max(max_r, 1)}"

        # Sample top row for headers
        headers: list[str] = []
        formula_count = 0
        has_formulas = False

        if max_r > 0 and max_c > 0:
            for col_idx in range(1, min(max_c + 1, 30)):
                val = ws.cell(row=1, column=col_idx).value
                if val is not None:
                    headers.append(str(val))

            # Inspect sample range for formulas
            for r in range(1, min(max_r + 1, 100)):
                for c in range(1, min(max_c + 1, 30)):
                    cval = ws.cell(row=r, column=c).value
                    if isinstance(cval, str) and cval.startswith("="):
                        has_formulas = True
                        formula_count += 1

        cells_populated = max_r * max_c
        total_cells += cells_populated

        sheet_outlines.append(
            SheetOutline(
                name=name,
                max_row=max_r,
                max_column=max_c,
                dimension=dim,
                headers=headers,
                has_formulas=has_formulas,
                formula_count=formula_count,
            )
        )

    wb.close()

    return WorkbookOutline(
        filename=filename,
        sheet_count=len(sheet_outlines),
        sheets=sheet_outlines,
        total_cells_populated=total_cells,
    )


def batch_write_matrix(
    filename: str,
    sheet_name: str | None,
    start_cell: str,
    data_matrix: list[list[Any]],
) -> BatchWriteResult:
    """Write a 2D matrix of values into a sheet starting from a given cell in one atomic pass."""
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Workbook not found at {filename}")

    wb = openpyxl.load_workbook(filename)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active

    # Parse start cell coordinate (e.g., 'B3')
    from openpyxl.utils.cell import coordinate_to_tuple
    start_row, start_col = coordinate_to_tuple(start_cell)

    rows_affected = len(data_matrix)
    cols_affected = max(len(r) for r in data_matrix) if data_matrix else 0
    total_cells = 0

    for r_offset, row in enumerate(data_matrix):
        for c_offset, val in enumerate(row):
            ws.cell(row=start_row + r_offset, column=start_col + c_offset, value=val)
            total_cells += 1

    end_cell = f"{get_column_letter(start_col + cols_affected - 1)}{start_row + rows_affected - 1}"
    range_written = f"{start_cell}:{end_cell}"

    wb.save(filename)

    return BatchWriteResult(
        filename=filename,
        sheet_name=ws.title,
        range_written=range_written,
        rows_affected=rows_affected,
        columns_affected=cols_affected,
        cells_updated=total_cells,
    )
