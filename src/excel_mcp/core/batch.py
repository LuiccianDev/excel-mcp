"""Single-pass batch workbook building and lightweight outline inspection.

Enables generation of entire multi-sheet Excel workbooks with typography,
formulas, zebra striping, currency formatting, and charts in a single MCP tool call,
saving >90% tokens and preventing chatty micro-iterations.
"""

from __future__ import annotations

from typing import Any
import openpyxl
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from excel_mcp.core._io import resolve_sheet, workbook_session
from excel_mcp.core.calculations import (
    FormulaSecurityError,
    validate_cell_value_secure,
)
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

SUPPORTED_TOTAL_FUNCTIONS = frozenset(
    {"SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX", "SUBTOTAL"}
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
    warnings: list[str] | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
) -> int:
    """Build a complete formatted sheet from structured data."""
    headers = sheet_data.get("headers", [])
    rows = sheet_data.get("rows", [])
    title_banner = sheet_data.get("title")
    auto_totals = sheet_data.get("auto_totals", False)
    currency_cols = set(sheet_data.get("currency_columns", []))
    numeric_cols = set(sheet_data.get("numeric_columns", []))
    chart_config = sheet_data.get("chart")
    charts_config = sheet_data.get("charts") or []
    data_validations = sheet_data.get("data_validations") or []
    freeze_panes = sheet_data.get("freeze_panes")
    auto_filter = sheet_data.get("auto_filter", False)

    preset = get_preset(preset_name)
    effective_font_name = font_name or preset["font_name"]
    effective_font_size = (
        font_size if font_size is not None else preset["header_font_size"]
    )

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
        apply_header_style(
            banner_cell,
            preset_name,
            font_name=font_name,
            font_size=font_size,
        )
        banner_cell.font = openpyxl.styles.Font(
            name=effective_font_name,
            size=min(effective_font_size + 1, 12),
            bold=True,
            color=preset["header_font_color"],
        )
        ws.row_dimensions[current_row].height = 24
        current_row += 2

    # 2. Table Headers
    header_start_row = current_row
    if headers:
        ws.row_dimensions[current_row].height = 22
        for col_idx, header_text in enumerate(headers, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=header_text)
            apply_header_style(
                cell,
                preset_name,
                font_name=font_name,
                font_size=font_size,
            )
        current_row += 1

    data_start_row = current_row

    # 3. Data Rows
    for r_idx, row_values in enumerate(rows, start=1):
        ws.row_dimensions[current_row].height = 20
        for col_idx, val in enumerate(row_values, start=1):
            is_valid, msg = validate_cell_value_secure(val)
            if not is_valid:
                raise FormulaSecurityError(
                    f"Formula security validation failed in sheet '{ws.title}' row {current_row} col {col_idx}: {msg}"
                )
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            is_curr = (col_idx - 1 in currency_cols) or (
                headers
                and col_idx - 1 < len(headers)
                and headers[col_idx - 1] in currency_cols
            )
            is_num = (
                is_curr or (col_idx - 1 in numeric_cols) or isinstance(val, int | float)
            )
            apply_data_cell_style(
                cell=cell,
                row_idx=r_idx,
                preset_name=preset_name,
                is_total=False,
                is_numeric=is_num,
                is_currency=is_curr,
                font_name=font_name,
                font_size=font_size,
            )
        current_row += 1

    data_end_row = current_row - 1

    # 4. Optional Auto-Totals Row (supports bool, list of columns, or {col: agg_func} dict)
    if auto_totals and rows:
        ws.row_dimensions[current_row].height = 22
        first_cell = ws.cell(row=current_row, column=1, value="Total")
        apply_data_cell_style(
            first_cell,
            row_idx=1,
            preset_name=preset_name,
            is_total=True,
            font_name=font_name,
            font_size=font_size,
        )

        total_map: dict[int, str] = {}
        if isinstance(auto_totals, dict):
            for key, raw_func in auto_totals.items():
                func_upper = str(raw_func).strip().upper()
                if func_upper not in SUPPORTED_TOTAL_FUNCTIONS:
                    if warnings is not None:
                        warnings.append(
                            f"Sheet '{ws.title}' unsupported total function '{raw_func}', defaulting to SUM."
                        )
                    func_upper = "SUM"
                if isinstance(key, int):
                    total_map[key] = func_upper
                elif headers and key in headers:
                    total_map[headers.index(key)] = func_upper
        elif isinstance(auto_totals, list):
            for key in auto_totals:
                if isinstance(key, int):
                    total_map[key] = "SUM"
                elif headers and key in headers:
                    total_map[headers.index(key)] = "SUM"
        else:
            for idx in range(1, num_cols):
                total_map[idx] = "SUM"

        for col_idx in range(2, num_cols + 1):
            cell = ws.cell(row=current_row, column=col_idx)
            is_curr = (col_idx - 1 in currency_cols) or (
                headers
                and col_idx - 1 < len(headers)
                and headers[col_idx - 1] in currency_cols
            )

            agg_func = total_map.get(col_idx - 1)
            if agg_func is not None:
                col_letter = get_column_letter(col_idx)
                cell_range = f"{col_letter}{data_start_row}:{col_letter}{data_end_row}"
                if agg_func == "SUBTOTAL":
                    cell.value = f"=SUBTOTAL(9,{cell_range})"
                else:
                    cell.value = f"={agg_func}({cell_range})"
                apply_data_cell_style(
                    cell,
                    row_idx=1,
                    preset_name=preset_name,
                    is_total=True,
                    is_numeric=True,
                    is_currency=is_curr,
                    font_name=font_name,
                    font_size=font_size,
                )
            else:
                apply_data_cell_style(
                    cell,
                    row_idx=1,
                    preset_name=preset_name,
                    is_total=True,
                    font_name=font_name,
                    font_size=font_size,
                )
        current_row += 1

    # 4.5 Optional Sheet Navigation (Freeze Panes & Auto-Filter)
    if freeze_panes:
        ws.freeze_panes = freeze_panes
    if auto_filter and headers:
        filter_end_row = max(data_end_row, header_start_row)
        ws.auto_filter.ref = (
            f"A{header_start_row}:{get_column_letter(num_cols)}{filter_end_row}"
        )

    # 4.75 Optional Inline Data Validations
    for dv_spec in data_validations:
        try:
            target_range = dv_spec.get("range") or dv_spec.get("cell_range")
            dv_type = dv_spec.get("type", "list")
            if not target_range or not dv_spec.get("formula1"):
                continue
            dv = DataValidation(
                type=dv_type,
                operator=None
                if dv_type == "list"
                else dv_spec.get("operator", "between"),
                formula1=dv_spec["formula1"],
                formula2=dv_spec.get("formula2"),
                allow_blank=dv_spec.get("allow_blank", True),
            )
            if dv_spec.get("prompt_title") or dv_spec.get("prompt"):
                dv.promptTitle = dv_spec.get("prompt_title")
                dv.prompt = dv_spec.get("prompt")
                dv.showInputMessage = True
            if dv_spec.get("error_title") or dv_spec.get("error_message"):
                dv.errorTitle = dv_spec.get("error_title")
                dv.error = dv_spec.get("error_message")
                dv.showErrorMessage = True
            dv.add(target_range)
            ws.add_data_validation(dv)
        except Exception as e:
            if warnings is not None:
                warnings.append(f"Sheet '{ws.title}' data validation warning: {e}")

    # 5. Optional Chart Generation (supports single `chart` and/or list `charts`)
    all_charts = ([chart_config] if chart_config else []) + list(charts_config)
    for chart_spec in all_charts:
        if not rows:
            continue
        try:
            _render_chart(
                ws=ws,
                chart_config=chart_spec,
                start_row=header_start_row,
                end_row=data_end_row,
                num_cols=num_cols,
            )
        except Exception as e:
            if warnings is not None:
                warnings.append(f"Sheet '{ws.title}' chart generation warning: {e}")

    # 6. Auto-fit column widths
    if auto_fit:
        auto_fit_column_widths(ws)

    return len(rows)


def build_single_pass_workbook(
    filename: str,
    sheets: list[dict[str, Any]],
    preset: str = "corporate",
    auto_fit: bool = True,
    font_name: str | None = None,
    font_size: float | None = None,
) -> BuildWorkbookResult:
    """Construct an entire Excel workbook in one atomic pass."""
    sheets_created: list[str] = []
    total_rows = 0
    charts_count = 0
    warnings: list[str] = []

    with workbook_session(filename, mode="create", initial_sheet=None) as wb:
        default_sheet = wb.active

        for i, s_data in enumerate(sheets):
            sheet_name = s_data.get("name", f"Sheet{i + 1}")
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
                warnings=warnings,
                font_name=font_name,
                font_size=font_size,
            )
            total_rows += rows_inserted
            if s_data.get("rows"):
                if s_data.get("chart"):
                    charts_count += 1
                charts_count += len(s_data.get("charts") or [])

    return BuildWorkbookResult(
        filename=filename,
        sheets_created=sheets_created,
        total_rows_inserted=total_rows,
        preset_applied=preset,
        charts_created=charts_count,
        warnings=warnings,
        message=f"Successfully built workbook with {len(sheets_created)} sheets and {total_rows} rows.",
    )


def inspect_workbook_outline(filename: str) -> WorkbookOutline:
    """Inspect structural outline of workbook without loading full cell grids into context."""
    sheet_outlines: list[SheetOutline] = []
    total_cells = 0

    with workbook_session(filename, mode="read", data_only=False) as wb:
        for name in wb.sheetnames:
            ws = wb[name]
            max_r = ws.max_row or 0
            max_c = ws.max_column or 0
            dim = (
                getattr(ws, "dimensions", None)
                or f"A1:{get_column_letter(max(max_c, 1))}{max(max_r, 1)}"
            )

            headers: list[str] = []
            header_row_idx = 1
            formula_count = 0
            has_formulas = False
            sheet_populated = 0
            sheet_warnings: list[str] = []

            if max_r > 0 and max_c > 0:
                # Find the most likely header row among the first 5 rows
                # (skipping single-cell title banners or leading empty rows)
                best_headers: list[str] = []
                best_row = 1
                for r_idx, row in enumerate(
                    ws.iter_rows(
                        min_row=1,
                        max_row=min(max_r, 5),
                        max_col=min(max_c, 50),
                        values_only=True,
                    ),
                    start=1,
                ):
                    row_vals = [
                        str(v).strip() for v in row if v is not None and str(v).strip()
                    ]
                    if len(row_vals) > len(best_headers):
                        best_headers = row_vals
                        best_row = r_idx
                    if len(best_headers) >= 2:
                        break

                headers = best_headers
                header_row_idx = best_row

                # Count formulas and populated cells within a reasonable inspection window (first 500 rows, 50 cols)
                for row_cells in ws.iter_rows(
                    min_row=1,
                    max_row=min(max_r, 500),
                    max_col=min(max_c, 50),
                    values_only=False,
                ):
                    for cell in row_cells:
                        val = cell.value
                        if val is not None:
                            sheet_populated += 1
                            if isinstance(val, str) and val.startswith("="):
                                has_formulas = True
                                formula_count += 1

                if max_r > 500 or max_c > 50:
                    sheet_warnings.append(
                        f"Inspection capped at 500 rows and 50 cols (sheet dimensions: {max_r}x{max_c})"
                    )

            total_cells += sheet_populated

            sheet_outlines.append(
                SheetOutline(
                    name=name,
                    max_row=max_r,
                    max_column=max_c,
                    dimension=dim,
                    headers=headers,
                    header_row=header_row_idx,
                    has_formulas=has_formulas,
                    formula_count=formula_count,
                    warnings=sheet_warnings,
                )
            )
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
    from openpyxl.utils.cell import coordinate_to_tuple

    with workbook_session(filename, mode="edit") as wb:
        ws = resolve_sheet(wb, sheet_name)

        start_row, start_col = coordinate_to_tuple(start_cell)

        rows_affected = len(data_matrix)
        cols_affected = max(len(r) for r in data_matrix) if data_matrix else 0
        total_cells = 0

        for r_offset, row in enumerate(data_matrix):
            for c_offset, val in enumerate(row):
                is_valid, msg = validate_cell_value_secure(val)
                if not is_valid:
                    raise FormulaSecurityError(
                        f"Formula security validation failed at row offset {r_offset}, col offset {c_offset}: {msg}"
                    )
                ws.cell(
                    row=start_row + r_offset, column=start_col + c_offset, value=val
                )
                total_cells += 1

        end_cell = f"{get_column_letter(start_col + max(cols_affected - 1, 0))}{start_row + max(rows_affected - 1, 0)}"
        range_written = f"{start_cell}:{end_cell}"
        title = ws.title

    return BatchWriteResult(
        filename=filename,
        sheet_name=title,
        range_written=range_written,
        rows_affected=rows_affected,
        columns_affected=cols_affected,
        cells_updated=total_cells,
    )
