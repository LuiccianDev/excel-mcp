from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import openpyxl
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils import get_column_letter

from excel_mcp.core._io import resolve_sheet, workbook_session
from excel_mcp.core.calculations import assert_safe_cell_value
from excel_mcp.core.data_validation import apply_data_validation_to_sheet
from excel_mcp.exceptions.exception_core import ValidationError
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
    ctype = chart_config.get("type", "bar").lower()
    title = chart_config.get("title", "Chart")
    chart_cell = chart_config.get("position", f"B{end_row + 3}")

    data_col = chart_config.get("data_column", 2)
    label_col = chart_config.get("label_column", 1)

    for col_name, col_val in (("data_column", data_col), ("label_column", label_col)):
        if (
            isinstance(col_val, bool)
            or not isinstance(col_val, int)
            or col_val < 1
            or col_val > num_cols
        ):
            raise ValueError(
                f"chart {col_name} {col_val!r} out of bounds (1-{num_cols})"
            )

    chart_cls = {"line": LineChart, "pie": PieChart}.get(ctype, BarChart)
    chart = chart_cls()
    chart.title = title
    chart.style = 10
    chart.height = chart_config.get("height", 12)
    chart.width = chart_config.get("width", 18)

    data = Reference(ws, min_col=data_col, min_row=start_row, max_row=end_row)
    labels = Reference(ws, min_col=label_col, min_row=start_row + 1, max_row=end_row)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    ws.add_chart(chart, chart_cell)


def _resolve_col_index(key: Any, headers: list[str]) -> int | None:
    if isinstance(key, bool):
        return None
    if isinstance(key, int):
        return key if key >= 0 else None
    if headers and key in headers:
        return headers.index(key)
    if isinstance(key, str) and key.isdigit():
        return int(key)
    return None


def _resolve_col_indices(keys: Any, headers: list[str]) -> set[int]:
    resolved: set[int] = set()
    for key in keys or ():
        idx = _resolve_col_index(key, headers)
        if idx is not None:
            resolved.add(idx)
    return resolved


def _build_total_map(
    auto_totals: Any,
    headers: list[str],
    num_cols: int,
    sheet_title: str,
    warnings: list[str] | None,
) -> dict[int, str]:
    total_map: dict[int, str] = {}
    if isinstance(auto_totals, dict):
        items = list(auto_totals.items())
    elif isinstance(auto_totals, list):
        items = [(key, "SUM") for key in auto_totals]
    else:
        return dict.fromkeys(range(1, num_cols), "SUM")

    for key, raw_func in items:
        func_upper = str(raw_func).strip().upper()
        if func_upper not in SUPPORTED_TOTAL_FUNCTIONS:
            if warnings is not None:
                warnings.append(
                    f"Sheet '{sheet_title}' unsupported total function '{raw_func}', defaulting to SUM."
                )
            func_upper = "SUM"

        idx = _resolve_col_index(key, headers)
        if idx is None or idx == 0 or idx >= num_cols:
            if warnings is not None:
                warnings.append(
                    f"Sheet '{sheet_title}' could not apply auto_totals to column '{key}'."
                )
            continue
        total_map[idx] = func_upper
    return total_map


def _apply_sheet_data_validations(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    data_validations: Any,
    warnings: list[str] | None,
) -> None:
    specs = (
        [data_validations] if isinstance(data_validations, dict) else data_validations
    )
    for dv_spec in specs:
        if not isinstance(dv_spec, dict):
            if warnings is not None:
                warnings.append(
                    f"Sheet '{ws.title}' skipped invalid data_validation entry."
                )
            continue
        target_range = dv_spec.get("range") or dv_spec.get("cell_range")
        if not target_range:
            if warnings is not None:
                warnings.append(
                    f"Sheet '{ws.title}' skipped data validation missing range."
                )
            continue
        dv_type = dv_spec.get("type") or dv_spec.get("validation_type", "list")
        try:
            apply_data_validation_to_sheet(
                ws=ws,
                cell_range=str(target_range),
                validation_type=str(dv_type),
                operator=str(dv_spec.get("operator", "between")),
                allow_blank=bool(dv_spec.get("allow_blank", True)),
                **{
                    k: dv_spec.get(k)
                    for k in (
                        "formula1",
                        "formula2",
                        "prompt_title",
                        "prompt",
                        "error_title",
                        "error_message",
                    )
                },
            )
        except (ValidationError, ValueError, TypeError) as e:
            if warnings is not None:
                warnings.append(f"Sheet '{ws.title}' data validation warning: {e}")


@dataclass(frozen=True, slots=True)
class _SheetStyle:
    preset_name: str
    preset: dict[str, Any]
    font_name: str | None
    font_size: float | None

    @property
    def header_kwargs(self) -> dict[str, Any]:
        return {
            "preset_name": self.preset_name,
            "font_name": self.font_name,
            "font_size": self.font_size,
        }


def _write_title_banner(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    title_banner: Any,
    num_cols: int,
    style: _SheetStyle,
) -> int:
    if not title_banner:
        return 1
    assert_safe_cell_value(title_banner, f"in sheet '{ws.title}' title banner")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(num_cols, 1))
    banner_cell = ws.cell(row=1, column=1, value=title_banner)
    apply_header_style(banner_cell, **style.header_kwargs)
    effective_name = style.font_name or style.preset["font_name"]
    effective_size = (
        style.font_size
        if style.font_size is not None
        else style.preset["header_font_size"]
    )
    banner_cell.font = openpyxl.styles.Font(
        name=effective_name,
        size=min(effective_size + 1, 12),
        bold=True,
        color=style.preset["header_font_color"],
    )
    ws.row_dimensions[1].height = 24
    return 3


def _write_headers(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    headers: list[Any],
    current_row: int,
    style: _SheetStyle,
) -> int:
    if not headers:
        return current_row
    ws.row_dimensions[current_row].height = 22
    for col_idx, header_text in enumerate(headers, start=1):
        assert_safe_cell_value(
            header_text, f"in sheet '{ws.title}' header col {col_idx}"
        )
        cell = ws.cell(row=current_row, column=col_idx, value=header_text)
        apply_header_style(cell, **style.header_kwargs)
    return current_row + 1


def _write_data_rows(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    rows: list[list[Any]],
    current_row: int,
    currency_indices: set[int],
    numeric_indices: set[int],
    style: _SheetStyle,
) -> int:
    for r_idx, row_values in enumerate(rows, start=1):
        ws.row_dimensions[current_row].height = 20
        for col_idx, val in enumerate(row_values, start=1):
            assert_safe_cell_value(
                val, f"in sheet '{ws.title}' row {current_row} col {col_idx}"
            )
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            zero_idx = col_idx - 1
            is_curr = zero_idx in currency_indices
            is_num = (
                is_curr or (zero_idx in numeric_indices) or isinstance(val, int | float)
            )
            apply_data_cell_style(
                cell=cell,
                row_idx=r_idx,
                is_total=False,
                is_numeric=is_num,
                is_currency=is_curr,
                **style.header_kwargs,
            )
        current_row += 1
    return current_row


def _write_auto_totals(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    auto_totals: Any,
    headers: list[Any],
    num_cols: int,
    current_row: int,
    data_start_row: int,
    data_end_row: int,
    currency_indices: set[int],
    style: _SheetStyle,
    warnings: list[str] | None,
) -> int:
    if not auto_totals:
        return current_row
    ws.row_dimensions[current_row].height = 22
    first_cell = ws.cell(row=current_row, column=1, value="Total")
    apply_data_cell_style(first_cell, row_idx=1, is_total=True, **style.header_kwargs)
    total_map = _build_total_map(auto_totals, headers, num_cols, ws.title, warnings)
    for col_idx in range(2, num_cols + 1):
        cell = ws.cell(row=current_row, column=col_idx)
        zero_idx = col_idx - 1
        agg_func = total_map.get(zero_idx)
        if agg_func is not None:
            col_letter = get_column_letter(col_idx)
            cell_range = f"{col_letter}{data_start_row}:{col_letter}{data_end_row}"
            cell.value = (
                f"=SUBTOTAL(9,{cell_range})"
                if agg_func == "SUBTOTAL"
                else f"={agg_func}({cell_range})"
            )
        apply_data_cell_style(
            cell,
            row_idx=1,
            is_total=True,
            is_numeric=agg_func is not None,
            is_currency=(zero_idx in currency_indices) and (agg_func is not None),
            **style.header_kwargs,
        )
    return current_row + 1


def _render_sheet_charts(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    chart_config: Any,
    charts_config: list[Any],
    header_start_row: int,
    data_end_row: int,
    num_cols: int,
    warnings: list[str] | None,
) -> int:
    charts_rendered = 0
    all_charts = ([chart_config] if chart_config else []) + list(charts_config)
    for chart_spec in all_charts:
        try:
            _render_chart(
                ws=ws,
                chart_config=chart_spec,
                start_row=header_start_row,
                end_row=data_end_row,
                num_cols=num_cols,
            )
            charts_rendered += 1
        except Exception as e:
            if warnings is not None:
                warnings.append(f"Sheet '{ws.title}' chart generation warning: {e}")
    return charts_rendered


def _build_sheet_with_stats(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    sheet_data: dict[str, Any],
    preset_name: str = "corporate",
    auto_fit: bool = True,
    warnings: list[str] | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
) -> tuple[int, int]:
    headers = sheet_data.get("headers", [])
    rows = sheet_data.get("rows", [])
    currency_indices = _resolve_col_indices(
        sheet_data.get("currency_columns", []), headers
    )
    numeric_indices = _resolve_col_indices(
        sheet_data.get("numeric_columns", []), headers
    )
    chart_config = sheet_data.get("chart")
    charts_config = sheet_data.get("charts") or []
    if isinstance(charts_config, dict):
        charts_config = [charts_config]
    data_validations = sheet_data.get("data_validations") or []
    freeze_panes = sheet_data.get("freeze_panes")
    auto_filter = sheet_data.get("auto_filter", False)

    style = _SheetStyle(
        preset_name=preset_name,
        preset=get_preset(preset_name),
        font_name=font_name,
        font_size=font_size,
    )
    num_cols = len(headers) if headers else (len(rows[0]) if rows else 1)

    header_start_row = _write_title_banner(ws, sheet_data.get("title"), num_cols, style)
    data_start_row = _write_headers(ws, headers, header_start_row, style)
    current_row = _write_data_rows(
        ws, rows, data_start_row, currency_indices, numeric_indices, style
    )
    data_end_row = current_row - 1

    if rows:
        _write_auto_totals(
            ws,
            sheet_data.get("auto_totals", False),
            headers,
            num_cols,
            current_row,
            data_start_row,
            data_end_row,
            currency_indices,
            style,
            warnings,
        )

    if freeze_panes:
        ws.freeze_panes = freeze_panes
    if auto_filter and headers:
        filter_end_row = max(data_end_row, header_start_row)
        ws.auto_filter.ref = (
            f"A{header_start_row}:{get_column_letter(num_cols)}{filter_end_row}"
        )

    if data_validations:
        _apply_sheet_data_validations(ws, data_validations, warnings)

    charts_rendered = (
        _render_sheet_charts(
            ws,
            chart_config,
            charts_config,
            header_start_row,
            data_end_row,
            num_cols,
            warnings,
        )
        if rows
        else 0
    )

    if auto_fit:
        auto_fit_column_widths(ws)

    return len(rows), charts_rendered


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
    rows_inserted, _ = _build_sheet_with_stats(
        ws, sheet_data, preset_name, auto_fit, warnings, font_name, font_size
    )
    return rows_inserted


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
            rows_inserted, charts_rendered = _build_sheet_with_stats(
                ws=ws,
                sheet_data=s_data,
                preset_name=preset,
                auto_fit=auto_fit,
                warnings=warnings,
                font_name=font_name,
                font_size=font_size,
            )
            total_rows += rows_inserted
            charts_count += charts_rendered

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
                assert_safe_cell_value(
                    val, f"at row offset {r_offset}, col offset {c_offset}"
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
