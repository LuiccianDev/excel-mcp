"""Single-pass Excel workbook builder with presets, totals, validations, and charts."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import openpyxl
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    ValidationError as PydanticValidationError,
)

from excel_mcp.core._io import validate_sheet_name, workbook_session
from excel_mcp.core.calculations import assert_safe_cell_value
from excel_mcp.core.data_validation import apply_data_validation_to_sheet
from excel_mcp.core.presets import (
    apply_data_cell_style,
    apply_header_style,
    auto_fit_column_widths,
    get_preset,
)
from excel_mcp.exceptions.exception_core import ValidationError
from excel_mcp.models.response_models import BuildWorkbookResult
from excel_mcp.utils.cell_utils import get_column_letter

SUPPORTED_TOTAL_FUNCTIONS = frozenset(
    {"SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX", "SUBTOTAL"}
)


class SheetSpec(BaseModel):
    """Typed specification for a single worksheet in a single-pass build."""

    model_config = ConfigDict(extra="forbid")

    name: str | None = None
    title: Any = None
    headers: list[Any] = Field(default_factory=list)
    rows: list[list[Any]] = Field(default_factory=list)
    auto_totals: bool | list[Any] | dict[Any, Any] = False
    currency_columns: list[Any] = Field(default_factory=list)
    numeric_columns: list[Any] = Field(default_factory=list)
    chart: dict[str, Any] | None = None
    charts: list[dict[str, Any]] | dict[str, Any] | None = None
    data_validations: list[dict[str, Any]] | dict[str, Any] | None = None
    freeze_panes: str | None = None
    auto_filter: bool = False


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


def _resolve_col_index(key: Any, headers: list[Any]) -> int | None:
    if isinstance(key, bool):
        return None
    if isinstance(key, int):
        return key if key >= 0 else None
    if headers and key in headers:
        return headers.index(key)
    if isinstance(key, str) and key.isdigit():
        return int(key)
    return None


def _resolve_col_indices(keys: Any, headers: list[Any]) -> set[int]:
    resolved: set[int] = set()
    for key in keys or ():
        idx = _resolve_col_index(key, headers)
        if idx is not None:
            resolved.add(idx)
    return resolved


def _build_total_map(
    auto_totals: Any,
    headers: list[Any],
    num_cols: int,
    sheet_title: str,
    warnings: list[str] | None,
) -> dict[int, str]:
    if isinstance(auto_totals, dict):
        items = list(auto_totals.items())
    elif isinstance(auto_totals, list):
        items = [(key, "SUM") for key in auto_totals]
    else:
        return dict.fromkeys(range(1, num_cols), "SUM")

    total_map: dict[int, str] = {}
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
        dv_type = str(dv_spec.get("type") or dv_spec.get("validation_type", "list"))
        raw_op = dv_spec.get("operator")
        if raw_op is not None:
            effective_op: str | None = str(raw_op)
        elif dv_type in {"list", "custom"} or "operator" in dv_spec:
            effective_op = None
        else:
            effective_op = "between"

        try:
            apply_data_validation_to_sheet(
                ws=ws,
                cell_range=str(target_range),
                validation_type=dv_type,
                operator=effective_op,
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
    spec: SheetSpec,
    preset_name: str = "corporate",
    auto_fit: bool = True,
    warnings: list[str] | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
) -> tuple[int, int]:
    headers = spec.headers
    rows = spec.rows
    currency_indices = _resolve_col_indices(spec.currency_columns, headers)
    numeric_indices = _resolve_col_indices(spec.numeric_columns, headers)
    charts_config = (
        [spec.charts] if isinstance(spec.charts, dict) else (spec.charts or [])
    )

    style = _SheetStyle(
        preset_name=preset_name,
        preset=get_preset(preset_name),
        font_name=font_name,
        font_size=font_size,
    )
    num_cols = len(headers) if headers else (len(rows[0]) if rows else 1)

    header_start_row = _write_title_banner(ws, spec.title, num_cols, style)
    data_start_row = _write_headers(ws, headers, header_start_row, style)
    current_row = _write_data_rows(
        ws, rows, data_start_row, currency_indices, numeric_indices, style
    )
    data_end_row = current_row - 1

    if rows:
        _write_auto_totals(
            ws,
            spec.auto_totals,
            headers,
            num_cols,
            current_row,
            data_start_row,
            data_end_row,
            currency_indices,
            style,
            warnings,
        )

    if spec.freeze_panes:
        ws.freeze_panes = spec.freeze_panes
    if spec.auto_filter and headers:
        filter_end_row = max(data_end_row, header_start_row)
        ws.auto_filter.ref = (
            f"A{header_start_row}:{get_column_letter(num_cols)}{filter_end_row}"
        )

    if spec.data_validations:
        _apply_sheet_data_validations(ws, spec.data_validations, warnings)

    charts_rendered = (
        _render_sheet_charts(
            ws,
            spec.chart,
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


def build_single_pass_workbook(
    filename: str,
    sheets: list[dict[str, Any]],
    preset: str = "corporate",
    auto_fit: bool = True,
    font_name: str | None = None,
    font_size: float | None = None,
) -> BuildWorkbookResult:
    """Construct an entire Excel workbook in one atomic pass."""
    validated_specs: list[tuple[str, SheetSpec]] = []
    seen_names: set[str] = set()
    for i, s_data in enumerate(sheets):
        try:
            spec = SheetSpec.model_validate(s_data)
        except PydanticValidationError as exc:
            raise ValidationError(f"Invalid sheet specification: {exc}") from exc
        sheet_name = spec.name if spec.name is not None else f"Sheet{i + 1}"
        validate_sheet_name(sheet_name)
        lowered = sheet_name.lower()
        if lowered in seen_names:
            raise ValidationError(
                f"Duplicate sheet name '{sheet_name}' in workbook specification"
            )
        seen_names.add(lowered)
        validated_specs.append((sheet_name, spec))

    sheets_created: list[str] = []
    total_rows = 0
    charts_count = 0
    warnings: list[str] = []

    with workbook_session(filename, mode="create", initial_sheet=None) as wb:
        default_sheet = wb.active

        for i, (sheet_name, spec) in enumerate(validated_specs):
            if i == 0 and default_sheet is not None:
                ws = default_sheet
                ws.title = sheet_name
            else:
                ws = wb.create_sheet(title=sheet_name)

            sheets_created.append(sheet_name)
            rows_inserted, charts_rendered = _build_sheet_with_stats(
                ws=ws,
                spec=spec,
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
