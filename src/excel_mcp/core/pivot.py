"""Pivot table generation with single-pass grouping and strict bounds."""

from __future__ import annotations

from collections import defaultdict
import math
from typing import Any, Final
import uuid

from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.core._io import atomic_save, resolve_sheet, workbook_session
from excel_mcp.core.data import SheetData
from excel_mcp.exceptions.exception_core import (
    PivotError,
    SheetNotFoundError,
    ValidationError,
)
from excel_mcp.utils.cell_utils import parse_cell_range

MAX_PIVOT_ROW_COMBOS: Final[int] = 50_000
MAX_PIVOT_COL_COMBOS: Final[int] = 1_000
MAX_PIVOT_CELLS: Final[int] = 250_000
VALID_AGG_FUNCS: Final[frozenset[str]] = frozenset(
    {"sum", "average", "count", "min", "max"}
)
_AGG_SUFFIXES: Final[tuple[str, ...]] = (
    " (sum)",
    " (average)",
    " (count)",
    " (min)",
    " (max)",
)


def _clean_field_name(field: str) -> str:
    text = str(field).strip()
    lower = text.lower()
    for suffix in _AGG_SUFFIXES:
        if lower.endswith(suffix):
            return text[: -len(suffix)]
    return text


def _key_token(value: Any) -> tuple[bool, Any]:
    if value is None:
        return (False, "")
    return (isinstance(value, bool), value)


def _sort_key(value: Any) -> tuple[int, Any]:
    if value is None or value == "":
        return (0, "")
    if isinstance(value, bool):
        return (1, int(value))
    if isinstance(value, int | float):
        return (2, float(value))
    return (3, str(value).lower())


def _make_pivot_sheet_name(sheet_name: str) -> str:
    return f"{sheet_name[:25]}_pivot"


def _convert_sheetdata_to_dicts(sheet_data: SheetData) -> list[dict[str, Any]]:
    """Convert SheetData (list of lists) into row dictionaries keyed by header row."""
    if not sheet_data:
        return []

    headers = [
        str(h) if h is not None else f"Column_{i}" for i, h in enumerate(sheet_data[0])
    ]
    result: list[dict[str, Any]] = []
    for row in sheet_data[1:]:
        record: dict[str, Any] = {}
        for i, value in enumerate(row):
            if i < len(headers):
                record[headers[i]] = value
        result.append(record)
    return result


def _read_range_from_sheet(
    ws: Worksheet, start_row: int, start_col: int, end_row: int, end_col: int
) -> SheetData:
    ws_max_row = ws.max_row or 0
    ws_max_col = ws.max_column or 0
    eff_end_row = min(end_row, ws_max_row) if ws_max_row > 0 else end_row
    eff_end_col = min(end_col, ws_max_col) if ws_max_col > 0 else end_col
    if eff_end_row < start_row or eff_end_col < start_col:
        return []
    rows: SheetData = []
    for row in ws.iter_rows(
        min_row=start_row,
        max_row=eff_end_row,
        min_col=start_col,
        max_col=eff_end_col,
        values_only=True,
    ):
        row_data = list(row)
        if any(v is not None for v in row_data):
            rows.append(row_data)
    return rows


def _resolve_field_list(
    fields: list[str],
    field_lookup: dict[str, str],
    field_type: str,
) -> list[str]:
    resolved: list[str] = []
    for field in fields:
        cleaned = _clean_field_name(field)
        canonical = field_lookup.get(cleaned.lower())
        if canonical is None:
            available = ", ".join(sorted(field_lookup.keys()))
            raise ValidationError(
                f"Invalid {field_type} field '{field}'. Available fields: {available}"
            )
        resolved.append(canonical)
    return resolved


def _get_combinations(
    field_values: dict[str, list[Any]], *, max_combos: int = MAX_PIVOT_ROW_COMBOS
) -> list[dict[str, Any]]:
    """Expand bounded Cartesian combinations of field values."""
    if not field_values:
        return [{}]
    sizes = [len(vals) for vals in field_values.values()]
    total_combos = math.prod(sizes)
    if total_combos > max_combos:
        raise PivotError(
            f"Pivot combination count ({total_combos}) exceeds maximum allowed ({max_combos})."
        )
    result: list[dict[str, Any]] = [{}]
    for field, values in list(field_values.items()):
        sorted_vals = sorted(values, key=_sort_key)
        next_result: list[dict[str, Any]] = []
        for combo in result:
            for value in sorted_vals:
                new_combo = combo.copy()
                new_combo[field] = value
                next_result.append(new_combo)
        result = next_result
    return result


def _filter_data(
    data: list[dict[str, Any]],
    row_filters: dict[str, Any],
    col_filters: dict[str, Any],
) -> list[dict[str, Any]]:
    """Filter data rows matching row and column field values."""
    result: list[dict[str, Any]] = []
    for record in data:
        if all(
            _key_token(record.get(f)) == _key_token(v) for f, v in row_filters.items()
        ) and all(
            _key_token(record.get(f)) == _key_token(v) for f, v in col_filters.items()
        ):
            result.append(record)
    return result


def _aggregate_numeric(nums: list[int | float], agg_func: str) -> float:
    if not nums:
        return 0.0
    func = agg_func.lower()
    if func == "average":
        return float(sum(nums) / len(nums))
    if func == "count":
        return float(len(nums))
    if func == "min":
        return float(min(nums))
    if func == "max":
        return float(max(nums))
    return float(sum(nums))


def _aggregate_values(data: list[dict[str, Any]], field: str, agg_func: str) -> float:
    """Aggregate numeric values (excluding booleans) using the requested function."""
    nums = [
        record[field]
        for record in data
        if field in record
        and isinstance(record[field], int | float)
        and not isinstance(record[field], bool)
    ]
    return _aggregate_numeric(nums, agg_func)


def _distinct_field_values(
    data: list[dict[str, Any]], fields: list[str]
) -> dict[str, list[Any]]:
    distinct: dict[str, list[Any]] = {}
    for field in fields:
        by_token: dict[tuple[bool, Any], Any] = {}
        for record in data:
            raw = record.get(field)
            norm_val = "" if raw is None else raw
            tok = _key_token(norm_val)
            if tok not in by_token:
                by_token[tok] = norm_val
        distinct[field] = sorted(by_token.values(), key=_sort_key)
    return distinct


def _build_pivot_buckets(
    data: list[dict[str, Any]],
    row_fields: list[str],
    col_fields: list[str],
    value_fields: list[str],
) -> dict[
    tuple[tuple[tuple[bool, Any], ...], tuple[tuple[bool, Any], ...]],
    dict[str, list[int | float]],
]:
    buckets: dict[
        tuple[tuple[tuple[bool, Any], ...], tuple[tuple[bool, Any], ...]],
        dict[str, list[int | float]],
    ] = defaultdict(lambda: {vf: [] for vf in value_fields})
    for record in data:
        r_key = tuple(_key_token(record.get(f)) for f in row_fields)
        c_key = tuple(_key_token(record.get(f)) for f in col_fields)
        bucket = buckets[(r_key, c_key)]
        for vf in value_fields:
            val = record.get(vf)
            if isinstance(val, int | float) and not isinstance(val, bool):
                bucket[vf].append(val)
    return buckets


def _write_pivot_sheet(
    wb: Any,
    pivot_sheet_name: str,
    row_fields: list[str],
    col_fields: list[str],
    value_fields: list[str],
    row_combos: list[dict[str, Any]],
    col_combos: list[dict[str, Any]],
    buckets: dict[
        tuple[tuple[tuple[bool, Any], ...], tuple[tuple[bool, Any], ...]],
        dict[str, list[int | float]],
    ],
    agg_func: str,
) -> None:
    col_headers: list[tuple[dict[str, Any], str, str]] = []
    if col_fields:
        for combo in col_combos:
            prefix = " | ".join(str(combo[c]) for c in col_fields)
            for vf in value_fields:
                col_headers.append((combo, vf, f"{prefix} | {vf} ({agg_func})"))
    else:
        for vf in value_fields:
            col_headers.append(({}, vf, f"{vf} ({agg_func})"))

    total_rows = len(row_combos) + 1
    total_cols = len(row_fields) + len(col_headers)
    if total_rows * total_cols > MAX_PIVOT_CELLS:
        raise PivotError(
            f"Pivot output ({total_rows * total_cols} cells) exceeds maximum allowed ({MAX_PIVOT_CELLS})."
        )

    if pivot_sheet_name in wb.sheetnames:
        wb.remove(wb[pivot_sheet_name])
    pivot_ws = wb.create_sheet(pivot_sheet_name)

    bold_font = Font(bold=True)
    header_titles = [*row_fields, *(hdr for _, _, hdr in col_headers)]
    for col_idx, text in enumerate(header_titles, start=1):
        cell = pivot_ws.cell(row=1, column=col_idx, value=text)
        cell.font = bold_font

    for row_offset, r_combo in enumerate(row_combos, start=2):
        for col_offset, rf in enumerate(row_fields, start=1):
            pivot_ws.cell(row=row_offset, column=col_offset, value=r_combo[rf])
        r_key = tuple(_key_token(r_combo[rf]) for rf in row_fields)
        out_col = len(row_fields) + 1
        for c_combo, vf, _ in col_headers:
            c_key = tuple(_key_token(c_combo[cf]) for cf in col_fields)
            nums = buckets.get((r_key, c_key), {}).get(vf, [])
            pivot_ws.cell(
                row=row_offset,
                column=out_col,
                value=_aggregate_numeric(nums, agg_func),
            )
            out_col += 1

    pivot_range = f"A1:{get_column_letter(total_cols)}{total_rows}"
    pivot_table = Table(
        displayName=f"PivotTable_{uuid.uuid4().hex[:8]}", ref=pivot_range
    )
    pivot_table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=True,
    )
    pivot_ws.add_table(pivot_table)


def create_pivot_table(
    filename: str,
    sheet_name: str,
    data_range: str,
    rows: list[str],
    values: list[str],
    columns: list[str] | None = None,
    agg_func: str = "sum",
) -> dict[str, Any]:
    """Create a pivot summary worksheet in a single workbook session."""
    try:
        if not rows or not values:
            raise ValidationError("Both 'rows' and 'values' must be non-empty.")
        if agg_func.lower() not in VALID_AGG_FUNCS:
            raise ValidationError(
                f"Invalid aggregation function. Must be one of: {', '.join(sorted(VALID_AGG_FUNCS))}"
            )
        if ":" not in data_range:
            raise ValidationError("Data range must be in format 'A1:B2'")
        try:
            start_row, start_col, end_row, end_col = parse_cell_range(data_range)
        except ValueError as exc:
            raise ValidationError(f"Invalid data range format: {exc}") from exc
        if end_row is None or end_col is None:
            raise ValidationError("Invalid data range format: missing end coordinates")

        data_range_str = (
            f"{get_column_letter(start_col)}{start_row}"
            f":{get_column_letter(end_col)}{end_row}"
        )

        with workbook_session(filename, mode="edit", save=False) as wb:
            try:
                ws = resolve_sheet(wb, sheet_name)
            except SheetNotFoundError as exc:
                raise ValidationError(f"Sheet '{sheet_name}' not found") from exc
            sheet_data = _read_range_from_sheet(
                ws, start_row, start_col, end_row, end_col
            )
            if not sheet_data:
                raise PivotError("No data found in range")
            data = _convert_sheetdata_to_dicts(sheet_data)
            if not data:
                raise PivotError("No data found after converting to dictionary format")

            field_lookup = {_clean_field_name(k).lower(): k for k in data[0].keys()}
            cleaned_rows = _resolve_field_list(rows, field_lookup, "row")
            cleaned_values = _resolve_field_list(values, field_lookup, "value")
            cleaned_columns = (
                _resolve_field_list(columns, field_lookup, "column") if columns else []
            )

            row_combos = _get_combinations(
                _distinct_field_values(data, cleaned_rows),
                max_combos=MAX_PIVOT_ROW_COMBOS,
            )
            col_combos = (
                _get_combinations(
                    _distinct_field_values(data, cleaned_columns),
                    max_combos=MAX_PIVOT_COL_COMBOS,
                )
                if cleaned_columns
                else [{}]
            )
            buckets = _build_pivot_buckets(
                data, cleaned_rows, cleaned_columns, cleaned_values
            )
            pivot_sheet_name = _make_pivot_sheet_name(sheet_name)
            _write_pivot_sheet(
                wb,
                pivot_sheet_name,
                cleaned_rows,
                cleaned_columns,
                cleaned_values,
                row_combos,
                col_combos,
                buckets,
                agg_func,
            )
            atomic_save(wb, filename)

        return {
            "message": "Summary table created successfully",
            "details": {
                "source_range": data_range_str,
                "pivot_sheet": pivot_sheet_name,
                "rows": cleaned_rows,
                "columns": cleaned_columns,
                "values": cleaned_values,
                "aggregation": agg_func,
            },
        }
    except (ValidationError, PivotError) as exc:
        return {"error": str(exc), "message": "Failed to create pivot table"}
    except Exception as exc:
        raise PivotError(str(exc)) from exc
