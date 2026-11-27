"""Search and profiling operations for Excel MCP Server.

Provides cell discovery without full grid extraction and column-level profiling
for fast, token-efficient sheet inspection.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
import logging
from pathlib import Path
import re
from typing import Any, Final

from excel_mcp.core._io import resolve_sheet, workbook_session
from excel_mcp.exceptions.exception_core import ValidationError
from excel_mcp.models.response_models import SheetOutline, WorkbookOutline
from excel_mcp.utils.cell_utils import get_column_letter

logger = logging.getLogger(__name__)

MAX_REGEX_LENGTH: Final[int] = 256
MAX_SEARCH_CELLS: Final[int] = 500_000
MAX_PROFILE_ROWS: Final[int] = 50_000
MAX_PROFILE_COLUMNS: Final[int] = 500
MAX_PROFILE_CELLS: Final[int] = 500_000


@dataclass
class _GroupFrame:
    has_quantifier: bool = False
    has_alt: bool = False


@dataclass
class _ColumnStats:
    non_null: int = 0
    null_count: int = 0
    types: set[str] = field(default_factory=set)
    numeric_min: int | float | None = None
    numeric_max: int | float | None = None
    distinct_samples: set[str] = field(default_factory=set)


def _read_brace_quantifier_end(pattern: str, start: int) -> int | None:
    if start >= len(pattern) or pattern[start] != "{":
        return None
    end = pattern.find("}", start + 1)
    if end == -1:
        return None
    body = pattern[start + 1 : end]
    if re.fullmatch(r"\d+(?:,\d*)?", body):
        return end
    return None


def _has_nested_quantifier(pattern: str) -> bool:
    stack: list[_GroupFrame] = []
    in_char_class = False
    escaped = False
    last_closed_risky = False
    prev_unbounded_quant = False
    pending_atom = False
    idx = 0
    length = len(pattern)

    while idx < length:
        ch = pattern[idx]
        if escaped:
            if pending_atom:
                prev_unbounded_quant = False
            pending_atom = True
            escaped = False
            last_closed_risky = False
            idx += 1
            continue
        if ch == "\\":
            escaped = True
            last_closed_risky = False
            idx += 1
            continue
        if in_char_class:
            if ch == "]":
                in_char_class = False
                pending_atom = True
            idx += 1
            continue
        if ch == "[":
            if pending_atom:
                prev_unbounded_quant = False
            in_char_class = True
            pending_atom = False
            last_closed_risky = False
            idx += 1
            continue
        if ch == "(":
            if pending_atom:
                prev_unbounded_quant = False
            stack.append(_GroupFrame())
            pending_atom = False
            last_closed_risky = False
            idx += 1
            continue
        if ch == ")":
            frame = stack.pop() if stack else _GroupFrame()
            last_closed_risky = frame.has_quantifier or frame.has_alt
            pending_atom = True
            if stack:
                if frame.has_quantifier:
                    stack[-1].has_quantifier = True
                if frame.has_alt:
                    stack[-1].has_alt = True
            idx += 1
            continue
        if ch == "|" and stack:
            stack[-1].has_alt = True
            last_closed_risky = False
            prev_unbounded_quant = False
            pending_atom = False
            idx += 1
            continue

        quant_end: int | None = None
        is_unbounded = False
        if ch in "+*":
            quant_end = idx
            is_unbounded = True
        elif ch == "{":
            quant_end = _read_brace_quantifier_end(pattern, idx)
            if quant_end is not None:
                is_unbounded = pattern[idx + 1 : quant_end].endswith(",")

        if quant_end is not None:
            if last_closed_risky or (is_unbounded and prev_unbounded_quant):
                return True
            for frame in stack:
                frame.has_quantifier = True
            last_closed_risky = False
            prev_unbounded_quant = is_unbounded
            pending_atom = False
            idx = quant_end + 1
            continue

        if ch == "?":
            pending_atom = False
        else:
            last_closed_risky = False
            if pending_atom:
                prev_unbounded_quant = False
            pending_atom = True
        idx += 1

    return False


def find_cells_in_workbook(
    filename: str,
    query: str,
    sheet_name: str | None = None,
    match_case: bool = False,
    exact_match: bool = False,
    use_regex: bool = False,
    max_results: int = 100,
) -> dict[str, Any]:
    """Search for cells containing text or pattern across one or all worksheets.

    Returns matching cell coordinates, sheet name, and values without
    dumping entire matrices into LLM context.
    """
    if isinstance(max_results, bool) or max_results < 1:
        raise ValidationError("max_results must be >= 1.")

    if not query:
        path = Path(filename).resolve()
        if not path.exists():
            raise FileNotFoundError(f"Workbook not found: {path}")
        return {"total_matches": 0, "matches": [], "limit_reached": False}

    pattern: re.Pattern[str] | None = None
    if use_regex:
        if len(query) > MAX_REGEX_LENGTH:
            raise ValidationError(
                f"Regex query exceeds maximum length of {MAX_REGEX_LENGTH} characters."
            )
        if _has_nested_quantifier(query):
            raise ValidationError(
                "Regex query contains nested quantifiers that may cause catastrophic backtracking."
            )
        flags = 0 if match_case else re.IGNORECASE
        try:
            pattern = re.compile(query, flags)
        except re.error as exc:
            raise ValidationError(f"Invalid regex pattern: {exc}") from exc
    else:
        norm_query = query if match_case else query.lower()

    with workbook_session(filename, mode="read") as wb:
        if sheet_name:
            sheets_to_search: list[Any] = [resolve_sheet(wb, sheet_name)]
        else:
            sheets_to_search = wb.worksheets

        matches: list[dict[str, Any]] = []
        cells_scanned = 0

        for ws in sheets_to_search:
            for row in ws.iter_rows():
                for cell in row:
                    cells_scanned += 1
                    if cells_scanned > MAX_SEARCH_CELLS:
                        return {
                            "total_matches": len(matches),
                            "matches": matches,
                            "limit_reached": False,
                            "scan_truncated": True,
                        }
                    val = cell.value
                    if val is None:
                        continue

                    val_str = str(val)
                    matched = False

                    if use_regex and pattern:
                        matched = bool(pattern.search(val_str))
                    elif exact_match:
                        matched = (
                            (val_str == norm_query)
                            if match_case
                            else (val_str.lower() == norm_query)
                        )
                    else:
                        matched = (
                            (norm_query in val_str)
                            if match_case
                            else (norm_query in val_str.lower())
                        )

                    if matched:
                        is_formula = isinstance(val, str) and val.startswith("=")
                        matches.append(
                            {
                                "sheet": ws.title,
                                "cell": cell.coordinate,
                                "value": val,
                                "is_formula": is_formula,
                                "data_type": cell.data_type,
                            }
                        )
                        if len(matches) >= max_results:
                            return {
                                "total_matches": len(matches),
                                "matches": matches,
                                "limit_reached": True,
                            }

        return {
            "total_matches": len(matches),
            "matches": matches,
            "limit_reached": False,
        }


def _record_cell_stat(stats: _ColumnStats, val: Any) -> None:
    if val is None or val == "":
        stats.null_count += 1
        return

    stats.non_null += 1
    if isinstance(val, bool):
        stats.types.add("boolean")
    elif isinstance(val, int | float):
        stats.types.add("integer" if isinstance(val, int) else "float")
        if stats.numeric_min is None or val < stats.numeric_min:
            stats.numeric_min = val
        if stats.numeric_max is None or val > stats.numeric_max:
            stats.numeric_max = val
    elif isinstance(val, datetime | date):
        stats.types.add("datetime")
    else:
        stats.types.add("string")

    if len(stats.distinct_samples) < 6:
        stats.distinct_samples.add(str(val))


def _build_column_summaries(
    headers: list[str], col_stats: list[_ColumnStats]
) -> list[dict[str, Any]]:
    summaries: list[dict[str, Any]] = []
    for c_idx, stats in enumerate(col_stats):
        letter = get_column_letter(c_idx + 1)
        raw_header = headers[c_idx] if headers[c_idx] else f"Column {letter}"
        types_set = stats.types
        if not types_set:
            detected_type = "empty"
        elif len(types_set) == 1:
            detected_type = next(iter(types_set))
        elif types_set == {"integer", "float"}:
            detected_type = "numeric"
        else:
            detected_type = "mixed"

        col_info: dict[str, Any] = {
            "column": letter,
            "header": raw_header,
            "detected_type": detected_type,
            "non_null_count": stats.non_null,
            "null_count": stats.null_count,
            "samples": sorted(stats.distinct_samples)[:5],
        }
        if stats.numeric_min is not None:
            col_info["min"] = stats.numeric_min
            col_info["max"] = stats.numeric_max
        summaries.append(col_info)
    return summaries


def profile_worksheet(
    filename: str,
    sheet_name: str | None = None,
    sample_rows: int | None = None,
) -> dict[str, Any]:
    """Generate lightweight column-level summary and statistical profiling for a worksheet.

    Token-efficient: inspects data types, null counts, min/max, and distinct samples
    without transmitting full row payloads.
    """
    if sample_rows is not None and (isinstance(sample_rows, bool) or sample_rows < 1):
        raise ValidationError("sample_rows must be >= 1.")

    with workbook_session(filename, mode="read", data_only=True) as wb:
        ws = resolve_sheet(wb, sheet_name)
        row_iter = ws.iter_rows(values_only=True)
        try:
            first_row = next(row_iter)
        except StopIteration:
            return {
                "sheet_name": ws.title,
                "total_rows": 0,
                "total_columns": 0,
                "columns": [],
                "sampled_rows": 0,
            }

        columns_truncated = len(first_row) > MAX_PROFILE_COLUMNS
        headers = [str(c or "").strip() for c in first_row[:MAX_PROFILE_COLUMNS]]
        num_cols = len(headers)
        col_stats = [_ColumnStats() for _ in range(num_cols)]

        effective_row_limit = (
            min(sample_rows, MAX_PROFILE_ROWS)
            if sample_rows is not None
            else MAX_PROFILE_ROWS
        )
        total_sampled = 0
        cells_scanned = 0
        scan_truncated = False
        for row in row_iter:
            if total_sampled >= effective_row_limit:
                if sample_rows is None or sample_rows > MAX_PROFILE_ROWS:
                    scan_truncated = True
                break
            if cells_scanned + num_cols > MAX_PROFILE_CELLS:
                scan_truncated = True
                break
            total_sampled += 1
            cells_scanned += num_cols
            for c_idx in range(num_cols):
                _record_cell_stat(
                    col_stats[c_idx], row[c_idx] if c_idx < len(row) else None
                )

        result: dict[str, Any] = {
            "sheet_name": ws.title,
            "total_columns": num_cols,
            "sampled_data_rows": total_sampled,
            "columns": _build_column_summaries(headers, col_stats),
        }
        if scan_truncated:
            result["scan_truncated"] = True
        if columns_truncated:
            result["columns_truncated"] = True
        return result


def _outline_single_sheet(ws: Any, name: str) -> tuple[SheetOutline, int]:
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
            row_vals = [str(v).strip() for v in row if v is not None and str(v).strip()]
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

    return (
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
        ),
        sheet_populated,
    )


def inspect_workbook_outline(filename: str) -> WorkbookOutline:
    """Inspect structural outline of workbook without loading full cell grids into context."""
    sheet_outlines: list[SheetOutline] = []
    total_cells = 0

    with workbook_session(filename, mode="read", data_only=False) as wb:
        for name in wb.sheetnames:
            outline, sheet_populated = _outline_single_sheet(wb[name], name)
            total_cells += sheet_populated
            sheet_outlines.append(outline)

    return WorkbookOutline(
        filename=filename,
        sheet_count=len(sheet_outlines),
        sheets=sheet_outlines,
        total_cells_populated=total_cells,
    )
