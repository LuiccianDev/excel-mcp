"""Search and profiling operations for Excel MCP Server.

Provides cell discovery without full grid extraction and column-level profiling
for fast, token-efficient sheet inspection.
"""

from __future__ import annotations

from datetime import date, datetime
import logging
import os
import re
from typing import Any

import openpyxl
from openpyxl.utils.cell import get_column_letter

from excel_mcp.exceptions.exception_core import SheetNotFoundError

logger = logging.getLogger(__name__)


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
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Workbook not found: {filename}")
    if not query:
        return {"total_matches": 0, "matches": [], "limit_reached": False}

    pattern: re.Pattern[str] | None = None
    if use_regex:
        flags = 0 if match_case else re.IGNORECASE
        pattern = re.compile(query, flags)
    else:
        norm_query = query if match_case else query.lower()

    wb = openpyxl.load_workbook(filename, data_only=False, read_only=False)
    try:
        sheets_to_search: list[Any] = []
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                raise SheetNotFoundError(
                    f"Worksheet '{sheet_name}' not found in workbook."
                )
            sheets_to_search = [wb[sheet_name]]
        else:
            sheets_to_search = wb.worksheets

        matches: list[dict[str, Any]] = []

        for ws in sheets_to_search:
            for row in ws.iter_rows():
                for cell in row:
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
    finally:
        wb.close()


def profile_worksheet(
    filename: str,
    sheet_name: str | None = None,
    sample_rows: int | None = None,
) -> dict[str, Any]:
    """Generate lightweight column-level summary and statistical profiling for a worksheet.

    Token-efficient: inspects data types, null counts, min/max, and distinct samples
    without transmitting full row payloads.
    """
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Workbook not found: {filename}")

    wb = openpyxl.load_workbook(filename, data_only=True, read_only=True)
    try:
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                raise SheetNotFoundError(
                    f"Worksheet '{sheet_name}' not found in workbook."
                )
            ws = wb[sheet_name]
        else:
            ws = wb.active

        # Stream rows
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

        headers: list[str] = [str(c or "").strip() for c in first_row]
        num_cols = len(headers)

        col_stats: list[dict[str, Any]] = [
            {
                "non_null": 0,
                "null_count": 0,
                "types": set(),
                "numeric_min": None,
                "numeric_max": None,
                "distinct_samples": set(),
            }
            for _ in range(num_cols)
        ]

        total_sampled = 0
        for row in row_iter:
            total_sampled += 1
            if sample_rows is not None and total_sampled > sample_rows:
                break

            for c_idx in range(num_cols):
                val = row[c_idx] if c_idx < len(row) else None
                stats = col_stats[c_idx]

                if val is None or val == "":
                    stats["null_count"] += 1
                    continue

                stats["non_null"] += 1

                if isinstance(val, bool):
                    stats["types"].add("boolean")
                elif isinstance(val, int):
                    stats["types"].add("integer")
                    if stats["numeric_min"] is None or val < stats["numeric_min"]:
                        stats["numeric_min"] = val
                    if stats["numeric_max"] is None or val > stats["numeric_max"]:
                        stats["numeric_max"] = val
                elif isinstance(val, float):
                    stats["types"].add("float")
                    if stats["numeric_min"] is None or val < stats["numeric_min"]:
                        stats["numeric_min"] = val
                    if stats["numeric_max"] is None or val > stats["numeric_max"]:
                        stats["numeric_max"] = val
                elif isinstance(val, datetime | date):
                    stats["types"].add("datetime")
                else:
                    stats["types"].add("string")

                if len(stats["distinct_samples"]) < 6:
                    stats["distinct_samples"].add(str(val))

        # Format column summaries
        column_summaries = []
        for c_idx in range(num_cols):
            letter = get_column_letter(c_idx + 1)
            raw_header = (
                headers[c_idx]
                if c_idx < len(headers) and headers[c_idx]
                else f"Column {letter}"
            )
            stats = col_stats[c_idx]
            types_set = stats["types"]

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
                "non_null_count": stats["non_null"],
                "null_count": stats["null_count"],
                "samples": sorted(list(stats["distinct_samples"]))[:5],
            }

            if stats["numeric_min"] is not None:
                col_info["min"] = stats["numeric_min"]
                col_info["max"] = stats["numeric_max"]

            column_summaries.append(col_info)

        return {
            "sheet_name": ws.title,
            "total_columns": num_cols,
            "sampled_data_rows": total_sampled,
            "columns": column_summaries,
        }
    finally:
        wb.close()
