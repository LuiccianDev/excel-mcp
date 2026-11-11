#!/usr/bin/env python3
"""Generate comprehensive docs/TOOLS.md documentation from registered tools.

Introspects the FastMCP tool registry in excel_mcp.tools.register_tools
and produces markdown documentation including annotations, signatures,
and descriptions.
"""

from __future__ import annotations

import inspect
import os
import sys
from pathlib import Path
from typing import Any

# Ensure src is on sys.path
repo_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(repo_root / "src"))

from excel_mcp.tools.register_tools import register_all_tools


class DocCollectorMCP:
    """Mock FastMCP collector to capture tool definitions and metadata."""

    def __init__(self) -> None:
        self.tools: dict[str, dict[str, Any]] = {}

    def tool(self, name: str | None = None, annotations: Any = None) -> Any:
        def decorator(fn: Any) -> Any:
            tool_name = name or getattr(fn, "__name__", "unnamed")
            # Determine annotation type
            read_only = False
            destructive = False
            if annotations:
                kwargs = getattr(annotations, "kwargs", {})
                read_only = getattr(annotations, "read_only_hint", False) or kwargs.get(
                    "read_only_hint", False
                )
                destructive = getattr(
                    annotations, "destructive_hint", False
                ) or kwargs.get("destructive_hint", False)

            self.tools[tool_name] = {
                "fn": fn,
                "read_only": read_only,
                "destructive": destructive,
                "doc": inspect.getdoc(fn) or "No description provided.",
                "sig": inspect.signature(fn),
            }
            return fn

        return decorator


def categorize_tools(tool_names: list[str]) -> dict[str, list[str]]:
    """Group canonical tools into logical sections."""
    categories: dict[str, list[str]] = {
        "Batch & Single-Pass Tools": [
            "excel_build_workbook",
            "excel_get_workbook_outline",
            "excel_batch_write_data",
            "excel_append_rows",
        ],
        "Structure & Worksheet Management": [
            "excel_create_workbook",
            "excel_create_worksheet",
            "excel_copy_worksheet",
            "excel_delete_worksheet",
            "excel_rename_worksheet",
            "excel_list_documents",
            "excel_get_workbook_metadata",
            "excel_insert_rows_columns",
            "excel_delete_rows_columns",
        ],
        "Formatting, Tables & Layout": [
            "excel_format_range",
            "excel_format_sheet",
            "excel_create_table",
            "excel_sort_range",
            "excel_merge_cells",
            "excel_unmerge_cells",
            "excel_copy_range",
            "excel_delete_range",
            "excel_validate_range",
        ],
        "Search & Analytical Profiling": [
            "excel_find_cells",
            "excel_profile_sheet",
        ],
        "Formulas & Analytics": [
            "excel_validate_formula_syntax",
            "excel_apply_formula",
            "excel_create_chart",
            "excel_create_pivot_table",
        ],
        "Direct Content I/O": [
            "excel_read_data",
            "excel_write_data",
        ],
        "Validation & Print Layout": [
            "excel_add_data_validation",
            "excel_protect_sheet",
            "excel_configure_page_setup",
        ],
    }

    # Safety net: a newly registered tool must never disappear from the
    # generated reference just because nobody updated this mapping yet.
    categorized = {name for names in categories.values() for name in names}
    uncategorized = [name for name in tool_names if name not in categorized]
    if uncategorized:
        categories["Other Tools"] = uncategorized

    return categories


def generate_tools_markdown() -> str:
    collector = DocCollectorMCP()
    os.environ["EXCEL_MCP_LEGACY_ALIASES"] = "0"
    register_all_tools(collector)

    lines: list[str] = [
        "# Excel MCP Tools Reference",
        "",
        "> [!NOTE]",
        "> This document is automatically generated from the registered FastMCP tool signatures.",
        f"> Total Canonical Tools: **{len(collector.tools)}**.",
        "",
        "## Tool Inventory",
        "",
        "| Canonical Tool Name | Type | Description |",
        "| :--- | :--- | :--- |",
    ]

    for name, meta in sorted(collector.tools.items()):
        mode = (
            "Destructive"
            if meta["destructive"]
            else ("Read-Only" if meta["read_only"] else "Write")
        )
        first_line = meta["doc"].strip().split("\n")[0].replace("|", "\\|")
        lines.append(
            f"| [`{name}`](#{name.replace('_', '-')}) | `{mode}` | {first_line} |"
        )

    lines.append("")
    categories = categorize_tools(list(collector.tools.keys()))

    for cat_name, t_names in categories.items():
        lines.append(f"## {cat_name}")
        lines.append("")
        for name in t_names:
            if name not in collector.tools:
                continue
            meta = collector.tools[name]
            mode = (
                "Destructive"
                if meta["destructive"]
                else ("Read-Only" if meta["read_only"] else "Write")
            )
            lines.append(f"### `{name}`")
            lines.append("")
            lines.append(f"- **Access Mode:** `{mode}`")
            lines.append(f"- **Signature:** `{name}{meta['sig']}`")
            lines.append("")
            lines.append("```text")
            lines.append(meta["doc"])
            lines.append("```")
            lines.append("")

    return "\n".join(lines)


def main() -> None:
    output_path = repo_root / "docs" / "TOOLS.md"
    content = generate_tools_markdown()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")
    print(f"Successfully generated {output_path}")


if __name__ == "__main__":
    main()
