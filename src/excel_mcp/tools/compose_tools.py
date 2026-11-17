"""Composed domain-grouped MCP tools for compact tool-surface mode.

Groups the 30 granular Excel operations into 5 action-dispatched domain tools
(`excel_workbook`, `excel_data`, `excel_format`, `excel_formula`, `excel_analytics`)
while leaving all underlying tool modules and single-pass tools intact.
"""

from __future__ import annotations

import inspect
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any, Literal

from excel_mcp.tools import (
    analysis_tools,
    content_tools,
    excel_tools,
    format_tools,
    formulas_excel_tools,
    graphics_tools,
    page_tools,
    protection_tools,
    structure_tools,
    validation_tools,
)


@dataclass(frozen=True)
class ComposeActionSpec:
    """Metadata and target handler for a composed tool action."""

    handler: Callable[..., Any]
    summary: str
    read_only: bool = False
    full_name: str = ""
    destructive: bool = False

    def __post_init__(self) -> None:
        if self.read_only and self.destructive:
            raise ValueError(
                f"ComposeActionSpec '{self.full_name}' cannot be both read_only and destructive."
            )
        if not self.full_name.startswith("excel_"):
            raise ValueError(
                f"ComposeActionSpec full_name '{self.full_name}' must start with 'excel_'."
            )


COMPOSE_REGISTRY: dict[str, dict[str, ComposeActionSpec]] = {
    "workbook": {
        "create": ComposeActionSpec(
            excel_tools.create_excel_workbook,
            "Create a new empty .xlsx workbook.",
            full_name="excel_create_workbook",
        ),
        "list": ComposeActionSpec(
            excel_tools.list_excel_documents,
            "List .xlsx files in allowed directories.",
            read_only=True,
            full_name="excel_list_documents",
        ),
        "metadata": ComposeActionSpec(
            format_tools.get_workbook_metadata,
            "Inspect workbook sheets, dimensions, and optional used ranges.",
            read_only=True,
            full_name="excel_get_workbook_metadata",
        ),
        "create_sheet": ComposeActionSpec(
            excel_tools.create_excel_worksheet,
            "Add a new worksheet to an existing workbook.",
            full_name="excel_create_worksheet",
        ),
        "copy_sheet": ComposeActionSpec(
            format_tools.copy_worksheet,
            "Duplicate an existing worksheet within a workbook.",
            full_name="excel_copy_worksheet",
        ),
        "delete_sheet": ComposeActionSpec(
            format_tools.delete_worksheet,
            "Delete a worksheet from a workbook.",
            full_name="excel_delete_worksheet",
            destructive=True,
        ),
        "rename_sheet": ComposeActionSpec(
            format_tools.rename_worksheet,
            "Rename an existing worksheet.",
            full_name="excel_rename_worksheet",
        ),
    },
    "data": {
        "read": ComposeActionSpec(
            content_tools.read_data_from_excel,
            "Read cell values or formatted metadata from a worksheet range.",
            read_only=True,
            full_name="excel_read_data",
        ),
        "write": ComposeActionSpec(
            content_tools.write_data_to_excel,
            "Write a 2D matrix of values to a worksheet starting at start_cell.",
            full_name="excel_write_data",
        ),
        "append_rows": ComposeActionSpec(
            structure_tools.append_rows,
            "Append rows after the last populated row of a worksheet.",
            full_name="excel_append_rows",
        ),
        "copy_range": ComposeActionSpec(
            format_tools.copy_range,
            "Copy a cell range (values and formatting) to another location.",
            full_name="excel_copy_range",
        ),
        "delete_range": ComposeActionSpec(
            format_tools.delete_range,
            "Delete a cell range and shift remaining cells.",
            full_name="excel_delete_range",
            destructive=True,
        ),
        "merge_cells": ComposeActionSpec(
            format_tools.merge_cells,
            "Merge a rectangular cell range.",
            full_name="excel_merge_cells",
        ),
        "unmerge_cells": ComposeActionSpec(
            format_tools.unmerge_cells,
            "Unmerge a previously merged cell range.",
            full_name="excel_unmerge_cells",
        ),
        "insert_rows_columns": ComposeActionSpec(
            structure_tools.insert_rows_columns,
            "Insert rows or columns at a given index.",
            full_name="excel_insert_rows_columns",
        ),
        "delete_rows_columns": ComposeActionSpec(
            structure_tools.delete_rows_columns,
            "Delete rows or columns starting at a given index.",
            full_name="excel_delete_rows_columns",
            destructive=True,
        ),
        "sort_range": ComposeActionSpec(
            structure_tools.sort_range,
            "Sort a range by one or more columns.",
            full_name="excel_sort_range",
        ),
        "validate_range": ComposeActionSpec(
            format_tools.validate_excel_range,
            "Validate that a cell range is within worksheet bounds.",
            read_only=True,
            full_name="excel_validate_range",
        ),
        "find_cells": ComposeActionSpec(
            analysis_tools.find_cells,
            "Search for matching cell values or regex patterns.",
            read_only=True,
            full_name="excel_find_cells",
        ),
        "profile_sheet": ComposeActionSpec(
            analysis_tools.profile_sheet,
            "Compute column types, non-null counts, and summary statistics.",
            read_only=True,
            full_name="excel_profile_sheet",
        ),
    },
    "format": {
        "format_range": ComposeActionSpec(
            format_tools.format_range_excel,
            "Apply font, fill, border, alignment, and number formatting to a range.",
            full_name="excel_format_range",
        ),
        "format_sheet": ComposeActionSpec(
            structure_tools.format_sheet,
            "Apply sheet-level styling, column widths, row heights, or freeze panes.",
            full_name="excel_format_sheet",
        ),
        "create_table": ComposeActionSpec(
            structure_tools.create_table,
            "Convert a cell range into a native Excel Table.",
            full_name="excel_create_table",
        ),
        "add_data_validation": ComposeActionSpec(
            validation_tools.add_data_validation,
            "Attach dropdown, numeric, or custom data validation rules.",
            full_name="excel_add_data_validation",
        ),
        "protect_sheet": ComposeActionSpec(
            protection_tools.protect_sheet,
            "Enable or disable worksheet protection.",
            full_name="excel_protect_sheet",
        ),
        "configure_page_setup": ComposeActionSpec(
            page_tools.configure_page_setup,
            "Configure print orientation, paper size, margins, and scaling.",
            full_name="excel_configure_page_setup",
        ),
    },
    "formula": {
        "validate": ComposeActionSpec(
            formulas_excel_tools.validate_formula_syntax,
            "Validate Excel formula syntax without modifying the workbook.",
            read_only=True,
            full_name="excel_validate_formula_syntax",
        ),
        "apply": ComposeActionSpec(
            formulas_excel_tools.apply_formula_excel,
            "Validate and write a formula into a target cell.",
            full_name="excel_apply_formula",
        ),
    },
    "analytics": {
        "create_chart": ComposeActionSpec(
            graphics_tools.create_chart,
            "Create a native bar, line, pie, or scatter chart.",
            full_name="excel_create_chart",
        ),
        "create_pivot_table": ComposeActionSpec(
            graphics_tools.create_pivot_table,
            "Create a pivot table summary from a source range.",
            full_name="excel_create_pivot_table",
        ),
    },
}


def _normalize_result(result: Any) -> dict[str, Any]:
    """Normalize dict, string, or scalar tool returns into a structured dict."""
    if isinstance(result, dict):
        if "error" in result and "status" not in result:
            return {"status": "error", "message": str(result["error"]), **result}
        return result
    if isinstance(result, str):
        if result.startswith("Error:"):
            return {"status": "error", "message": result}
        return {"status": "success", "message": result}
    return {"status": "success", "data": result}


async def dispatch_composed_action(
    group: str,
    action: str,
    filename: str | None = None,
    params: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Validate and dispatch a composed tool action directly to its handler."""
    group_actions = COMPOSE_REGISTRY.get(group)
    if group_actions is None:
        return {
            "status": "error",
            "message": f"Unknown compose group '{group}'.",
        }

    spec = group_actions.get(action)
    if spec is None:
        return {
            "status": "error",
            "message": f"Unknown action '{action}' for excel_{group}.",
            "available_actions": sorted(group_actions.keys()),
        }

    sig = inspect.signature(spec.handler)
    accepts_kwargs = any(
        p.kind == inspect.Parameter.VAR_KEYWORD for p in sig.parameters.values()
    )

    call_kwargs = dict(params or {})
    if not accepts_kwargs:
        unknown = sorted(k for k in call_kwargs if k not in sig.parameters)
        if unknown:
            return {
                "status": "error",
                "message": (
                    f"Unknown parameter(s) for '{group}.{action}': {', '.join(unknown)}"
                ),
                "accepted_parameters": list(sig.parameters.keys()),
            }

    if filename is not None:
        if "filename" in sig.parameters and call_kwargs.get("filename") is None:
            call_kwargs["filename"] = filename
        elif "file_path" in sig.parameters and call_kwargs.get("file_path") is None:
            call_kwargs["file_path"] = filename

    if (
        "filename" in sig.parameters
        and "file_path" in sig.parameters
        and call_kwargs.get("filename") is None
        and call_kwargs.get("file_path") is None
    ):
        return {
            "status": "error",
            "message": f"Missing required parameter 'filename' for '{group}.{action}'.",
            "accepted_parameters": list(sig.parameters.keys()),
        }

    missing = [
        name
        for name, param in sig.parameters.items()
        if param.default is inspect.Parameter.empty
        and param.kind
        not in (inspect.Parameter.VAR_POSITIONAL, inspect.Parameter.VAR_KEYWORD)
        and name not in call_kwargs
    ]
    if missing:
        return {
            "status": "error",
            "message": (
                f"Missing required parameter(s) for '{group}.{action}': "
                f"{', '.join(missing)}"
            ),
            "accepted_parameters": list(sig.parameters.keys()),
        }

    try:
        result = spec.handler(**call_kwargs)
        if inspect.isawaitable(result):
            result = await result
    except Exception as exc:
        return {"status": "error", "message": str(exc)}
    return _normalize_result(result)


async def excel_workbook(
    action: Literal[
        "create",
        "list",
        "metadata",
        "create_sheet",
        "copy_sheet",
        "delete_sheet",
        "rename_sheet",
    ],
    filename: str | None = None,
    params: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Manage Excel workbooks and worksheets in one compact tool.

    Actions:
        - `create`: Create a new workbook (`filename`).
        - `list`: List `.xlsx` files in allowed directories (no `filename` required).
        - `metadata`: Inspect sheets and dimensions (`filename`, optional `include_ranges`).
        - `create_sheet`: Add a worksheet (`filename`, `sheet_name`).
        - `copy_sheet`: Copy a worksheet (`filename`, `source_sheet`, `target_sheet`).
        - `delete_sheet`: Delete a worksheet (`filename`, `sheet_name`).
        - `rename_sheet`: Rename a worksheet (`filename`, `old_name`, `new_name`).
    """
    return await dispatch_composed_action("workbook", action, filename, params)


async def excel_data(
    action: Literal[
        "read",
        "write",
        "append_rows",
        "copy_range",
        "delete_range",
        "merge_cells",
        "unmerge_cells",
        "insert_rows_columns",
        "delete_rows_columns",
        "sort_range",
        "validate_range",
        "find_cells",
        "profile_sheet",
    ],
    filename: str | None = None,
    params: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Read, write, transform, search, and profile worksheet data.

    Actions:
        - `read`, `write`, `append_rows`, `copy_range`, `delete_range`
        - `merge_cells`, `unmerge_cells`, `insert_rows_columns`, `delete_rows_columns`
        - `sort_range`, `validate_range`, `find_cells`, `profile_sheet`
    """
    return await dispatch_composed_action("data", action, filename, params)


async def excel_format(
    action: Literal[
        "format_range",
        "format_sheet",
        "create_table",
        "add_data_validation",
        "protect_sheet",
        "configure_page_setup",
    ],
    filename: str | None = None,
    params: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Apply cell/sheet formatting, tables, data validation, protection, and page setup.

    Actions:
        - `format_range`, `format_sheet`, `create_table`
        - `add_data_validation`, `protect_sheet`, `configure_page_setup`
    """
    return await dispatch_composed_action("format", action, filename, params)


async def excel_formula(
    action: Literal["validate", "apply"],
    filename: str | None = None,
    params: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Validate or apply Excel formulas to worksheet cells.

    Actions:
        - `validate`: Check formula syntax without writing (`filename`, `sheet_name`, `cell`, `formula`).
        - `apply`: Validate and write formula (`filename`, `sheet_name`, `cell`, `formula`).
    """
    return await dispatch_composed_action("formula", action, filename, params)


async def excel_analytics(
    action: Literal["create_chart", "create_pivot_table"],
    filename: str | None = None,
    params: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Create native Excel charts and pivot tables.

    Actions:
        - `create_chart`: Add a bar, line, pie, or scatter chart.
        - `create_pivot_table`: Summarize a data range into a pivot table.
    """
    return await dispatch_composed_action("analytics", action, filename, params)
