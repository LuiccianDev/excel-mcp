"""Compound domain-grouped MCP tools for compact tool-surface mode.

Groups the 30 granular Excel operations into 5 action-dispatched domain tools
(`excel_workbook`, `excel_data`, `excel_format`, `excel_formula`, `excel_analytics`)
while leaving all underlying tool modules and single-pass tools intact.
"""

from __future__ import annotations

import inspect
from collections.abc import Callable
from dataclasses import dataclass
from functools import cached_property
from typing import Any, Literal

from excel_mcp.exceptions.exception_tools import ExceptionTool

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
class CompoundActionSpec:
    """Metadata and target handler for a compound tool action."""

    handler: Callable[..., Any]
    summary: str
    read_only: bool = False
    full_name: str = ""
    destructive: bool = False

    def __post_init__(self) -> None:
        if self.read_only and self.destructive:
            raise ValueError(
                f"CompoundActionSpec '{self.full_name}' cannot be both read_only and destructive."
            )
        if not self.full_name.startswith("excel_"):
            raise ValueError(
                f"CompoundActionSpec full_name '{self.full_name}' must start with 'excel_'."
            )

    @cached_property
    def signature_meta(
        self,
    ) -> tuple[tuple[str, ...], tuple[str, ...], bool, bool]:
        """Cache handler parameter names, required params, kwargs flag, and filename flag."""
        sig = inspect.signature(self.handler)
        params = sig.parameters
        accepted = tuple(params.keys())
        required = tuple(
            name
            for name, param in params.items()
            if param.default is inspect.Parameter.empty
            and param.kind
            not in (inspect.Parameter.VAR_POSITIONAL, inspect.Parameter.VAR_KEYWORD)
        )
        accepts_kwargs = any(
            p.kind == inspect.Parameter.VAR_KEYWORD for p in params.values()
        )
        accepts_filename = "filename" in params
        return (
            accepted,
            required,
            accepts_kwargs,
            accepts_filename,
        )


COMPOUND_REGISTRY: dict[str, dict[str, CompoundActionSpec]] = {
    "workbook": {
        "create": CompoundActionSpec(
            excel_tools.create_excel_workbook,
            "Create a new empty .xlsx workbook.",
            full_name="excel_create_workbook",
            destructive=True,
        ),
        "list": CompoundActionSpec(
            excel_tools.list_excel_documents,
            "List .xlsx files in allowed directories.",
            read_only=True,
            full_name="excel_list_documents",
        ),
        "metadata": CompoundActionSpec(
            format_tools.get_workbook_metadata,
            "Inspect workbook sheets, dimensions, and optional used ranges.",
            read_only=True,
            full_name="excel_get_workbook_metadata",
        ),
        "create_sheet": CompoundActionSpec(
            excel_tools.create_excel_worksheet,
            "Add a new worksheet to an existing workbook.",
            full_name="excel_create_worksheet",
        ),
        "copy_sheet": CompoundActionSpec(
            format_tools.copy_worksheet,
            "Duplicate an existing worksheet within a workbook.",
            full_name="excel_copy_worksheet",
        ),
        "delete_sheet": CompoundActionSpec(
            format_tools.delete_worksheet,
            "Delete a worksheet from a workbook.",
            full_name="excel_delete_worksheet",
            destructive=True,
        ),
        "rename_sheet": CompoundActionSpec(
            format_tools.rename_worksheet,
            "Rename an existing worksheet.",
            full_name="excel_rename_worksheet",
        ),
    },
    "data": {
        "read": CompoundActionSpec(
            content_tools.read_data_from_excel,
            "Read cell values or formatted metadata from a worksheet range.",
            read_only=True,
            full_name="excel_read_data",
        ),
        "write": CompoundActionSpec(
            content_tools.write_data_to_excel,
            "Write a 2D matrix of values to a worksheet starting at start_cell.",
            full_name="excel_write_data",
        ),
        "append_rows": CompoundActionSpec(
            structure_tools.append_rows,
            "Append rows after the last populated row of a worksheet.",
            full_name="excel_append_rows",
        ),
        "copy_range": CompoundActionSpec(
            format_tools.copy_range,
            "Copy a cell range (values and formatting) to another location.",
            full_name="excel_copy_range",
        ),
        "delete_range": CompoundActionSpec(
            format_tools.delete_range,
            "Delete a cell range and shift remaining cells.",
            full_name="excel_delete_range",
            destructive=True,
        ),
        "merge_cells": CompoundActionSpec(
            format_tools.merge_cells,
            "Merge a rectangular cell range.",
            full_name="excel_merge_cells",
        ),
        "unmerge_cells": CompoundActionSpec(
            format_tools.unmerge_cells,
            "Unmerge a previously merged cell range.",
            full_name="excel_unmerge_cells",
        ),
        "insert_rows_columns": CompoundActionSpec(
            structure_tools.insert_rows_columns,
            "Insert rows or columns at a given index.",
            full_name="excel_insert_rows_columns",
        ),
        "delete_rows_columns": CompoundActionSpec(
            structure_tools.delete_rows_columns,
            "Delete rows or columns starting at a given index.",
            full_name="excel_delete_rows_columns",
            destructive=True,
        ),
        "sort_range": CompoundActionSpec(
            structure_tools.sort_range,
            "Sort a range by one or more columns.",
            full_name="excel_sort_range",
        ),
        "validate_range": CompoundActionSpec(
            format_tools.validate_excel_range,
            "Validate that a cell range is within worksheet bounds.",
            read_only=True,
            full_name="excel_validate_range",
        ),
        "find_cells": CompoundActionSpec(
            analysis_tools.find_cells,
            "Search for matching cell values or regex patterns.",
            read_only=True,
            full_name="excel_find_cells",
        ),
        "profile_sheet": CompoundActionSpec(
            analysis_tools.profile_sheet,
            "Compute column types, non-null counts, and summary statistics.",
            read_only=True,
            full_name="excel_profile_sheet",
        ),
    },
    "format": {
        "format_range": CompoundActionSpec(
            format_tools.format_range_excel,
            "Apply font, fill, border, alignment, and number formatting to a range.",
            full_name="excel_format_range",
        ),
        "format_sheet": CompoundActionSpec(
            structure_tools.format_sheet,
            "Apply sheet-level styling, column widths, row heights, or freeze panes.",
            full_name="excel_format_sheet",
        ),
        "create_table": CompoundActionSpec(
            structure_tools.create_table,
            "Convert a cell range into a native Excel Table.",
            full_name="excel_create_table",
        ),
        "add_data_validation": CompoundActionSpec(
            validation_tools.add_data_validation,
            "Attach dropdown, numeric, or custom data validation rules.",
            full_name="excel_add_data_validation",
        ),
        "protect_sheet": CompoundActionSpec(
            protection_tools.protect_sheet,
            "Enable or disable worksheet protection.",
            full_name="excel_protect_sheet",
        ),
        "configure_page_setup": CompoundActionSpec(
            page_tools.configure_page_setup,
            "Configure print orientation, paper size, margins, and scaling.",
            full_name="excel_configure_page_setup",
        ),
    },
    "formula": {
        "validate": CompoundActionSpec(
            formulas_excel_tools.validate_formula_syntax,
            "Validate Excel formula syntax without modifying the workbook.",
            read_only=True,
            full_name="excel_validate_formula_syntax",
        ),
        "apply": CompoundActionSpec(
            formulas_excel_tools.apply_formula_excel,
            "Validate and write a formula into a target cell.",
            full_name="excel_apply_formula",
        ),
    },
    "analytics": {
        "create_chart": CompoundActionSpec(
            graphics_tools.create_chart,
            "Create a native bar, line, pie, or scatter chart.",
            full_name="excel_create_chart",
        ),
        "create_pivot_table": CompoundActionSpec(
            graphics_tools.create_pivot_table,
            "Create a pivot table summary from a source range.",
            full_name="excel_create_pivot_table",
        ),
    },
}


def _normalize_and_validate_kwargs(
    group: str,
    action: str,
    filename: str | None,
    params: dict[str, Any] | None,
    spec: CompoundActionSpec,
) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
    (
        accepted_params,
        required_params,
        accepts_kwargs,
        accepts_filename,
    ) = spec.signature_meta

    call_kwargs = dict(params or {})
    if (
        "range_string" in call_kwargs
        and "cell_range" in call_kwargs
        and (
            "range_string" not in accepted_params or "cell_range" not in accepted_params
        )
    ):
        return None, {
            "status": "error",
            "message": "Provide either 'range_string' or 'cell_range', not both.",
        }
    for source_key, target_key in (
        ("range_string", "cell_range"),
        ("cell_range", "range_string"),
    ):
        if (
            source_key in call_kwargs
            and source_key not in accepted_params
            and target_key in accepted_params
        ):
            call_kwargs[target_key] = call_kwargs.pop(source_key)
            break

    if not accepts_kwargs:
        unknown = sorted(k for k in call_kwargs if k not in accepted_params)
        if unknown:
            return None, {
                "status": "error",
                "message": (
                    f"Unknown parameter(s) for '{group}.{action}': {', '.join(unknown)}"
                ),
                "accepted_parameters": list(accepted_params),
            }

    param_filename = call_kwargs.get("filename")
    param_file_path = call_kwargs.get("file_path")
    if (
        filename is not None
        and param_filename is not None
        and param_filename != filename
    ):
        return None, {
            "status": "error",
            "message": (
                f"Conflicting 'filename' values between top-level argument and params "
                f"for '{group}.{action}'."
            ),
        }
    if (
        filename is not None
        and param_file_path is not None
        and param_file_path != filename
    ):
        return None, {
            "status": "error",
            "message": (
                f"Conflicting 'filename' and 'file_path' values for '{group}.{action}'."
            ),
        }
    if (
        param_filename is not None
        and param_file_path is not None
        and param_filename != param_file_path
    ):
        return None, {
            "status": "error",
            "message": (
                f"Conflicting 'filename' and 'file_path' values in params for '{group}.{action}'."
            ),
        }

    if filename is not None and accepts_filename and param_filename is None:
        call_kwargs["filename"] = filename

    missing = [name for name in required_params if name not in call_kwargs]
    if missing:
        return None, {
            "status": "error",
            "message": (
                f"Missing required parameter(s) for '{group}.{action}': "
                f"{', '.join(missing)}"
            ),
            "accepted_parameters": list(accepted_params),
        }

    return call_kwargs, None


def dispatch_compound_action(
    group: str,
    action: str,
    filename: str | None = None,
    params: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Validate and dispatch a compound tool action directly to its handler."""
    group_actions = COMPOUND_REGISTRY.get(group)
    if group_actions is None:
        return {
            "status": "error",
            "message": f"Unknown compound group '{group}'.",
        }

    spec = group_actions.get(action)
    if spec is None:
        return {
            "status": "error",
            "message": f"Unknown action '{action}' for excel_{group}.",
            "available_actions": sorted(group_actions.keys()),
        }

    call_kwargs, err = _normalize_and_validate_kwargs(
        group, action, filename, params, spec
    )
    if err is not None or call_kwargs is None:
        return err or {"status": "error", "message": "Invalid compound parameters."}

    try:
        result: dict[str, Any] = spec.handler(**call_kwargs)
    except Exception as exc:
        return ExceptionTool.handle_error(
            exc,
            filename=call_kwargs.get("filename") or filename,
            operation=f"excel_{group}.{action}",
        )
    return result


def excel_workbook(
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
    return dispatch_compound_action("workbook", action, filename, params)


def excel_data(
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

    Actions (`params` keys):
        - `read`: (`sheet_name`, `start_cell`, optional `end_cell`, `preview_only`)
        - `write`: (`sheet_name`, `data`, optional `start_cell`)
        - `append_rows`: (`rows`, optional `sheet_name`, `start_column`)
        - `copy_range`: (`sheet_name`, `source_start`, `source_end`, `target_start`, optional `target_sheet`)
        - `delete_range`: (`sheet_name`, `start_cell`, optional `end_cell`, `shift_direction`)
        - `merge_cells`: (`sheet_name`, `start_cell`, `end_cell`)
        - `unmerge_cells`: (`sheet_name`, `start_cell`, `end_cell`)
        - `insert_rows_columns`: (`dimension`: `'row'|'column'`, `index`, optional `amount`, `sheet_name`)
        - `delete_rows_columns`: (`dimension`: `'row'|'column'`, `index`, optional `amount`, `sheet_name`)
        - `sort_range`: (`range_string`, `key_column`, optional `ascending`, `has_header`, `sheet_name`)
        - `validate_range`: (`sheet_name`, `start_cell`, optional `end_cell`)
        - `find_cells`: (`query`, optional `sheet_name`, `match_case`, `exact_match`, `use_regex`, `max_results`)
        - `profile_sheet`: (optional `sheet_name`, `sample_rows`)
    """
    return dispatch_compound_action("data", action, filename, params)


def excel_format(
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

    Actions (`params` keys):
        - `format_range`: (`sheet_name`, `start_cell`, optional `end_cell`, `bold`, `italic`, `font_size`, `bg_color`, `border_style`, `number_format`, `alignment`, `wrap_text`, `conditional_format`)
        - `format_sheet`: (optional `sheet_name`, `freeze_panes`, `auto_filter`, `show_grid_lines`, `column_widths`)
        - `create_table`: (`range_string`, optional `table_name`, `style_name`, `sheet_name`)
        - `add_data_validation`: (`cell_range`, `validation_type`, `formula1`, optional `operator`, `formula2`, `allow_blank`, `prompt_title`, `prompt`, `error_title`, `error_message`, `sheet_name`)
        - `protect_sheet`: (optional `enabled`, `password`, `sheet_name`, `allow_*` flags)
        - `configure_page_setup`: (optional `orientation`, `paper_size`, `fit_to_width`, `fit_to_height`, `print_area`, `print_title_rows`, `header`, `footer`, `center_horizontally`, `sheet_name`)
    """
    return dispatch_compound_action("format", action, filename, params)


def excel_formula(
    action: Literal["validate", "apply"],
    filename: str | None = None,
    params: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Validate or apply Excel formulas to worksheet cells.

    Actions:
        - `validate`: Check formula syntax without writing (`filename`, `sheet_name`, `cell`, `formula`).
        - `apply`: Validate and write formula (`filename`, `sheet_name`, `cell`, `formula`).
    """
    return dispatch_compound_action("formula", action, filename, params)


def excel_analytics(
    action: Literal["create_chart", "create_pivot_table"],
    filename: str | None = None,
    params: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Create native Excel charts and pivot tables.

    Actions (`params` keys):
        - `create_chart`: (`sheet_name`, `data_range`, `chart_type`, `target_cell`, optional `title`, `x_axis`, `y_axis`, `style`)
        - `create_pivot_table`: (`sheet_name`, `data_range`, `rows`, `values`, optional `columns`, `agg_func`)
    """
    return dispatch_compound_action("analytics", action, filename, params)
