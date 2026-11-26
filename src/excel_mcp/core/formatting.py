import re
from typing import Any, Final, Literal, TypedDict

from openpyxl.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
    IconSetRule,
)
from openpyxl.styles import (
    Alignment,
    Border,
    Color,
    Font,
    PatternFill,
    Protection,
    Side,
)
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.core._io import atomic_save, workbook_session
from excel_mcp.utils.cell_utils import (
    assert_range_cell_limit,
    parse_cell_range,
    validate_cell_reference,
)
from excel_mcp.utils.validation_utils import validate_formula_secure

MAX_FORMAT_CELLS: Final[int] = 100_000

BorderStyle = Literal[
    "dashDot",
    "dashDotDot",
    "dashed",
    "dotted",
    "double",
    "hair",
    "medium",
    "mediumDashDot",
    "mediumDashDotDot",
    "mediumDashed",
    "slantDashDot",
    "thick",
    "thin",
    "none",
    None,
]

UnderlineStyle = Literal[
    "single", "double", "singleAccounting", "doubleAccounting", "none"
]


class FontArgs(TypedDict, total=False):
    bold: bool
    italic: bool
    underline: UnderlineStyle | None
    size: int
    color: Color


_CONDITIONAL_RULE_BUILDERS: Final[dict[str, Any]] = {
    "color_scale": ColorScaleRule,
    "data_bar": DataBarRule,
    "icon_set": IconSetRule,
    "formula": FormulaRule,
    "cell_is": CellIsRule,
}


_HEX_COLOR_RE: Final[re.Pattern[str]] = re.compile(r"^[0-9A-F]{6}([0-9A-F]{2})?$")
_CF_FORMULA_SCALAR_KEYS: Final[tuple[str, ...]] = (
    "start_value",
    "mid_value",
    "end_value",
    "min_value",
    "max_value",
)


def _normalize_argb(color: str) -> str:
    if not isinstance(color, str):
        raise ValueError(f"Invalid hex color {color!r}: expected hex string.")
    cleaned = color.strip().lstrip("#").upper()
    if not _HEX_COLOR_RE.fullmatch(cleaned):
        raise ValueError(
            f"Invalid hex color '{color}': expected 6-digit RRGGBB or 8-digit AARRGGBB."
        )
    return f"FF{cleaned}" if len(cleaned) == 6 else cleaned


def _validate_conditional_format_formulas(
    rule_type: str, params: dict[str, Any]
) -> str | None:
    candidates: list[str] = []
    formula_val = params.get("formula")
    if isinstance(formula_val, str):
        if formula_val:
            candidates.append(formula_val)
    elif isinstance(formula_val, list | tuple):
        candidates.extend(item for item in formula_val if isinstance(item, str))

    if rule_type == "formula" and not any(item.strip() for item in candidates):
        return (
            "Conditional format formula rule requires a non-empty 'formula' parameter."
        )

    for key in _CF_FORMULA_SCALAR_KEYS:
        val = params.get(key)
        if isinstance(val, str) and val != "":
            candidates.append(val)

    values_val = params.get("values")
    if isinstance(values_val, list | tuple):
        candidates.extend(
            item for item in values_val if isinstance(item, str) and item != ""
        )

    for expr in candidates:
        ok, msg = validate_formula_secure(expr)
        if not ok:
            return f"Conditional format formula security validation failed: {msg}"
    return None


def _build_font(
    bold: bool,
    italic: bool,
    underline: bool,
    font_size: int | None,
    font_color: str | None,
) -> Font:
    font_args: FontArgs = {
        "bold": bold,
        "italic": italic,
        "underline": "single" if underline else None,
    }
    if font_size is not None:
        font_args["size"] = font_size
    if font_color is not None:
        font_args["color"] = Color(rgb=_normalize_argb(font_color))
    return Font(**font_args)


def _build_fill(bg_color: str | None) -> PatternFill | None:
    if bg_color is None:
        return None
    normalized = _normalize_argb(bg_color)
    return PatternFill(
        start_color=Color(rgb=normalized),
        end_color=Color(rgb=normalized),
        fill_type="solid",
    )


def _build_border(border_style: BorderStyle, border_color: str | None) -> Border | None:
    if border_style is None or border_style == "none":
        return None
    side = Side(
        style=border_style,
        color=Color(rgb=_normalize_argb(border_color or "000000")),
    )
    return Border(left=side, right=side, top=side, bottom=side)


def _apply_conditional_formatting(
    sheet: Worksheet, range_str: str, conditional_format: dict[str, Any]
) -> dict[str, Any] | None:
    rule_type = conditional_format.get("type")
    if not rule_type:
        return {
            "status": "error",
            "message": "Conditional format type not specified",
        }
    builder = _CONDITIONAL_RULE_BUILDERS.get(rule_type)
    if builder is None:
        return {
            "status": "error",
            "message": f"Invalid conditional format type: {rule_type}",
        }
    params = dict(conditional_format.get("params", {}))
    formula_err = _validate_conditional_format_formulas(rule_type, params)
    if formula_err is not None:
        return {"status": "error", "message": formula_err}
    if rule_type in ("cell_is", "formula") and isinstance(params.get("fill"), dict):
        raw_fg = params["fill"].get("fgColor")
        fill_color = _normalize_argb(raw_fg if raw_fg is not None else "FFC7CE")
        params["fill"] = PatternFill(
            start_color=fill_color,
            end_color=fill_color,
            fill_type="solid",
        )
    try:
        sheet.conditional_formatting.add(range_str, builder(**params))
    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to apply conditional formatting: {str(e)}",
        }
    return None


def format_range(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str | None = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: int | None = None,
    font_color: str | None = None,
    bg_color: str | None = None,
    border_style: BorderStyle = None,
    border_color: str | None = None,
    number_format: str | None = None,
    alignment: str | None = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: dict[str, Any] | None = None,
    conditional_format: dict[str, Any] | None = None,
) -> dict[str, Any]:
    if not validate_cell_reference(start_cell):
        return {
            "status": "error",
            "message": f"Invalid start cell reference: {start_cell}",
        }
    if end_cell and not validate_cell_reference(end_cell):
        return {"status": "error", "message": f"Invalid end cell reference: {end_cell}"}
    try:
        with workbook_session(
            filename, mode="upsert", initial_sheet=None, save=False
        ) as wb:
            if sheet_name not in wb.sheetnames:
                return {"status": "error", "message": f"Sheet '{sheet_name}' not found"}
            sheet = wb[sheet_name]
            try:
                start_row, start_col, end_row, end_col = parse_cell_range(
                    start_cell, end_cell
                )
            except ValueError as e:
                return {"status": "error", "message": f"Invalid cell range: {str(e)}"}
            eff_end_row = end_row if end_row is not None else start_row
            eff_end_col = end_col if end_col is not None else start_col
            try:
                assert_range_cell_limit(
                    start_row,
                    start_col,
                    eff_end_row,
                    eff_end_col,
                    MAX_FORMAT_CELLS,
                    ValueError,
                    label="formatting",
                )
            except ValueError as e:
                return {"status": "error", "message": str(e)}

            font = _build_font(bold, italic, underline, font_size, font_color)
            fill = _build_fill(bg_color)
            border = _build_border(border_style, border_color)
            align = (
                Alignment(horizontal=alignment, vertical="center", wrap_text=wrap_text)
                if alignment is not None or wrap_text
                else None
            )
            protect = Protection(**protection) if protection else None

            for row in range(start_row, eff_end_row + 1):
                for col in range(start_col, eff_end_col + 1):
                    cell = sheet.cell(row=row, column=col)
                    cell.font = font
                    if fill:
                        cell.fill = fill
                    if border:
                        cell.border = border
                    if align:
                        cell.alignment = align
                    if protect:
                        cell.protection = protect
                    if number_format:
                        cell.number_format = number_format

            if merge_cells and end_cell:
                try:
                    sheet.merge_cells(f"{start_cell}:{end_cell}")
                except ValueError as e:
                    return {
                        "status": "error",
                        "message": f"Failed to merge cells: {str(e)}",
                    }

            range_str = f"{start_cell}:{end_cell}" if end_cell else start_cell
            if conditional_format:
                cf_err = _apply_conditional_formatting(
                    sheet, range_str, conditional_format
                )
                if cf_err is not None:
                    return cf_err

            atomic_save(wb, filename)
            return {
                "status": "success",
                "message": f"Applied formatting to range {range_str}",
                "range": range_str,
            }
    except Exception as e:
        return {"status": "error", "message": str(e)}
