"""Correctness, layering, and structural invariants for core modules and utils."""

from __future__ import annotations

import ast
from copy import copy
from pathlib import Path
import re
import zipfile

import openpyxl
from openpyxl.styles import Font, GradientFill
import pytest

from excel_mcp.core._io import workbook_session
from excel_mcp.core.calculations import apply_formula
from excel_mcp.core.chart import create_chart_in_sheet
from excel_mcp.core.data import read_excel_range
from excel_mcp.core.data_validation import add_data_validation
from excel_mcp.core.formatting import format_range
from excel_mcp.core.pivot import _get_combinations, create_pivot_table
from excel_mcp.core.ranges import (
    copy_range,
    copy_range_operation,
    delete_range_operation,
    merge_range,
)
from excel_mcp.core.workbook import get_workbook_info
from excel_mcp.exceptions.exception_core import PivotError, RangeError, ValidationError
from excel_mcp.utils.cell_utils import parse_cell_range, validate_cell_reference

ROOT = Path(__file__).resolve().parents[1]
UTILS_DIR = ROOT / "src" / "excel_mcp" / "utils"
CORE_DIR = ROOT / "src" / "excel_mcp" / "core"


def test_utils_has_zero_openpyxl_or_core_imports() -> None:
    violations: list[str] = []
    for py_file in sorted(UTILS_DIR.glob("*.py")):
        tree = ast.parse(py_file.read_text(encoding="utf-8"), filename=str(py_file))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    if alias.name == "openpyxl" or alias.name.startswith(
                        ("openpyxl.", "excel_mcp.core")
                    ):
                        violations.append(f"{py_file.name}:{node.lineno}:{alias.name}")
            elif isinstance(node, ast.ImportFrom):
                mod = node.module or ""
                if mod == "openpyxl" or mod.startswith(("openpyxl.", "excel_mcp.core")):
                    violations.append(f"{py_file.name}:{node.lineno}:{mod}")
    assert violations == []


def test_pivot_and_chart_functions_under_150_lines() -> None:
    oversized: list[str] = []
    for mod_name in ("pivot.py", "chart.py", "ranges.py"):
        path = CORE_DIR / mod_name
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef | ast.AsyncFunctionDef):
                end_lineno = getattr(node, "end_lineno", node.lineno)
                length = end_lineno - node.lineno + 1
                if length >= 150:
                    oversized.append(f"{mod_name}:{node.name}={length}")
    assert oversized == []


def test_inverted_column_ranges_and_bounds_rejected(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="precede"):
        parse_cell_range("XFD1:A1048576")
    with pytest.raises(ValueError, match="precede"):
        parse_cell_range("C1", "B5")

    assert validate_cell_reference("ZZZZ99999999") is False
    assert validate_cell_reference("α1") is False

    target = tmp_path / "inverted.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = "ok"
    wb.save(target)
    wb.close()

    with pytest.raises(RangeError, match="precede"):
        merge_range(str(target), "Sheet1", "XFD1", "A1048576")

    fmt_res = format_range(str(target), "Sheet1", "XFD1", "A1048576", merge_cells=True)
    assert fmt_res["status"] == "error"

    with zipfile.ZipFile(target, "r") as zin:
        items = [(info, zin.read(info.filename)) for info in zin.infolist()]
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename == "xl/worksheets/sheet1.xml":
                data = data.replace(
                    b"</worksheet>",
                    b'<mergeCells count="1"><mergeCell ref="XFD1:A1048576"/></mergeCells></worksheet>',
                )
            zout.writestr(info.filename, data)

    with pytest.raises(ValidationError, match="mergeCell area limit"):
        with workbook_session(target, mode="read"):
            pass


def test_pivot_correctness_keys_bool_sheet_limit_and_cartesian_bound(
    tmp_path: Path,
) -> None:
    path = tmp_path / "pivot_correctness.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    long_name = "Sales_Analytics_Sheet_Over_31"
    ws.title = long_name
    ws.append(["Region", "Year", "Active", "Amount"])
    ws.append(["Region", 2024, True, 100])
    ws.append(["Region", 2024, 1, 50])
    ws.append(["West", 2025, False, 200])
    ws.append(["West", 2025, True, 50])
    wb.save(path)
    wb.close()

    res = create_pivot_table(
        str(path),
        long_name,
        "A1:D5",
        rows=["region", "year"],
        values=["amount"],
        agg_func="sum",
    )
    assert "error" not in res
    pivot_name = res["details"]["pivot_sheet"]
    assert len(pivot_name) <= 31

    wb_out = openpyxl.load_workbook(path)
    pws = wb_out[pivot_name]
    rows_out = [
        tuple(pws.cell(row=r, column=c).value for c in range(1, 4))
        for r in range(2, pws.max_row + 1)
    ]
    wb_out.close()
    assert ("Region", 2024, 150) in rows_out
    assert ("West", 2025, 250) in rows_out

    res_bool = create_pivot_table(
        str(path),
        long_name,
        "A1:D5",
        rows=["active"],
        values=["amount"],
        agg_func="sum",
    )
    assert "error" not in res_bool
    wb_bool = openpyxl.load_workbook(path)
    pws_bool = wb_bool[res_bool["details"]["pivot_sheet"]]
    bool_rows = [
        (pws_bool.cell(row=r, column=1).value, pws_bool.cell(row=r, column=2).value)
        for r in range(2, pws_bool.max_row + 1)
    ]
    wb_bool.close()
    assert any(type(k) is bool and k is False and v == 200 for k, v in bool_rows)
    assert any(type(k) is bool and k is True and v == 150 for k, v in bool_rows)
    assert any(type(k) is int and k == 1 and v == 50 for k, v in bool_rows)

    with pytest.raises(PivotError, match="combination count"):
        _get_combinations(
            {"A": list(range(300)), "B": list(range(300))}, max_combos=50_000
        )


def test_ranges_isolation_self_copy_gradient_fill_bounds_and_style_clear(
    tmp_path: Path,
) -> None:
    path = tmp_path / "ranges_test.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Grid"
    ws.append(["A1", "B1", "C1"])
    ws.append(["A2", "B2", "C2"])
    ws.append(["A3", "B3", "C3"])
    ws["A1"].fill = GradientFill(stop=("FF0000", "00FF00"))
    ws["B1"].font = Font(bold=True)
    wb.save(path)
    wb.close()

    res_del_up = delete_range_operation(
        str(path), "Grid", "B1", "B1", shift_direction="up"
    )
    assert res_del_up["status"] == "success"
    wb_up = openpyxl.load_workbook(path)
    ws_up = wb_up["Grid"]
    assert [ws_up[f"B{r}"].value for r in (1, 2, 3)] == ["B2", "B3", None]
    assert ws_up["B1"].font.bold is False
    assert [ws_up[f"A{r}"].value for r in (1, 2, 3)] == ["A1", "A2", "A3"]
    assert [ws_up[f"C{r}"].value for r in (1, 2, 3)] == ["C1", "C2", "C3"]
    wb_up.close()

    res_del_left = delete_range_operation(
        str(path), "Grid", "A2", "A2", shift_direction="left"
    )
    assert res_del_left["status"] == "success"
    wb_left = openpyxl.load_workbook(path)
    ws_left = wb_left["Grid"]
    assert [ws_left.cell(row=2, column=c).value for c in (1, 2, 3)] == [
        "B3",
        "C2",
        None,
    ]
    assert [ws_left.cell(row=1, column=c).value for c in (1, 2, 3)] == [
        "A1",
        "B2",
        "C1",
    ]
    wb_left.close()

    res_copy = copy_range_operation(str(path), "Grid", "A1", "A1", "D1")
    assert res_copy["status"] == "success"
    wb_grad = openpyxl.load_workbook(path)
    ws_grad = wb_grad["Grid"]
    assert isinstance(copy(ws_grad["D1"].fill), GradientFill)
    assert ws_grad["D1"].fill.type == "linear"

    ws_self = wb_grad.create_sheet("SelfCopy")
    ws_self["A1"], ws_self["B1"] = 1, 2
    ws_self["A2"], ws_self["B2"] = 3, 4
    copy_range(ws_self, ws_self, "A1:B2", "B2")
    assert (
        ws_self["B2"].value,
        ws_self["C2"].value,
        ws_self["B3"].value,
        ws_self["C3"].value,
    ) == (1, 2, 3, 4)

    with pytest.raises(RangeError, match="column limit"):
        copy_range(ws_self, ws_self, "A1:B1", "XFD1")
    with pytest.raises(RangeError, match="row limit"):
        copy_range(ws_self, ws_self, "A1:A2", "A1048576")
    wb_grad.close()

    with pytest.raises(RangeError, match="merging"):
        merge_range(str(path), "Grid", "A1", "XFD1048576")


def test_cross_sheet_quoted_chart_places_chart_on_target_sheet(tmp_path: Path) -> None:
    path = tmp_path / "cross_chart.xlsx"
    wb = openpyxl.Workbook()
    summary = wb.active
    assert summary is not None
    summary.title = "Summary"
    sales = wb.create_sheet("Sales Data")
    sales.append(["Month", "Revenue"])
    sales.append(["Jan", 100])
    sales.append(["Feb", 200])
    wb.save(path)
    wb.close()

    res = create_chart_in_sheet(
        filename=str(path),
        sheet_name="Summary",
        data_range="'Sales Data'!$A$1:$B$3",
        chart_type="bar",
        target_cell="B2",
        title="Monthly Revenue",
    )
    assert res["status"] == "success"

    wb_after = openpyxl.load_workbook(path)
    assert len(wb_after["Summary"]._charts) == 1
    assert len(wb_after["Sales Data"]._charts) == 0
    wb_after.close()


def test_data_validation_custom_empty_formula2_and_whitespace_formula1(
    tmp_path: Path,
) -> None:
    path = tmp_path / "custom_dv.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = "Value"
    wb.save(path)
    wb.close()

    res = add_data_validation(
        filename=str(path),
        cell_range="A2:A10",
        validation_type="custom",
        formula1="=A2>0",
        formula2="",
        sheet_name="Sheet1",
    )
    assert res["operator"] is None

    with pytest.raises(ValidationError, match="formula1"):
        add_data_validation(
            filename=str(path),
            cell_range="A2:A10",
            validation_type="custom",
            formula1="   ",
            sheet_name="Sheet1",
        )

    wb_after = openpyxl.load_workbook(path)
    dv = wb_after["Sheet1"].data_validations.dataValidation[0]
    assert dv.type == "custom"
    assert dv.operator is None
    assert dv.formula1 == "=A2>0"
    wb_after.close()


def test_apply_formula_strips_leading_whitespace_before_equals(tmp_path: Path) -> None:
    path = tmp_path / "formula_ws.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 10
    ws["A2"] = 20
    wb.save(path)
    wb.close()

    res = apply_formula(str(path), "Sheet1", "A3", " =SUM(A1:A2)")
    assert res["formula"] == "=SUM(A1:A2)"


def test_read_and_info_succeed_when_worksheet_lacks_dimension_tag(
    tmp_path: Path,
) -> None:
    path = tmp_path / "no_dimension.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws.append(["Name", "Score"])
    ws.append(["Alice", 95])
    wb.save(path)
    wb.close()

    with zipfile.ZipFile(path, "r") as zin:
        items = [(info, zin.read(info.filename)) for info in zin.infolist()]
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename == "xl/worksheets/sheet1.xml":
                data = re.sub(rb"<dimension\b[^>]*/?>", b"", data)
            zout.writestr(info.filename, data)

    rows = read_excel_range(str(path), "Sheet1", "A1", None)
    assert rows == [["Name", "Score"], ["Alice", 95]]
    info = get_workbook_info(str(path), include_ranges=True)
    assert info["sheets"] == ["Sheet1"]
