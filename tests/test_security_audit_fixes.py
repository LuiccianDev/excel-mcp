from __future__ import annotations

import asyncio
from pathlib import Path
import zipfile

import openpyxl
import pytest

from excel_mcp.core import search as search_module
from excel_mcp.core._io import (
    MAX_XLSX_ENTRIES,
    atomic_save,
    workbook_session,
)
from excel_mcp.core.batch import batch_write_matrix, build_single_pass_workbook
from excel_mcp.core.calculations import (
    FormulaSecurityError,
    apply_formula,
    validate_cell_value_secure,
    validate_formula_secure,
)
from excel_mcp.core.data import (
    MAX_READ_CELLS,
    PREVIEW_MAX_ROWS,
    read_excel_range,
    write_data,
)
from excel_mcp.core.data_validation import add_data_validation
from excel_mcp.core.formatting import format_range
from excel_mcp.core.search import find_cells_in_workbook
from excel_mcp.core.structure import append_rows_to_sheet
from excel_mcp.exceptions.exception_core import RangeError, ValidationError
from excel_mcp.tools.excel_tools import create_excel_workbook
from excel_mcp.tools.formulas_excel_tools import (
    apply_formula_excel,
    validate_formula_syntax,
)
from excel_mcp.tools.register_tools import register_all_tools
from excel_mcp.utils import metrics as metrics_module
from excel_mcp.utils.paths import resolve_allowed_directories, validate_path_safety


def _make_workbook(path: Path, sheet_title: str = "Sheet1") -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = sheet_title
    ws["A1"] = "hello"
    wb.save(path)
    wb.close()
    return path


def test_apply_formula_and_tools_enforce_secure_validation(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_MCP_DIRS", str(tmp_path))
    target = _make_workbook(tmp_path / "formulas.xlsx")

    for payload in (
        '=CALL("Kernel32","WinExec","JCJ","calc.exe",1)',
        "=cmd|'/C calc'!A0",
        "=msexcel|a!b",
        "=_xlfn.CALL('lib','fn')",
        "=_xlfn.HYPERLINK('http://evil.example')",
    ):
        with pytest.raises(FormulaSecurityError):
            apply_formula(str(target), "Sheet1", "B1", payload)

        res_apply = asyncio.run(
            apply_formula_excel(str(target), "Sheet1", "B1", payload)
        )
        assert res_apply["status"] == "error"

        res_val = asyncio.run(
            validate_formula_syntax(str(target), "Sheet1", "B1", payload)
        )
        assert res_val["status"] == "error"


def test_cell_value_secure_blocks_bypasses_and_allows_prose(tmp_path: Path) -> None:
    for unsafe in (
        '+HYPERLINK("http://evil.example","x")',
        '-WEBSERVICE("http://evil.example")',
        '@INDIRECT("A1")',
        '+1+CALL("lib","fn")',
        "+msexcel|a!b",
        "=_xlfn.CALL('a','b')",
    ):
        ok, _ = validate_cell_value_secure(unsafe)
        assert ok is False, f"Expected rejection for {unsafe}"

    ok_xlfn, _ = validate_formula_secure("=_xlfn.WEBSERVICE('http://evil.example')")
    assert ok_xlfn is False

    for safe in (
        "+1-800-555-0199",
        "-50.25",
        "+1 (555) 123-4567",
        "@alice",
        "+SUM(A1:A10)",
        "- Please call (555) 123-4567",
        "- Open (9am-5pm)",
        "- https://example.com",
        "- Status (WIP",
    ):
        ok, _ = validate_cell_value_secure(safe)
        assert ok is True, f"Expected safe allowance for {safe}"

    target = _make_workbook(tmp_path / "cell_sinks.xlsx")
    with pytest.raises(FormulaSecurityError):
        write_data(target, "Sheet1", [['+HYPERLINK("http://evil.example","x")']])
    with pytest.raises(FormulaSecurityError):
        batch_write_matrix(str(target), "Sheet1", "A1", [["=_xlfn.CALL('a','b')"]])
    with pytest.raises(FormulaSecurityError):
        append_rows_to_sheet(target, [["+msexcel|a!b"]], sheet_name="Sheet1")


def test_data_validation_and_batch_headers_enforce_formula_security(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_MCP_DIRS", str(tmp_path))
    target = _make_workbook(tmp_path / "dv.xlsx")

    with pytest.raises(FormulaSecurityError):
        add_data_validation(
            filename=str(target),
            cell_range="A2:A5",
            validation_type="custom",
            operator="equal",
            formula1='=WEBSERVICE("http://evil.example")',
            sheet_name="Sheet1",
        )

    with pytest.raises(FormulaSecurityError):
        add_data_validation(
            filename=str(target),
            cell_range="A2:A5",
            validation_type="list",
            formula1='WEBSERVICE("http://evil.example")',
            sheet_name="Sheet1",
        )

    with pytest.raises(FormulaSecurityError):
        add_data_validation(
            filename=str(target),
            cell_range="A2:A5",
            validation_type="whole",
            operator="between",
            formula1="1",
            formula2="CALL('a','b')",
            sheet_name="Sheet1",
        )

    with pytest.raises(FormulaSecurityError):
        build_single_pass_workbook(
            str(tmp_path / "bad_title.xlsx"),
            sheets=[{"name": "S1", "title": "=cmd|'/C calc'!A0", "rows": [[1]]}],
        )

    with pytest.raises(FormulaSecurityError):
        build_single_pass_workbook(
            str(tmp_path / "bad_header.xlsx"),
            sheets=[{"name": "S1", "headers": ['=EXEC("calc")'], "rows": [[1]]}],
        )

    with pytest.raises(FormulaSecurityError):
        build_single_pass_workbook(
            str(tmp_path / "bad_dv.xlsx"),
            sheets=[
                {
                    "name": "S1",
                    "rows": [[1]],
                    "data_validations": [
                        {
                            "range": "A1:A1",
                            "type": "custom",
                            "operator": "equal",
                            "formula1": "=CALL('lib','fn')",
                        }
                    ],
                }
            ],
        )


def test_archive_validation_and_atomic_save_size_cap(tmp_path: Path) -> None:
    bomb_ratio_zip = tmp_path / "ratio_bomb.xlsx"
    with zipfile.ZipFile(bomb_ratio_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("xl/sharedStrings.xml", b"A" * (2 * 1024 * 1024))

    with pytest.raises(ValidationError, match="compression ratio"):
        with workbook_session(bomb_ratio_zip, mode="read"):
            pass

    many_entries_zip = tmp_path / "many_entries.xlsx"
    with zipfile.ZipFile(many_entries_zip, "w") as zf:
        for idx in range(MAX_XLSX_ENTRIES + 1):
            zf.writestr(f"e_{idx}.txt", b"x")

    with pytest.raises(ValidationError, match="entries"):
        with workbook_session(many_entries_zip, mode="read"):
            pass

    wb = openpyxl.Workbook()
    out = tmp_path / "capped.xlsx"
    with pytest.raises(ValidationError, match="exceeds maximum allowed size"):
        atomic_save(wb, out, max_file_size=10)
    assert not out.exists()


def test_range_bounds_and_preview_clamp(tmp_path: Path) -> None:
    target = tmp_path / "ranges.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Data"
    for r in range(1, PREVIEW_MAX_ROWS + 25):
        ws.cell(row=r, column=1, value=r)
    wb.save(target)
    wb.close()

    fmt_err = format_range(str(target), "Data", "A1", "XFD1048576", bold=True)
    assert fmt_err["status"] == "error"
    assert "too large" in fmt_err["message"].lower()

    preview = read_excel_range(str(target), "Data", "A1", None, preview_only=True)
    assert len(preview) == PREVIEW_MAX_ROWS

    sparse_target = tmp_path / "sparse.xlsx"
    sparse_wb = openpyxl.Workbook()
    sparse_ws = sparse_wb.active
    assert sparse_ws is not None
    sparse_ws.title = "Sparse"
    sparse_ws["A1"] = "top"
    sparse_ws["Z15000"] = "far"
    sparse_wb.save(sparse_target)
    sparse_wb.close()

    sparse_preview = read_excel_range(
        str(sparse_target), "Sparse", "A1", None, preview_only=True
    )
    assert sparse_preview == [["top"]]

    with pytest.raises(RangeError, match=str(MAX_READ_CELLS)):
        read_excel_range(str(sparse_target), "Sparse", "A1", "Z15000")


def test_search_regex_guards_truncation_and_metrics_rotation(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    target = tmp_path / "search.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws.append(["alpha", "beta"])
    ws.append(["gamma", "delta"])
    wb.save(target)
    wb.close()

    for bad_regex in ("(a+)+$", "((a)+)+$", "((?:a+)+)$", "(a{1,})+$", "(a|aa)+$"):
        with pytest.raises(ValidationError, match="nested quantifiers"):
            find_cells_in_workbook(str(target), query=bad_regex, use_regex=True)

    with pytest.raises(ValidationError, match="maximum length"):
        find_cells_in_workbook(str(target), query="a" * 300, use_regex=True)

    with pytest.raises(ValidationError, match="Invalid regex pattern"):
        find_cells_in_workbook(str(target), query="[unclosed", use_regex=True)

    ok_regex = find_cells_in_workbook(
        str(target), query="^(alpha|beta)$", use_regex=True
    )
    assert ok_regex["total_matches"] == 2

    monkeypatch.setattr(search_module, "MAX_SEARCH_CELLS", 2)
    truncated = find_cells_in_workbook(str(target), query="delta")
    assert truncated["scan_truncated"] is True
    assert truncated["limit_reached"] is False

    log_path = tmp_path / "metrics.jsonl"
    monkeypatch.setenv("EXCEL_MCP_METRICS_LOG", str(log_path))
    monkeypatch.setattr(metrics_module, "DEFAULT_MAX_METRICS_LOG_BYTES", 50)

    metrics_module.record_tool_metric("excel_data", "read", 1.5, "success")
    assert log_path.exists()
    assert log_path.stat().st_size >= 50
    metrics_module.record_tool_metric("excel_data", "read", 2.0, "success")

    rotated = tmp_path / "metrics.jsonl.1"
    assert rotated.exists()
    assert log_path.exists()


def test_path_placeholders_reserved_space_dir_check_and_annotations(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_MCP_DIRS", f"${{user_config:directory}}, {tmp_path}")
    assert resolve_allowed_directories() == [str(tmp_path.resolve())]

    assert validate_path_safety("CON .xlsx") is not None

    monkeypatch.setenv("EXCEL_MCP_DIRS", str(tmp_path))
    sub = tmp_path / "folder.xlsx"
    sub.mkdir()
    res = asyncio.run(create_excel_workbook(str(sub)))
    assert res["status"] == "error"
    assert "directory" in res["message"].lower()

    class _Recorder:
        def __init__(self) -> None:
            self.annotations: dict[str, object] = {}

        def tool(self, name: str, annotations: object = None):
            self.annotations[name] = annotations
            return lambda fn: fn

    compound_rec = _Recorder()
    register_all_tools(compound_rec, mode="compound")
    assert (
        getattr(
            compound_rec.annotations["excel_build_workbook"], "destructive_hint", False
        )
        is True
    )
    assert (
        getattr(compound_rec.annotations["excel_workbook"], "destructive_hint", False)
        is True
    )

    full_rec = _Recorder()
    register_all_tools(full_rec, mode="full")
    assert (
        getattr(full_rec.annotations["excel_build_workbook"], "destructive_hint", False)
        is True
    )
    assert (
        getattr(
            full_rec.annotations["excel_create_workbook"], "destructive_hint", False
        )
        is True
    )
