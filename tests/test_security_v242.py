"""Security hardening regression suite for path, formula, archive, search, and range guards."""

from __future__ import annotations

import asyncio
from pathlib import Path
from unittest.mock import patch
import zipfile

import openpyxl
import pytest
import yaml

from excel_mcp.core import search as search_module
from excel_mcp.core._io import (
    MAX_XLSX_HYPERLINKS,
    MAX_XLSX_MERGE_CELLS,
    workbook_session,
)
from excel_mcp.core.formatting import _normalize_argb, format_range
from excel_mcp.core.search import find_cells_in_workbook, profile_worksheet
from excel_mcp.core.structure import (
    MAX_APPEND_CELLS,
    MAX_DIMENSION_AMOUNT,
    MAX_EXCEL_ROWS,
    append_rows_to_sheet,
    create_excel_table,
    delete_worksheet_dimension,
    insert_worksheet_dimension,
    sort_worksheet_range,
)
from excel_mcp.exceptions.exception_core import ValidationError
from excel_mcp.tools.compound_tools import excel_data
from excel_mcp.utils.file_utils import list_excel_files_in_directory
from excel_mcp.utils.paths import validate_and_resolve_path, validate_path_safety
from excel_mcp.utils.validation_utils import (
    validate_cell_value_secure,
    validate_formula,
    validate_formula_secure,
)


def _make_workbook(path: Path, sheet_title: str = "Sheet1") -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = sheet_title
    ws["A1"] = "Header"
    wb.save(path)
    wb.close()
    return path


def _inject_sheet_xml(
    path: Path, snippet: bytes, *, target_member: str = "xl/worksheets/sheet1.xml"
) -> None:
    wb = openpyxl.Workbook()
    wb.save(path)
    wb.close()
    with zipfile.ZipFile(path, "r") as zin:
        items = [(info, zin.read(info.filename)) for info in zin.infolist()]
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename == "xl/worksheets/sheet1.xml":
                data = data.replace(b"</worksheet>", snippet + b"</worksheet>")
                zout.writestr(target_member, data)
            else:
                zout.writestr(info.filename, data)


def test_mixed_slash_and_leading_space_unc_paths_rejected(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_MCP_DIRS", str(tmp_path))
    for bad in (
        r"\\evil\share\a.xlsx",
        "//evil/share/a.xlsx",
        r"\/evil\share\a.xlsx",
        r"/\evil\share\a.xlsx",
        r"  \/evil\share\a.xlsx",
        "  //evil/share/a.xlsx",
    ):
        err = validate_path_safety(bad)
        assert err is not None, f"Expected UNC rejection for {bad!r}"
        resolved, res_err = validate_and_resolve_path(bad)
        assert resolved is None
        assert res_err is not None


def test_conditional_formatting_rejects_formula_injection_and_normalizes_argb(
    tmp_path: Path,
) -> None:
    target = _make_workbook(tmp_path / "cf.xlsx")
    wb = openpyxl.load_workbook(target)
    ws = wb["Sheet1"]
    ws["B1"] = 10
    wb.save(target)
    wb.close()

    assert _normalize_argb("#FF0000") == "FFFF0000"
    assert _normalize_argb("FF0000") == "FFFF0000"
    assert _normalize_argb("80FF0000") == "80FF0000"
    assert _normalize_argb("ff00aa") == "FFFF00AA"
    with pytest.raises(ValueError, match="Invalid hex color"):
        _normalize_argb("ZZZZZZ")

    fmt_ok = format_range(
        str(target),
        "Sheet1",
        "B1",
        font_color="#FF0000",
        bg_color="FF0000",
    )
    assert fmt_ok["status"] == "success"
    wb_after = openpyxl.load_workbook(target)
    cell_after = wb_after["Sheet1"]["B1"]
    assert cell_after.font.color.rgb == "FFFF0000"
    assert cell_after.fill.fgColor.rgb == "FFFF0000"
    wb_after.close()

    for bad_cf in (
        {"type": "formula", "params": {"formula": ['=CALL("Kernel32","WinExec","J")']}},
        {"type": "formula", "params": {"formula": [None]}},
        {
            "type": "formula",
            "params": {"formula": '=WEBSERVICE("http://evil.example")'},
        },
        {
            "type": "cell_is",
            "params": {"operator": "equal", "formula": ["+cmd|'/C calc'!A0"]},
        },
        {
            "type": "color_scale",
            "params": {
                "start_type": "formula",
                "start_value": '=CALL("lib","fn")',
                "start_color": "FF0000",
                "end_type": "num",
                "end_value": 100,
                "end_color": "00FF00",
            },
        },
        {
            "type": "data_bar",
            "params": {
                "start_type": "num",
                "start_value": 0,
                "end_type": "formula",
                "end_value": '=EXEC("calc")',
                "color": "0000FF",
            },
        },
    ):
        res = format_range(str(target), "Sheet1", "B1", conditional_format=bad_cf)
        assert res["status"] == "error"

    valid_cf = format_range(
        str(target),
        "Sheet1",
        "B1",
        conditional_format={
            "type": "formula",
            "params": {"formula": ["=B1>5"], "fill": {"fgColor": None}},
        },
    )
    assert valid_cf["status"] == "success"


def test_formula_blocklist_parentheses_os_boundary_and_external_udfs() -> None:
    for blocked in (
        '=IMAGE("https://evil.example/img.png")',
        '=IMPORTXML("https://evil.example","//a")',
        '=IMPORTDATA("https://evil.example/data.csv")',
        '=SQL.REQUEST("DSN=db")',
        '=EVALUATE("1+1")',
        "=_xll.CALL('lib','fn')",
        "=_xll.custom_addin(1)",
        "=_xludf.evil_func(A1)",
        "=_xlfn._xll.custom_addin(1)",
        "=_xlws._xludf.evil_func(A1)",
        "='\\\\evil\\share\\[a.xlsx]Sheet1'!A1",
        "='\\/evil\\share\\[a.xlsx]Sheet1'!A1",
    ):
        ok_sec, _ = validate_formula_secure(blocked)
        assert ok_sec is False, (
            f"Expected validate_formula_secure rejection for {blocked}"
        )
        assert validate_formula(blocked)[0] is False

    for allowed in (
        '="Photos.jpg"',
        '="Nos. 1-5"',
        '=IF(A1>0, "a)", "b(")',
        '=IF(A1>0, "say ""(hi)""", "ok")',
    ):
        ok_sec, msg = validate_formula_secure(allowed)
        assert ok_sec is True, f"Expected allowance for {allowed}: {msg}"


def test_prefixed_cell_value_bypasses_rejected_and_prose_preserved() -> None:
    for unsafe in (
        '+(HYPERLINK("http://evil.example","x"))',
        '+1+2+CALL("lib","fn")',
        '+-EXEC("calc")',
        '+=CALL("lib","fn")',
        '+"a"&WEBSERVICE("http://evil.example")',
        '@(IMAGE("http://evil.example/x.png"))',
        "+'\\\\evil\\share\\[a.xlsx]S1'!A1",
        "+1+'\\\\evil\\share\\[a.xlsx]S1'!A1",
        "+1+'//evil/share/[a.xlsx]S1'!A1",
        '+1+CALL ("lib","fn")',
        '+CALL ("lib","fn")',
        '+A1+CALL ("lib","fn")',
        '-B2*EXEC ("calc")',
    ):
        ok, _ = validate_cell_value_secure(unsafe)
        assert ok is False, f"Expected cell value rejection for {unsafe}"

    for prose in (
        "- Please call (555) 123-4567",
        "- Open (9am-5pm)",
        "- https://example.com",
        "- Status (WIP",
        "+1-800-555-0199",
        "+1 (555) 123-4567",
        "@alice",
    ):
        ok, msg = validate_cell_value_secure(prose)
        assert ok is True, f"Expected prose allowance for {prose}: {msg}"


def test_archive_validation_rejects_aggregate_ratio_and_xml_amplification(
    tmp_path: Path,
) -> None:
    agg_bomb = tmp_path / "agg_bomb.xlsx"
    chunk = b"A" * (600 * 1024)
    with zipfile.ZipFile(agg_bomb, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("xl/part1.xml", chunk)
        zf.writestr("xl/part2.xml", chunk)
    with pytest.raises(ValidationError, match="aggregate compression ratio"):
        with workbook_session(agg_bomb, mode="read"):
            pass

    merge_area_bomb = tmp_path / "merge_area.xlsx"
    _inject_sheet_xml(
        merge_area_bomb,
        b'<mergeCells count="1"><mergeCell ref = "A1:XFD1048576"/></mergeCells>',
    )
    with pytest.raises(ValidationError, match="mergeCell area limit"):
        with workbook_session(merge_area_bomb, mode="read"):
            pass

    custom_path_bomb = tmp_path / "custom_path_merge.xlsx"
    _inject_sheet_xml(
        custom_path_bomb,
        b'<mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells>',
        target_member="xl/sheets/sheet1.xml",
    )
    with pytest.raises(ValidationError, match="mergeCell area limit"):
        with workbook_session(custom_path_bomb, mode="read"):
            pass

    merge_count_bomb = tmp_path / "merge_count.xlsx"
    many_merges = "".join(
        f'<mergeCell ref="A{i}:B{i}"/>' for i in range(1, MAX_XLSX_MERGE_CELLS + 2)
    ).encode("ascii")
    _inject_sheet_xml(
        merge_count_bomb, b"<mergeCells>" + many_merges + b"</mergeCells>"
    )
    with pytest.raises(ValidationError, match="mergeCell limit"):
        with workbook_session(merge_count_bomb, mode="read"):
            pass

    hyperlink_bomb = tmp_path / "hyperlink_count.xlsx"
    many_links = (b'<hyperlink ref="A1" r:id="rId1"/>') * (MAX_XLSX_HYPERLINKS + 1)
    _inject_sheet_xml(hyperlink_bomb, b"<hyperlinks>" + many_links + b"</hyperlinks>")
    with pytest.raises(ValidationError, match="hyperlink limit"):
        with workbook_session(hyperlink_bomb, mode="read"):
            pass


def test_search_redos_nested_alternation_and_profile_limits(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    target = tmp_path / "search_profile.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Data"
    ws.append(["ColA", "ColB", "ColC"])
    for idx in range(1, 5):
        ws.append([f"alpha_{idx}", idx, idx * 10])
    wb.save(target)
    wb.close()

    for bad_regex in ("((a|a))+$", "(((a|b)))+$", "((a|a)?)+$"):
        with pytest.raises(ValidationError, match="nested quantifiers"):
            find_cells_in_workbook(str(target), query=bad_regex, use_regex=True)

    with pytest.raises(ValidationError, match="max_results"):
        find_cells_in_workbook(str(target), query="alpha", max_results=0)
    with pytest.raises(ValidationError, match="max_results"):
        find_cells_in_workbook(str(target), query="alpha", max_results=True)

    ok_alt = find_cells_in_workbook(
        str(target), query="^(alpha_1|alpha_2)$", use_regex=True
    )
    assert ok_alt["total_matches"] == 2

    with pytest.raises(ValidationError, match="sample_rows"):
        profile_worksheet(str(target), sheet_name="Data", sample_rows=0)
    with pytest.raises(ValidationError, match="sample_rows"):
        profile_worksheet(str(target), sheet_name="Data", sample_rows=True)

    exact = profile_worksheet(str(target), sheet_name="Data", sample_rows=2)
    assert exact["sampled_data_rows"] == 2
    assert "scan_truncated" not in exact

    monkeypatch.setattr(search_module, "MAX_PROFILE_ROWS", 2)
    monkeypatch.setattr(search_module, "MAX_PROFILE_COLUMNS", 2)
    capped = profile_worksheet(str(target), sheet_name="Data")
    assert capped["sampled_data_rows"] == 2
    assert capped["total_columns"] == 2
    assert capped["scan_truncated"] is True
    assert capped["columns_truncated"] is True


def test_structure_bounds_relative_key_column_and_sparse_append(
    tmp_path: Path,
) -> None:
    target = tmp_path / "structure.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Grid"
    ws["B1"] = "Item"
    ws["C1"] = "Rank"
    ws["D1"] = "Note"
    ws["B2"], ws["C2"], ws["D2"] = "First", 30, "x"
    ws["B3"], ws["C3"], ws["D3"] = "Second", 10, "y"
    ws["B4"], ws["C4"], ws["D4"] = "Third", 20, "z"
    wb.save(target)
    wb.close()

    with pytest.raises(ValidationError, match="sorting"):
        sort_worksheet_range(
            str(target), "A1:XFD1048576", key_column=1, sheet_name="Grid"
        )
    with pytest.raises(ValidationError, match="boolean"):
        sort_worksheet_range(str(target), "B1:D4", key_column=True, sheet_name="Grid")
    with pytest.raises(ValidationError, match="table creation"):
        create_excel_table(str(target), "A1:XFD1048576", sheet_name="Grid")
    with pytest.raises(ValidationError, match="maximum allowed"):
        insert_worksheet_dimension(
            str(target), "row", 1, amount=MAX_DIMENSION_AMOUNT + 1
        )
    with pytest.raises(ValidationError, match="maximum allowed"):
        delete_worksheet_dimension(
            str(target), "column", 1, amount=MAX_DIMENSION_AMOUNT + 1
        )
    with pytest.raises(ValidationError, match="start_column"):
        append_rows_to_sheet(str(target), [[1]], start_column=True)

    res = sort_worksheet_range(
        str(target),
        "B1:D4",
        key_column=2,
        ascending=True,
        has_header=True,
        sheet_name="Grid",
    )
    assert res["sorted_rows"] == 3
    wb_sorted = openpyxl.load_workbook(target)
    ws_sorted = wb_sorted["Grid"]
    assert [ws_sorted[f"C{r}"].value for r in (2, 3, 4)] == [10, 20, 30]
    assert [ws_sorted[f"B{r}"].value for r in (2, 3, 4)] == [
        "Second",
        "Third",
        "First",
    ]
    wb_sorted.close()

    sparse = tmp_path / "sparse_append.xlsx"
    wb_s = openpyxl.Workbook()
    ws_s = wb_s.active
    assert ws_s is not None
    ws_s["A1"] = "head"
    ws_s["Z10000"] = "tail"
    wb_s.save(sparse)
    wb_s.close()

    appended = append_rows_to_sheet(str(sparse), [["new_row"]])
    assert appended["start_row"] == 10001

    with pytest.raises(ValidationError, match="Append payload"):
        append_rows_to_sheet(str(sparse), [[1] * (MAX_APPEND_CELLS + 1)])
    with pytest.raises(ValidationError, match="Excel row limit"):
        append_rows_to_sheet(str(sparse), [[] for _ in range(MAX_EXCEL_ROWS + 1)])


def test_list_directory_excludes_symlinks_and_compound_rejects_conflicting_filename(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_MCP_DIRS", str(tmp_path))
    real_file = tmp_path / "real.xlsx"
    _make_workbook(real_file)

    outside_file = tmp_path.parent / "outside_secret.xlsx"
    _make_workbook(outside_file)
    link_file = tmp_path / "escape_link.xlsx"
    try:
        link_file.symlink_to(outside_file)
        listed_real = list_excel_files_in_directory()
        assert [item["filename"] for item in listed_real] == ["real.xlsx"]
    except OSError:

        class _FakeEntry:
            def __init__(self, name: str, is_link: bool) -> None:
                self.name = name
                self.path = str(tmp_path / name)
                self._is_link = is_link

            def is_symlink(self) -> bool:
                return self._is_link

            def is_file(self, *, follow_symlinks: bool = True) -> bool:
                return not self._is_link or follow_symlinks

            def stat(self, *, follow_symlinks: bool = True):
                return real_file.stat()

        class _ScanCtx:
            def __init__(self, entries: list[_FakeEntry]) -> None:
                self._entries = entries

            def __enter__(self) -> list[_FakeEntry]:
                return self._entries

            def __exit__(self, *args: object) -> None:
                return None

        fake_entries = [
            _FakeEntry("real.xlsx", is_link=False),
            _FakeEntry("escape_link.xlsx", is_link=True),
        ]
        with patch("os.scandir", return_value=_ScanCtx(fake_entries)):
            listed = list_excel_files_in_directory()
        assert [item["filename"] for item in listed] == ["real.xlsx"]

    mismatch = asyncio.run(
        excel_data(
            "read",
            str(real_file),
            {
                "filename": str(tmp_path / "other.xlsx"),
                "sheet_name": "Sheet1",
                "start_cell": "A1",
            },
        )
    )
    assert mismatch["status"] == "error"
    assert "Conflicting 'filename'" in mismatch["message"]

    params_mismatch = asyncio.run(
        excel_data(
            "sort_range",
            None,
            {
                "filename": str(real_file),
                "file_path": str(tmp_path / "other.xlsx"),
                "range_string": "A1:A1",
                "key_column": 1,
            },
        )
    )
    assert params_mismatch["status"] == "error"
    assert "Conflicting 'filename' and 'file_path'" in params_mismatch["message"]

    matched = asyncio.run(
        excel_data(
            "read",
            str(real_file),
            {"filename": str(real_file), "sheet_name": "Sheet1", "start_cell": "A1"},
        )
    )
    assert matched["status"] == "success"


def test_ci_workflow_hardens_version_bump_guard() -> None:
    ci_data = yaml.safe_load(
        (
            Path(__file__).resolve().parents[1] / ".github" / "workflows" / "ci.yml"
        ).read_text(encoding="utf-8")
    )
    steps = ci_data["jobs"]["test"]["steps"]
    checkout = next(s for s in steps if s.get("name") == "Checkout repository")
    assert checkout["with"]["fetch-depth"] == 0

    guard = next(
        s for s in steps if s.get("name") == "Verify PR does not bump version directly"
    )
    assert "BASE_REF" in guard["env"]
    assert "set -euo pipefail" in guard["run"]
