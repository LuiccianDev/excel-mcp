"""Tests for compact compound tool-surface mode."""

from __future__ import annotations

import asyncio
from pathlib import Path
from typing import Any

import pytest

from excel_mcp.tools.compound_tools import (
    COMPOUND_REGISTRY,
    excel_analytics,
    excel_data,
    excel_format,
    excel_formula,
    excel_workbook,
)
from excel_mcp.tools.register_tools import get_tool_mode, register_all_tools


class _ToolCollector:
    """Minimal FastMCP stand-in that records registered tool names."""

    def __init__(self) -> None:
        self.tools: dict[str, Any] = {}

    def tool(self, name: str, **_kwargs: Any) -> Any:
        def _decorator(fn: Any) -> Any:
            self.tools[name] = fn
            return fn

        return _decorator


def test_compound_mode_registers_eight_tools(monkeypatch: pytest.MonkeyPatch) -> None:
    """Default compound mode exposes 8 tools and EXCEL_MCP_TOOLS=full switches to 33."""
    monkeypatch.delenv("EXCEL_MCP_TOOLS", raising=False)
    assert get_tool_mode() == "compound"

    collector = _ToolCollector()
    register_all_tools(collector)

    assert set(collector.tools.keys()) == {
        "excel_build_workbook",
        "excel_get_workbook_outline",
        "excel_batch_write_data",
        "excel_workbook",
        "excel_data",
        "excel_format",
        "excel_formula",
        "excel_analytics",
    }

    monkeypatch.setenv("EXCEL_MCP_TOOLS", "full")
    assert get_tool_mode() == "full"
    full_collector = _ToolCollector()
    register_all_tools(full_collector)
    assert len(full_collector.tools) == 33


def test_compound_registry_covers_all_granular_tools() -> None:
    """Every non-single-pass canonical tool in full mode must exist in COMPOUND_REGISTRY."""
    full_collector = _ToolCollector()
    register_all_tools(full_collector, mode="full")

    single_pass = {
        "excel_build_workbook",
        "excel_get_workbook_outline",
        "excel_batch_write_data",
    }
    full_handlers = {
        fn for name, fn in full_collector.tools.items() if name not in single_pass
    }
    compound_handlers = {
        spec.handler
        for actions in COMPOUND_REGISTRY.values()
        for spec in actions.values()
    }
    assert compound_handlers == full_handlers


def test_compound_dispatcher_literals_match_registry() -> None:
    """Every compound dispatcher's Literal action annotation must match COMPOUND_REGISTRY."""
    from typing import get_args, get_type_hints
    from excel_mcp.tools import compound_tools

    for group, actions in COMPOUND_REGISTRY.items():
        fn = getattr(compound_tools, f"excel_{group}")
        hints = get_type_hints(fn)
        assert set(get_args(hints["action"])) == set(actions.keys())


def test_invalid_mode_raises_value_error() -> None:
    """Explicit invalid mode passed to register_all_tools must fail loudly."""
    collector = _ToolCollector()
    with pytest.raises(ValueError, match="Invalid tool mode"):
        register_all_tools(collector, mode="invalid")


def test_compound_tools_end_to_end(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Compound tools must execute real workbook, data, formula, and format flows."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    monkeypatch.setenv("EXCEL_MCP_DIRS", str(docs_dir))

    res_create = asyncio.run(excel_workbook("create", "sales.xlsx"))
    assert res_create.get("status") == "success"

    res_sheet = asyncio.run(
        excel_workbook("create_sheet", "sales.xlsx", {"sheet_name": "Q1"})
    )
    assert res_sheet.get("status") == "success"

    res_list = asyncio.run(excel_workbook("list"))
    assert res_list.get("status") == "success"

    res_meta = asyncio.run(excel_workbook("metadata", "sales.xlsx"))
    assert "Q1" in res_meta.get("sheets", [])

    res_write = asyncio.run(
        excel_data(
            "write",
            "sales.xlsx",
            {
                "sheet_name": "Q1",
                "data": [["Item", "Amount"], ["Ore", 100], ["Gold", 250]],
                "start_cell": "A1",
            },
        )
    )
    assert res_write.get("status") == "success"

    res_read = asyncio.run(
        excel_data("read", "sales.xlsx", {"sheet_name": "Q1", "start_cell": "A1"})
    )
    assert res_read.get("status") == "success"

    res_val = asyncio.run(
        excel_formula(
            "validate",
            "sales.xlsx",
            {"sheet_name": "Q1", "cell": "B4", "formula": "=SUM(B2:B3)"},
        )
    )
    assert res_val.get("status") == "success"

    res_apply = asyncio.run(
        excel_formula(
            "apply",
            "sales.xlsx",
            {"sheet_name": "Q1", "cell": "B4", "formula": "=SUM(B2:B3)"},
        )
    )
    assert res_apply.get("status") == "success"

    res_table = asyncio.run(
        excel_format(
            "create_table",
            "sales.xlsx",
            {"sheet_name": "Q1", "range_string": "A1:B3", "table_name": "SalesTbl"},
        )
    )
    assert res_table.get("status") == "success"


def test_compound_tools_reject_invalid_actions_and_params() -> None:
    """Invalid actions, unknown params, and missing required params must return errors."""
    bad_action = asyncio.run(excel_workbook("nonexistent"))  # type: ignore[arg-type]
    assert bad_action["status"] == "error"
    assert "available_actions" in bad_action

    bad_param = asyncio.run(excel_formula("validate", "a.xlsx", {"not_a_param": 1}))
    assert bad_param["status"] == "error"
    assert "Unknown parameter" in bad_param["message"]

    missing_param = asyncio.run(excel_formula("validate", "a.xlsx", {}))
    assert missing_param["status"] == "error"
    assert "Missing required parameter" in missing_param["message"]

    bad_chart = asyncio.run(excel_analytics("create_chart", None, {}))
    assert bad_chart["status"] == "error"

    bad_create = asyncio.run(excel_workbook("create", "payload.exe"))
    assert bad_create["status"] == "error"
    assert "message" in bad_create


def test_compound_metadata_and_exception_normalization(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Dict results without status are normalized and handler exceptions use ExceptionTool."""
    from excel_mcp.tools.compound_tools import CompoundActionSpec

    monkeypatch.setenv("EXCEL_MCP_DIRS", str(tmp_path))
    created = asyncio.run(excel_workbook("create", "meta.xlsx"))
    assert created["status"] == "success"

    meta = asyncio.run(excel_workbook("metadata", "meta.xlsx"))
    assert meta["status"] == "success"
    assert "sheets" in meta

    def _raising_handler(filename: str, sheet_name: str, cell: str, formula: str):
        raise RuntimeError("boom")

    orig = COMPOUND_REGISTRY["formula"]["validate"]
    monkeypatch.setitem(
        COMPOUND_REGISTRY["formula"],
        "validate",
        CompoundActionSpec(
            handler=_raising_handler,
            summary=orig.summary,
            read_only=orig.read_only,
            full_name=orig.full_name,
        ),
    )
    err = asyncio.run(
        excel_formula(
            "validate",
            "meta.xlsx",
            {"sheet_name": "Sheet1", "cell": "A1", "formula": "=1+1"},
        )
    )
    assert err["status"] == "error"
    assert err["error_type"] == "RuntimeError"
    assert err["operation"] == "excel_formula.validate"


def test_compound_dispatcher_annotations_reflect_group_actions() -> None:
    """Compound dispatcher annotations must reflect destructive vs write group actions."""
    recorded: dict[str, object] = {}

    class _AnnotatedMCP:
        def tool(self, name: str, annotations: object = None):
            recorded[name] = annotations
            return lambda fn: fn

    register_all_tools(_AnnotatedMCP(), mode="compound")
    assert getattr(recorded["excel_workbook"], "destructive_hint", False) is True
    assert getattr(recorded["excel_data"], "destructive_hint", False) is True
    assert getattr(recorded["excel_format"], "destructive_hint", False) is not True


def test_compound_range_alias_normalization(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """range_string alias must transparently normalize to cell_range across format actions."""
    from excel_mcp.tools.batch_tools import build_workbook

    monkeypatch.setenv("EXCEL_MCP_DIRS", str(tmp_path))
    target = tmp_path / "aliases.xlsx"
    built = build_workbook(
        filename=str(target),
        sheets=[
            {
                "name": "Data",
                "headers": ["Item", "Status"],
                "rows": [["Alpha", "Open"], ["Beta", "Closed"]],
            }
        ],
    )
    assert built["status"] == "success"

    table_res = asyncio.run(
        excel_format(
            "create_table",
            str(target),
            {"range_string": "A1:B3", "table_name": "OrdersTbl", "sheet_name": "Data"},
        )
    )
    assert table_res["status"] == "success"

    dv_res = asyncio.run(
        excel_format(
            "add_data_validation",
            str(target),
            {
                "range_string": "B2:B3",
                "validation_type": "list",
                "formula1": '"Open,Closed"',
                "sheet_name": "Data",
            },
        )
    )
    assert dv_res["status"] == "success"

    dual_err = asyncio.run(
        excel_format(
            "add_data_validation",
            str(target),
            {
                "range_string": "B2:B3",
                "cell_range": "A2:A3",
                "validation_type": "list",
                "formula1": '"Open,Closed"',
                "sheet_name": "Data",
            },
        )
    )
    assert dual_err["status"] == "error"
    assert "not both" in dual_err["message"]


def test_compound_chart_style_and_optional_delete_end_cell(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Verify excel_analytics accepts style on create_chart and delete_range allows omitting end_cell."""
    from excel_mcp.tools.batch_tools import build_workbook

    monkeypatch.setenv("EXCEL_MCP_DIRS", str(tmp_path))
    target = tmp_path / "analytics_delete.xlsx"
    built = build_workbook(
        filename=str(target),
        sheets=[
            {
                "name": "Metrics",
                "headers": ["Month", "Sales"],
                "rows": [["Jan", 10], ["Feb", 20], ["Temp", 99]],
            }
        ],
    )
    assert built["status"] == "success"

    chart_res = asyncio.run(
        excel_analytics(
            "create_chart",
            str(target),
            {
                "sheet_name": "Metrics",
                "data_range": "A1:B3",
                "chart_type": "bar",
                "target_cell": "D2",
                "title": "Sales",
                "style": {"show_legend": False},
            },
        )
    )
    assert chart_res["status"] == "success"

    del_res = asyncio.run(
        excel_data(
            "delete_range",
            str(target),
            {"sheet_name": "Metrics", "start_cell": "A4", "shift_direction": "up"},
        )
    )
    assert del_res["status"] == "success"
