"""Tests for compact compose tool-surface mode."""

from __future__ import annotations

import asyncio
from pathlib import Path
from typing import Any

import pytest

from excel_mcp.tools.compose_tools import (
    COMPOSE_REGISTRY,
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


def test_compose_mode_registers_eight_tools(monkeypatch: pytest.MonkeyPatch) -> None:
    """Default compose mode exposes 8 tools and EXCEL_MCP_MODE=full switches to 33."""
    monkeypatch.delenv("EXCEL_MCP_MODE", raising=False)
    monkeypatch.delenv("EXCEL_MCP_LEGACY_ALIASES", raising=False)
    assert get_tool_mode() == "compose"

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

    monkeypatch.setenv("EXCEL_MCP_MODE", "full")
    assert get_tool_mode() == "full"
    full_collector = _ToolCollector()
    register_all_tools(full_collector)
    assert len(full_collector.tools) == 33


def test_compose_registry_covers_all_granular_tools() -> None:
    """Every non-single-pass canonical tool in full mode must exist in COMPOSE_REGISTRY."""
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
    composed_handlers = {
        spec.handler
        for actions in COMPOSE_REGISTRY.values()
        for spec in actions.values()
    }
    assert composed_handlers == full_handlers


def test_invalid_mode_raises_value_error() -> None:
    """Explicit invalid mode passed to register_all_tools must fail loudly."""
    collector = _ToolCollector()
    with pytest.raises(ValueError, match="Invalid tool mode"):
        register_all_tools(collector, mode="invalid")


def test_compose_tools_end_to_end(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Composed tools must execute real workbook, data, formula, and format flows."""
    docs_dir = tmp_path / "Documents"
    docs_dir.mkdir(parents=True)
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(docs_dir))

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


def test_compose_tools_reject_invalid_actions_and_params() -> None:
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
