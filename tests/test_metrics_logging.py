"""Tests for optional JSONL tool execution metrics via EXCEL_MCP_METRICS_LOG."""

import asyncio
import json
from pathlib import Path

from excel_mcp.utils.guards import tool_guard


def test_metrics_disabled_by_default(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.delenv("EXCEL_MCP_METRICS_LOG", raising=False)
    metrics_file = tmp_path / "metrics.jsonl"

    @tool_guard("sample_tool")
    def ok_tool() -> dict[str, str]:
        return {"status": "success"}

    result = ok_tool()
    assert result == {"status": "success"}
    assert not metrics_file.exists()


def test_metrics_records_success_and_error(monkeypatch, tmp_path: Path) -> None:
    metrics_file = tmp_path / "nested" / "metrics.jsonl"
    monkeypatch.setenv("EXCEL_MCP_METRICS_LOG", str(metrics_file))

    @tool_guard("excel_data")
    def ok_tool(action: str = "read") -> dict[str, str]:
        return {"status": "success", "action": action}

    @tool_guard("excel_data")
    def fail_tool(action: str = "write") -> dict[str, str]:
        raise ValueError("bad input")

    assert ok_tool(action="read")["status"] == "success"
    err = fail_tool(action="write")
    assert err["status"] == "error"

    lines = [
        json.loads(line)
        for line in metrics_file.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    assert len(lines) == 2

    assert lines[0]["tool"] == "excel_data"
    assert lines[0]["action"] == "read"
    assert lines[0]["status"] == "success"
    assert isinstance(lines[0]["duration_ms"], float)
    assert "timestamp" in lines[0]

    assert lines[1]["tool"] == "excel_data"
    assert lines[1]["action"] == "write"
    assert lines[1]["status"] == "error"
    assert lines[1]["error_type"] == "ValueError"
    assert isinstance(lines[1]["duration_ms"], float)


def test_metrics_async_tool_and_non_dict_return(monkeypatch, tmp_path: Path) -> None:
    metrics_file = tmp_path / "async_metrics.jsonl"
    monkeypatch.setenv("EXCEL_MCP_METRICS_LOG", str(metrics_file))

    @tool_guard("async_sample")
    async def async_ok() -> str:
        return "plain-text"

    result = asyncio.run(async_ok())
    assert result == "plain-text"

    records = [
        json.loads(line)
        for line in metrics_file.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    assert len(records) == 1
    assert records[0]["tool"] == "async_sample"
    assert records[0]["status"] == "success"


def test_metrics_unwritable_path_does_not_fail_tool(
    monkeypatch, tmp_path: Path
) -> None:
    # Point EXCEL_MCP_METRICS_LOG to an existing directory so open("a") raises IsADirectoryError/PermissionError
    monkeypatch.setenv("EXCEL_MCP_METRICS_LOG", str(tmp_path))

    @tool_guard("resilient_tool")
    def ok_tool() -> dict[str, str]:
        return {"status": "success"}

    assert ok_tool() == {"status": "success"}


def test_server_middleware_records_single_pass_and_io_phases(
    monkeypatch, tmp_path: Path
) -> None:
    from excel_mcp.server import create_server

    metrics_file = tmp_path / "server_metrics.jsonl"
    xlsx_file = tmp_path / "sample.xlsx"
    monkeypatch.setenv("EXCEL_MCP_METRICS_LOG", str(metrics_file))
    monkeypatch.setenv("EXCEL_MCP_DIRS", str(tmp_path))

    async def run_calls() -> None:
        mcp = create_server()
        await mcp.call_tool(
            "excel_build_workbook",
            {
                "filename": str(xlsx_file),
                "sheets": [
                    {
                        "name": "Ventas",
                        "headers": ["Producto", "Cantidad"],
                        "rows": [["Laptop", 3]],
                    }
                ],
            },
        )
        await mcp.call_tool("excel_get_workbook_outline", {"filename": str(xlsx_file)})

    asyncio.run(run_calls())

    records = [
        json.loads(line)
        for line in metrics_file.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    assert len(records) == 2
    assert records[0]["tool"] == "excel_build_workbook"
    assert records[0]["status"] == "success"
    assert "io_save_ms" in records[0]["phases"]

    assert records[1]["tool"] == "excel_get_workbook_outline"
    assert records[1]["status"] == "success"
    assert "io_read_ms" in records[1]["phases"]
