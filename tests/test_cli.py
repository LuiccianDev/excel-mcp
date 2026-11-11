"""CLI and server construction tests."""

from __future__ import annotations

import sys

import excel_mcp
import excel_mcp.server as server_module


class _FakeServer:
    def run(self, transport: str = "stdio") -> None:
        return None


class _Recorder:
    """Minimal FastMCP stand-in that captures constructor kwargs."""

    captured: dict = {}

    def __init__(self, **kwargs):
        type(self).captured = kwargs

    def tool(self, *args, **kwargs):
        def decorator(fn):
            return fn

        return decorator


def test_main_keeps_stdout_clean(monkeypatch, capsys):
    from excel_mcp import __main__ as cli

    monkeypatch.setattr(sys, "argv", ["excel-mcp"])
    monkeypatch.setattr(cli, "run_server", lambda: _FakeServer())
    cli.main()
    captured = capsys.readouterr()
    assert captured.out == ""


def test_create_server_reports_package_version(monkeypatch):
    monkeypatch.setattr(server_module, "FastMCP", _Recorder)
    server_module.create_server()
    assert _Recorder.captured["version"] == excel_mcp.__version__
    assert _Recorder.captured["name"] == "excel_mcp"


def test_instructions_describe_opt_in_aliases(monkeypatch):
    monkeypatch.setattr(server_module, "FastMCP", _Recorder)
    server_module.create_server()
    instructions = _Recorder.captured["instructions"]
    assert "also registered" not in instructions
    assert "EXCEL_MCP_LEGACY_ALIASES" in instructions
