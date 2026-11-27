"""CLI entrypoint and server construction tests."""

from __future__ import annotations

import os
import sys

import pytest

import excel_mcp
import excel_mcp.server as server_module
from excel_mcp import __main__ as cli
from excel_mcp.config import config_manager
from excel_mcp.tools.register_tools import normalize_tool_mode


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


def test_instructions_omit_legacy_aliases(monkeypatch):
    monkeypatch.setattr(server_module, "FastMCP", _Recorder)
    server_module.create_server()
    instructions = _Recorder.captured["instructions"]
    assert "also registered" not in instructions
    assert "EXCEL_MCP_LEGACY_ALIASES" not in instructions


@pytest.mark.parametrize(
    "extra_args",
    [
        ["server"],
        ["list"],
        ["--tool-mode", "full"],
        ["--path", "x"],
        ["--help"],
    ],
)
def test_main_rejects_any_cli_argument(monkeypatch, capsys, extra_args):
    monkeypatch.setattr(sys, "argv", ["excel-mcp", *extra_args])
    monkeypatch.setattr(cli, "run_server", lambda: _FakeServer())
    with pytest.raises(SystemExit) as exc:
        cli.main()
    assert exc.value.code == 2
    assert "takes no arguments" in capsys.readouterr().err


def test_main_uses_excel_mcp_tools_env(monkeypatch, capsys):
    monkeypatch.setenv("EXCEL_MCP_TOOLS", "full")
    monkeypatch.setattr(sys, "argv", ["excel-mcp"])
    monkeypatch.setattr(cli, "run_server", lambda: _FakeServer())
    cli.main()
    assert normalize_tool_mode(None) == "full"
    assert capsys.readouterr().out == ""


def test_main_rejects_dirs_env_pointing_to_file(monkeypatch, tmp_path):
    original_config = config_manager._config
    original_allowed = os.environ.get("EXCEL_MCP_DIRS")
    try:
        not_a_dir = tmp_path / "file.txt"
        not_a_dir.write_text("x", encoding="utf-8")
        monkeypatch.setenv("EXCEL_MCP_DIRS", str(not_a_dir))
        monkeypatch.setattr(sys, "argv", ["excel-mcp"])
        monkeypatch.setattr(cli, "run_server", lambda: _FakeServer())
        with pytest.raises(SystemExit) as exc:
            cli.main()
        assert exc.value.code != 0
    finally:
        config_manager._config = original_config
        if original_allowed is None:
            os.environ.pop("EXCEL_MCP_DIRS", None)
        else:
            os.environ["EXCEL_MCP_DIRS"] = original_allowed
