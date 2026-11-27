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
    monkeypatch.setattr(cli, "run_server", lambda mode=None: _FakeServer())
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


def test_main_supports_root_tool_mode_flag(monkeypatch):
    from excel_mcp import __main__ as cli

    captured_modes: list[str | None] = []

    def _fake_run_server(mode: str | None = None):
        captured_modes.append(mode)
        return _FakeServer()

    monkeypatch.setattr(sys, "argv", ["excel-mcp", "--tool-mode", "full"])
    monkeypatch.setattr(cli, "run_server", _fake_run_server)
    cli.main()
    assert captured_modes == ["full"]


def test_main_rejects_invalid_root_tool_mode(monkeypatch):
    import pytest
    from excel_mcp import __main__ as cli

    monkeypatch.setattr(sys, "argv", ["excel-mcp", "--tool-mode", "bogus"])
    with pytest.raises(SystemExit) as exc:
        cli.main()
    assert exc.value.code != 0


def test_main_rejects_placeholder_path(monkeypatch):
    import os
    import pytest
    from excel_mcp import __main__ as cli
    from excel_mcp.config import config_manager

    original_config = config_manager._config
    original_allowed = os.environ.get("EXCEL_MCP_DIRS")
    try:
        monkeypatch.setattr(sys, "argv", ["excel-mcp", "--path", "${workspaceFolder}"])
        monkeypatch.setattr(cli, "run_server", lambda mode=None: _FakeServer())
        with pytest.raises(SystemExit) as exc:
            cli.main()
        assert exc.value.code != 0
    finally:
        config_manager._config = original_config
        if original_allowed is None:
            os.environ.pop("EXCEL_MCP_DIRS", None)
        else:
            os.environ["EXCEL_MCP_DIRS"] = original_allowed


def test_main_rejects_path_that_is_a_file(monkeypatch, tmp_path):
    import os
    import pytest
    from excel_mcp import __main__ as cli
    from excel_mcp.config import config_manager

    original_config = config_manager._config
    original_allowed = os.environ.get("EXCEL_MCP_DIRS")
    try:
        not_a_dir = tmp_path / "file.txt"
        not_a_dir.write_text("x", encoding="utf-8")
        monkeypatch.setattr(sys, "argv", ["excel-mcp", "--path", str(not_a_dir)])
        monkeypatch.setattr(cli, "run_server", lambda mode=None: _FakeServer())
        with pytest.raises(SystemExit) as exc:
            cli.main()
        assert exc.value.code != 0
    finally:
        config_manager._config = original_config
        if original_allowed is None:
            os.environ.pop("EXCEL_MCP_DIRS", None)
        else:
            os.environ["EXCEL_MCP_DIRS"] = original_allowed


def test_main_rejects_removed_server_subcommand(monkeypatch):
    import pytest
    from excel_mcp import __main__ as cli

    monkeypatch.setattr(sys, "argv", ["excel-mcp", "server"])
    monkeypatch.setattr(cli, "run_server", lambda mode=None: _FakeServer())
    with pytest.raises(SystemExit) as exc:
        cli.main()
    assert exc.value.code != 0
