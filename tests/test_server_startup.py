"""Tests for MCP server startup optimizations and warning suppression."""

from __future__ import annotations

import os
import warnings
from unittest.mock import MagicMock

import excel_mcp.server as server_mod


def test_fastmcp_banner_and_update_check_disabled_by_default() -> None:
    """Server module must disable FastMCP update check and CLI banner by default."""
    assert os.environ.get("FASTMCP_CHECK_FOR_UPDATES") == "off"
    assert os.environ.get("FASTMCP_SHOW_SERVER_BANNER") == "false"


def test_openpyxl_user_warning_is_suppressed(recwarn: object) -> None:
    """Noisy openpyxl UserWarnings should be filtered out from stderr."""
    with warnings.catch_warnings(record=True) as caught:
        # Re-apply server warning filter inside catch_warnings
        server_mod.configure_warning_filters()
        warnings.warn_explicit(
            "Failed to load a conditional formatting rule. It will be discarded.",
            category=UserWarning,
            filename="openpyxl/worksheet/_reader.py",
            lineno=313,
            module="openpyxl.worksheet._reader",
        )
    assert caught == []


def test_run_stdio_passes_show_banner_false() -> None:
    """run_stdio helper passes show_banner=False when FastMCP.run accepts it."""
    mock_server = MagicMock()

    def _fake_run(transport: str = "stdio", show_banner: bool | None = None) -> None:
        assert transport == "stdio"
        assert show_banner is False

    mock_server.run = _fake_run
    server_mod.run_stdio(mock_server)


def test_run_stdio_falls_back_when_show_banner_unsupported() -> None:
    """run_stdio helper falls back cleanly if a mock/older FastMCP.run lacks show_banner."""
    called_with: list[str] = []

    class LegacyServer:
        def run(self, transport: str = "stdio") -> None:
            called_with.append(transport)

    server_mod.run_stdio(LegacyServer())
    assert called_with == ["stdio"]
