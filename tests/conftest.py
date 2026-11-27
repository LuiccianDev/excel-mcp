"""
Test configuration and fixtures for MCP Excel Office Server tests.

This module provides pytest fixtures and configuration for testing
the MCP Excel Office Server functionality.
"""

from collections.abc import Generator
import signal
import socket
import sys
from typing import Any

import pytest

from excel_mcp.config import ConfigurationManager


def _install_windows_sandbox_asyncio_fallback() -> None:
    if sys.platform != "win32":
        return
    probe: socket.socket | None = None
    try:
        probe = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        probe.bind(("127.0.0.1", 0))
        probe.listen(1)
        probe.settimeout(0.05)
        client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            client.settimeout(0.05)
            client.connect(probe.getsockname())
            conn, _ = probe.accept()
            conn.close()
            return
        finally:
            client.close()
    except OSError:
        pass
    finally:
        if probe is not None:
            probe.close()

    import asyncio.proactor_events as pe
    import asyncio.windows_events as we

    class _DummySelfPipeSocket:
        def fileno(self) -> int:
            return -1

        def close(self) -> None:
            return None

        def send(self, data: bytes) -> int:
            return len(data)

    orig_set_wakeup = signal.set_wakeup_fd
    signal_mod: Any = signal
    signal_mod.set_wakeup_fd = (
        lambda fd, **kw: -1 if fd == -1 else orig_set_wakeup(fd, **kw)
    )

    def _make_dummy_self_pipe(self: Any) -> None:
        self._ssock = _DummySelfPipeSocket()
        self._csock = _DummySelfPipeSocket()
        self._internal_fds += 1

    proactor_cls: Any = pe.BaseProactorEventLoop
    proactor_cls._make_self_pipe = _make_dummy_self_pipe
    proactor_cls._loop_self_reading = lambda self, f=None: None
    orig_select = we.IocpProactor.select
    iocp_cls: Any = we.IocpProactor
    iocp_cls.select = lambda self, timeout=None: orig_select(
        self, 0.005 if timeout is None else min(timeout, 0.005)
    )


_install_windows_sandbox_asyncio_fallback()


@pytest.fixture
def test_config(tmp_path) -> Generator[ConfigurationManager, None, None]:
    """Configure the system for testing with a temporary directory."""
    manager = ConfigurationManager()

    # Store original configuration
    original_config = manager.config

    try:
        # Set up test configuration using tmp_path
        manager.reload_configuration(directory=str(tmp_path), log_level="INFO")
        yield manager
    finally:
        # Restore original configuration
        manager._config = original_config


@pytest.fixture
def test_excel_file(tmp_path) -> str:
    """Create a test Excel file in the temporary directory."""
    import openpyxl

    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    # Create a test workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Add some test data
    test_data = [
        ["Name", "Age", "City"],
        ["Alice", 30, "New York"],
        ["Bob", 25, "Los Angeles"],
        ["Charlie", 35, "Chicago"],
    ]

    for row_idx, row_data in enumerate(test_data, 1):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    # Save to tmp_path directory
    test_file_path = tmp_path / "test_file.xlsx"
    wb.save(test_file_path)

    return str(test_file_path)


@pytest.fixture
def test_workbook_file(tmp_path) -> str:
    """Create a test workbook file in the temporary directory."""
    import openpyxl

    # Configure test environment
    manager = ConfigurationManager()
    manager.reload_configuration(directory=str(tmp_path), log_level="INFO")

    # Create a test workbook with multiple sheets
    wb = openpyxl.Workbook()

    # First sheet with data
    ws1 = wb.active
    ws1.title = "Data"
    data = [
        ["Product", "Sales", "Region"],
        ["Widget A", 100, "North"],
        ["Widget B", 150, "South"],
        ["Widget C", 200, "East"],
    ]

    for row_idx, row_data in enumerate(data, 1):
        for col_idx, value in enumerate(row_data, 1):
            ws1.cell(row=row_idx, column=col_idx, value=value)

    # Second sheet for formulas
    ws2 = wb.create_sheet("Formulas")
    ws2["A1"] = "Total Sales"
    ws2["B1"] = "=SUM(Data.B2:B4)"

    # Save to tmp_path directory
    test_file_path = tmp_path / "test_workbook.xlsx"
    wb.save(test_file_path)

    return str(test_file_path)
