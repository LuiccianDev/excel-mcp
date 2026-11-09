"""Tests for canonical tool registration and legacy aliases switch."""

from __future__ import annotations

import os
from unittest.mock import patch
import unittest

from excel_mcp.tools.register_tools import register_all_tools, should_register_legacy_aliases


class MockFastMCP:
    """Mock FastMCP server to record tool registrations."""

    def __init__(self) -> None:
        self.registered_tools: dict[str, object] = {}

    def tool(self, name: str | None = None, annotations: object = None) -> object:
        def decorator(fn: object) -> object:
            tool_name = name or getattr(fn, "__name__", "unnamed")
            self.registered_tools[tool_name] = fn
            return fn

        return decorator


class TestRegisterTools(unittest.TestCase):
    """Test suite for tool registration and legacy alias toggling."""

    def test_default_only_registers_canonical_excel_names(self) -> None:
        """Verify default configuration registers only canonical excel_* tools."""
        with patch.dict(os.environ, {}, clear=True):
            self.assertFalse(should_register_legacy_aliases())

            mock_mcp = MockFastMCP()
            register_all_tools(mock_mcp)

            tool_names = list(mock_mcp.registered_tools.keys())
            self.assertEqual(len(tool_names), 30)
            for name in tool_names:
                self.assertTrue(
                    name.startswith("excel_"),
                    f"Tool '{name}' should be prefixed with 'excel_' by default",
                )

            # Check specific canonical tools
            self.assertIn("excel_create_workbook", tool_names)
            self.assertIn("excel_create_worksheet", tool_names)
            self.assertIn("excel_format_range", tool_names)
            self.assertIn("excel_apply_formula", tool_names)
            self.assertIn("excel_append_rows", tool_names)
            self.assertIn("excel_format_sheet", tool_names)
            self.assertIn("excel_create_table", tool_names)
            self.assertIn("excel_sort_range", tool_names)
            self.assertIn("excel_insert_rows_columns", tool_names)
            self.assertIn("excel_delete_rows_columns", tool_names)
            self.assertIn("excel_find_cells", tool_names)
            self.assertIn("excel_profile_sheet", tool_names)
            self.assertNotIn("create_excel_workbook", tool_names)

    def test_legacy_aliases_enabled(self) -> None:
        """Verify EXCEL_MCP_LEGACY_ALIASES=1 enables legacy aliases."""
        with patch.dict(os.environ, {"EXCEL_MCP_LEGACY_ALIASES": "1"}, clear=True):
            self.assertTrue(should_register_legacy_aliases())

            mock_mcp = MockFastMCP()
            register_all_tools(mock_mcp)

            tool_names = list(mock_mcp.registered_tools.keys())
            self.assertGreater(len(tool_names), 30)

            # Both canonical and legacy aliases must exist
            self.assertIn("excel_create_workbook", tool_names)
            self.assertIn("create_excel_workbook", tool_names)
            self.assertIn("excel_create_excel_workbook", tool_names)
            self.assertIn("format_range_excel", tool_names)
            self.assertIn("build_workbook", tool_names)


if __name__ == "__main__":
    unittest.main()
