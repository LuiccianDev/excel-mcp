"""Tests for canonical tool registration and legacy aliases switch."""

from __future__ import annotations

import os
from unittest.mock import patch
import unittest

from excel_mcp.tools.register_tools import register_all_tools


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
    """Test suite for canonical tool registration."""

    def test_default_registers_compound_mode_and_full_mode_registers_canonical_names(
        self,
    ) -> None:
        """Verify default is compound (8 tools) and EXCEL_MCP_TOOLS=full registers 33 canonical tools."""
        with patch.dict(os.environ, {}, clear=True):
            compound_mcp = MockFastMCP()
            register_all_tools(compound_mcp)
            self.assertEqual(len(compound_mcp.registered_tools), 8)

        with patch.dict(os.environ, {"EXCEL_MCP_TOOLS": "full"}, clear=True):
            mock_mcp = MockFastMCP()
            register_all_tools(mock_mcp)

            tool_names = list(mock_mcp.registered_tools.keys())
            self.assertEqual(len(tool_names), 33)
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


if __name__ == "__main__":
    unittest.main()
