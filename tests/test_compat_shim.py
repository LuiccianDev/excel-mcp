"""Tests for backward-compatibility alias shim mcp_excel."""

from __future__ import annotations

import unittest


class TestCompatShim(unittest.TestCase):
    """Test suite ensuring mcp_excel imports dynamically resolve to excel_mcp."""

    def test_submodule_aliasing(self) -> None:
        """Verify mcp_excel submodules resolve to identical excel_mcp objects."""
        import mcp_excel
        from mcp_excel.core import data as m_data
        from excel_mcp.core import data as e_data
        self.assertIs(m_data, e_data)

        from mcp_excel.exceptions import exception_core as m_exc
        from excel_mcp.exceptions import exception_core as e_exc
        self.assertIs(m_exc, e_exc)

    def test_top_level_attribute_access(self) -> None:
        """Verify top-level getattr dynamic resolution."""
        import mcp_excel

        self.assertTrue(hasattr(mcp_excel, "core"))
        self.assertTrue(hasattr(mcp_excel, "exceptions"))


if __name__ == "__main__":
    unittest.main()
