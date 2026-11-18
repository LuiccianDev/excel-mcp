"""Tests for tool_guard execution decorator."""

from __future__ import annotations

import asyncio
import unittest

from excel_mcp.exceptions.exception_core import SheetNotFoundError, ValidationError
from excel_mcp.utils.guards import tool_guard


class TestToolGuard(unittest.TestCase):
    """Test suite for tool_guard error handling decorator."""

    def test_sync_success(self) -> None:
        """Verify successful synchronous tool execution."""

        @tool_guard()
        def sample_tool(val: int) -> dict[str, str]:
            return {"status": "success", "result": str(val * 2)}

        res = sample_tool(5)
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["result"], "10")

    def test_sync_exception_handling(self) -> None:
        """Verify synchronous exception translation into structured dict."""

        @tool_guard(tool_name="test_sheet_tool")
        def failing_tool(filename: str) -> None:
            raise SheetNotFoundError(f"Missing sheet in {filename}")

        res = failing_tool(filename="sample.xlsx")
        self.assertEqual(res["status"], "error")
        self.assertEqual(res["operation"], "test_sheet_tool")
        self.assertEqual(res["filename"], "sample.xlsx")
        self.assertEqual(res["error_type"], "SheetNotFoundError")
        self.assertIn("Missing sheet", res["message"])

    def test_async_success(self) -> None:
        """Verify successful asynchronous tool execution."""

        @tool_guard()
        async def async_tool(x: int) -> dict[str, int]:
            await asyncio.sleep(0.001)
            return {"status": "success", "computed": x + 1}

        res = asyncio.run(async_tool(9))
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["computed"], 10)

    def test_async_exception_handling(self) -> None:
        """Verify asynchronous exception translation into structured dict."""

        @tool_guard()
        async def async_failing_tool(file_path: str) -> None:
            await asyncio.sleep(0.001)
            raise ValidationError("Invalid input range")

        res = asyncio.run(async_failing_tool(file_path="/data/test.xlsx"))
        self.assertEqual(res["status"], "error")
        self.assertEqual(res["filename"], "/data/test.xlsx")
        self.assertEqual(res["error_type"], "ValidationError")
        self.assertIn("Invalid input range", res["message"])

    def test_guarded_legacy_tool_modules_return_structured_error(self) -> None:
        """Verify newly guarded tool functions convert unhandled errors via @tool_guard."""
        import os
        import tempfile
        from unittest.mock import patch

        from excel_mcp.tools.format_tools import delete_worksheet
        from excel_mcp.tools.formulas_excel_tools import validate_formula_syntax

        with tempfile.TemporaryDirectory() as tmp_dir:
            sample_path = os.path.join(tmp_dir, "sample.xlsx")
            with (
                patch.dict(os.environ, {"EXCEL_MCP_DIRS": tmp_dir}),
                patch(
                    "excel_mcp.tools.format_tools.delete_sheet",
                    side_effect=SheetNotFoundError("Missing sheet"),
                ),
            ):
                res = asyncio.run(
                    delete_worksheet(filename=sample_path, sheet_name="Missing")
                )
            self.assertEqual(res["status"], "error")
            self.assertEqual(res["error_type"], "SheetNotFoundError")
            self.assertEqual(res["operation"], "delete_worksheet")

            with (
                patch.dict(os.environ, {"EXCEL_MCP_DIRS": tmp_dir}),
                patch(
                    "excel_mcp.tools.formulas_excel_tools.validate_formula_in_cell_operation",
                    side_effect=ValidationError("Bad syntax"),
                ),
            ):
                res2 = asyncio.run(
                    validate_formula_syntax(
                        filename=sample_path,
                        sheet_name="Sheet1",
                        cell="A1",
                        formula="=1+",
                    )
                )
            self.assertEqual(res2["status"], "error")
            self.assertEqual(res2["error_type"], "ValidationError")
            self.assertEqual(res2["operation"], "validate_formula_syntax")


if __name__ == "__main__":
    unittest.main()
