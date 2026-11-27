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
        """Verify guarded tool functions convert real core errors via @tool_guard."""
        import os
        from pathlib import Path
        import tempfile
        from unittest.mock import patch

        import openpyxl

        from excel_mcp.tools.format_tools import delete_worksheet
        from excel_mcp.tools.formulas_excel_tools import apply_formula_excel

        with tempfile.TemporaryDirectory() as tmp_dir:
            sample_path = Path(tmp_dir) / "sample.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            assert ws is not None
            ws.title = "Sheet1"
            wb.save(sample_path)
            wb.close()

            with patch.dict(os.environ, {"EXCEL_MCP_DIRS": tmp_dir}):
                res = delete_worksheet(filename=str(sample_path), sheet_name="Missing")
                res2 = apply_formula_excel(
                    filename=str(sample_path),
                    sheet_name="Missing",
                    cell="A1",
                    formula="=1+1",
                )

            self.assertEqual(res["status"], "error")
            self.assertEqual(res["error_type"], "SheetNotFoundError")
            self.assertEqual(res["operation"], "delete_worksheet")
            self.assertEqual(res["filename"], str(sample_path.resolve()))

            self.assertEqual(res2["status"], "error")
            self.assertEqual(res2["error_type"], "ValidationError")
            self.assertEqual(res2["operation"], "apply_formula_excel")
            self.assertEqual(res2["filename"], str(sample_path.resolve()))

    def test_path_rejection_returns_error_and_logs_metric_in_full_and_compound(
        self,
    ) -> None:
        """Verify path rejections return PermissionError payloads and emit error metrics."""
        import json
        import os
        from pathlib import Path
        import tempfile
        from unittest.mock import patch

        from excel_mcp.tools.compound_tools import excel_data
        from excel_mcp.tools.content_tools import read_data_from_excel

        with tempfile.TemporaryDirectory() as tmp_dir:
            allowed_dir = Path(tmp_dir) / "allowed"
            allowed_dir.mkdir()
            metrics_path = Path(tmp_dir) / "metrics.jsonl"

            with patch.dict(
                os.environ,
                {
                    "EXCEL_MCP_DIRS": str(allowed_dir),
                    "EXCEL_MCP_METRICS_LOG": str(metrics_path),
                },
            ):
                full_res = read_data_from_excel(
                    filename="../outside.xlsx", sheet_name="Sheet1"
                )
                compound_res = excel_data(
                    "read",
                    "../outside.xlsx",
                    {"sheet_name": "Sheet1"},
                )

            self.assertEqual(full_res["status"], "error")
            self.assertEqual(full_res["error_type"], "PermissionError")
            self.assertEqual(full_res["operation"], "read_data_from_excel")

            self.assertEqual(compound_res["status"], "error")
            self.assertEqual(compound_res["error_type"], "PermissionError")
            self.assertEqual(compound_res["operation"], "read_data_from_excel")

            records = [
                json.loads(line)
                for line in metrics_path.read_text(encoding="utf-8").splitlines()
                if line.strip()
            ]
            self.assertGreaterEqual(len(records), 2)
            self.assertTrue(all(r["status"] == "error" for r in records))
            self.assertTrue(all(r["error_type"] == "PermissionError" for r in records))

    def test_src_modules_and_functions_respect_size_limits(self) -> None:
        """Verify all src/excel_mcp modules stay <=600 lines and functions <=150 lines."""
        import ast
        from pathlib import Path

        src_root = Path(__file__).resolve().parents[1] / "src" / "excel_mcp"
        oversized_modules: list[tuple[str, int]] = []
        oversized_functions: list[tuple[str, str, int]] = []

        for py_file in sorted(src_root.rglob("*.py")):
            text = py_file.read_text(encoding="utf-8")
            line_count = len(text.splitlines())
            if line_count > 600:
                oversized_modules.append((py_file.name, line_count))
            tree = ast.parse(text, filename=str(py_file))
            for node in ast.walk(tree):
                if isinstance(node, ast.FunctionDef | ast.AsyncFunctionDef):
                    assert node.end_lineno is not None
                    fn_lines = node.end_lineno - node.lineno + 1
                    if fn_lines > 150:
                        oversized_functions.append((py_file.name, node.name, fn_lines))

        self.assertEqual(oversized_modules, [])
        self.assertEqual(oversized_functions, [])


if __name__ == "__main__":
    unittest.main()
