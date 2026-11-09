"""Tests for single-pass batch workbook building, presets, and outline inspection."""

from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path

import openpyxl

from excel_mcp.core.batch import (
    batch_write_matrix,
    build_single_pass_workbook,
    inspect_workbook_outline,
)
from excel_mcp.core.presets import PRESETS, get_preset
from excel_mcp.tools.batch_tools import (
    batch_write_data,
    build_workbook,
    get_workbook_outline,
)


class TestBatchTools(unittest.TestCase):
    """Test suite for batch builder, presets, and lightweight outline inspection."""

    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        os.environ["MCP_ALLOWED_DIRECTORIES"] = self.temp_dir.name

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_presets_exist(self) -> None:
        """Verify all standard presets exist with required styling attributes."""
        for name in ["corporate", "financial", "modern", "compact"]:
            self.assertIn(name, PRESETS)
            preset = get_preset(name)
            self.assertIn("font_name", preset)
            self.assertIn("header_fill", preset)
            self.assertIn("header_font_color", preset)
            self.assertIn("zebra_fill", preset)
            self.assertIn("currency_format", preset)

    def test_build_single_pass_workbook(self) -> None:
        """Verify single-pass workbook creation with headers, data, zebra styles, and totals."""
        wb_path = os.path.join(self.temp_dir.name, "test_report.xlsx")
        sheets = [
            {
                "name": "Revenue",
                "title": "ANNUAL REVENUE 2026",
                "headers": ["Quarter", "Target", "Actual"],
                "rows": [
                    ["Q1", 100000, 115000],
                    ["Q2", 120000, 128000],
                    ["Q3", 130000, 135000],
                ],
                "currency_columns": [1, 2],
                "auto_totals": [1, 2],
            },
            {
                "name": "Headcount",
                "headers": ["Department", "Employees"],
                "rows": [
                    ["Engineering", 45],
                    ["Sales", 30],
                    ["Marketing", 15],
                ],
                "auto_totals": [1],
            },
        ]

        result = build_single_pass_workbook(
            filename=wb_path,
            sheets=sheets,
            preset="financial",
            auto_fit=True,
        )

        self.assertEqual(result.total_rows_inserted, 6)
        self.assertEqual(len(result.sheets_created), 2)
        self.assertTrue(os.path.exists(wb_path))

        # Inspect generated workbook with openpyxl
        wb = openpyxl.load_workbook(wb_path)
        self.assertIn("Revenue", wb.sheetnames)
        self.assertIn("Headcount", wb.sheetnames)

        rev_ws = wb["Revenue"]
        # Title banner check
        self.assertEqual(rev_ws.cell(row=1, column=1).value, "ANNUAL REVENUE 2026")
        # Check auto totals formula
        # Row 1 is banner, Row 2 is blank, Row 3 is headers, Rows 4,5,6 are data, Row 7 is Total
        total_cell = rev_ws.cell(row=7, column=2)
        self.assertTrue(str(total_cell.value).startswith("=SUM(B4:B6)"))

    def test_inspect_workbook_outline(self) -> None:
        """Verify lightweight outline inspection without loading entire cell matrices."""
        wb_path = os.path.join(self.temp_dir.name, "outline_test.xlsx")
        sheets = [
            {
                "name": "DataSheet",
                "headers": ["ID", "Name", "Score"],
                "rows": [
                    [1, "Alice", 95],
                    [2, "Bob", 88],
                ],
            }
        ]
        build_single_pass_workbook(filename=wb_path, sheets=sheets, preset="modern")

        outline = inspect_workbook_outline(wb_path)
        self.assertEqual(outline.sheet_count, 1)
        self.assertEqual(outline.sheets[0].name, "DataSheet")
        self.assertIn("ID", outline.sheets[0].headers)
        self.assertIn("Score", outline.sheets[0].headers)

    def test_batch_write_matrix(self) -> None:
        """Verify atomic 2D matrix writing."""
        wb_path = os.path.join(self.temp_dir.name, "batch_test.xlsx")
        build_single_pass_workbook(
            filename=wb_path,
            sheets=[{"name": "Grid", "headers": ["A", "B", "C"], "rows": [[1, 2, 3]]}],
        )

        write_res = batch_write_matrix(
            filename=wb_path,
            sheet_name="Grid",
            start_cell="B5",
            data_matrix=[
                ["Data 1", 100],
                ["Data 2", 200],
            ],
        )

        self.assertEqual(write_res.rows_affected, 2)
        self.assertEqual(write_res.columns_affected, 2)
        self.assertEqual(write_res.cells_updated, 4)

        wb = openpyxl.load_workbook(wb_path)
        ws = wb["Grid"]
        self.assertEqual(ws["B5"].value, "Data 1")
        self.assertEqual(ws["C5"].value, 100)
        self.assertEqual(ws["B6"].value, "Data 2")
        self.assertEqual(ws["C6"].value, 200)

    def test_tool_wrapper_security(self) -> None:
        """Verify MCP batch tool security rejects forbidden paths."""
        res = build_workbook(
            filename="/etc/forbidden_excel.xlsx",
            sheets=[{"name": "S", "rows": [[1]]}],
        )
        self.assertEqual(res["status"], "error")
        self.assertIn("Security restriction", res["message"])


if __name__ == "__main__":
    unittest.main()
