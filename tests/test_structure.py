"""Unit tests for structure tools and operations."""

from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl

from excel_mcp.core.calculations import FormulaSecurityError
from excel_mcp.core.structure import (
    append_rows_to_sheet,
    create_excel_table,
    delete_worksheet_dimension,
    format_worksheet_properties,
    insert_worksheet_dimension,
    sort_worksheet_range,
)
from excel_mcp.exceptions.exception_tools import ValidationError
from excel_mcp.tools.structure_tools import (
    append_rows,
    create_table,
    delete_rows_columns,
    format_sheet,
    insert_rows_columns,
    sort_range,
)


class TestStructureOperations(unittest.TestCase):
    """Test suite for worksheet structure and layout operations."""

    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.sample_workbook = str(Path(self.temp_dir.name) / "test_structure.xlsx")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Name", "Age", "Score"])
        ws.append(["Alice", 30, 85.5])
        ws.append(["Bob", 25, 92.0])
        ws.append(["Charlie", 35, 78.0])
        wb.save(self.sample_workbook)
        wb.close()

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_append_rows_success(self) -> None:
        res = append_rows_to_sheet(
            self.sample_workbook,
            rows=[["David", 28, 88.0], ["Eve", 22, 95.0]],
            sheet_name="Data",
        )
        self.assertEqual(res["rows_appended"], 2)
        self.assertEqual(res["start_row"], 5)
        self.assertEqual(res["end_row"], 6)

        wb = openpyxl.load_workbook(self.sample_workbook)
        ws = wb["Data"]
        self.assertEqual(ws.cell(5, 1).value, "David")
        self.assertEqual(ws.cell(6, 1).value, "Eve")
        self.assertEqual(ws.cell(6, 2).value, 22)
        wb.close()

    def test_append_rows_blocks_injected_formula(self) -> None:
        with self.assertRaises(FormulaSecurityError):
            append_rows_to_sheet(
                self.sample_workbook,
                rows=[["Malicious", 99, "=cmd|' /C calc'!A0"]],
                sheet_name="Data",
            )

    def test_insert_and_delete_rows_columns(self) -> None:
        # Insert 1 row at row 2
        res_ins = insert_worksheet_dimension(
            self.sample_workbook, dimension="row", idx=2, amount=1, sheet_name="Data"
        )
        self.assertEqual(res_ins["action"], "insert")

        wb = openpyxl.load_workbook(self.sample_workbook)
        ws = wb["Data"]
        self.assertIsNone(ws.cell(2, 1).value)
        self.assertEqual(ws.cell(3, 1).value, "Alice")
        wb.close()

        # Delete the inserted row
        res_del = delete_worksheet_dimension(
            self.sample_workbook, dimension="row", idx=2, amount=1, sheet_name="Data"
        )
        self.assertEqual(res_del["action"], "delete")

        wb = openpyxl.load_workbook(self.sample_workbook)
        ws = wb["Data"]
        self.assertEqual(ws.cell(2, 1).value, "Alice")
        wb.close()

    def test_format_worksheet_properties(self) -> None:
        res = format_worksheet_properties(
            self.sample_workbook,
            sheet_name="Data",
            freeze_panes="A2",
            auto_filter="A1:C4",
            show_grid_lines=True,
            column_widths={"A": 20.0, "B": 12.0},
        )
        self.assertIn("freeze_panes", res["properties_updated"])
        self.assertEqual(res["properties_updated"]["freeze_panes"], "A2")
        self.assertEqual(res["properties_updated"]["column_widths"]["A"], 20.0)

        wb = openpyxl.load_workbook(self.sample_workbook)
        ws = wb["Data"]
        self.assertEqual(ws.freeze_panes, "A2")
        self.assertEqual(ws.auto_filter.ref, "A1:C4")
        self.assertEqual(ws.column_dimensions["A"].width, 20.0)
        wb.close()

    def test_create_excel_table(self) -> None:
        res = create_excel_table(
            self.sample_workbook,
            range_string="A1:C4",
            table_name="UsersTable",
            sheet_name="Data",
        )
        self.assertEqual(res["table_name"], "UsersTable")

        wb = openpyxl.load_workbook(self.sample_workbook)
        ws = wb["Data"]
        self.assertIn("UsersTable", ws.tables)
        wb.close()

        # Duplicate table name should raise ValidationError
        with self.assertRaises(ValidationError):
            create_excel_table(
                self.sample_workbook,
                range_string="A1:C4",
                table_name="UsersTable",
                sheet_name="Data",
            )

    def test_sort_worksheet_range(self) -> None:
        # Sort by Age (column 2 / B), ascending
        res = sort_worksheet_range(
            self.sample_workbook,
            range_string="A1:C4",
            key_column=2,
            ascending=True,
            has_header=True,
            sheet_name="Data",
        )
        self.assertEqual(res["sorted_rows"], 3)

        wb = openpyxl.load_workbook(self.sample_workbook)
        ws = wb["Data"]
        # Row 2 should now be Bob (age 25)
        self.assertEqual(ws.cell(2, 1).value, "Bob")
        self.assertEqual(ws.cell(2, 2).value, 25)
        # Row 3 Alice (30)
        self.assertEqual(ws.cell(3, 1).value, "Alice")
        # Row 4 Charlie (35)
        self.assertEqual(ws.cell(4, 1).value, "Charlie")
        wb.close()

    def test_structure_tools_guards(self) -> None:
        with patch.dict(
            os.environ,
            {"MCP_ALLOWED_DIRECTORIES": os.path.dirname(self.sample_workbook)},
        ):
            r1 = append_rows(
                filename=self.sample_workbook,
                rows=[["Zara", 21, 99.0]],
                sheet_name="Data",
            )
            self.assertEqual(r1["status"], "success")

            r2 = format_sheet(
                filename=self.sample_workbook, sheet_name="Data", freeze_panes="A2"
            )
            self.assertEqual(r2["status"], "success")

            r3 = sort_range(
                filename=self.sample_workbook,
                range_string="A1:C5",
                key_column="Score",
                ascending=False,
            )
            self.assertEqual(r3["status"], "success")

            r4 = insert_rows_columns(
                filename=self.sample_workbook, dimension="column", index=1, amount=1
            )
            self.assertEqual(r4["status"], "success")

            r5 = delete_rows_columns(
                filename=self.sample_workbook, dimension="column", index=1, amount=1
            )
            self.assertEqual(r5["status"], "success")

            r6 = create_table(
                filename=self.sample_workbook,
                range_string="A1:C5",
                table_name="AutoTable",
            )
            self.assertEqual(r6["status"], "success")


if __name__ == "__main__":
    unittest.main()
