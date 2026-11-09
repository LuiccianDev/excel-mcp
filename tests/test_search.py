"""Unit tests for search and profiling tools and operations."""

from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl

from excel_mcp.core.search import find_cells_in_workbook, profile_worksheet
from excel_mcp.tools.analysis_tools import find_cells, profile_sheet


class TestSearchAndProfileOperations(unittest.TestCase):
    """Test suite for cell discovery and worksheet profiling."""

    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.sample_workbook = str(Path(self.temp_dir.name) / "test_analysis.xlsx")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Employees"
        ws.append(["Name", "Department", "Salary", "Active"])
        ws.append(["Alice Smith", "Engineering", 120000, True])
        ws.append(["Bob Jones", "Marketing", 95000, False])
        ws.append(["Charlie Brown", "Engineering", 110000, True])
        ws.append(["Diana Prince", "Finance", None, True])

        ws2 = wb.create_sheet(title="Summary")
        ws2.append(["Metric", "Formula"])
        ws2.append(["Avg Salary", "=AVERAGE(Employees!C2:C4)"])

        wb.save(self.sample_workbook)
        wb.close()

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_find_cells_substring_and_case(self) -> None:
        # Case-insensitive substring search
        res = find_cells_in_workbook(self.sample_workbook, query="engineering")
        self.assertEqual(res["total_matches"], 2)
        cells = [m["cell"] for m in res["matches"]]
        self.assertIn("B2", cells)
        self.assertIn("B4", cells)

        # Case-sensitive search
        res_cs = find_cells_in_workbook(self.sample_workbook, query="engineering", match_case=True)
        self.assertEqual(res_cs["total_matches"], 0)

        # Formula detection
        res_formula = find_cells_in_workbook(self.sample_workbook, query="AVERAGE")
        self.assertEqual(res_formula["total_matches"], 1)
        self.assertTrue(res_formula["matches"][0]["is_formula"])
        self.assertEqual(res_formula["matches"][0]["sheet"], "Summary")

    def test_find_cells_regex(self) -> None:
        # All sheets match (4 employee names + "Avg Salary")
        res_all = find_cells_in_workbook(
            self.sample_workbook, query=r"^[A-Z][a-z]+ [A-Z][a-z]+$", use_regex=True
        )
        self.assertEqual(res_all["total_matches"], 5)

        # Scoped to Employees sheet only
        res_emp = find_cells_in_workbook(
            self.sample_workbook,
            query=r"^[A-Z][a-z]+ [A-Z][a-z]+$",
            sheet_name="Employees",
            use_regex=True,
        )
        self.assertEqual(res_emp["total_matches"], 4)

    def test_profile_worksheet(self) -> None:
        res = profile_worksheet(self.sample_workbook, sheet_name="Employees")
        self.assertEqual(res["total_columns"], 4)
        self.assertEqual(res["sampled_data_rows"], 4)

        cols = {c["header"]: c for c in res["columns"]}
        self.assertEqual(cols["Name"]["detected_type"], "string")
        self.assertEqual(cols["Department"]["detected_type"], "string")
        self.assertEqual(cols["Salary"]["detected_type"], "integer")
        self.assertEqual(cols["Salary"]["min"], 95000)
        self.assertEqual(cols["Salary"]["max"], 120000)
        self.assertEqual(cols["Salary"]["null_count"], 1)
        self.assertEqual(cols["Active"]["detected_type"], "boolean")

    def test_analysis_tools_guards(self) -> None:
        with patch.dict(os.environ, {"MCP_ALLOWED_DIRECTORIES": os.path.dirname(self.sample_workbook)}):
            r1 = find_cells(filename=self.sample_workbook, query="Alice")
            self.assertEqual(r1["status"], "success")
            self.assertGreaterEqual(r1["data"]["total_matches"], 1)

            r2 = profile_sheet(filename=self.sample_workbook, sheet_name="Employees")
            self.assertEqual(r2["status"], "success")
            self.assertEqual(len(r2["data"]["columns"]), 4)


if __name__ == "__main__":
    unittest.main()
