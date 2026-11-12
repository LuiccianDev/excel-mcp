"""Tests for atomic file I/O operations in core/_io.py."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl

from excel_mcp.core._io import atomic_save


class TestAtomicSave(unittest.TestCase):
    """Test suite for atomic_save utility."""

    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.target_file = Path(self.temp_dir.name) / "test_wb.xlsx"

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_atomic_save_success(self) -> None:
        """Test successful atomic serialization."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Atomic Write Test"

        saved_path = atomic_save(wb, self.target_file)
        self.assertTrue(saved_path.exists())
        self.assertEqual(saved_path, self.target_file.resolve())

        # Verify no residual temp files
        temp_files = list(Path(self.temp_dir.name).glob(".tmp_*"))
        self.assertEqual(len(temp_files), 0)

        # Verify content
        loaded_wb = openpyxl.load_workbook(saved_path)
        self.assertEqual(loaded_wb.active["A1"].value, "Atomic Write Test")
        loaded_wb.close()

    def test_atomic_save_cleanup_on_failure(self) -> None:
        """Test temp file cleanup when serialization fails."""
        wb = openpyxl.Workbook()

        with patch.object(
            wb, "save", side_effect=OSError("Simulated disk write failure")
        ):
            with self.assertRaises(OSError):
                atomic_save(wb, self.target_file)

        # Destination file should not exist
        self.assertFalse(self.target_file.exists())
        # Temp file should be cleaned up
        temp_files = list(Path(self.temp_dir.name).glob(".tmp_*"))
        self.assertEqual(len(temp_files), 0)


if __name__ == "__main__":
    unittest.main()
