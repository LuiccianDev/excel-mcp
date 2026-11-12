"""Security regression tests for formula injection and active worksheet guards."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_mcp.core.batch import batch_write_matrix
from excel_mcp.core.calculations import (
    FormulaSecurityError,
    validate_cell_value_secure,
)
from excel_mcp.core.data import write_data


def test_validate_cell_value_secure_allows_normal_values() -> None:
    assert validate_cell_value_secure(42)[0] is True
    assert validate_cell_value_secure("Normal text")[0] is True
    assert validate_cell_value_secure("+1-800-555-0199")[0] is True
    assert validate_cell_value_secure("-50.25")[0] is True
    assert validate_cell_value_secure("=SUM(A1:A10)")[0] is True


def test_write_data_blocks_dangerous_formula(tmp_path: Path) -> None:
    target = tmp_path / "write_data_evil.xlsx"
    with pytest.raises(FormulaSecurityError):
        write_data(
            target, "Sheet1", [['=CALL("Kernel32","WinExec","JJ","calc.exe",1)']]
        )


def test_batch_write_blocks_leading_whitespace_formula(tmp_path: Path) -> None:
    target = tmp_path / "batch_whitespace_evil.xlsx"
    wb = Workbook()
    wb.save(target)
    with pytest.raises(FormulaSecurityError):
        batch_write_matrix(
            str(target),
            None,
            "A1",
            [['  =CALL("Kernel32","WinExec","JJ","calc.exe",1)']],
        )


def test_batch_write_blocks_dde_plus_prefix_injection(tmp_path: Path) -> None:
    target = tmp_path / "batch_dde_evil.xlsx"
    wb = Workbook()
    wb.save(target)
    with pytest.raises(FormulaSecurityError):
        batch_write_matrix(
            str(target),
            None,
            "A1",
            [["+cmd|'/C calc'!A0"]],
        )
