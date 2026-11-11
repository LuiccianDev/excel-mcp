"""Deterministic functional benchmarks for the main workbook operations."""

from __future__ import annotations

import json
import time
import tracemalloc
from pathlib import Path
from typing import Any, Callable

import openpyxl

from excel_mcp.core.batch import batch_write_matrix, build_single_pass_workbook
from excel_mcp.core.formatting import format_range


ROOT = Path(__file__).resolve().parents[2]
REPORT = ROOT / "docs" / "benchmarks" / "workbook-baseline.json"


def _measure(operation: Callable[[], Any]) -> tuple[float, int, Any]:
    tracemalloc.start()
    started = time.perf_counter()
    result = operation()
    elapsed_ms = (time.perf_counter() - started) * 1000
    _, peak_bytes = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    return elapsed_ms, peak_bytes, result


def _baseline(name: str) -> dict[str, float]:
    report = json.loads(REPORT.read_text(encoding="utf-8"))
    return next(item for item in report["scenarios"] if item["name"] == name)


def test_batch_write_10000_cells_is_functionally_complete(tmp_path: Path) -> None:
    filename = tmp_path / "batch.xlsx"
    openpyxl.Workbook().save(filename)
    matrix = [[row * 100 + column for column in range(100)] for row in range(100)]

    elapsed_ms, peak_bytes, result = _measure(
        lambda: batch_write_matrix(str(filename), None, "A1", matrix)
    )

    assert result.cells_updated == 10_000
    assert elapsed_ms > 0
    assert peak_bytes > 0
    baseline = _baseline("batch_write_10000_cells")
    assert elapsed_ms <= baseline["median_runtime_ms"] * 10
    assert peak_bytes <= baseline["peak_memory_bytes"] * 10
    assert openpyxl.load_workbook(filename).active["CV100"].value == 9_999


def test_multi_sheet_build_is_functionally_complete(tmp_path: Path) -> None:
    filename = tmp_path / "build.xlsx"
    rows = [[f"item-{index}", index] for index in range(100)]
    sheets = [
        {"name": "First", "headers": ["Name", "Value"], "rows": rows},
        {"name": "Second", "headers": ["Name", "Value"], "rows": rows},
    ]

    elapsed_ms, peak_bytes, result = _measure(
        lambda: build_single_pass_workbook(str(filename), sheets, auto_fit=False)
    )

    assert result.sheets_created == ["First", "Second"]
    assert result.total_rows_inserted == 200
    assert elapsed_ms > 0
    assert peak_bytes > 0
    baseline = _baseline("multi_sheet_build")
    assert elapsed_ms <= baseline["median_runtime_ms"] * 10
    assert peak_bytes <= baseline["peak_memory_bytes"] * 10
    assert set(openpyxl.load_workbook(filename).sheetnames) == {"First", "Second"}


def test_format_and_save_is_functionally_complete(tmp_path: Path) -> None:
    filename = tmp_path / "format.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active.append(["Name", "Value"])
    workbook.active.append(["item", 10])
    workbook.save(filename)

    elapsed_ms, peak_bytes, result = _measure(
        lambda: format_range(
            str(filename),
            "Sheet",
            "A1",
            "B2",
            bold=True,
            bg_color="D9EAF7",
            number_format="0.00",
        )
    )

    assert result["status"] == "success"
    assert elapsed_ms > 0
    assert peak_bytes > 0
    baseline = _baseline("format_and_save")
    assert elapsed_ms <= baseline["median_runtime_ms"] * 10
    assert peak_bytes <= baseline["peak_memory_bytes"] * 10
    assert openpyxl.load_workbook(filename).active["A1"].font.bold is True


def test_baseline_report_has_measured_scenarios() -> None:
    assert REPORT.exists()
    report = json.loads(REPORT.read_text(encoding="utf-8"))
    assert all(item["median_runtime_ms"] > 0 for item in report["scenarios"])
    assert all(item["peak_memory_bytes"] > 0 for item in report["scenarios"])
    assert {item["name"] for item in report["scenarios"]} == {
        "batch_write_10000_cells",
        "multi_sheet_build",
        "format_and_save",
    }
