# Security Hardening, Quality Gates and Tool Expansion Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Unify path security on `MCP_ALLOWED_DIRECTORIES`, remove protocol-breaking stdout writes, replace pre-commit with uv-native quality gates, raise test coverage, and add three publication-grade Excel tools.

**Architecture:** A single `utils/paths.py` resolver becomes the only source of truth for allowed directories; `security.py` and `file_utils.py` delegate to it. CLI startup output moves to stderr so stdio JSON-RPC stays pure. CI runs ruff (lint + format), mypy and pytest with a coverage floor across Python 3.11-3.13. Three new openpyxl-native tools (data validation, conditional formatting, page setup) follow the existing `core/` + `tools/` + `register_tools.py` pattern.

**Tech Stack:** Python 3.11+, FastMCP, openpyxl, pytest, pytest-cov, ruff, mypy, uv, GitHub Actions.

**Spec:** `docs/superpowers/specs/2026-11-10-security-hardening-design.md`

## Global Constraints

- Repository language: Technical English for code, comments, commits and docs. User replies in Spanish.
- No informal emojis anywhere.
- Runtime code lives only in `src/excel_mcp/`; tests in `tests/`; scripts in `scripts/` (kebab-case filenames).
- Tools are registered with the `excel_` canonical prefix; legacy aliases stay opt-in via `EXCEL_MCP_LEGACY_ALIASES=1`.
- `MCP_ALLOWED_DIRECTORIES` takes precedence over `DIRECTORY`, which takes precedence over the configured default `./documents`.
- Every task ends green: `uv run pytest -q`, `uv run ruff check .`, `uv run ruff format --check .`, `uv run mypy src`.
- Commit messages follow Conventional Commits; documentary-only commits end with `[skip ci]`.

## Review Focus

- A path outside every allowed directory (including `..` traversal and symlink escape) must be rejected by every tool, not only by the ones using `security.py`.
- A stdio client that sends `initialize` must receive only JSON-RPC lines on stdout, on the first launch after install.
- `MCP_ALLOWED_DIRECTORIES` with several comma-separated roots must authorize all of them, and a root with a trailing separator must behave identically.
- New tools must reject malformed input (bad range syntax, unknown enum value, invalid hex color) with a structured `{"status": "error"}` response instead of raising.
- `manifest.json` must never advertise a tool name that `register_all_tools` does not register.

---

### Task 1: Unified allowed-directory resolver

**Files:**
- Create: `src/excel_mcp/utils/paths.py`
- Modify: `src/excel_mcp/utils/security.py`
- Modify: `src/excel_mcp/utils/file_utils.py:25-67`
- Test: `tests/test_path_security.py`

**Interfaces:**
- Consumes: `excel_mcp.config.get_directory()`.
- Produces: `resolve_allowed_directories() -> list[str]` and `is_path_allowed(path: str | Path) -> tuple[bool, str | None]` in `excel_mcp.utils.paths`; `security.py` keeps `get_allowed_directories()`, `is_path_in_allowed_directories()`, `resolve_and_check_path()`.

- [ ] **Step 1: Write the failing tests**

```python
"""Tests for the unified allowed-directory resolver."""

from __future__ import annotations

import asyncio

import openpyxl
import pytest

from excel_mcp.utils.paths import is_path_allowed, resolve_allowed_directories


def test_mcp_allowed_directories_takes_precedence(monkeypatch, tmp_path):
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    monkeypatch.setenv("DIRECTORY", str(tmp_path / "legacy"))
    assert resolve_allowed_directories() == [str(tmp_path.resolve())]


def test_directory_env_is_used_as_fallback(monkeypatch, tmp_path):
    monkeypatch.delenv("MCP_ALLOWED_DIRECTORIES", raising=False)
    monkeypatch.setenv("DIRECTORY", str(tmp_path))
    assert resolve_allowed_directories() == [str(tmp_path.resolve())]


def test_multiple_roots_and_trailing_separator(monkeypatch, tmp_path):
    first, second = tmp_path / "a", tmp_path / "b"
    first.mkdir()
    second.mkdir()
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", f"{first}{os.sep}, {second}")
    assert resolve_allowed_directories() == [str(first.resolve()), str(second.resolve())]


def test_traversal_outside_root_is_rejected(monkeypatch, tmp_path):
    root = tmp_path / "root"
    root.mkdir()
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(root))
    allowed, _ = is_path_allowed(str(root / ".." / "escape.xlsx"))
    assert allowed is False


def test_read_tool_accepts_path_under_mcp_allowed_directories(monkeypatch, tmp_path):
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    monkeypatch.delenv("DIRECTORY", raising=False)
    workbook = tmp_path / "book.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.save(workbook)

    from excel_mcp.tools.content_tools import read_data_from_excel

    result = asyncio.run(
        read_data_from_excel(filename=str(workbook), sheet_name="Sheet")
    )
    assert result["status"] == "success"
```

Add `import os` to the test module.

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_path_security.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'excel_mcp.utils.paths'`.

- [ ] **Step 3: Implement `src/excel_mcp/utils/paths.py`**

```python
"""Single source of truth for filesystem sandboxing."""

from __future__ import annotations

import os
from pathlib import Path

from excel_mcp.config import get_directory


def _split(value: str) -> list[str]:
    return [item.strip() for item in value.split(",") if item.strip()]


def resolve_allowed_directories() -> list[str]:
    """Return resolved, de-duplicated allowed roots.

    Precedence: MCP_ALLOWED_DIRECTORIES, then DIRECTORY, then the configured
    default directory. A trailing separator is normalized by Path.resolve().
    """
    raw = os.environ.get("MCP_ALLOWED_DIRECTORIES", "").strip()
    candidates = _split(raw)
    if not candidates:
        candidates = _split(os.environ.get("DIRECTORY", "").strip())
    if not candidates:
        candidates = [get_directory()]

    resolved: list[str] = []
    for candidate in candidates:
        path = str(Path(candidate).resolve())
        if path not in resolved:
            resolved.append(path)
    return resolved


def is_path_allowed(path: str | Path) -> tuple[bool, str | None]:
    """Return whether path lives inside one allowed root, plus an error message."""
    target = Path(path).resolve()
    allowed = resolve_allowed_directories()
    for root in allowed:
        root_path = Path(root)
        if target == root_path or target.is_relative_to(root_path):
            return True, None
    return False, f"Path '{path}' is not in allowed directories: {', '.join(allowed)}"
```

- [ ] **Step 4: Delegate `security.py` and `file_utils.py` to the resolver**

In `security.py`, `get_allowed_directories()` returns `resolve_allowed_directories()` and `is_path_in_allowed_directories()` returns `is_path_allowed(file_path)`. In `file_utils.py`, `_get_allowed_directories()` returns `resolve_allowed_directories()` and `_is_path_in_allowed_directories()` returns `is_path_allowed(file_path)`. Remove the now-duplicated environment parsing and the unused `get_directory` import.

- [ ] **Step 5: Run the tests to verify they pass**

Run: `uv run pytest tests/test_path_security.py tests/test_file_security.py tests/test_file_utils.py tests/test_batch_tools.py -q`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add src/excel_mcp/utils/paths.py src/excel_mcp/utils/security.py src/excel_mcp/utils/file_utils.py tests/test_path_security.py
git commit -m "fix(security): unify allowed-directory resolution on MCP_ALLOWED_DIRECTORIES"
```

---

### Task 2: Pure stdio transport and accurate server version

**Files:**
- Modify: `src/excel_mcp/__main__.py:109-126, 224-238`
- Modify: `src/excel_mcp/server.py:70-82`
- Modify: `src/excel_mcp/__init__.py:7-15`
- Test: `tests/test_cli.py`

**Interfaces:**
- Consumes: `excel_mcp.__version__`, `excel_mcp.server.create_server()`.
- Produces: `main()` writes nothing to stdout when starting the server; `create_server()` passes `version=excel_mcp.__version__` to `FastMCP`.

- [ ] **Step 1: Write the failing tests**

```python
"""CLI and server construction tests."""

from __future__ import annotations

import sys

import excel_mcp
import excel_mcp.server as server_module


class _FakeServer:
    def run(self, transport: str = "stdio") -> None:
        return None


def test_main_keeps_stdout_clean(monkeypatch, capsys):
    from excel_mcp import __main__ as cli

    monkeypatch.setattr(sys, "argv", ["excel-mcp"])
    monkeypatch.setattr(cli, "run_server", lambda: _FakeServer())
    cli.main()
    captured = capsys.readouterr()
    assert captured.out == ""
    assert "MCP" in captured.err or captured.err == ""


def test_create_server_reports_package_version(monkeypatch):
    captured = {}

    class Recorder:
        def __init__(self, **kwargs):
            captured.update(kwargs)

        def tool(self, *args, **kwargs):
            def decorator(fn):
                return fn

            return decorator

    monkeypatch.setattr(server_module, "FastMCP", Recorder)
    server_module.create_server()
    assert captured["version"] == excel_mcp.__version__
    assert captured["name"] == "excel_mcp"


def test_instructions_do_not_claim_aliases_are_enabled(monkeypatch):
    captured = {}

    class Recorder:
        def __init__(self, **kwargs):
            captured.update(kwargs)

        def tool(self, *args, **kwargs):
            def decorator(fn):
                return fn

            return decorator

    monkeypatch.setattr(server_module, "FastMCP", Recorder)
    server_module.create_server()
    assert "also registered" not in captured["instructions"]
    assert "EXCEL_MCP_LEGACY_ALIASES" in captured["instructions"]
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_cli.py -q`
Expected: FAIL (stdout contains "Starting MCP Excel Office Server...", `version` missing).

- [ ] **Step 3: Move startup messages to stderr and pass the version**

In `__main__.py`, delete the five `typer.echo(...)` calls in `run_mcp_server` and `main()`; log them with `logger.info(...)` instead. Keep `list_excel_files` printing to stdout. In `server.py`, import `from excel_mcp import __version__` and call `FastMCP(name="excel_mcp", instructions=instructions, version=__version__, on_duplicate="error")`. Replace the final "Tool Naming" paragraph with a sentence stating that legacy aliases are opt-in via `EXCEL_MCP_LEGACY_ALIASES=1`.

- [ ] **Step 4: Make the metadata fallback honest**

In `__init__.py`, change the final fallback to `"0.0.0+unknown"` so a stale hard-coded version can never be reported.

- [ ] **Step 5: Run the tests to verify they pass**

Run: `uv run pytest tests/test_cli.py -q`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add src/excel_mcp/__main__.py src/excel_mcp/server.py src/excel_mcp/__init__.py tests/test_cli.py
git commit -m "fix(server): keep stdio stdout pure and report the package version"
```

---

### Task 3: Coverage for chart and formatting cores

**Files:**
- Test: `tests/test_chart.py`
- Test: `tests/test_formatting.py`

**Interfaces:**
- Consumes: `excel_mcp.core.chart.create_chart_from_data`, `excel_mcp.core.formatting.*`.

- [ ] **Step 1: Read both modules and list their public callables**

Run: `uv run python -c "import excel_mcp.core.chart as c, excel_mcp.core.formatting as f; print([n for n in dir(c) if not n.startswith('_')]); print([n for n in dir(f) if not n.startswith('_')])"`

- [ ] **Step 2: Write failing tests**

One test per public callable: build a workbook in `tmp_path`, apply the callable, reopen with `openpyxl.load_workbook` and assert the persisted property (chart count/type, font name, fill color, border style, number format). Tests must assert persisted workbook state, not internal dictionaries.

- [ ] **Step 3: Run the tests to verify they fail**

Run: `uv run pytest tests/test_chart.py tests/test_formatting.py -q`
Expected: FAIL with `ImportError` or `AttributeError` for names that do not exist; fix the test to the real signature and rerun until the failure is an assertion, not an import error.

- [ ] **Step 4: Fix any real defects the tests expose**

If a test reveals a genuine bug, fix the module. Do not weaken the assertion to match broken behavior.

- [ ] **Step 5: Run the tests to verify they pass**

Run: `uv run pytest tests/test_chart.py tests/test_formatting.py -q --cov=excel_mcp.core.chart --cov=excel_mcp.core.formatting --cov-report=term-missing`
Expected: PASS, each module at or above 70 %.

- [ ] **Step 6: Commit**

```bash
git add tests/test_chart.py tests/test_formatting.py src/excel_mcp/core
git commit -m "test(core): cover chart and formatting persistence paths"
```

---

### Task 4: uv-native quality gates and pre-commit removal

**Files:**
- Delete: `.pre-commit-config.yaml`
- Modify: `pyproject.toml:35-43, 122-128`
- Modify: `.github/workflows/ci.yml`
- Modify: `scripts/bump-version.py:21`
- Modify: `scripts/generate-tools-doc.py:21`
- Modify: `README.md:169-184`
- Modify: `CHANGELOG.md`

- [ ] **Step 1: Remove pre-commit and add the coverage floor**

Delete `.pre-commit-config.yaml`. In `pyproject.toml`, drop `pre-commit` from `[dependency-groups].dev` and append:

```toml
[tool.coverage.run]
source = ["excel_mcp"]
branch = false

[tool.coverage.report]
fail_under = 75
show_missing = true
skip_covered = true
```

Then run `uv lock` and `uv sync --all-groups`.

- [ ] **Step 2: Fix the two lint errors**

Delete the unused `import sys` in `scripts/bump-version.py`. Append `# noqa: E402` to the `from excel_mcp.tools.register_tools import register_all_tools` line in `scripts/generate-tools-doc.py`.

- [ ] **Step 3: Rewrite the CI job**

`ci.yml` keeps `paths-ignore` and `concurrency`; the `test` job becomes a matrix over `["3.11", "3.12", "3.13"]` with `fail-fast: false`, and runs, in order: `uv sync --all-groups`, `uv run ruff check .`, `uv run ruff format --check .`, `uv run mypy src`, `uv run pytest --cov --cov-report=term-missing`. Remove the `.pre-commit-config.yaml` entry from both `paths-ignore` blocks. Name the job `Test & Lint (Python ${{ matrix.python-version }})`.

- [ ] **Step 4: Format the repository**

Run: `uv run ruff format .` then `uv run ruff check . --fix`.
Expected: `ruff check` reports zero errors.

- [ ] **Step 5: Verify the gates locally**

Run: `uv run ruff check . && uv run ruff format --check . && uv run mypy src && uv run pytest -q --cov --cov-report=term`
Expected: all pass; coverage at or above 75 %. If coverage is below 75 %, add tests to the weakest module rather than lowering the floor; record a ruling if 75 proves unreachable.

- [ ] **Step 6: Update README and CHANGELOG**

README development section documents `uv run ruff check .`, `uv run ruff format .`, `uv run mypy src`, `uv run pytest --cov` and states that pre-commit was removed in favour of uv-native gates. CHANGELOG `[Unreleased]` gains a `Changed` entry for the CI matrix and the removal of pre-commit, and a `Fixed` entry for the lint errors in `scripts/`.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "chore(ci): replace pre-commit with uv-native lint, format, type and coverage gates"
```

---

### Task 5: `excel_add_data_validation`

**Files:**
- Create: `src/excel_mcp/core/data_validation.py`
- Create: `src/excel_mcp/tools/validation_tools.py`
- Modify: `src/excel_mcp/tools/register_tools.py:70-79`
- Test: `tests/test_data_validation.py`

**Interfaces:**
- Produces: `core.data_validation.add_data_validation(filename: str, cell_range: str, validation_type: str, operator: str = "between", formula1: str | None = None, formula2: str | None = None, allow_blank: bool = True, prompt_title: str | None = None, prompt: str | None = None, error_title: str | None = None, error_message: str | None = None, sheet_name: str | None = None) -> dict[str, Any]` returning `{"status": "success", "data": {"sheet": str, "range": str, "validation_type": str}}`; tool `excel_add_data_validation` with the same parameters plus `filename`/`file_path` aliases.

- [ ] **Step 1: Write the failing tests**

Cases: list validation persists `DataValidation.type == "list"` and `formula1 == '"Open,Closed"'`; whole-number between persists `operator == "between"`; a `sheet_name` that does not exist returns `{"status": "error"}`; an unknown `validation_type` returns `{"status": "error"}`; a missing `formula1` returns `{"status": "error"}`. Reopen the workbook with `openpyxl.load_workbook` and assert against `ws.data_validations.dataValidation`.

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_data_validation.py -q`
Expected: FAIL with `ModuleNotFoundError`.

- [ ] **Step 3: Implement the core module and tool**

Use `openpyxl.worksheet.datavalidation.DataValidation` with `add(cell_range)`, `ws.add_data_validation(dv)`, and `atomic_save` from `excel_mcp.core._io`. Validate `validation_type` against `{"list", "whole", "decimal", "date", "time", "textLength", "custom"}` and `operator` against `{"between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"}`, raising `ValidationError` from `excel_mcp.exceptions.exception_tools` for unknown values. The tool wrapper uses `@tool_guard()` and `resolve_and_check_path`.

- [ ] **Step 4: Register the tool**

Add `_reg("excel_add_data_validation", validation_tools.add_data_validation, write, ["add_data_validation"])` under a new `# 8. VALIDATION & LAYOUT TOOLS` section.

- [ ] **Step 5: Run the tests to verify they pass**

Run: `uv run pytest tests/test_data_validation.py tests/test_register_tools.py -q`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add src/excel_mcp/core/data_validation.py src/excel_mcp/tools/validation_tools.py src/excel_mcp/tools/register_tools.py tests/test_data_validation.py
git commit -m "feat(tools): add excel_add_data_validation"
```

---

### Task 6: `excel_protect_sheet` (swapped from `excel_add_conditional_format`)

> [!NOTE]
> The original Task 6 proposed `excel_add_conditional_format`, but conditional
> formatting is already reachable through `excel_format_range(conditional_format=...)`
> (cellIs, formula, colorScale, dataBar, iconSet). Adding a second tool for the
> same capability would duplicate the surface. This task therefore ships sheet
> protection instead, which had no coverage at all.

**Files:**
- Create: `src/excel_mcp/core/protection.py`
- Create: `src/excel_mcp/tools/protection_tools.py`
- Modify: `src/excel_mcp/tools/register_tools.py`
- Test: `tests/test_protection.py`

**Interfaces:**
- Produces: `core.protection.protect_sheet(filename: str, enabled: bool = True, password: str | None = None, sheet_name: str | None = None, allow_select_locked_cells: bool = True, allow_select_unlocked_cells: bool = True, allow_format_cells: bool = False, allow_format_columns: bool = False, allow_format_rows: bool = False, allow_insert_columns: bool = False, allow_insert_rows: bool = False, allow_insert_hyperlinks: bool = False, allow_delete_columns: bool = False, allow_delete_rows: bool = False, allow_sort: bool = False, allow_auto_filter: bool = False, allow_pivot_tables: bool = False, allow_edit_objects: bool = False, allow_edit_scenarios: bool = False) -> dict[str, Any]`; tool `excel_protect_sheet`.

- [ ] **Step 1: Write the failing tests**

Cases: `enabled=True` with a password persists `ws.protection.sheet is True` and a hashed password; `allow_sort=True` persists `ws.protection.sort is False`; `enabled=False` removes protection from an already protected sheet; a `sheet_name` that does not exist returns `{"status": "error"}`; a missing filename returns `{"status": "error"}`.

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_protection.py -q`
Expected: FAIL with `ModuleNotFoundError`.

- [ ] **Step 3: Implement the core module and tool**

Use `openpyxl.worksheet.protection.SheetProtection` through `ws.protection`. In openpyxl the boolean attributes mean "locked/disallowed", so each `allow_*` flag maps to the negation of its attribute. Hash the password with openpyxl's built-in `SheetProtection.set_password` (assigning `ws.protection.password` does this automatically). Use `atomic_save` from `excel_mcp.core._io`, raise `SheetNotFoundError` for unknown sheets, and wrap the tool with `@tool_guard()` plus `resolve_and_check_path`.

- [ ] **Step 4: Register the tool and run the tests**

Add `_reg("excel_protect_sheet", protection_tools.protect_sheet, write, ["protect_sheet"])` under `# 8. VALIDATION & LAYOUT TOOLS`. Run: `uv run pytest tests/test_protection.py tests/test_register_tools.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/excel_mcp/core/protection.py src/excel_mcp/tools/protection_tools.py src/excel_mcp/tools/register_tools.py tests/test_protection.py
git commit -m "feat(tools): add excel_protect_sheet"
```

---

### Task 7: `excel_configure_page_setup`

**Files:**
- Create: `src/excel_mcp/core/page_setup.py`
- Create: `src/excel_mcp/tools/page_tools.py`
- Modify: `src/excel_mcp/tools/register_tools.py`
- Test: `tests/test_page_setup.py`

**Interfaces:**
- Produces: `core.page_setup.configure_page_setup(filename: str, orientation: str = "portrait", paper_size: str = "A4", fit_to_width: int | None = 1, fit_to_height: int | None = None, print_area: str | None = None, print_title_rows: str | None = None, header: str | None = None, footer: str | None = None, center_horizontally: bool = False, sheet_name: str | None = None) -> dict[str, Any]`; tool `excel_configure_page_setup`.

- [ ] **Step 1: Write the failing tests**

Cases: landscape orientation and `fit_to_width=1` persist with `ws.page_setup.orientation == "landscape"` and `ws.sheet_properties.pageSetUpPr.fitToPage is True`; `paper_size="A3"` maps to the openpyxl constant; `print_area="A1:H50"` persists as `"'Sheet'!$A$1:$H$50"`; `header="&CReport"` persists in `ws.oddHeader.center.text`; an invalid orientation returns `{"status": "error"}`.

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_page_setup.py -q`
Expected: FAIL with `ModuleNotFoundError`.

- [ ] **Step 3: Implement the core module and tool**

Map paper sizes `{"A4": "9", "A3": "8", "Letter": "1", "Legal": "5"}` to `ws.page_setup.paperSize`, set `ws.sheet_properties.pageSetUpPr.fitToPage = True` when a fit dimension is given, and validate ranges with `excel_mcp.utils.validation_utils`. Use `atomic_save`.

- [ ] **Step 4: Register the tool and run the tests**

Run: `uv run pytest tests/test_page_setup.py tests/test_register_tools.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/excel_mcp/core/page_setup.py src/excel_mcp/tools/page_tools.py src/excel_mcp/tools/register_tools.py tests/test_page_setup.py
git commit -m "feat(tools): add excel_configure_page_setup"
```

---

### Task 8: Manifest, documentation and skills synchronization

**Files:**
- Modify: `manifest.json`
- Modify: `docs/TOOLS.md` (regenerated)
- Modify: `scripts/generate-tools-doc.py:52-98`
- Modify: `skills/excel-tools-reference/SKILL.md`
- Modify: `skills/excel-spreadsheet-editor/SKILL.md`
- Modify: `README.md:122-158`
- Modify: `tests/test_skills_in_sync.py:40-42`
- Test: `tests/test_manifest.py`

**Interfaces:**
- Consumes: the 33 canonical tool names from Task 5-7.

- [ ] **Step 1: Write the failing tests**

`tests/test_manifest.py` asserts: `manifest.json["version"]` equals the `pyproject.toml` version; every `manifest.json["tools"][i]["name"]` is in the registered canonical tool set; the canonical count is 33.

- [ ] **Step 2: Run the tests to verify they fail**

Run: `uv run pytest tests/test_manifest.py tests/test_skills_in_sync.py -q`
Expected: FAIL on the stale version `2.0.0`, on `build_sheet`, and on the count `30`.

- [ ] **Step 3: Update the manifest, the generator and the docs**

Set `manifest.json` version to the `pyproject.toml` value, replace the tool list with the 33 canonical names, and keep `tools_generated: true`. In `generate-tools-doc.py`, add a `"Validation & Print Layout"` category for the three new tools and make `categorize_tools` emit any uncategorized tool under `"Other Tools"` so future tools can never disappear from the document. Run `uv run python scripts/generate-tools-doc.py`. Add the three tools to the skills reference and the editor skill, and add three rows to the README tool table.

- [ ] **Step 4: Update the drift-guard count and run the tests**

Change `test_canonical_tool_count` to 33. Run: `uv run pytest tests/test_manifest.py tests/test_skills_in_sync.py -q`
Expected: PASS.

- [ ] **Step 5: Update CHANGELOG and commit**

Add `[Unreleased]` entries: `Added` for the three tools, `Fixed` for the sandbox variable, stdout purity, version reporting and manifest drift. Then:

```bash
git add -A
git commit -m "docs(tools): synchronize manifest, tool reference and skills with 33 canonical tools"
```

---

### Task 9: Full verification

- [ ] **Step 1: Run every gate**

Run: `uv run ruff check . && uv run ruff format --check . && uv run mypy src && uv run pytest -q --cov --cov-report=term-missing`
Expected: zero lint errors, zero type errors, all tests pass, coverage at or above 75 %.

- [ ] **Step 2: Reproduce the original defects against the fixed code**

Run the three reproductions from the audit: only `MCP_ALLOWED_DIRECTORIES` set (read + list must succeed), stdio `initialize` probe (stdout must contain only the JSON-RPC line), and `manifest.json` version comparison. Expected: all three now behave correctly.

- [ ] **Step 3: Commit any residual fix**

```bash
git add -A
git commit -m "chore: final verification pass"
```
