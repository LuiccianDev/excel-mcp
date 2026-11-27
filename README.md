<div align="center">
<h1>Excel MCP Server</h1>
<p><em>High-performance Model Context Protocol (MCP) server for programmatic Excel (.xlsx) spreadsheet generation, analysis, and manipulation.</em></p>
</div>

<p align="center">
  <a href="https://pypi.org/project/luiccian-excel-mcp/"><img src="https://img.shields.io/pypi/v/luiccian-excel-mcp.svg?logo=pypi&logoColor=FFE873" alt="PyPI version" /></a>
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.11+-3776AB?style=flat&logo=python&logoColor=white" alt="Python 3.11+" /></a>
  <a href="https://github.com/astral-sh/uv"><img src="https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/uv/main/assets/badge/v0.json" alt="uv" /></a>
  <a href="https://github.com/LuiccianDev/excel-mcp/actions/workflows/ci.yml"><img src="https://github.com/LuiccianDev/excel-mcp/actions/workflows/ci.yml/badge.svg" alt="CI" /></a>
  <a href="https://github.com/LuiccianDev/excel-mcp/actions/workflows/codeql.yml"><img src="https://github.com/LuiccianDev/excel-mcp/actions/workflows/codeql.yml/badge.svg" alt="CodeQL" /></a>
  <a href="https://github.com/LuiccianDev/excel-mcp/blob/master/SECURITY.md"><img src="https://img.shields.io/badge/security-policy-blue.svg" alt="Security Policy" /></a>
  <a href="https://github.com/LuiccianDev/excel-mcp/blob/master/LICENSE"><img src="https://img.shields.io/badge/License-Apache_2.0-blue.svg" alt="License" /></a>
</p>

Connects AI coding agents (Antigravity, Claude Desktop, Claude Code, Cursor, VS Code, Windsurf, Cline) to local `.xlsx` workbooks with **33 canonical tools across 5 domain dispatchers and 3 single-pass builders**, selectable via `EXCEL_MCP_TOOLS` (`compound` by default with 8 tools, or `full` with 33 tools) and sandboxed to `EXCEL_MCP_DIRS`.

---

## Quick Start

### 1. Connect the MCP Server (`uvx`, Recommended)

Requires [`uv`](https://docs.astral.sh/uv/). Runs [`luiccian-excel-mcp`](https://pypi.org/project/luiccian-excel-mcp/) on demand over `stdio` with zero local virtual environment setup:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "uvx",
      "args": ["luiccian-excel-mcp"],
      "env": {
        "EXCEL_MCP_DIRS": "/path/to/your/spreadsheets"
      }
    }
  }
}
```

> **Note:** On Windows, use forward slashes or escaped backslashes (for example `"C:/Users/YourName/Documents/Spreadsheets"`). You can allow multiple roots separated by commas. Set `"EXCEL_MCP_TOOLS": "full"` in `env` or pass `server --tool-mode full` when a client requires all 33 granular tools.

### 2. CLI Commands & Modes

```bash
# Start MCP server on stdio (default compound mode, 8 tools)
uvx luiccian-excel-mcp

# Start MCP server exposing all 33 canonical tools
uvx luiccian-excel-mcp server --tool-mode full

# List .xlsx files in configured EXCEL_MCP_DIRS (or --path)
uvx luiccian-excel-mcp list --path /path/to/your/spreadsheets
```

---

## Architecture & Tool Modes

### Selective Tool Loading (`EXCEL_MCP_TOOLS`)

| Mode | Tools Exposed | Behavior |
| :--- | :--- | :--- |
| `compound` *(Default)* | 8 tools | Exposes 3 single-pass tools (`excel_build_workbook`, `excel_get_workbook_outline`, `excel_batch_write_data`) plus 5 action-dispatched domain routers (`excel_workbook`, `excel_data`, `excel_format`, `excel_formula`, `excel_analytics`) taking `(action, filename=None, params=None)`. |
| `full` | 33 tools | Opt-in via `EXCEL_MCP_TOOLS=full` or `server --tool-mode full`; exposes all 33 canonical `excel_*` tool schemas directly. |

### Environment Variables

| Variable | Default | Description |
| :--- | :--- | :--- |
| `EXCEL_MCP_DIRS` | `~/Documents` *(or `./documents` fallback)* | Comma-separated list of allowed root directories for reading, writing, and listing `.xlsx` files. |
| `EXCEL_MCP_TOOLS` | `compound` | Tool surface mode: `compound` (8 tools) or `full` (33 tools). |
| `LOG_LEVEL` | `INFO` | Logging verbosity (`DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`). |
| `EXCEL_MCP_METRICS_LOG` | *(unset)* | Optional path to a rotating JSONL file for development tool-call metrics logging. |

### Visual Design Presets

| Preset | Header Fill | Header Text | Zebra Row Fill | Default Font | Borders / Totals | Best For |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `corporate` *(Default)* | Navy Blue (`#1F4E79`) | White | Ice Blue (`#EDF2F8`) | Calibri 11pt | Thin gray, soft blue total | Business proposals, management reports, executive summaries |
| `financial` | Forest Green (`#1E4620`) | White | Pale Mint (`#F0F6F1`) | Calibri 11pt | Double bottom line on totals | Balance sheets, P&L statements, budgeting, cash flow |
| `modern` | Slate / Charcoal (`#2F3542`) | White | Cool Gray (`#F8F9FA`) | Segoe UI 10.5pt | Minimalist soft borders | SaaS metrics, tech product analytics, modern dashboards |
| `compact` | Steel Gray (`#4A6572`) | White | Off-White (`#F5F5F5`) | Arial 9.5pt | Dense grid, high data density | Inventory lists, master catalogs, high-volume exports |

> **Tip:** Pass `font_name` and `font_size` to `excel_build_workbook` to override typography uniformly across a preset (for example `preset="modern"`, `font_name="Arial"`, `font_size=11`), and configure `auto_totals`, inline `data_validations`, `freeze_panes`, `auto_filter`, and multiple `charts` in one pass.

### Domain Catalog (8 Compound Tools / 33 Canonical Actions)

| Tool (`compound` Mode) | Type | Actions / Purpose |
| :--- | :--- | :--- |
| `excel_build_workbook` | `Destructive` | Build full multi-sheet workbooks with styles, uniform fonts, navigation, formulas, validations, and charts in 1 call |
| `excel_get_workbook_outline` | `Read-Only` | Lightweight outline inspection (sheet names, dimensions, headers, formula counts) without loading full cell matrices |
| `excel_batch_write_data` | `Write` | Write a 2D matrix of values or formulas into a sheet range atomically |
| `excel_workbook` | `Destructive` / `Read-Only` | `create`, `list`, `metadata`, `create_sheet`, `copy_sheet`, `delete_sheet`, `rename_sheet` |
| `excel_data` | `Destructive` / `Read-Only` | `read`, `write`, `append_rows`, `copy_range`, `delete_range`, `merge_cells`, `unmerge_cells`, `insert_rows_columns`, `delete_rows_columns`, `sort_range`, `validate_range`, `find_cells`, `profile_sheet` |
| `excel_format` | `Write` | `format_range`, `format_sheet`, `create_table`, `add_data_validation`, `protect_sheet`, `configure_page_setup` |
| `excel_formula` | `Write` / `Read-Only` | `validate`, `apply` |
| `excel_analytics` | `Write` | `create_chart`, `create_pivot_table` |

### Canonical Tool Inventory (`EXCEL_MCP_TOOLS=full`, 33 Tools)

| Category | Canonical Tool Name | Type | Description |
| :--- | :--- | :--- | :--- |
| **Batch & Single-Pass** | `excel_build_workbook` | `Destructive` | Build full multi-sheet workbooks with styles, uniform fonts, navigation, formulas, and charts in 1 call |
| | `excel_get_workbook_outline` | `Read-Only` | Lightweight outline inspection without loading entire cell matrices |
| | `excel_batch_write_data` | `Write` | Write 2D matrix of values or formulas into a sheet range atomically |
| | `excel_append_rows` | `Write` | Append batch rows below existing data automatically detecting last row |
| **Search & Profiling** | `excel_profile_sheet` | `Read-Only` | Column data types, null counts, min/max, and distinct samples |
| | `excel_find_cells` | `Read-Only` | Search cells by text, exact match, or regex without full grid dump |
| **Formatting & Tables** | `excel_create_table` | `Write` | Convert cell range into native Excel Table with filter drop-downs |
| | `excel_format_sheet` | `Write` | Configure freeze panes, autofilter, gridlines, and custom column widths |
| | `excel_sort_range` | `Write` | Sort 2D cell range in-place by key column with header preservation |
| | `excel_format_range` | `Write` | Apply fonts, fills, borders, alignments, and number formats |
| | `excel_merge_cells` | `Write` | Merge rectangular cell range |
| | `excel_unmerge_cells` | `Write` | Unmerge cell range |
| | `excel_copy_range` | `Write` | Copy cells, formats, and formulas across ranges |
| | `excel_delete_range` | `Destructive` | Delete range of cells with optional cell shifting |
| | `excel_validate_range` | `Read-Only` | Validate range coordinate syntax |
| **Worksheet Management** | `excel_create_workbook` | `Destructive` | Create a new Excel workbook file |
| | `excel_create_worksheet` | `Write` | Add a new worksheet to a workbook |
| | `excel_copy_worksheet` | `Write` | Duplicate a worksheet within a workbook |
| | `excel_delete_worksheet` | `Destructive` | Permanently delete a worksheet |
| | `excel_rename_worksheet` | `Write` | Rename an existing worksheet |
| | `excel_insert_rows_columns` | `Write` | Insert empty rows or columns at 1-based index |
| | `excel_delete_rows_columns` | `Destructive` | Delete rows or columns starting at 1-based index |
| | `excel_list_documents` | `Read-Only` | List available Excel files in allowed directories |
| | `excel_get_workbook_metadata` | `Read-Only` | Retrieve sheet names, dimensions, and named ranges |
| **Formulas & Analytics** | `excel_validate_formula_syntax` | `Read-Only` | Validate formula syntax before applying |
| | `excel_apply_formula` | `Write` | Apply calculation formulas to specific cells |
| | `excel_create_chart` | `Write` | Generate native bar, line, pie, or scatter charts from data |
| | `excel_create_pivot_table` | `Write` | Create analytical pivot tables from ranges |
| **Content I/O** | `excel_read_data` | `Read-Only` | Read tabular data from worksheet range |
| | `excel_write_data` | `Write` | Write data starting at target cell |
| **Validation & Print Layout** | `excel_add_data_validation` | `Write` | Add list, numeric, date, or custom validation rules to a cell range |
| | `excel_protect_sheet` | `Write` | Protect or unprotect a worksheet with action-level permissions |
| | `excel_configure_page_setup` | `Write` | Configure print orientation, scaling, area, headers, footers, and centering |

---

## Alternative Client Configurations & Local Development

<details>
<summary>Cursor, VS Code, Antigravity, Windsurf, Cline, Claude Code CLI, and Local Python Development (`uv run`)</summary>

### Cursor (`.cursor/mcp.json`)

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "uvx",
      "args": ["luiccian-excel-mcp"],
      "env": {
        "EXCEL_MCP_DIRS": "${workspaceFolder}"
      }
    }
  }
}
```

### VS Code / GitHub Copilot (`.vscode/mcp.json`)

```json
{
  "servers": {
    "excel-mcp": {
      "type": "stdio",
      "command": "uvx",
      "args": ["luiccian-excel-mcp"],
      "env": {
        "EXCEL_MCP_DIRS": "${workspaceFolder}"
      }
    }
  }
}
```

### Antigravity / Windsurf (`mcp_config.json`) & Cline (`cline_mcp_settings.json`)

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "uvx",
      "args": ["luiccian-excel-mcp"],
      "env": {
        "EXCEL_MCP_DIRS": "/path/to/your/spreadsheets"
      }
    }
  }
}
```

### Claude Code CLI

```bash
claude mcp add excel-mcp \
  -e EXCEL_MCP_DIRS="/path/to/your/spreadsheets" \
  -- uvx luiccian-excel-mcp
```

### Local Development (`uv run`)

```bash
git clone https://github.com/LuiccianDev/excel-mcp.git
cd excel-mcp
uv sync --all-groups

uv run ruff check .
uv run ruff format --check .
uv run mypy src
uv run python scripts/check-repository-consistency.py
uv run pytest --cov --cov-report=term-missing
uv run python -m build
```

</details>

---

## Security

- **Restricted Directories**: Access is confined to `EXCEL_MCP_DIRS` (supporting multiple comma-separated roots, mixed-slash UNC blocking, symlink traversal checks, and filtering of unexpanded `${...}` placeholders).
- **Formula & DDE Sanitization**: All formula inputs, conditional formatting formulas, `data_validation` rules, and batch banners/headers pass `validate_formula_secure`, blocking leading-whitespace smuggling, `_xlfn.`/`_xlws.`/`_xlpm.`/`_xleta.` prefixes, DDE command prefixes (`cmd|`, `powershell|`), and dangerous functions.
- **Resource & Regex Bounds**: `workbook_session` validates ZIP entry counts, uncompressed sizes, and aggregate compression ratios against archive bombs; range operations enforce cell limits; search queries enforce depth-aware ReDoS protection.
- **Crash-Safe Atomic Writes**: Mutating operations write to isolated temporary files and commit atomically via `os.replace`, with `destructive_hint=True` annotations on file-creating and destructive tools.
- **Trusted Publishing**: Tag-driven PyPI releases (`v*.*.*`) use GitHub Actions OIDC and PEP 740 artifact attestation.

---

## Documentation, Skills & Governance

| Resource | Purpose |
| :--- | :--- |
| [MCP Tools Reference](https://github.com/LuiccianDev/excel-mcp/blob/master/docs/TOOLS.md) | Generated parameter signatures and access modes for all 33 canonical tools |
| [MCPB Packaging Guide](https://github.com/LuiccianDev/excel-mcp/blob/master/docs/MCPB.md) | Zero-install `uvx` execution and bundle manifest (`manifest.json`) setup |
| [Excel Tools Reference Skill](https://github.com/LuiccianDev/excel-mcp/blob/master/skills/excel-tools-reference/SKILL.md) | Decision matrix, `compound`/`full` inventory, and `params` conventions |
| [Excel Spreadsheet Creator Skill](https://github.com/LuiccianDev/excel-mcp/blob/master/skills/excel-spreadsheet-creator/SKILL.md) | Single-pass layout patterns, uniform typography, and dynamic formulas |
| [Excel Spreadsheet Editor Skill](https://github.com/LuiccianDev/excel-mcp/blob/master/skills/excel-spreadsheet-editor/SKILL.md) | Structural outline inspection, atomic matrix writes, and in-place sorting |
| [Excel Spreadsheet Analyzer Skill](https://github.com/LuiccianDev/excel-mcp/blob/master/skills/excel-spreadsheet-analyzer/SKILL.md) | Column profiling, regex/text cell search, and formula hygiene audits |
| [Security Policy](https://github.com/LuiccianDev/excel-mcp/blob/master/SECURITY.md) | Vulnerability reporting and supported security controls |
| [Changelog](https://github.com/LuiccianDev/excel-mcp/blob/master/CHANGELOG.md) | Version history following Keep a Changelog and SemVer |

Licensed under the [Apache License 2.0](https://github.com/LuiccianDev/excel-mcp/blob/master/LICENSE). See [CONTRIBUTING.md](https://github.com/LuiccianDev/excel-mcp/blob/master/CONTRIBUTING.md) and [AGENTS.md](https://github.com/LuiccianDev/excel-mcp/blob/master/AGENTS.md).
