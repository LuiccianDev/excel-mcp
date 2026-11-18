<div align="center">
  <h1>Excel MCP Server</h1>
  <p>
    <em>High-performance Model Context Protocol (MCP) server for programmatic Excel (.xlsx) spreadsheet generation and manipulation</em>
  </p>

  [![PyPI Version](https://img.shields.io/pypi/v/luiccian-excel-mcp.svg)](https://pypi.org/project/luiccian-excel-mcp/)
  [![Python Version](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/)
  [![uv](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/uv/main/assets/badge/v0.json)](https://github.com/astral-sh/uv)
  [![CodeQL](https://github.com/LuiccianDev/excel-mcp/actions/workflows/codeql.yml/badge.svg)](https://github.com/LuiccianDev/excel-mcp/actions/workflows/codeql.yml)
  [![Security Policy](https://img.shields.io/badge/security-policy-blue.svg)](SECURITY.md)
  [![Rules for AI](https://img.shields.io/badge/rules-AGENTS.md-blue.svg)](AGENTS.md)
  [![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
  [![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-brightgreen)](https://modelcontextprotocol.io)

</div>

**Excel MCP Server** (`excel-mcp`, PyPI: [`luiccian-excel-mcp`](https://pypi.org/project/luiccian-excel-mcp/)) is a high-performance Python server implementing the Model Context Protocol (MCP) to provide advanced, token-optimized Excel spreadsheet (`.xlsx`) creation, batch synthesis, formula automation, and inspection capabilities for AI assistants (Claude Desktop, Claude Code, Cursor, VS Code / GitHub Copilot, Antigravity, Windsurf, Cline).

Designed from the ground up for maximum token efficiency, it replaces chatty micro-iterations with single-pass workbook construction—saving over 90% of model tokens and conversational turns.

---

## Key Architectural Highlights

- **Compact Compound Mode by Default**: Registers 8 tools by default (3 single-pass tools plus 5 action-dispatched domain tools: `excel_workbook`, `excel_data`, `excel_format`, `excel_formula`, `excel_analytics`) to minimize tool-schema token overhead, with opt-in to all 33 granular tools via `EXCEL_MCP_TOOLS=full` or `server --tool-mode full`.
- **Single-Pass Workbook Builder**: Create complete, multi-sheet, publication-ready workbooks (compact `<=12pt` title banners, table headers, alternating row colors, accounting borders, configurable column aggregations via `auto_totals`, currency formatting, inline `data_validations`, per-sheet `freeze_panes` / `auto_filter`, and multiple native `charts`) in **one single MCP call** (`excel_build_workbook`).
- **Uniform Typography & Design Presets**: Built-in visual standards (`corporate`, `financial`, `modern`, `compact`) paired with optional `font_name` and `font_size` overrides to enforce uniform `Calibri 11pt`, `Arial 11pt`, or any user-requested font across the entire workbook.
- **Dynamic Formula Discipline**: Native support for validated live Excel formulas (`SUM`, `AVERAGE`, `IFERROR`, `SUMIF`, cross-sheet references) without hardcoded static results.
- **Lightweight Outline Inspection**: Extract sheet names, populated dimensions (`A1:H50`), column headers, and formula presence without dumping massive cell grids (`excel_get_workbook_outline`), saving thousands of context tokens.
- **Batch Data Ingestion**: Append or overwrite 2D matrices of values into existing sheets in a single atomic operation (`excel_batch_write_data`, `excel_append_rows`).
- **Zero-Install Execution**: Instant launch with `uvx luiccian-excel-mcp` (see also [docs/MCPB.md](docs/MCPB.md)).

---

## Quick Start with `uvx` (Recommended)

No manual clone or virtual environment management required. Run the server directly using Astral [`uv`](https://docs.astral.sh/uv/):

```bash
# Run with stdio transport (default for AI IDEs)
uvx luiccian-excel-mcp
```

---

## MCP Client Configuration

Configure your AI client by passing environment variables directly in the client's `env` block (no `.env` file is required). By default, the server starts in `compound` mode (8 tools); set `"EXCEL_MCP_TOOLS": "full"` in `env` or pass `--tool-mode full` (e.g., `["luiccian-excel-mcp", "--tool-mode", "full"]` or `server --tool-mode full`) to expose all 33 granular tools.

### 1. Claude Desktop (`claude_desktop_config.json`)

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

> [!TIP]
> **Windows path example:** Use forward slashes or escaped backslashes (e.g., `"C:/Users/YourName/Documents/Spreadsheets"`). You can allow multiple roots separated by commas.
>
> **Local checkout:** If running from a local git repository instead of PyPI:
> ```json
> {
>   "mcpServers": {
>     "excel-mcp": {
>       "command": "uv",
>       "args": ["run", "--directory", "/path/to/excel-mcp", "excel-mcp"],
>       "env": {
>         "EXCEL_MCP_DIRS": "/path/to/your/spreadsheets"
>       }
>     }
>   }
> }
> ```

### 2. Cursor (`.cursor/mcp.json`)

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

### 3. VS Code / GitHub Copilot (`.vscode/mcp.json`)

VS Code uses a top-level `"servers"` key and explicit `"type": "stdio"`:

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

### 4. Antigravity / Windsurf (`mcp_config.json`)

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

### 5. Cline / Roo Code (`cline_mcp_settings.json`)

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

### 6. Claude Code CLI

```bash
claude mcp add excel-mcp \
  -e EXCEL_MCP_DIRS="/path/to/your/spreadsheets" \
  -- uvx luiccian-excel-mcp
```

---

## Environment Variables

All runtime configuration is supplied via standard OS environment variables or the MCP client's `env` object:

| Variable | Default | Description |
| :--- | :--- | :--- |
| `EXCEL_MCP_DIRS` | *(current directory)* | Comma-separated list of allowed root directories for reading, writing, and listing `.xlsx` files. |
| `EXCEL_MCP_TOOLS` | `compound` | Tool surface mode: `compound` (8 tools: 3 single-pass + 5 domain dispatchers) or `full` (all 33 granular tools). |
| `LOG_LEVEL` | `INFO` | Logging verbosity (`DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`). |
| `DEV_MODE` | `false` | Enables verbose development diagnostics when set to `true`. |

---

## Visual Design Presets

| Preset | Header Fill | Header Text | Zebra Row Fill | Default Font | Borders / Totals | Best For |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `corporate` *(Default)* | Navy Blue (`#1F4E79`) | White | Ice Blue (`#EDF2F8`) | Calibri 11pt | Thin gray, soft blue total | Business proposals, management reports, executive summaries |
| `financial` | Forest Green (`#1E4620`) | White | Pale Mint (`#F0F6F1`) | Calibri 11pt | Double bottom line on totals | Balance sheets, P&L statements, budgeting, cash flow |
| `modern` | Slate / Charcoal (`#2F3542`) | White | Cool Gray (`#F8F9FA`) | Segoe UI 10.5pt | Minimalist soft borders | SaaS metrics, tech product analytics, modern dashboards |
| `compact` | Steel Gray (`#4A6572`) | White | Off-White (`#F5F5F5`) | Arial 9.5pt | Dense grid, high data density | Inventory lists, master catalogs, high-volume exports |

> [!TIP]
> Pass `font_name` and `font_size` to `excel_build_workbook` to apply any color preset with uniform typography (e.g., `preset: "modern", font_name: "Arial", font_size: 11`) or to honor a user's explicit font request.

---

## Modular Agent Skills

The repository includes pre-built agent skills under `skills/`:

- **`excel-tools-reference`**: Authoritative encyclopedia and decision matrix covering the default `compound` surface (8 tools) and all 33 canonical tools in `full` mode, signatures, spreadsheet engineering standards, and safety rules.
- **`excel-spreadsheet-creator`**: Single-pass synthesis featuring 3 distinct layout patterns (Operations Ledger, Multi-Tab Model with `Assumptions`, and KPI Dashboard), uniform `11pt` default typography (honoring user font overrides), uncluttered layouts, and 100% dynamic Excel formulas.
- **`excel-spreadsheet-editor`**: Lightweight structural inspection, atomic batch matrix writes, row appending, typography preservation, uncluttered organization, and in-place sorting.
- **`excel-spreadsheet-analyzer`**: Context-efficient sampling, statistical column profiling (`excel_profile_sheet`), targeted cell discovery (`excel_find_cells`), and hardcoded-calculation / layout hygiene audits.

---

## Tool Taxonomy

The server defaults to `compound` mode (8 tools) and exposes all 33 granular `excel_*` tools when `EXCEL_MCP_TOOLS=full` or `server --tool-mode full` is set. Each domain dispatcher accepts `(action, filename=None, params=None)` and transparently normalizes `range_string` / `cell_range` parameter aliases; for the underlying `full`-mode parameter schemas (`params` keys), see [docs/TOOLS.md](docs/TOOLS.md).

### Default `compound` Mode (8 Tools)

| Category | Tool Name | Type | Actions / Purpose |
| :--- | :--- | :--- | :--- |
| **Single-Pass** | `excel_build_workbook` | `Write` | Build full multi-sheet workbooks with styles, uniform fonts, navigation, formulas, and charts in 1 call |
| | `excel_get_workbook_outline` | `Read-Only` | Lightweight outline inspection without loading entire cell matrices |
| | `excel_batch_write_data` | `Write` | Write 2D matrix of values or formulas into a sheet range atomically |
| **Domain Dispatchers** | `excel_workbook` | `Write` / `Read-Only` | `create`, `list`, `metadata`, `create_sheet`, `copy_sheet`, `delete_sheet`, `rename_sheet` |
| | `excel_data` | `Write` / `Read-Only` | `read`, `write`, `append_rows`, `copy_range`, `delete_range`, `merge_cells`, `unmerge_cells`, `insert_rows_columns`, `delete_rows_columns`, `sort_range`, `validate_range`, `find_cells`, `profile_sheet` |
| | `excel_format` | `Write` | `format_range`, `format_sheet`, `create_table`, `add_data_validation`, `protect_sheet`, `configure_page_setup` |
| | `excel_formula` | `Write` / `Read-Only` | `validate`, `apply` |
| | `excel_analytics` | `Write` | `create_chart`, `create_pivot_table` |

### Full Mode Granular Inventory (`EXCEL_MCP_TOOLS=full` / `server --tool-mode full`, 33 Tools)

| Category | Canonical Tool Name | Type | Description |
| :--- | :--- | :--- | :--- |
| **Batch & Single-Pass** | `excel_build_workbook` | `Write` | Build full multi-sheet workbooks with styles, uniform fonts, navigation, formulas, and charts in 1 call |
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
| **Worksheet Management**| `excel_create_workbook` | `Write` | Create a new Excel workbook file |
| | `excel_create_worksheet` | `Write` | Add a new worksheet to a workbook |
| | `excel_copy_worksheet` | `Write` | Duplicate a worksheet within a workbook |
| | `excel_delete_worksheet` | `Destructive` | Permanently delete a worksheet |
| | `excel_rename_worksheet` | `Write` | Rename an existing worksheet |
| | `excel_insert_rows_columns` | `Write` | Insert empty rows or columns at 1-based index |
| | `excel_delete_rows_columns` | `Destructive` | Delete rows or columns starting at 1-based index |
| | `excel_list_documents` | `Read-Only` | List available Excel files in allowed directories |
| | `excel_get_workbook_metadata` | `Read-Only` | Retrieve sheet names, dimensions, and named ranges |
| **Formulas & Analytics**| `excel_validate_formula_syntax`| `Read-Only` | Validate formula syntax before applying |
| | `excel_apply_formula` | `Write` | Apply calculation formulas to specific cells |
| | `excel_create_chart` | `Write` | Generate native bar, line, or pie charts from data |
| | `excel_create_pivot_table` | `Write` | Create analytical pivot tables from ranges |
| **Content I/O** | `excel_read_data` | `Read-Only` | Read tabular data from worksheet range |
| | `excel_write_data` | `Write` | Write data starting at target cell |
| **Validation & Print Layout** | `excel_add_data_validation` | `Write` | Add list, numeric, date, or custom validation rules to a cell range |
| | `excel_protect_sheet` | `Write` | Protect or unprotect a worksheet with action-level permissions |
| | `excel_configure_page_setup` | `Write` | Configure print orientation, scaling, area, headers, footers, and centering |

---

## Security

- **Restricted Directories**: Access is strictly limited to paths listed in `EXCEL_MCP_DIRS` (supporting multiple comma-separated roots and symlink traversal protection).
- **Formula & DDE Sanitization**: All formula inputs pass strict AST validation (`validate_cell_value_secure`) blocking leading-whitespace formula smuggling, DDE command prefixes (`cmd|`, `powershell|`), and dangerous functions.
- **Crash-Safe Atomic Writes**: Mutating operations write to isolated temporary files and commit atomically via `os.replace`.
- **Trusted Publishing**: Tag-driven PyPI releases (`v*.*.*`) use GitHub Actions OIDC and PEP 740 artifact attestation.

---

## Development

```bash
# Clone the repository
git clone https://github.com/LuiccianDev/excel-mcp.git
cd excel-mcp

# Install dependencies with Astral uv
uv sync --all-groups

# Run linting and formatting
uv run ruff check .
uv run ruff format --check .

# Run static type checking
uv run mypy src

# Verify repository consistency (tools, manifest, docs)
uv run python scripts/check-repository-consistency.py

# Run test suite with coverage gate (80% minimum)
uv run pytest --cov --cov-report=term-missing

# Build the wheel and source distribution
uv run python -m build
```

> [!NOTE]
> The commands above run through `uv` and are the local source of truth for CI across Python 3.11–3.14. A clean checkout is reproducible with `uv sync --all-groups`.

---

## License

This project is licensed under the [Apache License 2.0](LICENSE).
