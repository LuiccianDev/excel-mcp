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
  [![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
  [![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-brightgreen)](https://modelcontextprotocol.io)

</div>

**Excel MCP Server** (`excel-mcp`, PyPI: [`luiccian-excel-mcp`](https://pypi.org/project/luiccian-excel-mcp/)) is a high-performance Python server implementing the Model Context Protocol (MCP) to provide advanced, token-optimized Excel spreadsheet (`.xlsx`) creation, batch synthesis, formula automation, and inspection capabilities for AI assistants (Claude, Cursor, Copilot, Cline).

Designed from the ground up for maximum token efficiency, it replaces chatty micro-iterations with single-pass workbook construction—saving over 90% of model tokens and conversational turns.

---

## Key Architectural Highlights

- **Single-Pass Workbook Builder**: Create complete, multi-sheet, publication-ready workbooks (title banners, table headers, alternating row colors, accounting borders, auto-calculated totals, currency formatting, and native charts) in **one single MCP call** (`excel_build_workbook`).
- **Visual Design Presets**: Built-in typography and visual standards (`corporate`, `financial`, `modern`, `compact`).
- **Lightweight Outline Inspection**: Extract sheet names, populated dimensions (`A1:H50`), column headers, and formula summaries without dumping massive cell grids (`excel_get_workbook_outline`), saving thousands of context tokens.
- **Batch Data Ingestion**: Append or overwrite 2D matrices of values into existing sheets in a single atomic operation (`excel_batch_write_data`).
- **Formula & Analytics Engine**: Apply validated formulas (`SUM`, `AVERAGE`, `VLOOKUP`), generate native bar/line/pie charts, and build analytical pivot tables.
- **Zero-Install Execution**: Instant launch with `uvx luiccian-excel-mcp`.

---

## Quick Start with `uvx` (Recommended)

No manual clone or virtual environment management required. You can run the server directly using [`uv`](https://docs.astral.sh/uv/):

```bash
# Run with stdio transport (default for AI IDEs)
uvx luiccian-excel-mcp
```

---

## MCP Client Configuration

### 1. Claude Desktop

Add this configuration to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "uvx",
      "args": ["luiccian-excel-mcp"],
      "env": {
        "MCP_ALLOWED_DIRECTORIES": "/path/to/your/spreadsheets"
      }
    }
  }
}
```

> [!TIP]
> **Local checkout:** If running from a local git repository:
> ```json
> {
>   "mcpServers": {
>     "excel-mcp": {
>       "command": "uv",
>       "args": ["run", "--directory", "/path/to/excel-mcp", "excel-mcp"],
>       "env": {
>         "MCP_ALLOWED_DIRECTORIES": "/path/to/your/spreadsheets"
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
        "MCP_ALLOWED_DIRECTORIES": "${workspaceFolder}"
      }
    }
  }
}
```

---

## Visual Design Presets

| Preset | Header Fill | Header Text | Zebra Row Fill | Typography | Borders / Totals | Best For |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `corporate` (Default) | Navy Blue (`#1F4E79`) | White | Ice Blue (`#EDF2F8`) | Calibri 11pt | Thin gray, soft blue total | Business proposals, management reports, executive summaries |
| `financial` | Forest Green (`#1E4620`) | White | Pale Mint (`#F0F6F1`) | Calibri 11pt, currency format | Double bottom line on totals | Balance sheets, P&L statements, budgeting, cash flow |
| `modern` | Slate / Charcoal (`#2F3542`) | White | Cool Gray (`#F8F9FA`) | Segoe UI 10.5pt | Minimalist soft borders | SaaS metrics, tech product analytics, modern dashboards |
| `compact` | Steel Gray (`#4A6572`) | White | Off-White (`#F5F5F5`) | Arial 9.5pt | Dense grid, high data density | Inventory lists, master catalogs, high-volume exports |

---

## Modular Agent Skills

The repository includes pre-built agent skills under `skills/`:

- **`excel-spreadsheet-creator`**: Direct single-pass synthesis for reports, budgets, and financial statements.
- **`excel-spreadsheet-editor`**: Lightweight structural inspection, atomic batch matrix writes, formatting, and pivot tables.
- **`excel-spreadsheet-analyzer`**: Context-efficient sampling, formula verification, and structural auditing.

---

## Tool Taxonomy

All tools support both canonical `excel_` prefixes and legacy names for 100% backward compatibility:

| Category | Primary Tool | Description |
| :--- | :--- | :--- |
| **Batch / Single-Pass** | `excel_build_workbook` | Build full multi-sheet workbooks with styles, formulas, and charts in 1 call |
| | `excel_get_workbook_outline` | Lightweight outline inspection without loading entire cell matrices |
| | `excel_batch_write_data` | Write 2D matrix of values or formulas into a sheet range atomically |
| **Workbook Operations** | `excel_create_excel_workbook` | Create a new Excel workbook file |
| | `excel_create_excel_worksheet` | Add a new worksheet to a workbook |
| | `excel_list_excel_documents` | List available Excel files in allowed directories |
| | `excel_copy_worksheet` | Duplicate a worksheet within a workbook |
| | `excel_delete_worksheet` | Delete a worksheet |
| | `excel_rename_worksheet` | Rename an existing worksheet |
| | `excel_get_workbook_metadata` | Retrieve sheet names, active sheet, and file size |
| **Content & Ranges** | `excel_read_data_from_excel` | Read specific cell range data |
| | `excel_write_data_to_excel` | Write individual cell values |
| | `excel_copy_range` | Copy cells, formats, and formulas across ranges |
| | `excel_delete_range` | Delete range of cells |
| | `excel_merge_cells` | Merge cell ranges |
| | `excel_unmerge_cells` | Unmerge cell ranges |
| **Formatting** | `excel_format_range_excel` | Apply custom fonts, fills, borders, alignments, and number formats |
| **Formulas** | `excel_apply_formula_excel` | Apply calculation formulas to cells |
| | `excel_validate_formula_syntax`| Validate formula syntax before applying |
| **Graphics & Analytics** | `excel_create_chart` | Generate native bar, line, or pie charts from data |
| | `excel_create_pivot_table` | Create analytical pivot tables |

---

## Security

- **Restricted Directories**: Access is strictly limited to paths listed in `MCP_ALLOWED_DIRECTORIES` (or `DIRECTORY`).
- **Formula Sanitization**: Formula inputs are validated to avoid dangerous injection patterns.
- **Trusted Publishing**: Automated PyPI releases use GitHub Actions OIDC and PEP 740 artifact attestation.

---

## Development

```bash
# Clone the repository
git clone https://github.com/LuiccianDev/excel-mcp.git
cd excel-mcp

# Install dependencies with Astral uv
uv sync --all-groups

# Run linting
uv run ruff check src

# Run test suite
uv run pytest
```

---

## License

This project is licensed under the [MIT License](LICENSE).
