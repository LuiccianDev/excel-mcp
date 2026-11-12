---
name: excel-spreadsheet-editor
description: "Edit, update, format, and enhance existing Excel spreadsheets (.xlsx) efficiently. Performs lightweight outline inspection, batch matrix writes, row appending, dimension management, table styling, and in-place sorting."
argument-hint: "Target workbook path, worksheet name, edit operation, and data/ranges to update"
user-invocable: true
disable-model-invocation: false
---

# Excel Spreadsheet Editor

Use this skill to modify, format, and enhance existing Microsoft Excel workbooks while preserving token efficiency.

---

## Core Rules

1. **Always Inspect Outline Before Modifying:**
   - Always call `excel_get_workbook_outline(filename)` first. It returns dimensions (e.g., `A1:F50`), detected column headers, and sheet names without loading full cell grids into context.
2. **Prefer Atomic Batch Operations:**
   - Use `excel_append_rows` to add new records below existing content automatically without manually tracking row indices.
   - Use `excel_batch_write_data` to update existing matrices in a single atomic pass.
3. **Preserve Uniform Typography:**
   - Match the workbook's uniform font (default `Calibri 11pt` or `Arial 11pt`, or the user's explicitly requested font) via `excel_format_range`. Never introduce mixed font families or oversized (>12pt) cells.
4. **Maintain Uncluttered Sheet Structure:**
   - Do not cram unrelated tables or configuration blocks onto an already-crowded worksheet. Create a dedicated worksheet (`excel_create_worksheet`) or leave clear separator rows, and apply `freeze_panes` + `auto_filter` via `excel_format_sheet`.
5. **100% Dynamic Calculations (Zero Hardcoded Results):**
   - When adding or updating totals, subtotals, variances, or KPIs via `excel_batch_write_data`, `excel_append_rows`, or `excel_apply_formula`, always write live Excel formulas (`=SUM(...)`, `=IFERROR(...)`) instead of static numbers.
6. **Format Sheets Holistically:**
   - Use `excel_format_sheet` to freeze headers, enable autofilters, and set column widths in one call.
   - Use `excel_create_table` to convert ranges into native Excel tables with banded styling and filter drop-downs.

---

## Canonical Editing Protocols

### Protocol 1: Appending Rows Automatically

Append new records below existing table data in one call:

```json
{
  "filename": "/path/to/inventory.xlsx",
  "sheet_name": "Stock",
  "rows": [
    ["SKU-9021", "USB-C Docking Station", "Hardware", 120, 89.50],
    ["SKU-9022", "Mechanical Keyboard", "Hardware", 85, 115.00],
    ["SKU-9023", "27-inch 4K Monitor", "Hardware", 40, 320.00]
  ]
}
```
**Tool:** `excel_append_rows`

---

### Protocol 2: Native Excel Table Conversion

Turn a raw tabular range into a professional Excel Table:

```json
{
  "filename": "/path/to/inventory.xlsx",
  "sheet_name": "Stock",
  "range_string": "A1:E43",
  "table_name": "InventoryTable",
  "style_name": "TableStyleMedium9"
}
```
**Tool:** `excel_create_table`

---

### Protocol 3: Sheet-Level Layout & Display Formatting

Freeze top rows and set custom column widths in a single call:

```json
{
  "filename": "/path/to/inventory.xlsx",
  "sheet_name": "Stock",
  "freeze_panes": "A2",
  "auto_filter": true,
  "show_grid_lines": true,
  "column_widths": {
    "A": 15.0,
    "B": 30.0,
    "C": 18.0,
    "D": 12.0,
    "E": 14.0
  }
}
```
**Tool:** `excel_format_sheet`

---

### Protocol 4: In-Place Range Sorting

Sort tabular rows by key column while preserving header rows:

```json
{
  "filename": "/path/to/inventory.xlsx",
  "sheet_name": "Stock",
  "range_string": "A1:E43",
  "key_column": "Unit Price",
  "ascending": false,
  "has_header": true
}
```
**Tool:** `excel_sort_range`

---

### Protocol 5: Protecting and Preparing a Printable Worksheet

Apply worksheet protection and print layout in dedicated operations:

```json
{
  "filename": "/path/to/report.xlsx",
  "sheet_name": "Report",
  "enabled": true,
  "password": "template-secret",
  "allow_sort": true
}
```
**Tool:** `excel_protect_sheet`

```json
{
  "filename": "/path/to/report.xlsx",
  "sheet_name": "Report",
  "orientation": "landscape",
  "paper_size": "A4",
  "fit_to_width": 1,
  "print_area": "A1:H50",
  "header": "&CQuarterly Report"
}
```
**Tool:** `excel_configure_page_setup`

Use `excel_add_data_validation` to constrain editable cells before protecting a template.

---

### Protocol 6: Generating Analytical Pivot Tables

Summarize large data sheets using native pivot tables:

```json
{
  "filename": "/path/to/sales_ledger.xlsx",
  "source_sheet": "Transactions",
  "target_sheet": "Pivot_Department_Summary",
  "source_range": "A1:F500",
  "target_cell": "A3",
  "rows": ["Department"],
  "columns": ["Quarter"],
  "values": ["Revenue"],
  "aggfunc": "sum"
}
```
**Tool:** `excel_create_pivot_table`
