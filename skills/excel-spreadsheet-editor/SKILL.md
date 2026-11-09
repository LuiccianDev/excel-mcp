---
name: excel-spreadsheet-editor
description: "Edit, update, format, and enhance existing Excel spreadsheets (.xlsx) efficiently. Performs lightweight outline inspection, batch matrix writes, formula additions, range formatting, and pivot table generation."
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
2. **Prefer Atomic Batch Writes:**
   - Use `excel_batch_write_data` to update tables, append rows, or insert matrices in a single call instead of issuing repetitive cell-by-cell write operations.
3. **Validate Complex Formulas First:**
   - For custom formulas or conditional expressions, verify validity with `excel_validate_formula_syntax` before inserting into worksheets.

---

## Canonical Editing Protocols

### Protocol 1: Inspecting & Appending Batch Data

```json
// Step 1: Inspect structure
{
  "filename": "/path/to/inventory.xlsx"
}
// Tool: excel_get_workbook_outline
// Response indicates sheet "Stock" spans A1:E40. New rows should start at A41.

// Step 2: Write batch rows starting at A41
{
  "filename": "/path/to/inventory.xlsx",
  "sheet_name": "Stock",
  "start_cell": "A41",
  "data_matrix": [
    ["SKU-9021", "USB-C Docking Station", "Hardware", 120, 89.50],
    ["SKU-9022", "Mechanical Keyboard", "Hardware", 85, 115.00],
    ["SKU-9023", "27-inch 4K Monitor", "Hardware", 40, 320.00]
  ]
}
// Tool: excel_batch_write_data
```

---

### Protocol 2: Formatting & Auto-Fitting Columns

Apply professional currency styling to column `E` and auto-fit column widths:

```json
// Format currency range
{
  "filename": "/path/to/inventory.xlsx",
  "sheet_name": "Stock",
  "start_cell": "E2",
  "end_cell": "E43",
  "number_format": "$#,##0.00"
}
// Tool: excel_format_range_excel

// Auto-fit all columns in the sheet
{
  "filename": "/path/to/inventory.xlsx",
  "sheet_name": "Stock"
}
// Tool: excel_auto_fit_columns
```

---

### Protocol 3: Generating Analytical Pivot Tables

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
// Tool: excel_create_pivot_table
```

*Valid `aggfunc` options:* `"sum"`, `"mean"`, `"count"`, `"min"`, `"max"`, `"std"`.

---

### Protocol 4: Worksheet Management

- **Duplicate template sheet:** `excel_copy_sheet(filename, "Template", "Q1_Report")`
- **Rename sheet:** `excel_rename_sheet(filename, "Sheet1", "Dashboard")`
- **Delete obsolete sheet:** `excel_delete_sheet(filename, "Old_Draft")`
