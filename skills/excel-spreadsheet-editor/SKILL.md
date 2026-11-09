---
name: excel-spreadsheet-editor
description: "Edit, update, format, and enhance existing Excel spreadsheets (.xlsx) efficiently. Performs lightweight outline inspection, batch matrix writes, formula additions, range formatting, and pivot table generation."
argument-hint: "Target workbook path, worksheet name, edit operation, and data/ranges to update"
user-invocable: true
disable-model-invocation: false
---

# Excel Spreadsheet Editor

Use this skill to modify and format existing Microsoft Excel workbooks while preserving token efficiency.

## Core Rules

1. **Inspect before modifying:**
   - Always call `excel_get_workbook_outline(filename)` first to understand the sheet structure, dimensions, and headers without dumping thousands of rows into context.
2. **Prefer atomic batch writes over cell-by-cell loops:**
   - Use `excel_batch_write_data` to write 2D arrays or rows at once rather than repeatedly calling `write_data_to_excel`.
3. **Always validate formulas before bulk application:**
   - Use `excel_validate_formula_syntax` if creating complex calculation expressions.

## Canonical Editing Workflow

### 1. Structure Inspection

```json
{
  "filename": "/path/to/financial_model.xlsx"
}
```
Tool: `excel_get_workbook_outline`

### 2. Batch Data Append / Update

Write multiple rows or a table update starting from cell `A10`:

```json
{
  "filename": "/path/to/financial_model.xlsx",
  "sheet_name": "Revenue Summary",
  "start_cell": "A10",
  "data_matrix": [
    ["November 2026", 85000, 92000, "=C10-B10"],
    ["December 2026", 95000, 104000, "=C11-B11"]
  ]
}
```
Tool: `excel_batch_write_data`

### 3. Visual Styling & Formatting

Apply consistent currency formatting or headers:

```json
{
  "filename": "/path/to/financial_model.xlsx",
  "sheet_name": "Revenue Summary",
  "start_cell": "B10",
  "end_cell": "D11",
  "number_format": "$#,##0.00"
}
```
Tool: `excel_format_range_excel`

### 4. Pivot Table Summary

Generate an analytical pivot table from tabular data:

```json
{
  "filename": "/path/to/financial_model.xlsx",
  "source_sheet": "Transactions",
  "target_sheet": "Pivot_Summary",
  "source_range": "A1:E500",
  "target_cell": "A3",
  "rows": ["Department"],
  "columns": ["Quarter"],
  "values": ["Amount"],
  "aggfunc": "sum"
}
```
Tool: `excel_create_pivot_table`
