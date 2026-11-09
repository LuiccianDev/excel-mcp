---
name: excel-spreadsheet-analyzer
description: "Inspect, audit, and extract insights from Excel workbooks (.xlsx) without exhausting model context. Uses lightweight structural outlines, targeted range sampling, formula verification, and metadata extraction."
argument-hint: "Target workbook path, worksheet to analyze, and audit objective"
user-invocable: true
disable-model-invocation: false
---

# Excel Spreadsheet Analyzer

Use this skill to extract structural insights, audit formulas, and sample data from Microsoft Excel workbooks without consuming excessive tokens.

## Core Rules

1. **Never read the entire sheet blindly:**
   - Avoid requesting unbounded ranges like `A1:Z50000`.
   - Start by calling `excel_get_workbook_outline` to determine the exact populated bounds (`dimensions`), total rows, and column headers.
2. **Sample targeted ranges:**
   - Once headers and dimensions are identified, read only representative slices (e.g. `A1:H10` for sample rows, or specific summary columns).
3. **Audit calculations:**
   - Verify formula definitions and syntax with `excel_validate_formula_syntax`.

## Analysis Protocol

### Step 1: Structural Audit
Execute `excel_get_workbook_outline(filename)`.
Inspect:
- Total sheet count and sheet names
- Dimensions per sheet (e.g., `A1:F120`)
- Detected headers
- Whether formulas exist and how many formula cells are present

### Step 2: Targeted Range Extraction
Extract only relevant sample rows or summary ranges:

```json
{
  "filename": "/path/to/large_dataset.xlsx",
  "sheet_name": "Sales",
  "start_cell": "A1",
  "end_cell": "F25"
}
```
Tool: `excel_read_data_from_excel`

### Step 3: Synthesis
Provide the user with:
- Summary of data dimensions and column schemas
- Statistical highlights and formula checks
- Recommendations for charts, cleanup, or pivot tables
