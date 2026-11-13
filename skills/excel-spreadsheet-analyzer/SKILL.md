---
name: excel-spreadsheet-analyzer
description: "Inspect, audit, and extract insights from Excel workbooks (.xlsx) without exhausting model context. Uses lightweight structural outlines, column statistical profiling, targeted cell search, and metadata extraction."
argument-hint: "Target workbook path, worksheet to analyze, and audit objective"
user-invocable: true
disable-model-invocation: false
---

# Excel Spreadsheet Analyzer

Use this skill to extract structural insights, audit formulas, profile data columns, and search cells across Microsoft Excel workbooks without consuming excessive tokens.

---

## Core Rules

1. **Strict Context Preservation (Never Read Unbounded Sheets):**
   - Do NOT issue unbounded queries like `excel_read_data(start_cell="A1")` on unknown workbooks.
   - Always call `excel_get_workbook_outline(filename)` first to determine dimensions (e.g. `A1:M1500`), sheet names, and detected column headers.
2. **Profile Columns Statistically:**
   - Use `excel_profile_sheet` to inspect data types (`integer`, `float`, `string`, `datetime`, `boolean`), null counts, min/max for numerics, and distinct sample values without dumping row matrices into context.
3. **Targeted Cell Discovery:**
   - Use `excel_find_cells` to locate specific text, numbers, formulas, or regular expression patterns across sheets.

---

## 3-Phase Audit Protocol

### Phase 1: Structural Audit & Outline Inspection

```json
{
  "filename": "/path/to/enterprise_financials.xlsx"
}
```
**Tool:** `excel_get_workbook_outline`

**Evaluate:**
- Total sheet count and sheet names.
- Populated bounds per sheet (e.g., `Sheet1: A1:H250`, `Summary: A1:D15`).
- Column header integrity (check for blank or missing headers).
- Formula presence: Outline summary detects whether formula cells are present.

---

### Phase 2: Column Profiling & Statistical Health Check

Extract column-level data types, min/max values, null counts, and distinct samples in one lightweight pass:

```json
{
  "filename": "/path/to/enterprise_financials.xlsx",
  "sheet_name": "Revenue",
  "sample_rows": 100
}
```
**Tool:** `excel_profile_sheet`

**Data Health Checklist:**
- Are numbers stored as numeric types rather than strings?
- Are there unexpected null values in required columns?
- Do min and max bounds fall within plausible financial or operational thresholds?
- Are derived columns and summary rows using live Excel formulas (`=`) rather than hardcoded static numbers?
- Is the worksheet free of cluttered/overlapping tables and inconsistent or oversized (>12pt) fonts?

---

### Phase 3: Targeted Search & Formula Verification

1. **Locate Specific Values or Anomalies:**
   Find cells containing specific text or formula keywords:
   ```json
   {
     "filename": "/path/to/enterprise_financials.xlsx",
     "query": "Error",
     "match_case": false
   }
   ```
   **Tool:** `excel_find_cells`

2. **Audit Summary & Total Rows for Hardcoded Calculations:**
   Inspect summary/total ranges (using `excel_read_data` on targeted total rows, e.g., `A25:F25`) to verify that totals, variances, and KPIs contain `=` formulas rather than static values.

3. **Validate Formula Syntax:**
   Verify formula syntax before applying modifications:
   ```json
   {
     "formula": "=SUM(C4:C25)"
   }
   ```
   **Tool:** `excel_validate_formula_syntax`

---

## Summary Deliverable Format

When reporting analysis findings to the user:
- **Structural Bounds:** Present sheet names, row/column dimensions, and populated ranges in a table.
- **Column Health:** Report column data types, completeness (non-null counts), and min/max ranges.
- **Data Quality:** Highlight any anomalies (formula errors like `#REF!`, missing values).
- **Recommendations:** Suggest formatting improvements, native tables, or chart visual additions.
