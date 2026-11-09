---
name: excel-spreadsheet-analyzer
description: "Inspect, audit, and extract insights from Excel workbooks (.xlsx) without exhausting model context. Uses lightweight structural outlines, targeted range sampling, formula verification, and metadata extraction."
argument-hint: "Target workbook path, worksheet to analyze, and audit objective"
user-invocable: true
disable-model-invocation: false
---

# Excel Spreadsheet Analyzer

Use this skill to extract structural insights, audit formulas, and sample data from Microsoft Excel workbooks without consuming excessive tokens.

---

## Core Rules

1. **Strict Context Preservation (Never Read Unbounded Sheets):**
   - Do NOT issue unbounded queries like `read_data_from_excel(start_cell="A1")` on unknown workbooks.
   - Always call `excel_get_workbook_outline(filename)` first to determine dimensions (e.g. `A1:M1500`), sheet names, and detected column headers.
2. **Sample Slices Rather Than Full Matrices:**
   - Once dimensions are known, read a representative sample (e.g., first 10 rows: `A1:M10`) to understand column data types and patterns.
3. **Use Native Calculators for Summary Metrics:**
   - Instead of extracting 1,000 numbers into context to compute sums or averages, use `excel_calculate_range_sum` or `excel_calculate_range_average`.

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
- Formula density: Count of formula cells vs raw value cells.

---

### Phase 2: Targeted Range Sampling

Extract only the schema and sample rows:

```json
{
  "filename": "/path/to/enterprise_financials.xlsx",
  "sheet_name": "Revenue",
  "start_cell": "A1",
  "end_cell": "H10"
}
```
**Tool:** `excel_read_data_from_excel`

**Data Health Checklist:**
- Are numbers stored as numeric types rather than strings?
- Are currency columns formatted consistently?
- Are date columns standardized to ISO (`yyyy-mm-dd`)?

---

### Phase 3: Formula & Calculation Integrity Verification

1. **Syntax Validation:**
   ```json
   {
     "formula": "=SUM(C4:C25)"
   }
   ```
   **Tool:** `excel_validate_formula_syntax`

2. **Metric Verification:**
   Verify whether reported totals match the underlying cells:
   ```json
   {
     "filename": "/path/to/enterprise_financials.xlsx",
     "sheet_name": "Revenue",
     "start_cell": "C4",
     "end_cell": "C25"
   }
   ```
   **Tool:** `excel_calculate_range_sum`

---

## Summary Deliverable Format

When reporting analysis findings to the user:
- **Structural Bounds:** Present sheet names, row/column dimensions, and populated ranges in a table.
- **Data Quality:** Highlight any anomalies (missing headers, formula errors like `#REF!` or `#VALUE!`).
- **Recommendations:** Suggest formatting improvements, pivot summaries, or chart visual additions.
