---
name: excel-spreadsheet-creator
description: "Create high-quality, professional Excel workbooks (.xlsx) in a single pass using Excel MCP. Generates complete workbooks with uniform typography, uncluttered multi-sheet layouts, dynamic formulas, zebra-striping, currency formatting, auto-totals, and native charts in 1 tool call."
argument-hint: "Workbook purpose, layout pattern, design preset ('corporate', 'financial', 'modern', 'compact'), optional font_name/font_size, headers, data, and target filename"
user-invocable: true
disable-model-invocation: false
---

# Excel Spreadsheet Creator

Use this skill to create complete, cleanly structured, and visually uniform Microsoft Excel workbooks in **1 single tool call**, avoiding chatty micro-iterations, rigid repetitive layouts, and hardcoded calculations.

---

## Core Rules (Non-Negotiable)

1. **Mandatory Single-Pass Construction:**
   - **`excel_build_workbook`**: Always use this tool for creating new workbooks. Pass the complete structure inside `sheets`. Do NOT chain `create_workbook` &rarr; `write_data` &rarr; `format_range`.
2. **Uniform, Readable Typography:**
   - **Recommended Default:** Unless the user specifies otherwise, standardize on a single clean font across the entire workbook: **`Calibri 11pt`** or **`Arial 11pt`** (e.g., `preset: "corporate"` or `font_name: "Arial", font_size: 11`).
   - **User Font Priority:** If the user explicitly requests a custom font family or size, honor their preference across the entire workbook using `font_name` and `font_size` in `excel_build_workbook`.
   - **No Oversized or Mixed Fonts:** Never mix incompatible font families within a workbook, and keep titles/banners compact (`<= 12pt`) rather than oversized.
3. **Uncluttered Architecture (No Cramming):**
   - **Separation of Concerns:** Never cram configuration parameters, assumptions, raw transaction logs, and summary KPIs onto a single crowded sheet.
   - Split distinct concerns into dedicated logical worksheets (e.g., `Assumptions`, `Operations`, `Summary`) or use clean spacing, `freeze_panes`, and `auto_filter` so every tab is immediately readable.
4. **100% Dynamic Formulas (Zero Hardcoded Calculations):**
   - **Strict Rule:** NEVER write pre-calculated static numbers into totals, subtotals, variances, margins, percentages, averages, counts, or summary cells.
   - Every derived value MUST be a live Excel formula (`=SUM(...)`, `=AVERAGE(...)`, `=IFERROR(...)`, `=SUMIF(...)`, `='Operations'!C8`, etc.).
5. **Adapt Layout to Purpose:**
   - Choose the structural pattern below that best fits the user's domain instead of repeating a single merged-banner template for every workbook.
6. **Paths:** Always use absolute paths located within `MCP_ALLOWED_DIRECTORIES`.

---

## Design Presets & Font Overrides

| Preset | Header Fill | Header Text | Zebra Row Fill | Default Font | Borders / Totals | Recommended For |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`corporate`** *(Default)* | Navy Blue (`#1F4E79`) | White, Bold | Ice Blue (`#EDF2F8`) | Calibri 11pt | Thin gray, soft blue total banner | Executive reports, proposals, multi-tab summaries |
| **`financial`** | Forest Green (`#1E4620`) | White, Bold | Pale Mint (`#F0F6F1`) | Calibri 11pt | Double-line accounting bottom border | Financial statements, budgets, cash flow models |
| **`modern`** | Slate Charcoal (`#2F3542`) | White, Bold | Cool Gray (`#F8F9FA`) | Segoe UI 10.5pt | Minimalist hairline borders | SaaS metrics, product analytics, KPI dashboards |
| **`compact`** | Steel Gray (`#4A6572`) | White, Bold | Off-White (`#F5F5F5`) | Arial 9.5pt | High-density grid lines | Inventory ledgers, transaction logs, data exports |

> [!TIP]
> You can combine any color preset with uniform typography by passing `font_name` and `font_size` (for example, `preset: "modern", font_name: "Arial", font_size: 11`), or honor any specific font requested by the user.

---

## Structural Patterns (Choose by Use Case)

### Pattern A: Clean Tabular Operations Ledger (No Title Banner)

Best for operational logs, inventory lists, expenses, and datasets where users immediately filter, sort, and freeze the top header row (`row 1`):

```json
{
  "filename": "/path/to/operations_ledger.xlsx",
  "preset": "corporate",
  "font_name": "Arial",
  "font_size": 11,
  "auto_fit": true,
  "sheets": [
    {
      "name": "Transactions",
      "headers": ["Date", "Category", "Units", "Unit Price", "Subtotal", "Tax (18%)", "Total"],
      "rows": [
        ["2026-11-01", "Cloud Hosting", 4, 250.00, "=C2*D2", "=E2*0.18", "=E2+F2"],
        ["2026-11-03", "API Licenses", 10, 45.00, "=C3*D3", "=E3*0.18", "=E3+F3"],
        ["2026-11-05", "Storage Tier", 2, 120.00, "=C4*D4", "=E4*0.18", "=E4+F4"]
      ],
      "currency_columns": ["Unit Price", "Subtotal", "Tax (18%)", "Total"],
      "auto_totals": [4, 5, 6],
      "freeze_panes": "A2",
      "auto_filter": true
    }
  ]
}
```

---

### Pattern B: Structured Multi-Tab Model (Separated Assumptions + Calculations + Summary)

Best for budgets, financial models, pricing calculators, and projections. Keeps configuration parameters isolated on an `Assumptions` tab and links all calculations dynamically across sheets:

```json
{
  "filename": "/path/to/fy2026_budget_model.xlsx",
  "preset": "financial",
  "font_name": "Calibri",
  "font_size": 11,
  "auto_fit": true,
  "sheets": [
    {
      "name": "Assumptions",
      "title": "MODEL PARAMETERS & ASSUMPTIONS",
      "headers": ["Parameter", "Value", "Notes"],
      "rows": [
        ["Corporate Tax Rate", 0.21, "Applied to positive operating income"],
        ["Target Growth Rate", 0.12, "Annual revenue expansion target"]
      ],
      "freeze_panes": "A4"
    },
    {
      "name": "Operations",
      "title": "QUARTERLY REVENUE & COST MODEL",
      "headers": ["Quarter", "Revenue", "Direct Costs", "Gross Profit", "Estimated Tax", "Net Profit"],
      "rows": [
        ["Q1", 150000, 62000, "=B4-C4", "=MAX(0,D4)*Assumptions!$B$4", "=D4-E4"],
        ["Q2", 168000, 65000, "=B5-C5", "=MAX(0,D5)*Assumptions!$B$4", "=D5-E5"],
        ["Q3", 182000, 70000, "=B6-C6", "=MAX(0,D6)*Assumptions!$B$4", "=D6-E6"],
        ["Q4", 210000, 78000, "=B7-C7", "=MAX(0,D7)*Assumptions!$B$4", "=D7-E7"]
      ],
      "currency_columns": ["Revenue", "Direct Costs", "Gross Profit", "Estimated Tax", "Net Profit"],
      "auto_totals": [1, 2, 3, 4, 5],
      "freeze_panes": "A4",
      "auto_filter": true
    },
    {
      "name": "Executive Summary",
      "title": "ANNUAL PERFORMANCE SUMMARY",
      "headers": ["Metric", "FY2026 Total", "Margin / Share"],
      "rows": [
        ["Total Revenue", "=Operations!B8", "=IFERROR(B4/B4,0)"],
        ["Total Direct Costs", "=Operations!C8", "=IFERROR(B5/$B$4,0)"],
        ["Gross Profit", "=Operations!D8", "=IFERROR(B6/$B$4,0)"],
        ["Net Profit", "=Operations!F8", "=IFERROR(B7/$B$4,0)"]
      ],
      "currency_columns": ["FY2026 Total"]
    }
  ]
}
```

---

### Pattern C: KPI Summary & Chart Report

Best for management dashboards where a clean summary table is paired with a native visual chart:

```json
{
  "filename": "/path/to/regional_kpi_report.xlsx",
  "preset": "modern",
  "font_name": "Calibri",
  "font_size": 11,
  "auto_fit": true,
  "sheets": [
    {
      "name": "Regional KPIs",
      "title": "REGIONAL SALES PERFORMANCE",
      "headers": ["Region", "Quota", "Actual Sales", "Attainment %", "Variance"],
      "rows": [
        ["North", 450000, 485000, "=IFERROR(C4/B4,0)", "=C4-B4"],
        ["South", 380000, 362000, "=IFERROR(C5/B5,0)", "=C5-B5"],
        ["East", 520000, 560000, "=IFERROR(C6/B6,0)", "=C6-B6"],
        ["West", 410000, 435000, "=IFERROR(C7/B7,0)", "=C7-B7"]
      ],
      "currency_columns": ["Quota", "Actual Sales", "Variance"],
      "auto_totals": [1, 2, 4],
      "chart": {
        "type": "bar",
        "title": "Actual Sales by Region",
        "data_column": 3,
        "label_column": 1
      }
    }
  ]
}
```

### Formula Coordinate Mapping Note
- When a sheet omits `title`, headers sit on **row 1** and data starts on **row 2** (`=C2*D2`).
- When a sheet includes a `title` banner, the banner sits on **row 1**, row 2 is a blank separator, headers sit on **row 3**, and data starts on **row 4** (`=B4-C4`).

---

## Post-Creation Verification (QA)

Immediately following workbook construction:
1. Call `excel_get_workbook_outline(filename)` to confirm sheets, dimensions, headers, and `has_formulas` without loading the full cell matrix.
2. Verify that every derived calculation cell uses a formula and report the verified outline to the user alongside the file path.
