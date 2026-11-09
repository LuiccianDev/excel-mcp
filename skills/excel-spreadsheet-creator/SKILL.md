---
name: excel-spreadsheet-creator
description: "Create high-quality, professional Excel workbooks (.xlsx) in a single pass using Excel MCP. Generates complete workbooks with multiple sheets, typography, zebra-striping, currency formatting, auto-totals, formulas, and charts in 1 tool call, saving 90% tokens."
argument-hint: "Workbook purpose, sheet structures, design preset ('corporate', 'financial', 'modern', 'compact'), headers, data, and target filename"
user-invocable: true
disable-model-invocation: false
---

# Excel Spreadsheet Creator

Use this skill to create complete, visually styled Microsoft Excel workbooks in **1 single tool call**, avoiding chatty micro-iterations and preserving model context.

## Core Rules

1. **Always use Single-Pass Construction for new workbooks:**
   - **`excel_build_workbook`**: Primary tool. Generates workbooks with multiple sheets, banners, table headers, alternating row colors (zebra striping), accounting borders, auto-calculated totals, and native charts.
2. **Never make micro-calls for new workbooks:**
   - Do NOT chain `create_workbook` -> `set_cell_value` 50 times -> `set_cell_format` 20 times.
   - Send the entire workbook structure inside the `sheets` array in **one call**.
3. **Always use absolute allowed paths** within `MCP_ALLOWED_DIRECTORIES`.

## Design Presets

| Preset | Header Fill | Header Text | Zebra Row Fill | Typography | Borders / Totals | Best For |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `corporate` (Default) | Navy Blue (`#1F4E79`) | White | Ice Blue (`#EDF2F8`) | Calibri 11pt | Thin gray, soft blue total | Business proposals, management reports, executive summaries |
| `financial` | Forest Green (`#1E4620`) | White | Pale Mint (`#F0F6F1`) | Calibri 11pt, currency format | Double bottom line on totals | Balance sheets, P&L statements, budgeting, cash flow |
| `modern` | Slate / Charcoal (`#2F3542`) | White | Cool Gray (`#F8F9FA`) | Segoe UI 10.5pt | Minimalist soft borders | SaaS metrics, tech product analytics, modern dashboards |
| `compact` | Steel Gray (`#4A6572`) | White | Off-White (`#F5F5F5`) | Arial 9.5pt | Dense grid, high data density | Inventory lists, master catalogs, high-volume exports |

## Canonical Usage Pattern: `excel_build_workbook`

```json
{
  "filename": "/path/to/annual_financial_report_2026.xlsx",
  "preset": "financial",
  "auto_fit": true,
  "sheets": [
    {
      "name": "Revenue Summary",
      "title": "ANNUAL REVENUE & PERFORMANCE REPORT - FY2026",
      "headers": ["Quarter", "Target Revenue", "Actual Revenue", "Variance", "Growth Rate"],
      "rows": [
        ["Q1 2026", 150000, 168000, "=C4-B4", "=D4/B4"],
        ["Q2 2026", 160000, 172000, "=C5-B5", "=D5/B5"],
        ["Q3 2026", 175000, 189000, "=C6-B6", "=D6/B6"],
        ["Q4 2026", 200000, 224000, "=C7-B7", "=D7/B7"]
      ],
      "currency_columns": ["Target Revenue", "Actual Revenue", "Variance"],
      "auto_totals": [1, 2, 3],
      "chart": {
        "type": "bar",
        "title": "Target vs Actual Revenue by Quarter",
        "data_column": 3,
        "label_column": 1
      }
    },
    {
      "name": "Operating Expenses",
      "headers": ["Expense Category", "Budget", "Actual Spend", "Remaining"],
      "rows": [
        ["Payroll & Salaries", 420000, 415000, "=B2-C2"],
        ["Cloud & Infrastructure", 65000, 68500, "=B3-C3"],
        ["Sales & Marketing", 110000, 102000, "=B4-C4"],
        ["General & Administrative", 45000, 43200, "=B5-C5"]
      ],
      "currency_columns": ["Budget", "Actual Spend", "Remaining"],
      "auto_totals": [1, 2, 3]
    }
  ]
}
```

## Quick Verification (QA)

After generating the workbook with `excel_build_workbook`:
- Run `excel_get_workbook_outline(filename)` to confirm sheets, dimensions, and headers without dumping the entire grid.
- Present a concise, helpful summary to the user.
