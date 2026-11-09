---
name: excel-spreadsheet-creator
description: "Create high-quality, professional Excel workbooks (.xlsx) in a single pass using Excel MCP. Generates complete workbooks with multiple sheets, typography, zebra-striping, currency formatting, auto-totals, formulas, and charts in 1 tool call, saving 90% tokens."
argument-hint: "Workbook purpose, sheet structures, design preset ('corporate', 'financial', 'modern', 'compact'), headers, data, and target filename"
user-invocable: true
disable-model-invocation: false
---

# Excel Spreadsheet Creator

Use this skill to create complete, visually styled Microsoft Excel workbooks in **1 single tool call**, avoiding chatty micro-iterations and preserving model context.

---

## Core Rules

1. **Mandatory Single-Pass Construction for New Workbooks:**
   - **`excel_build_workbook`**: Always use this tool for creating workbooks from scratch. It builds sheets, banners, table headers, alternating row colors (zebra striping), accounting borders, auto-calculated totals (`SUM`), number formatting, and native charts in **one single call**.
2. **Prohibition of Micro-Calls:**
   - Do NOT chain `create_workbook` &rarr; 50 calls to `write_data_to_excel` &rarr; 20 calls to `format_range_excel`.
   - Pass the entire structure inside the `sheets` parameter.
3. **Paths:** Always use absolute paths located within `MCP_ALLOWED_DIRECTORIES`.

---

## Design Presets

| Preset | Header Fill | Header Text | Zebra Row Fill | Typography | Borders / Totals | Recommended For |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`corporate`** *(Default)* | Navy Blue (`#1F4E79`) | White, Bold | Ice Blue (`#EDF2F8`) | Calibri 11pt | Thin gray, soft blue total banner | Executive reports, board presentations, commercial proposals |
| **`financial`** | Forest Green (`#1E4620`) | White, Bold | Pale Mint (`#F0F6F1`) | Calibri 11pt | Double-line accounting bottom border | Income statements, balance sheets, cash flow models, budgets |
| **`modern`** | Slate Charcoal (`#2F3542`) | White, Bold | Cool Gray (`#F8F9FA`) | Segoe UI 10.5pt | Minimalist clean hairline borders | SaaS metrics, tech product analytics, modern KPI dashboards |
| **`compact`** | Steel Gray (`#4A6572`) | White, Bold | Off-White (`#F5F5F5`) | Arial 9.5pt | High-density grid lines | Large inventory logs, transaction journals, data exports |

---

## Multi-Tab Production Template

The following payload constructs an executive financial model with an **Executive Summary** tab and an **Operating Expenses** tab, complete with formulas, auto-totals, and a native bar chart:

```json
{
  "filename": "/path/to/fy2026_financial_model.xlsx",
  "preset": "financial",
  "auto_fit": true,
  "sheets": [
    {
      "name": "Executive Summary",
      "title": "EXECUTIVE SUMMARY - FY2026 PERFORMANCE",
      "headers": ["Quarter", "Revenue Target", "Actual Revenue", "Variance", "Growth %"],
      "rows": [
        ["Q1 2026", 1500000, 1680000, "=C4-B4", "=D4/B4"],
        ["Q2 2026", 1650000, 1720000, "=C5-B5", "=D5/B5"],
        ["Q3 2026", 1800000, 1950000, "=C6-B6", "=D6/B6"],
        ["Q4 2026", 2100000, 2300000, "=C7-B7", "=D7/B7"]
      ],
      "currency_columns": ["Revenue Target", "Actual Revenue", "Variance"],
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
      "title": "OPERATING EXPENSES BREAKDOWN - FY2026",
      "headers": ["Expense Category", "Budget", "Actual Spend", "Variance"],
      "rows": [
        ["R&D & Engineering", 850000, 820000, "=B4-C4"],
        ["Sales & Marketing", 620000, 645000, "=B5-C5"],
        ["Cloud & Infrastructure", 310000, 305000, "=B6-C6"],
        ["General & Administrative", 220000, 215000, "=B7-C7"]
      ],
      "currency_columns": ["Budget", "Actual Spend", "Variance"],
      "auto_totals": [1, 2, 3]
    }
  ]
}
```

### Formula Coordinate Mapping Note
- In `excel_build_workbook`, when a `title` banner is present, the banner occupies row 1, an empty separator occupies row 2, and the table headers occupy row 3.
- Consequently, the first data row starts on row 4. Formulas referencing preceding columns in row 4 should use `=C4-B4`, etc.

---

## Post-Creation Verification (QA)

Immediately following workbook construction:
1. Call `excel_get_workbook_outline(filename)` to confirm sheets, dimensions (`A1:E8`), and headers without loading the full cell matrix.
2. Report the verified outline to the user alongside the file path.
