# Excel MCP Server Tools

This document provides detailed information about all available tools in the Excel MCP server, organized by functional categories.

## Table of Contents

- [Content Tools](#content-tools)
- [Excel Tools](#excel-tools)
- [Format Tools](#format-tools)
- [Formula Tools](#formula-tools)
- [Graphics Tools](#graphics-tools)

---

## Content Tools

Tools for reading and writing data to Excel worksheets.

### read_data_from_excel

Read tabular data from a specified range in an Excel worksheet.

```python
async def read_data_from_excel(
    filename: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str | None = None,
    preview_only: bool = False,
) -> dict[str, Any]
```

**Context for AI/LLM:**
Use this tool to retrieve data from Excel files when you need to analyze, process, or present information to users. This is ideal for data extraction, reporting, or preparing information for further processing. The tool supports reading specific ranges and can return a preview for large datasets.

**When to use:**

- When you need to extract data from Excel for analysis or display
- When building reports that require Excel data
- When preparing data for further processing by other tools
- When users need to see the contents of specific cells or ranges

**Args:**

- `filename` (str): Path to the Excel file. Example: "reports/monthly_data.xlsx"
- `sheet_name` (str): Name of the worksheet to read from. Example: "SalesData"
- `start_cell` (str): Starting cell reference (default: "A1"). Example: "B2"
- `end_cell` (str | None): Ending cell reference (optional). Example: "D50"
- `preview_only` (bool): Return only a preview if True (default: False)

**Returns:**
`dict[str, Any]` containing:

- `status` (str): "success" or "error"
- `data` (list[list[Any]]): 2D array of cell values
- `message` (str): Operation result message

**Example:**

```json
{
  "status": "success",
  "data": [["Name", "Age", "City"], ["Alice", 30, "NYC"], ["Bob", 25, "LA"]],
  "message": "Data read successfully from A1:C3"
}
```

---

### write_data_to_excel

Write a 2D array of data to an Excel worksheet starting from a specified cell.

```python
async def write_data_to_excel(
    filename: str,
    sheet_name: str,
    data: list[list[Any]],
    start_cell: str = "A1",
) -> dict[str, Any]
```

**Context for AI/LLM:**
Use this tool to write or update data in Excel files. This is essential for creating reports, populating templates, or saving processed data. The tool handles data type preservation and automatically expands the worksheet as needed.

**When to use:**

- When creating new Excel files with data
- When updating existing data in workbooks
- When populating Excel templates programmatically
- When saving processed or generated data to Excel format

**Args:**

- `filename` (str): Path to the Excel file. Example: "output/report.xlsx"
- `sheet_name` (str): Target worksheet name. Example: "Results"
- `data` (list[list[Any]]): 2D array of data to write. Example: [["Name", "Score"], ["Alice", 95]]
- `start_cell` (str): Starting cell reference (default: "A1"). Example: "A1"

**Returns:**
`dict[str, Any]` containing:

- `status` (str): "success" or "error"
- `cells_written` (int): Number of cells written
- `message` (str): Operation result message

---

## Excel Tools

Basic workbook and worksheet operations.

### create_excel_workbook

Create a new Excel workbook (.xlsx) in a secure, validated path.

```python
async def create_excel_workbook(filename: str) -> dict[str, Any]
```

**Context for AI/LLM:**
Use this tool when you need to programmatically generate a new Excel workbook file as the starting point for a reporting, data collection, or automation workflow. This is typically used to initialize new datasets or reports, ensuring that the file is created only within authorized directories for security and compliance.

**When to use:**

- When starting a new workflow that requires a fresh Excel file
- When an automated agent needs to ensure the file is created securely
- When the filename or path may be user-supplied and must be validated for safety

**Args:**

- `filename` (str): Name or path of the workbook to create. Example: "data/reporte_diario"

**Returns:**
`dict[str, Any]` containing:

- `status` (str): "success" or "error"
- `filename` (str): Full path to the created file
- `message` (str): Operation result message

---

### create_excel_worksheet

Add a new worksheet to an existing Excel workbook.

```python
async def create_excel_worksheet(filename: str, sheet_name: str) -> dict[str, Any]
```

**Context for AI/LLM:**
Use this tool to organize or extend data in an existing Excel file by adding a new worksheet. This is ideal when you want to segment data by category, time period, or any logical grouping (e.g., adding a new month to a financial report or a new department to a tracking file).

**When to use:**

- When augmenting an existing Excel file with additional data sections
- When automating workflows that require dynamic worksheet creation
- When you need to ensure that sheet names remain unique and valid

**Args:**

- `filename` (str): Path to the existing Excel workbook. Example: "reports/monthly_report.xlsx"
- `sheet_name` (str): Name of the new worksheet to add. Must be unique and valid per Excel rules

**Returns:**
`dict[str, Any]` containing:

- `status` (str): "success" or "error"
- `sheet_name` (str): Name of the created sheet
- `message` (str): Operation result message

---

### list_excel_documents

List all .xlsx files in the specified directory.

```python
async def list_excel_documents() -> dict[str, Any]
```

**Context for AI/LLM:**
Use this tool to discover and enumerate all Excel files in a specific directory, which is helpful for inventory, audit, batch processing, or automated data discovery scenarios. The tool validates the directory for security, ensuring only authorized locations are scanned.

**When to use:**

- When an AI needs to present the user with available Excel files for further action
- When preparing to process, analyze, or summarize multiple Excel documents in a folder
- When auditing or verifying the presence and properties of .xlsx files in a given path

**Args:**

- None (uses configured directory)

**Returns:**
`dict[str, Any]` containing:

- `status` (str): "success" or "error"
- `count` (int): Number of Excel files found
- `files` (list[dict]): List of file metadata with name, size, modified date
- `message` (str): Operation result message

---

### copy_worksheet

Create a copy of an existing worksheet within the same workbook.

```python
async def copy_worksheet(
    filename: str,
    source_sheet: str,
    target_sheet: str,
) -> dict[str, Any]
```

**Args:**

- `filename` (str): Path to the Excel workbook
- `source_sheet` (str): Name of the sheet to copy
- `target_sheet` (str): Name for the new copy

**Returns:**
`dict[str, Any]` containing operation status and details

---

### delete_worksheet

Delete a worksheet from a workbook.

```python
async def delete_worksheet(filename: str, sheet_name: str) -> dict[str, Any]
```

**Args:**

- `filename` (str): Path to the Excel workbook
- `sheet_name` (str): Name of the worksheet to delete

**Returns:**
`dict[str, Any]` containing operation status and details

---

### rename_worksheet

Rename an existing worksheet.

```python
async def rename_worksheet(
    filename: str,
    old_name: str,
    new_name: str,
) -> dict[str, Any]
```

**Args:**

- `filename` (str): Path to the Excel workbook
- `old_name` (str): Current name of the sheet
- `new_name` (str): New name for the sheet

**Returns:**
`dict[str, Any]` containing operation status and details

---

### get_workbook_metadata

Get metadata about a workbook including sheets and their ranges.

```python
async def get_workbook_metadata(
    filename: str,
    include_ranges: bool = False,
) -> dict[str, Any]
```

**Args:**

- `filename` (str): Path to the Excel file
- `include_ranges` (bool): Include used ranges for each sheet (default: False)

**Returns:**
`dict[str, Any]` containing:

- `status` (str): "success" or "error"
- `filename` (str): Name of the file
- `sheets` (list[str]): List of sheet names
- `used_ranges` (dict[str, str] | None): Range info if requested
- `size` (int): File size in bytes
- `modified` (float): Last modification timestamp

---

## Format Tools

Cell formatting, merging, and range operations.

### format_range_excel

Apply formatting to a range of cells with comprehensive style options.

```python
async def format_range_excel(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str | None = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: int | None = None,
    font_color: str | None = None,
    bg_color: str | None = None,
    border_style: str | None = None,
    border_color: str | None = None,
    number_format: str | None = None,
    alignment: str | None = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: dict[str, Any] | None = None,
    conditional_format: dict[str, Any] | None = None,
) -> dict[str, Any]
```

**Args:**

- `filename` (str): Path to the Excel file
- `sheet_name` (str): Target worksheet name
- `start_cell` (str): Starting cell of range
- `end_cell` (str | None): Ending cell of range
- `bold` (bool): Apply bold formatting
- `italic` (bool): Apply italic formatting
- `underline` (bool): Apply underline
- `font_size` (int | None): Font size in points
- `font_color` (str | None): Font color (hex or color name)
- `bg_color` (str | None): Background color
- `border_style` (str | None): Border style (thin, medium, thick, etc.)
- `border_color` (str | None): Border color
- `number_format` (str | None): Number format (e.g., "0.00", "$#,##0.00")
- `alignment` (str | None): Text alignment (left, center, right)
- `wrap_text` (bool): Wrap text in cells
- `merge_cells` (bool): Merge the range after formatting
- `protection` (dict | None): Cell protection settings
- `conditional_format` (dict | None): Conditional formatting rules

**Returns:**
`dict[str, Any]` containing operation status and details

---

### merge_cells

Merge a range of cells into a single cell.

```python
async def merge_cells(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
) -> dict[str, Any]
```

**Args:**

- `filename` (str): Path to the Excel file
- `sheet_name` (str): Target worksheet name
- `start_cell` (str): Starting cell of range
- `end_cell` (str): Ending cell of range

**Returns:**
`dict[str, Any]` containing operation status and details

---

### unmerge_cells

Unmerge a previously merged range of cells.

```python
async def unmerge_cells(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
) -> dict[str, Any]
```

**Args:**

- `filename` (str): Path to the Excel file
- `sheet_name` (str): Target worksheet name
- `start_cell` (str): Starting cell of merged range
- `end_cell` (str): Ending cell of merged range

**Returns:**
`dict[str, Any]` containing operation status and details

---

### copy_range

Copy a range of cells to another location.

```python
async def copy_range(
    filename: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str | None = None,
    include_formatting: bool = True,
) -> dict[str, Any]
```

**Args:**

- `filename` (str): Path to the Excel file
- `sheet_name` (str): Source worksheet name
- `source_start` (str): Starting cell of source range
- `source_end` (str): Ending cell of source range
- `target_start` (str): Top-left cell of target range
- `target_sheet` (str | None): Target worksheet (defaults to source)
- `include_formatting` (bool): Copy cell formatting (default: True)

**Returns:**
`dict[str, Any]` containing operation status and details

---

### delete_range

Clear contents and optionally formatting from a range of cells.

```python
async def delete_range(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up",
) -> dict[str, Any]
```

**Args:**

- `filename` (str): Path to the Excel file
- `sheet_name` (str): Worksheet name
- `start_cell` (str): Starting cell of range
- `end_cell` (str): Ending cell of range
- `shift_direction` (str): "up" or "left" for deletion

**Returns:**
`dict[str, Any]` containing operation status and details

---

### validate_excel_range

Validate if a range reference is valid for the specified worksheet.

```python
async def validate_excel_range(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str | None = None,
) -> dict[str, Any]
```

**Args:**

- `filename` (str): Path to the Excel file
- `sheet_name` (str): Worksheet name
- `start_cell` (str): Starting cell to validate
- `end_cell` (str | None): Ending cell to validate

**Returns:**
`dict[str, Any]` containing:

- `status` (str): "valid" or "error"
- `is_valid` (bool): Whether the range is valid
- `message` (str): Validation result message

---

## Formula Tools

Excel formula operations.

### apply_formula_excel

Apply an Excel formula to a cell.

```python
async def apply_formula_excel(
    filename: str,
    sheet_name: str,
    cell: str,
    formula: str,
) -> dict[str, Any]
```

**Context for AI/LLM:**
Use this tool when you need to add calculations to Excel files. This is essential for financial modeling, data analysis, or creating calculated fields. The tool supports standard Excel formulas and handles formula syntax validation.

**When to use:**

- When adding calculations to Excel reports
- When creating computed columns or fields
- When building financial or scientific models
- When automating spreadsheet calculations

**Args:**

- `filename` (str): Path to the Excel file. Example: "financials/budget.xlsx"
- `sheet_name` (str): Target worksheet name. Example: "Summary"
- `cell` (str): Target cell reference. Example: "E2"
- `formula` (str): Excel formula. Example: "=SUM(A1:D1)"

**Returns:**
`dict[str, Any]` containing:

- `status` (str): "success" or "error"
- `cell` (str): Cell where formula was applied
- `formula` (str): The formula that was applied
- `message` (str): Operation result message

---

### validate_formula_syntax

Validate an Excel formula syntax without applying it.

```python
async def validate_formula_syntax(
    filename: str,
    sheet_name: str,
    cell: str,
    formula: str,
) -> dict[str, Any]
```

**Context for AI/LLM:**
Use this tool to verify that a formula is valid before applying it. This helps prevent errors in automated workflows and ensures formula correctness. The tool checks syntax and references without modifying the spreadsheet.

**When to use:**

- Before applying formulas in automated workflows
- When users want to verify formula correctness
- When debugging formula issues
- When validating user-provided formulas

**Args:**

- `filename` (str): Path to the Excel file
- `sheet_name` (str): Target worksheet name
- `cell` (str): Cell where formula would be placed
- `formula` (str): Formula to validate

**Returns:**
`dict[str, Any]` containing:

- `status` (str): "valid", "invalid", or "error"
- `is_valid` (bool): Whether the formula is valid
- `message` (str): Validation result and any error details

---

## Graphics Tools

Charts and pivot tables.

### create_chart

Create a chart in the specified worksheet.

```python
async def create_chart(
    filename: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = "",
) -> dict[str, Any]
```

**Context for AI/LLM:**
Use this tool to generate visualizations from Excel data. Charts help users understand data trends, comparisons, and distributions at a glance. The tool supports various chart types for different data visualization needs.

**When to use:**

- When creating data visualizations from spreadsheet data
- When generating reports with charts
- When analyzing trends or patterns in data
- When presenting data in a more digestible visual format

**Args:**

- `filename` (str): Path to the Excel file. Example: "sales/report.xlsx"
- `sheet_name` (str): Target worksheet name. Example: "Data"
- `data_range` (str): Range containing chart data. Example: "A1:B10"
- `chart_type` (str): Type of chart (line, bar, pie, scatter, area, column)
- `target_cell` (str): Cell for chart top-left corner. Example: "D2"
- `title` (str): Chart title (optional)
- `x_axis` (str): X-axis title (optional)
- `y_axis` (str): Y-axis title (optional)

**Returns:**
`dict[str, Any]` containing:

- `status` (str): "success" or "error"
- `chart_type` (str): Type of chart created
- `message` (str): Operation result message

---

### create_pivot_table

Create a pivot table for data analysis and summarization.

```python
async def create_pivot_table(
    filename: str,
    sheet_name: str,
    data_range: str,
    rows: list[str],
    values: list[str],
    columns: list[str] | None = None,
    agg_func: str = "mean",
) -> dict[str, Any]
```

**Context for AI/LLM:**
Use this tool to create pivot tables for data analysis and summarization. Pivot tables are powerful for summarizing large datasets, performing aggregations, and gaining insights from structured data.

**When to use:**

- When summarizing large datasets
- When performing group-by aggregations
- When creating cross-tabulations
- When analyzing data from multiple perspectives

**Args:**

- `filename` (str): Path to the Excel file
- `sheet_name` (str): Name of the sheet for pivot table
- `data_range` (str): Source data range. Example: "A1:D100"
- `rows` (list[str]): Field names for row labels
- `values` (list[str]): Field names for values to aggregate
- `columns` (list[str] | None): Field names for column labels
- `agg_func` (str): Aggregation function (mean, sum, count, min, max)

**Returns:**
`dict[str, Any]` containing:

- `status` (str): "success" or "error"
- `pivot_table_name` (str): Name of created pivot table
- `message` (str): Operation result message

---

## Return Value Conventions

All MCP tool functions return `dict[str, Any]` with consistent structure:

**Success Response:**

```json
{
  "status": "success",
  "message": "Operation completed successfully",
  ...additional_data
}
```

**Error Response:**

```json
{
  "status": "error",
  "message": "Description of the error"
}
```

---

## Error Handling

The MCP Excel server uses a hierarchical exception system:

- `ExcelMCPError` - Base exception
- `WorkbookError` - Workbook operations
- `SheetError` - Worksheet operations
- `DataError` - Data read/write operations
- `ValidationError` - Input validation
- `FormattingError` - Cell formatting
- `CalculationError` - Formula operations
- `PivotError` - Pivot table operations
- `ChartError` - Chart operations

All tool functions catch exceptions and return error dictionaries instead of raising exceptions.
