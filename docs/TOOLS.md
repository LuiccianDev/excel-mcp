# Excel MCP Tools Reference

> [!NOTE]
> This document is automatically generated from the registered FastMCP tool signatures.
> Total Canonical Tools: **32**.

## Tool Inventory

| Canonical Tool Name | Type | Description |
| :--- | :--- | :--- |
| [`excel_add_data_validation`](#excel-add-data-validation) | `Write` | Add a data validation rule to a worksheet range. |
| [`excel_append_rows`](#excel-append-rows) | `Write` | Append batch rows of data below the existing content in a worksheet in an atomic operation. |
| [`excel_apply_formula`](#excel-apply-formula) | `Write` | Apply a validated Excel formula to a specific cell in a worksheet. |
| [`excel_batch_write_data`](#excel-batch-write-data) | `Write` | Write a 2D matrix of values or formulas into a worksheet starting from a specific cell in one atomic operation. |
| [`excel_build_workbook`](#excel-build-workbook) | `Write` | Build a complete Excel workbook with typography, zebra striping, currency/date formats, formulas, and charts in a single pass. |
| [`excel_copy_range`](#excel-copy-range) | `Write` | Copy a range of cells, including their values and formatting, to a new location, potentially in a different worksheet. |
| [`excel_copy_worksheet`](#excel-copy-worksheet) | `Write` | Create a duplicate of an existing worksheet within the same Excel workbook. |
| [`excel_create_chart`](#excel-create-chart) | `Write` | Create a chart in a worksheet based on a specified data range. |
| [`excel_create_pivot_table`](#excel-create-pivot-table) | `Write` | Create a pivot table in a worksheet to summarize and analyze data from a given range. |
| [`excel_create_table`](#excel-create-table) | `Write` | Convert a range of cells into a native, styled Excel Table. |
| [`excel_create_workbook`](#excel-create-workbook) | `Write` | Create a new Excel workbook (.xlsx) in a secure, validated path. |
| [`excel_create_worksheet`](#excel-create-worksheet) | `Write` | Add a new worksheet to an existing Excel workbook. |
| [`excel_delete_range`](#excel-delete-range) | `Destructive` | Delete a range of cells and optionally shift the surrounding cells to fill the gap. |
| [`excel_delete_rows_columns`](#excel-delete-rows-columns) | `Destructive` | Delete rows or columns starting at the specified 1-based index in a worksheet. |
| [`excel_delete_worksheet`](#excel-delete-worksheet) | `Destructive` | Permanently delete a worksheet from a workbook. |
| [`excel_find_cells`](#excel-find-cells) | `Read-Only` | Search for cells containing text or regular expression patterns across sheets. |
| [`excel_format_range`](#excel-format-range) | `Write` | Apply a wide range of visual and data formatting styles to a specified cell range in an Excel worksheet. |
| [`excel_format_sheet`](#excel-format-sheet) | `Write` | Configure worksheet-level display settings and layout properties. |
| [`excel_get_workbook_metadata`](#excel-get-workbook-metadata) | `Read-Only` | Retrieve metadata about an Excel workbook, including a list of all worksheets and optional details about named ranges. |
| [`excel_get_workbook_outline`](#excel-get-workbook-outline) | `Read-Only` | Inspect structural outline of an Excel workbook without dumping massive cell grids into LLM context. |
| [`excel_insert_rows_columns`](#excel-insert-rows-columns) | `Write` | Insert empty rows or columns at the specified 1-based index in a worksheet. |
| [`excel_list_documents`](#excel-list-documents) | `Read-Only` | List all .xlsx files in the specified directory. |
| [`excel_merge_cells`](#excel-merge-cells) | `Write` | Merge a rectangular range of cells into a single, larger cell. |
| [`excel_profile_sheet`](#excel-profile-sheet) | `Read-Only` | Generate lightweight column-level profiling and statistical summaries for a worksheet. |
| [`excel_protect_sheet`](#excel-protect-sheet) | `Write` | Protect or unprotect a worksheet, with optional per-action exceptions. |
| [`excel_read_data`](#excel-read-data) | `Read-Only` | Read a specific range of tabular data from an Excel worksheet and return it in a structured format. |
| [`excel_rename_worksheet`](#excel-rename-worksheet) | `Write` | Change the name of an existing worksheet within a workbook. |
| [`excel_sort_range`](#excel-sort-range) | `Write` | Sort a 2D range of cells in-place by a specific key column. |
| [`excel_unmerge_cells`](#excel-unmerge-cells) | `Write` | Unmerge a previously merged cell range, reverting it to individual cells. |
| [`excel_validate_formula_syntax`](#excel-validate-formula-syntax) | `Read-Only` | Validate the syntax of an Excel formula without writing it to a cell. |
| [`excel_validate_range`](#excel-validate-range) | `Read-Only` | Validate that a given cell range is valid within a specific worksheet. |
| [`excel_write_data`](#excel-write-data) | `Write` | Write data to an Excel worksheet beginning at the specified cell. |

## Batch & Single-Pass Tools

### `excel_build_workbook`

- **Access Mode:** `Write`
- **Signature:** `excel_build_workbook(filename: 'str | None' = None, file_path: 'str | None' = None, sheets: 'list[dict[str, Any]] | None' = None, preset: 'str' = 'corporate', auto_fit: 'bool' = True) -> 'dict[str, Any]'`

```text
Build a complete Excel workbook with typography, zebra striping, currency/date formats, formulas, and charts in a single pass.

Args:
    filename: Path to the .xlsx file to create or overwrite.
    file_path: Alternative alias for filename.
    sheets: List of sheet definitions. Each sheet dict contains:
        - 'name': Sheet title (str).
        - 'title': Optional banner title header (str).
        - 'headers': List of column header names (list[str]).
        - 'rows': 2D array of data values (list[list[Any]]).
        - 'auto_totals': bool or list of 0-based column indices to calculate SUM.
        - 'currency_columns': list of column names or 0-based indices for currency formatting ($#,##0.00).
        - 'chart': Optional chart definition {'type': 'bar'|'line'|'pie', 'title': '...', 'data_column': 2, 'label_column': 1}.
    preset: Visual styling preset ('corporate', 'financial', 'modern', 'compact').
    auto_fit: If True, calculates optimal column widths automatically.

Returns:
    Structured result summary including sheets created and total rows.
```

### `excel_get_workbook_outline`

- **Access Mode:** `Read-Only`
- **Signature:** `excel_get_workbook_outline(filename: 'str | None' = None, file_path: 'str | None' = None) -> 'dict[str, Any]'`

```text
Inspect structural outline of an Excel workbook without dumping massive cell grids into LLM context.

Args:
    filename: Path to the target .xlsx file.
    file_path: Alternative alias for filename.

Returns:
    Lightweight outline containing sheet names, dimensions (e.g. A1:H45),
    detected header titles, and formula presence.
```

### `excel_batch_write_data`

- **Access Mode:** `Write`
- **Signature:** `excel_batch_write_data(filename: 'str | None' = None, file_path: 'str | None' = None, sheet_name: 'str | None' = None, start_cell: 'str' = 'A1', data_matrix: 'list[list[Any]] | None' = None) -> 'dict[str, Any]'`

```text
Write a 2D matrix of values or formulas into a worksheet starting from a specific cell in one atomic operation.

Args:
    filename: Path to the target .xlsx file.
    file_path: Alternative alias for filename.
    sheet_name: Worksheet name (defaults to active sheet if omitted).
    start_cell: Top-left coordinate (e.g., 'A1', 'B5').
    data_matrix: 2D array of values or formula strings to write.

Returns:
    Summary of updated cell range and row/column counts.
```

### `excel_append_rows`

- **Access Mode:** `Write`
- **Signature:** `excel_append_rows(filename: 'str | None' = None, file_path: 'str | None' = None, rows: 'list[list[Any]] | None' = None, sheet_name: 'str | None' = None, start_column: 'int' = 1) -> 'dict[str, Any]'`

```text
Append batch rows of data below the existing content in a worksheet in an atomic operation.

Detects the first unoccupied row automatically, validates formulas securely,
and commits changes atomically.

Args:
    filename: Path to the target .xlsx file.
    file_path: Alternative alias for filename.
    rows: 2D array of row values to append (e.g. [["val1", 10], ["val2", 20]]).
    sheet_name: Target sheet name (defaults to active sheet).
    start_column: Starting 1-based column index (defaults to 1 = column A).

Returns:
    Structured result with sheet name and appended row boundaries.
```

## Structure & Worksheet Management

### `excel_create_workbook`

- **Access Mode:** `Write`
- **Signature:** `excel_create_workbook(filename: str) -> dict[str, typing.Any]`

```text
Create a new Excel workbook (.xlsx) in a secure, validated path.

Context for LLM/AI: Use this tool when you need to programmatically generate a new Excel workbook file as the starting point for a reporting, data collection, or automation workflow. This is typically used to initialize new datasets or reports, ensuring that the file is created only within authorized directories for security and compliance. The tool will automatically add the `.xlsx` extension if missing and will validate the path to prevent unauthorized file creation.

When to use:
    - When starting a new workflow that requires a fresh Excel file.
    - When an automated agent needs to ensure the file is created securely.
    - When the filename or path may be user-supplied and must be validated for safety.

Args:
    filename (str): Name or path of the workbook to create. Example: "data/reporte_diario"

Returns:
    dict[str, Any]: A dictionary with status or error details.
        Example: {"status": "success", "filename": "/secure/path/reporte_diario.xlsx"}
    On failure, returns an error dictionary with a descriptive message.
```

### `excel_create_worksheet`

- **Access Mode:** `Write`
- **Signature:** `excel_create_worksheet(filename: str, sheet_name: str) -> dict[str, typing.Any]`

```text
Add a new worksheet to an existing Excel workbook.

Context for LLM/AI: Use this tool to organize or extend data in an existing Excel file by adding a new worksheet. This is ideal when you want to segment data by category, time period, or any logical grouping (e.g., adding a new month to a financial report or a new department to a tracking file). The tool ensures the file exists, is accessible, and that the new sheet name does not conflict with existing sheets. It enforces Excel naming conventions and security constraints.

When to use:
    - When augmenting an existing Excel file with additional data sections.
    - When automating workflows that require dynamic worksheet creation.
    - When you need to ensure that sheet names remain unique and valid.

Args:
    filename (str): Path to the existing Excel workbook. Example: "reports/monthly_report.xlsx"
    sheet_name (str): Name of the new worksheet to add. Must be unique and valid per Excel rules.

Returns:
    dict[str, Any]: Dictionary indicating success or describing the error.
        Example: {"status": "success", "sheet": "NewData"}
    On error, returns a dictionary with status "error" and a descriptive message.
```

### `excel_copy_worksheet`

- **Access Mode:** `Write`
- **Signature:** `excel_copy_worksheet(filename: str, source_sheet: str, target_sheet: str) -> dict[str, typing.Any]`

```text
Create a duplicate of an existing worksheet within the same Excel workbook.

Context for AI/LLM:
    Use this tool when you need to create a new worksheet based on an existing template or data sheet. It preserves all content, formatting, and formulas from the original.

Args:
    filename (str): Path to the Excel workbook.
    source_sheet (str): The name of the worksheet to be copied.
    target_sheet (str): The name for the newly created worksheet. Must not already exist.

Returns:
    dict[str, Any]: A status dictionary with a confirmation message or error details.
```

### `excel_delete_worksheet`

- **Access Mode:** `Destructive`
- **Signature:** `excel_delete_worksheet(filename: str, sheet_name: str) -> dict[str, typing.Any]`

```text
Permanently delete a worksheet from a workbook.

Context for AI/LLM:
    This is a destructive operation. Use it to remove temporary, outdated, or unnecessary worksheets from an Excel file as part of a cleanup or automation workflow. Ensure the agent confirms this action if the sheet contains data.

Args:
    filename (str): Path to the Excel workbook.
    sheet_name (str): The name of the worksheet to be deleted.

Returns:
    dict[str, Any]: A status dictionary indicating success or failure with a message.
```

### `excel_rename_worksheet`

- **Access Mode:** `Write`
- **Signature:** `excel_rename_worksheet(filename: str, old_name: str, new_name: str) -> dict[str, typing.Any]`

```text
Change the name of an existing worksheet within a workbook.

Context for AI/LLM:
    Use this tool to update worksheet names to be more descriptive or to follow a required naming convention as part of an automated process.

Args:
    filename (str): Path to the Excel workbook.
    old_name (str): The current name of the worksheet to be renamed.
    new_name (str): The new name for the worksheet. Must be unique and follow Excel's naming rules.

Returns:
    dict[str, Any]: A status dictionary with a confirmation or error message.
```

### `excel_list_documents`

- **Access Mode:** `Read-Only`
- **Signature:** `excel_list_documents() -> dict[str, typing.Any]`

```text
List all .xlsx files in the specified directory.

Context for LLM/AI: Use this tool to discover and enumerate all Excel files in a specific directory, which is helpful for inventory, audit, batch processing, or automated data discovery scenarios. The tool validates the directory for security, ensuring only authorized locations are scanned. It returns detailed metadata for each file, enabling downstream processing, selection, or reporting by an automated agent.

When to use:
    - When an AI needs to present the user with available Excel files for further action.
    - When preparing to process, analyze, or summarize multiple Excel documents in a folder.
    - When auditing or verifying the presence and properties of .xlsx files in a given path.

Args:
    directory (str): Path to the folder to scan. Defaults to the current directory (".")

Returns:
    dict[str, Any]: A dictionary with:
        - status (str): "success" or "error"
        - count (int): Number of .xlsx files found
        - directory (str): Absolute path to the searched directory
        - files (list[dict]): List of file metadata, e.g.:
            [{"name": "report.xlsx", "size": 2048, "modified": "2025-07-30T10:45:00"}]
    On error, returns a dictionary with status "error" and a descriptive message.
```

### `excel_get_workbook_metadata`

- **Access Mode:** `Read-Only`
- **Signature:** `excel_get_workbook_metadata(filename: str, include_ranges: bool = False) -> dict[str, typing.Any]`

```text
Retrieve metadata about an Excel workbook, including a list of all worksheets and optional details about named ranges.

Context for AI/LLM:
    Use this tool for discovery and inspection. It allows an agent to understand the structure of a workbook before performing read, write, or format operations. It's a crucial first step in many interactive or complex workflows.

Args:
    filename (str): Path to the Excel workbook to inspect.
    include_ranges (bool, optional): If True, the metadata will include details about named ranges. Defaults to False.

Returns:
    dict[str, Any]: A dictionary containing workbook metadata, such as a list of sheet names and properties.
```

### `excel_insert_rows_columns`

- **Access Mode:** `Write`
- **Signature:** `excel_insert_rows_columns(filename: 'str | None' = None, file_path: 'str | None' = None, dimension: "Literal['row', 'column']" = 'row', index: 'int' = 1, amount: 'int' = 1, sheet_name: 'str | None' = None) -> 'dict[str, Any]'`

```text
Insert empty rows or columns at the specified 1-based index in a worksheet.

Args:
    filename: Path to the target .xlsx file.
    file_path: Alternative alias for filename.
    dimension: 'row' to insert rows, or 'column' to insert columns.
    index: 1-based index where rows or columns should be inserted.
    amount: Number of empty rows or columns to insert (default 1).
    sheet_name: Target sheet name (defaults to active sheet).

Returns:
    Summary of the dimension inserted.
```

### `excel_delete_rows_columns`

- **Access Mode:** `Destructive`
- **Signature:** `excel_delete_rows_columns(filename: 'str | None' = None, file_path: 'str | None' = None, dimension: "Literal['row', 'column']" = 'row', index: 'int' = 1, amount: 'int' = 1, sheet_name: 'str | None' = None) -> 'dict[str, Any]'`

```text
Delete rows or columns starting at the specified 1-based index in a worksheet.

Args:
    filename: Path to the target .xlsx file.
    file_path: Alternative alias for filename.
    dimension: 'row' to delete rows, or 'column' to delete columns.
    index: 1-based starting index of rows or columns to delete.
    amount: Number of rows or columns to delete (default 1).
    sheet_name: Target sheet name (defaults to active sheet).

Returns:
    Summary of the dimension deleted.
```

## Formatting, Tables & Layout

### `excel_format_range`

- **Access Mode:** `Write`
- **Signature:** `excel_format_range(filename: str, sheet_name: str, start_cell: str, end_cell: str | None = None, bold: bool = False, italic: bool = False, underline: bool = False, font_size: int | None = None, font_color: str | None = None, bg_color: str | None = None, border_style: Literal['dashDot', 'dashDotDot', 'dashed', 'dotted', 'double', 'hair', 'medium', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin', 'none', None] = None, border_color: str | None = None, number_format: str | None = None, alignment: str | None = None, wrap_text: bool = False, merge_cells: bool = False, protection: dict[str, typing.Any] | None = None, conditional_format: dict[str, typing.Any] | None = None) -> dict[str, typing.Any]`

```text
Apply a wide range of visual and data formatting styles to a specified cell range in an Excel worksheet.

Context for AI/LLM:
    Use this comprehensive tool to programmatically style reports, highlight data, or enforce a consistent visual layout in an Excel sheet. It is ideal for automating the final presentation layer of data generation workflows.

Typical use cases:
    1. Styling table headers with bold text, background colors, and borders.
    2. Applying currency or date formats to numerical data columns.
    3. Setting up conditional formatting to highlight values above a certain threshold.
    4. Merging cells to create report titles or grouped headers.

Args:
    filename (str): Path to the Excel workbook.
    sheet_name (str): The name of the worksheet to format.
    start_cell (str): The top-left cell of the target range (e.g., "A1").
    end_cell (str | None): The bottom-right cell of the target range (e.g., "B2").
    bold (bool, optional): Apply bold font style. Defaults to False.
    italic (bool, optional): Apply italic font style. Defaults to False.
    underline (bool, optional): Apply underline. Defaults to False.
    font_size (int | None, optional): Set the font size. Defaults to None.
    font_color (str | None, optional): Font color as a hex code (e.g., "#FF0000"). Defaults to None.
    bg_color (str | None, optional): Cell background color as a hex code. Defaults to None.
    border_style (str | None, optional): Border style (e.g., "thin", "thick"). Defaults to None.
    border_color (str | None, optional): Border color as a hex code. Defaults to None.
    number_format (str | None, optional): Number format code (e.g., "0.00%", "YYYY-MM-DD"). Defaults to None.
    alignment (str | None, optional): Horizontal alignment ("left", "center", "right"). Defaults to None.
    wrap_text (bool, optional): Enable text wrapping within cells. Defaults to False.
    merge_cells (bool, optional): Merge the entire specified range into a single cell. Defaults to False.
    protection (dict[str, Any] | None, optional): Cell protection settings (e.g., `{"locked": True}`). Defaults to None.
    conditional_format (dict[str, Any] | None, optional): Rules for conditional formatting. Defaults to None.

Returns:
    dict[str, Any]: A dictionary indicating the status ("success" or "error") and a descriptive message.
```

### `excel_format_sheet`

- **Access Mode:** `Write`
- **Signature:** `excel_format_sheet(filename: 'str | None' = None, file_path: 'str | None' = None, sheet_name: 'str | None' = None, freeze_panes: 'str | None' = None, auto_filter: 'bool | str | None' = None, show_grid_lines: 'bool | None' = None, column_widths: 'dict[str, float] | None' = None) -> 'dict[str, Any]'`

```text
Configure worksheet-level display settings and layout properties.

Enables setting freeze panes (e.g. 'A2' to freeze top row), enabling autofilter
on the data range, toggling gridline visibility, and specifying custom column widths.

Args:
    filename: Path to the target .xlsx file.
    file_path: Alternative alias for filename.
    sheet_name: Target sheet name (defaults to active sheet).
    freeze_panes: Cell reference to freeze panes below/right (e.g. 'A2', 'B2', or empty to clear).
    auto_filter: True to enable on used range, range string (e.g. 'A1:E50'), or False to clear.
    show_grid_lines: True to display gridlines, False to hide.
    column_widths: Mapping of column letters to widths (e.g. {'A': 18.0, 'B': 25.0}).

Returns:
    Summary of sheet properties updated.
```

### `excel_create_table`

- **Access Mode:** `Write`
- **Signature:** `excel_create_table(filename: 'str | None' = None, file_path: 'str | None' = None, range_string: 'str | None' = None, table_name: 'str | None' = None, style_name: 'str' = 'TableStyleMedium9', sheet_name: 'str | None' = None) -> 'dict[str, Any]'`

```text
Convert a range of cells into a native, styled Excel Table.

Native Excel tables provide automatic filtering drop-downs, structured references,
and banded row formatting.

Args:
    filename: Path to the target .xlsx file.
    file_path: Alternative alias for filename.
    range_string: Cell range coordinate (e.g. 'A1:D15').
    table_name: Optional unique table identifier (alphanumeric). Auto-generated if omitted.
    style_name: Built-in Excel table style name (default 'TableStyleMedium9').
    sheet_name: Target sheet name (defaults to active sheet).

Returns:
    Summary with table name, range, and applied style.
```

### `excel_sort_range`

- **Access Mode:** `Write`
- **Signature:** `excel_sort_range(filename: 'str | None' = None, file_path: 'str | None' = None, range_string: 'str | None' = None, key_column: 'int | str | None' = None, ascending: 'bool' = True, has_header: 'bool' = True, sheet_name: 'str | None' = None) -> 'dict[str, Any]'`

```text
Sort a 2D range of cells in-place by a specific key column.

Supports numeric, date, and text ordering with type-safe comparator logic.

Args:
    filename: Path to the target .xlsx file.
    file_path: Alternative alias for filename.
    range_string: Range of cells to sort (e.g. 'A1:E50').
    key_column: Column index (1-based), column letter (e.g. 'B'), or header name.
    ascending: True for ascending order, False for descending.
    has_header: True if the first row of range contains column headers that should stay in place.
    sheet_name: Target sheet name (defaults to active sheet).

Returns:
    Summary with sorted row count and key column used.
```

### `excel_merge_cells`

- **Access Mode:** `Write`
- **Signature:** `excel_merge_cells(filename: str, sheet_name: str, start_cell: str, end_cell: str) -> dict[str, typing.Any]`

```text
Merge a rectangular range of cells into a single, larger cell.

Context for AI/LLM:
    Use this tool to create titles, headers, or grouped labels that span multiple columns or rows. This is a common formatting step in report generation.

Args:
    filename (str): Path to the Excel workbook.
    sheet_name (str): The name of the worksheet where the merge will occur.
    start_cell (str): The top-left cell of the range to merge.
    end_cell (str): The bottom-right cell of the range to merge.

Returns:
    dict[str, Any]: A status dictionary indicating success or failure.
```

### `excel_unmerge_cells`

- **Access Mode:** `Write`
- **Signature:** `excel_unmerge_cells(filename: str, sheet_name: str, start_cell: str, end_cell: str) -> dict[str, typing.Any]`

```text
Unmerge a previously merged cell range, reverting it to individual cells.

Context for AI/LLM:
    Use this tool to reverse a merge operation, typically as part of a re-formatting or data extraction workflow where individual cell access is required.

Args:
    filename (str): Path to the Excel workbook.
    sheet_name (str): The name of the worksheet containing the merged cells.
    start_cell (str): The top-left cell of the range to unmerge.
    end_cell (str): The bottom-right cell of the range to unmerge.

Returns:
    dict[str, Any]: A status dictionary indicating success or failure.
```

### `excel_copy_range`

- **Access Mode:** `Write`
- **Signature:** `excel_copy_range(filename: str, sheet_name: str, source_start: str, source_end: str, target_start: str, target_sheet: str | None = None) -> dict[str, typing.Any]`

```text
Copy a range of cells, including their values and formatting, to a new location, potentially in a different worksheet.

Context for AI/LLM:
    Use this tool to duplicate data or templates within a workbook. It's useful for creating summaries, staging data for charts, or replicating formatted tables.

Args:
    filename (str): Path to the Excel workbook.
    sheet_name (str): The name of the source worksheet.
    source_start (str): The top-left cell of the source range to copy.
    source_end (str): The bottom-right cell of the source range.
    target_start (str): The top-left cell of the destination.
    target_sheet (str | None, optional): The name of the destination worksheet. If None, the same sheet is used. Defaults to None.

Returns:
    dict[str, Any]: A status dictionary indicating success or failure.
```

### `excel_delete_range`

- **Access Mode:** `Destructive`
- **Signature:** `excel_delete_range(filename: str, sheet_name: str, start_cell: str, end_cell: str, shift_direction: str = 'up') -> dict[str, typing.Any]`

```text
Delete a range of cells and optionally shift the surrounding cells to fill the gap.

Context for AI/LLM:
    This is a destructive operation used to remove rows or columns of data programmatically. It's useful for cleaning up datasets by removing invalid or unnecessary records.

Args:
    filename (str): Path to the Excel workbook.
    sheet_name (str): The name of the worksheet.
    start_cell (str): The top-left cell of the range to delete.
    end_cell (str): The bottom-right cell of the range to delete.
    shift_direction (str, optional): Direction to shift cells after deletion ('up' or 'left'). Defaults to "up".

Returns:
    dict[str, Any]: A status dictionary indicating success or failure.
```

### `excel_validate_range`

- **Access Mode:** `Read-Only`
- **Signature:** `excel_validate_range(filename: str, sheet_name: str, start_cell: str, end_cell: str | None = None) -> dict[str, typing.Any]`

```text
Validate that a given cell range is valid within a specific worksheet.

Context for AI/LLM:
    Use this tool as a precondition check before attempting to read from or write to a range. This helps prevent errors by ensuring the target sheet and cell references are valid.

Args:
    filename (str): Path to the Excel workbook.
    sheet_name (str): The name of the worksheet to check.
    start_cell (str): The top-left cell of the range.
    end_cell (str | None, optional): The bottom-right cell of the range. Defaults to None.

Returns:
    dict[str, Any]: A dictionary with validation status ("success" or "error") and a descriptive message.
```

## Search & Analytical Profiling

### `excel_find_cells`

- **Access Mode:** `Read-Only`
- **Signature:** `excel_find_cells(filename: 'str | None' = None, file_path: 'str | None' = None, query: 'str | None' = None, sheet_name: 'str | None' = None, match_case: 'bool' = False, exact_match: 'bool' = False, use_regex: 'bool' = False, max_results: 'int' = 100) -> 'dict[str, Any]'`

```text
Search for cells containing text or regular expression patterns across sheets.

Token-efficient: returns only matching cell coordinates, sheet names, and values
without dumping the surrounding grid into LLM context.

Args:
    filename: Path to the target .xlsx file.
    file_path: Alternative alias for filename.
    query: String or regex pattern to search for.
    sheet_name: Specific sheet name to search, or None to search all sheets.
    match_case: If True, search is case-sensitive.
    exact_match: If True, cell value must match query exactly.
    use_regex: If True, interprets query as a regular expression.
    max_results: Maximum number of matches to return (default 100).

Returns:
    Structured list of matching cells with coordinates, values, and formula flags.
```

### `excel_profile_sheet`

- **Access Mode:** `Read-Only`
- **Signature:** `excel_profile_sheet(filename: 'str | None' = None, file_path: 'str | None' = None, sheet_name: 'str | None' = None, sample_rows: 'int | None' = None) -> 'dict[str, Any]'`

```text
Generate lightweight column-level profiling and statistical summaries for a worksheet.

Token-efficient: reports detected column data types (numeric, integer, float,
string, datetime, boolean), null counts, min/max for numerics, and distinct samples
without transmitting full row payloads.

Args:
    filename: Path to the target .xlsx file.
    file_path: Alternative alias for filename.
    sheet_name: Specific sheet name to profile (defaults to active sheet).
    sample_rows: Maximum data rows to sample (default all rows).

Returns:
    Structured summary with column data types, min/max, nulls, and sample values.
```

## Formulas & Analytics

### `excel_validate_formula_syntax`

- **Access Mode:** `Read-Only`
- **Signature:** `excel_validate_formula_syntax(filename: str, sheet_name: str, cell: str, formula: str) -> dict[str, typing.Any]`

```text
Validate the syntax of an Excel formula without writing it to a cell.

Context for AI/LLM:
    Use this as a crucial pre-flight check before applying a formula, especially if the formula is user-supplied or dynamically generated. This prevents writing invalid formulas to the sheet, which could corrupt calculations.

Args:
    filename (str): Path to the Excel workbook.
    sheet_name (str): The worksheet context for the validation.
    cell (str): The cell reference to use as context for the formula.
    formula (str): The Excel formula string to validate (must start with '=').

Returns:
    dict[str, Any]: A dictionary with validation status ("success" or "error") and a message.
```

### `excel_apply_formula`

- **Access Mode:** `Write`
- **Signature:** `excel_apply_formula(filename: str, sheet_name: str, cell: str, formula: str) -> dict[str, typing.Any]`

```text
Apply a validated Excel formula to a specific cell in a worksheet.

Context for AI/LLM:
    Use this tool to perform calculations within Excel by programmatically inserting formulas. It's ideal for automating summaries, financial models, or any task requiring dynamic calculations based on other cell values.

Typical use cases:
    1. Adding a `=SUM(A1:A10)` formula to cell A11 to total a column.
    2. Placing a `=VLOOKUP(...)` formula to link data between sheets.
    3. Automating the creation of calculated columns in a data table.

Args:
    filename (str): Path to the Excel workbook.
    sheet_name (str): The name of the worksheet where the formula will be applied.
    cell (str): The cell reference to write the formula to (e.g., "A1").
    formula (str): The Excel formula to apply (e.g., "=SUM(A1:A10)").

Returns:
    dict[str, Any]: A status dictionary indicating success or failure, with a descriptive message.

Notes:
    • The tool automatically validates the formula's syntax before attempting to apply it.
    • The cell's existing value or formula will be overwritten.
```

### `excel_create_chart`

- **Access Mode:** `Write`
- **Signature:** `excel_create_chart(filename: str, sheet_name: str, data_range: str, chart_type: str, target_cell: str, title: str = '', x_axis: str = '', y_axis: str = '') -> dict[str, typing.Any]`

```text
Create a chart in a worksheet based on a specified data range.

Context for AI/LLM:
    Use this tool to visually represent data within an Excel sheet. It is perfect for automating the creation of dashboards and reports, allowing an agent to generate charts like bar, line, or pie charts from existing data tables.

Args:
    filename (str): The path to the Excel workbook.
    sheet_name (str): The name of the worksheet where the chart will be created.
    data_range (str): The cell range containing the data for the chart (e.g., "A1:B10").
    chart_type (str): The type of chart to create (e.g., 'bar', 'line', 'pie', 'scatter').
    target_cell (str): The top-left cell where the chart will be anchored.
    title (str, optional): The title of the chart. Defaults to "".
    x_axis (str, optional): The label for the x-axis. Defaults to "".
    y_axis (str, optional): The label for the y-axis. Defaults to "".

Returns:
    dict[str, Any]: A status dictionary indicating success or failure, with a descriptive message.
```

### `excel_create_pivot_table`

- **Access Mode:** `Write`
- **Signature:** `excel_create_pivot_table(filename: str, sheet_name: str, data_range: str, rows: list[str], values: list[str], columns: list[str] | None = None, agg_func: str = 'mean') -> dict[str, typing.Any]`

```text
Create a pivot table in a worksheet to summarize and analyze data from a given range.

Context for AI/LLM:
    Use this powerful tool to perform data aggregation and summarization automatically. It's ideal for creating summary reports that group data by different categories and calculate metrics like sums, averages, or counts.

Args:
    filename (str): The path to the Excel workbook.
    sheet_name (str): The name of the worksheet where the pivot table will be created.
    data_range (str): The source data range for the pivot table (e.g., "A1:D100").
    rows (list[str]): A list of column headers from the data range to use as rows in the pivot table.
    values (list[str]): A list of column headers to use for aggregation in the pivot table.
    columns (list[str] | None, optional): A list of column headers to use as columns. Defaults to None.
    agg_func (str, optional): The aggregation function to apply to the values (e.g., 'sum', 'mean', 'count'). Defaults to "mean".

Returns:
    dict[str, Any]: A status dictionary indicating success or failure, with a descriptive message.
```

## Direct Content I/O

### `excel_read_data`

- **Access Mode:** `Read-Only`
- **Signature:** `excel_read_data(filename: str, sheet_name: str, start_cell: str = 'A1', end_cell: str | None = None, preview_only: bool = False) -> dict[str, typing.Any]`

```text
Read a specific range of tabular data from an Excel worksheet and return it in a structured format.

Context for AI/LLM:
    Use this tool to extract tabular data from a defined worksheet and cell range within an Excel workbook.
    Ideal for automated data ingestion, previewing spreadsheet contents, or validating required data before processing.

Typical use cases:
    1. Loading a data range for analysis or transformation.
    2. Previewing the first N rows of a sheet to confirm layout.
    3. Validating that a required sheet exists and contains data.

Args:
    filename (str): Absolute or relative path to the Excel workbook (.xlsx). The extension is enforced automatically.
    sheet_name (str): Name of the worksheet to read from.
    start_cell (str, optional): Top-left cell reference of the range to read. Defaults to "A1".
    end_cell (str | None, optional): Bottom-right cell reference. If None, reads until the first empty row/column. Defaults to None.
    preview_only (bool, optional): If True, returns only a small subset (e.g., first 100 rows) for quick inspection. Defaults to False.

Returns:
    dict[str, Any]: A dictionary containing:
        - status (str): "success" or "error".
        - data (list[list[Any]] | None): 2-D list of cell values from the worksheet range when status is "success".
        - message (str): Human-readable message or error description.

Raises:
    FileNotFoundError: If the workbook cannot be located.
    ValidationError | DataError: If the sheet name or range is invalid or missing.

Notes:
    • The specified worksheet must already exist; the function does not create new sheets.
    • If the target range is empty, status will be "error" with a descriptive message.
```

### `excel_write_data`

- **Access Mode:** `Write`
- **Signature:** `excel_write_data(filename: str, sheet_name: str, data: list[list[typing.Any]], start_cell: str = 'A1') -> str`

```text
Write data to an Excel worksheet beginning at the specified cell.

Context for AI/LLM:
    Employ this tool to programmatically populate or update a worksheet with tabular data produced by prior computations or external systems. Suitable for automated report generation, data export, or incremental updates to existing files.

Typical use cases:
    1. Dumping processed data frames into Excel for business users.
    2. Appending new monthly records to a reporting sheet.
    3. Overwriting a template region with fresh values for distribution.

Args:
    filename (str): Path to the workbook. The .xlsx extension is enforced.
    sheet_name (str): Target worksheet name. Must already exist.
    data (list[list[Any]]): 2-D list where each sub-list represents a row to write.
    start_cell (str, optional): Top-left cell where writing begins (e.g., "A1"). Defaults to "A1".

Returns:
    str: Success confirmation message, or an error description prefixed with "Error:".

Raises:
    ValidationError | DataError: When validation of input parameters or write operation fails.
    Exception: For unexpected I/O or library errors.

Notes:
    • Data will overwrite any existing content in the target range.
    • The function does not create new worksheets; the target sheet must already exist.
    • Input data is written as-is, without formula injection or type conversion.
```

## Validation & Print Layout

### `excel_add_data_validation`

- **Access Mode:** `Write`
- **Signature:** `excel_add_data_validation(filename: 'str | None' = None, file_path: 'str | None' = None, cell_range: 'str | None' = None, validation_type: "Literal['list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom']" = 'list', operator: "Literal['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']" = 'between', formula1: 'str | None' = None, formula2: 'str | None' = None, allow_blank: 'bool' = True, prompt_title: 'str | None' = None, prompt: 'str | None' = None, error_title: 'str | None' = None, error_message: 'str | None' = None, sheet_name: 'str | None' = None) -> 'dict[str, Any]'`

```text
Add a data validation rule to a worksheet range.

Restricts what users (or generated templates) may enter in a range: a
dropdown list, a whole/decimal number bound, a date or time window, a text
length limit, or a custom formula.

Args:
    filename: Path to the target .xlsx file.
    file_path: Alternative alias for filename.
    cell_range: Target range in A1 notation (e.g. "A2:A100").
    validation_type: Rule type: list, whole, decimal, date, time, textLength or custom.
    operator: Comparison operator for non-list rules.
    formula1: List source (e.g. '"Open,Closed"'), lower bound or custom formula.
    formula2: Upper bound for between/notBetween rules.
    allow_blank: Whether empty cells bypass the validation.
    prompt_title: Optional input-message title.
    prompt: Optional input message shown when the cell is selected.
    error_title: Optional error-dialog title.
    error_message: Optional error-dialog message.
    sheet_name: Target sheet name (defaults to active sheet).

Returns:
    Structured result with the sheet, range and applied validation type.
```

### `excel_protect_sheet`

- **Access Mode:** `Write`
- **Signature:** `excel_protect_sheet(filename: 'str | None' = None, file_path: 'str | None' = None, enabled: 'bool' = True, password: 'str | None' = None, sheet_name: 'str | None' = None, allow_select_locked_cells: 'bool' = True, allow_select_unlocked_cells: 'bool' = True, allow_format_cells: 'bool' = False, allow_format_columns: 'bool' = False, allow_format_rows: 'bool' = False, allow_insert_columns: 'bool' = False, allow_insert_rows: 'bool' = False, allow_insert_hyperlinks: 'bool' = False, allow_delete_columns: 'bool' = False, allow_delete_rows: 'bool' = False, allow_sort: 'bool' = False, allow_auto_filter: 'bool' = False, allow_pivot_tables: 'bool' = False, allow_edit_objects: 'bool' = False, allow_edit_scenarios: 'bool' = False) -> 'dict[str, Any]'`

```text
Protect or unprotect a worksheet, with optional per-action exceptions.

Use this when delivering a workbook template: lock the sheet so structure
and formulas survive, then open only the cells or actions the recipient
needs (for example, allow sorting and autofilter but not row deletion).

Args:
    filename: Path to the target .xlsx file.
    file_path: Alternative alias for filename.
    enabled: True to protect the sheet, False to remove existing protection.
    password: Optional password required to unprotect the sheet.
    sheet_name: Target sheet name (defaults to active sheet).
    allow_select_locked_cells: Allow selecting locked cells.
    allow_select_unlocked_cells: Allow selecting unlocked cells.
    allow_format_cells: Allow formatting cells.
    allow_format_columns: Allow formatting columns.
    allow_format_rows: Allow formatting rows.
    allow_insert_columns: Allow inserting columns.
    allow_insert_rows: Allow inserting rows.
    allow_insert_hyperlinks: Allow inserting hyperlinks.
    allow_delete_columns: Allow deleting columns.
    allow_delete_rows: Allow deleting rows.
    allow_sort: Allow sorting.
    allow_auto_filter: Allow using autofilters.
    allow_pivot_tables: Allow creating or refreshing pivot tables.
    allow_edit_objects: Allow editing embedded objects.
    allow_edit_scenarios: Allow editing scenarios.

Returns:
    Structured result with the sheet name and resulting protection state.
```
