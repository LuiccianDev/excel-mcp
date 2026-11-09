from typing import Any

from excel_mcp.core.chart import create_chart_in_sheet as create_chart_impl
from excel_mcp.core.pivot import create_pivot_table as create_pivot_table_impl

# Import exceptions
from excel_mcp.exceptions.exception_tools import ChartError, PivotError, ValidationError
from excel_mcp.utils.file_utils import ensure_xlsx_extension, validate_file_access


@validate_file_access("filename")
async def create_chart(
    filename: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = "",
) -> dict[str, Any]:
    """Create a chart in a worksheet based on a specified data range.

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
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result: dict[str, Any] = create_chart_impl(
            filename=filename,
            sheet_name=sheet_name,
            data_range=data_range,
            chart_type=chart_type,
            target_cell=target_cell,
            title=title,
            x_axis=x_axis,
            y_axis=y_axis,
        )
        return result
    except (ValidationError, ChartError) as e:
        return {"error": f"Error: {str(e)}"}
    except Exception as e:
        return {"error": f"Failed to create chart: {str(e)}"}


@validate_file_access("filename")
async def create_pivot_table(
    filename: str,
    sheet_name: str,
    data_range: str,
    rows: list[str],
    values: list[str],
    columns: list[str] | None = None,
    agg_func: str = "mean",
) -> dict[str, Any]:
    """Create a pivot table in a worksheet to summarize and analyze data from a given range.

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
    """
    filename = ensure_xlsx_extension(filename)

    try:
        result: dict[str, Any] = create_pivot_table_impl(
            filename=filename,
            sheet_name=sheet_name,
            data_range=data_range,
            rows=rows,
            values=values,
            columns=columns or [],
            agg_func=agg_func,
        )
        return result
    except (ValidationError, PivotError) as e:
        return {"error": f"Error: {str(e)}"}
    except Exception as e:
        return {"error": f"Failed to create pivot table: {str(e)}"}
