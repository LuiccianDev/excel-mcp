"""Core exceptions for the MCP Excel Office core module.

These exceptions are used by the core layer (core/*.py) for handling
errors related to workbooks, sheets, data, formulas, and validation.
"""


class CoreError(Exception):
    """Base exception for all core module errors.
    Used across core/*.py files.
    """

    pass


# * --- Workbook Exceptions ---
class WorkbookError(CoreError):
    """Base exception for all workbook-related errors.
    Used in core/workbook.py
    """

    pass


class WorkbookNotFoundError(WorkbookError):
    """Raised when a workbook file is not found.
    Used in core/workbook.py
    """

    pass


# * --- Sheet Exceptions ---
class SheetError(WorkbookError):
    """Base exception for sheet-related errors.
    Used in core/workbook.py, core/data.py
    """

    pass


class SheetNotFoundError(SheetError):
    """Raised when a specified sheet is not found.
    Used in core/workbook.py, core/data.py
    """

    pass


class SheetExistsError(SheetError):
    """Raised when trying to create a sheet that already exists.
    Used in core/workbook.py
    """

    pass


class WorksheetError(CoreError):
    """Raised when a worksheet operation fails.
    Used in core/workbook.py
    """

    pass


# * --- Data Exceptions ---
class DataError(CoreError):
    """Base exception for data-related errors.
    Used in core/data.py
    """

    pass


class InvalidDataError(DataError):
    """Raised when invalid data is provided.
    Used in core/data.py
    """

    pass


class CellReferenceError(DataError):
    """Base exception for cell reference related errors.
    Used in core/data.py
    """

    pass


class InvalidCellReferenceError(CellReferenceError):
    """Raised when a cell reference is invalid.
    Used in core/data.py
    """

    pass


class RangeError(DataError):
    """Raised when there's an error with cell ranges.
    Used in core/data.py
    """

    pass


# * --- Formula Exceptions ---
class FormulaError(CoreError):
    """Raised when there's an error with formula operations.
    Used in core/calculations.py
    """

    pass


# * --- Pivot Table Exceptions ---
class PivotError(CoreError):
    """Raised when pivot table operations fail.
    Used in core/pivot.py
    """

    pass


# * --- Validation Exceptions ---
class ValidationError(CoreError):
    """Raised when validation fails.
    Used across core/*.py files.
    """

    pass
