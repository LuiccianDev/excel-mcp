import re
from typing import Any, Final

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.utils.cell_utils import parse_cell_range, validate_cell_reference

FORMULA_PREFIX: Final[str] = "="
MAX_FORMULA_LENGTH: Final[int] = 8192  # Excel hard formula length limit

DANGEROUS_FUNCTIONS: Final[frozenset[str]] = frozenset(
    {
        "CALL",
        "EXEC",
        "RUN",
        "SYSTEM",
        "SHELL",
        "EVAL",
        "EVALUATE",
        "REGISTER.ID",
        "XLM.CALL",
        "XLM.EXEC",
        "XLM.RUN",
        "REGISTER",
        "CALL.LIBRARY",
        "VBA.CALL",
        "SQL.REQUEST",
    }
)
_UNSAFE_FUNCTIONS: Final[frozenset[str]] = frozenset(
    {
        "INDIRECT",
        "HYPERLINK",
        "WEBSERVICE",
        "DGET",
        "RTD",
        "IMAGE",
        "IMPORTXML",
        "IMPORTDATA",
        "IMPORTFEED",
        "IMPORTHTML",
        "IMPORTRANGE",
        "FILTERXML",
    }
)
_DANGEROUS_OR_UNSAFE_FUNCTIONS: Final[frozenset[str]] = (
    DANGEROUS_FUNCTIONS | _UNSAFE_FUNCTIONS
)

_UNC_EXFILTRATION_PATTERNS: Final[tuple[str, ...]] = (
    r"\\\\[^\\\s]+",
    r"(?:^|['\"(\s,=])(?:\\/|/\\|//)[a-z0-9_.\-]+",
)
_URL_EXFILTRATION_PATTERNS: Final[tuple[str, ...]] = (
    r"\b(?:http|ftp|https)://",
    r"file://",
    *_UNC_EXFILTRATION_PATTERNS,
)
_SCRIPT_CODE_PATTERNS: Final[tuple[str, ...]] = (
    r"<script",
    r"document\.write",
    r"eval\s*\(",
    r"__import__",
    r"\bos\.",
    r"subprocess",
    r"sys\.exit",
    r"open\s*\(",
)
_DDE_PATTERNS: Final[tuple[str, ...]] = (
    r"\bcmd\|",
    r"\bpowershell\|",
    r"\|\s*['\"]",
    r"\|[^|]*!",
)
INJECTION_PATTERNS: Final[list[str]] = [
    *_URL_EXFILTRATION_PATTERNS,
    *_SCRIPT_CODE_PATTERNS,
    *_DDE_PATTERNS,
]
_COMPILED_INJECTION_PATTERNS: Final[tuple[re.Pattern[str], ...]] = tuple(
    re.compile(pattern) for pattern in INJECTION_PATTERNS
)
_COMPILED_UNC_PATTERNS: Final[tuple[re.Pattern[str], ...]] = tuple(
    re.compile(pattern, re.IGNORECASE) for pattern in _UNC_EXFILTRATION_PATTERNS
)
_OOXML_FUNC_PREFIX_RE: Final[re.Pattern[str]] = re.compile(
    r"^(?:_?(?:XLFN|XLWS|XLPM|XLETA|XLL|XLUDF)\.)+"
)
_EXTERNAL_ADDIN_PREFIX_RE: Final[re.Pattern[str]] = re.compile(
    r"(?:^|\.)_?(?:XLL|XLUDF)\."
)
_FUNC_CALL_RE: Final[re.Pattern[str]] = re.compile(r"([A-Z_][A-Z0-9_.]*)\s*\(")
_STRICT_FUNC_CALL_RE: Final[re.Pattern[str]] = re.compile(
    r"(?<![A-Za-z0-9_])([A-Za-z_][A-Za-z0-9_.]*)\("
)
_OPERATOR_BOUND_FUNC_CALL_RE: Final[re.Pattern[str]] = re.compile(
    r"(?:^|[+\-*/^&=(,;])\s*([A-Za-z_][A-Za-z0-9_.]*)\s*\("
)
_PREFIXED_FUNC_CALL_RE: Final[re.Pattern[str]] = re.compile(
    r"^[+\-@]\s*(?:[A-Za-z_][A-Za-z0-9_.]*\(|[0-9A-Za-z_.$:]+\s*[+\-*/^&]\s*[A-Za-z_][A-Za-z0-9_.]*\()"
)
_PREFIXED_DDE_RE: Final[re.Pattern[str]] = re.compile(
    r"^[+\-@](?:[^|]*\|[^|]*!|[^|]*\|\s*['\"]|\s*(?:cmd|powershell|msexc)\s*\|)",
    re.IGNORECASE,
)


def _strip_ooxml_func_prefix(func_name: str) -> str:
    return _OOXML_FUNC_PREFIX_RE.sub("", func_name.upper())


def _has_balanced_parentheses(expr: str) -> tuple[bool, str]:
    open_count = 0
    close_count = 0
    depth = 0
    went_negative = False
    in_string = False
    idx = 0
    length = len(expr)

    while idx < length:
        ch = expr[idx]
        if ch == '"':
            if in_string and idx + 1 < length and expr[idx + 1] == '"':
                idx += 2
                continue
            in_string = not in_string
        elif not in_string:
            if ch == "(":
                open_count += 1
                depth += 1
            elif ch == ")":
                close_count += 1
                depth -= 1
                if depth < 0:
                    went_negative = True
        idx += 1

    if open_count != close_count:
        return False, f"Unbalanced parentheses: {open_count} open, {close_count} close"
    if went_negative:
        return False, "Unmatched closing parenthesis (unbalanced parentheses)"
    return True, ""


def validate_formula(formula: str) -> tuple[bool, str]:
    if not formula.startswith("="):
        return False, "Formula must start with '='"
    ok, msg = validate_formula_secure(formula)
    if not ok:
        return False, msg
    return True, "Formula is valid"


def validate_formula_secure(formula: str) -> tuple[bool, str]:
    """Comprehensive security validation for Excel formulas."""
    normalized = formula.strip()
    if not normalized:
        return False, "Formula cannot be empty"
    if not normalized.startswith("="):
        normalized = "=" + normalized

    raw_functions = set(_FUNC_CALL_RE.findall(normalized.upper()))
    external_addins = sorted(
        name for name in raw_functions if _EXTERNAL_ADDIN_PREFIX_RE.search(name)
    )
    if external_addins:
        return (
            False,
            f"Dangerous external add-in/UDF function detected: {external_addins[0]}",
        )

    functions_found = raw_functions | {
        _strip_ooxml_func_prefix(name) for name in raw_functions
    }

    dangerous_found = functions_found.intersection(DANGEROUS_FUNCTIONS)
    if dangerous_found:
        return False, f"Dangerous functions detected: {dangerous_found}"

    unsafe_found = functions_found.intersection(_UNSAFE_FUNCTIONS)
    if unsafe_found:
        return False, f"Unsafe function: {sorted(unsafe_found)[0]}"

    formula_lower = normalized.lower()
    for pattern in _COMPILED_INJECTION_PATTERNS:
        if pattern.search(formula_lower):
            return False, "Suspicious pattern detected (potential injection)"

    balanced, balance_msg = _has_balanced_parentheses(normalized)
    if not balanced:
        return False, balance_msg

    if len(normalized) > MAX_FORMULA_LENGTH:
        return (
            False,
            f"Formula exceeds maximum length of {MAX_FORMULA_LENGTH} characters",
        )

    return True, "Formula is valid and secure"


def _is_blocked_function_name(func_name: str) -> bool:
    upper = func_name.upper()
    if _EXTERNAL_ADDIN_PREFIX_RE.search(upper):
        return True
    stripped_name = _strip_ooxml_func_prefix(upper)
    return (
        upper in _DANGEROUS_OR_UNSAFE_FUNCTIONS
        or stripped_name in _DANGEROUS_OR_UNSAFE_FUNCTIONS
    )


def _extract_prefixed_formula_candidate(stripped: str) -> str | None:
    body = stripped.lstrip("+-@ \t\r\n")
    if not body:
        return None
    candidate = body if body.startswith("=") else "=" + body

    if (
        _PREFIXED_DDE_RE.match(stripped)
        or body.startswith("=")
        or _PREFIXED_FUNC_CALL_RE.match(stripped)
        or any(pattern.search(stripped) for pattern in _COMPILED_UNC_PATTERNS)
    ):
        return candidate

    for pattern in (_STRICT_FUNC_CALL_RE, _OPERATOR_BOUND_FUNC_CALL_RE):
        for match in pattern.finditer(body):
            if _is_blocked_function_name(match.group(1)):
                return candidate

    return None


def validate_cell_value_secure(value: Any) -> tuple[bool, str]:
    """Validate an arbitrary cell value against formula and DDE injection."""
    if not isinstance(value, str):
        return True, "Non-string value is safe"

    stripped = value.lstrip()
    if not stripped:
        return True, "Empty string is safe"

    if stripped.startswith("="):
        return validate_formula_secure(stripped)

    if stripped[0] in "+-@":
        candidate = _extract_prefixed_formula_candidate(stripped)
        if candidate is not None:
            return validate_formula_secure(candidate)

    return True, "String value is safe"


def validate_formula_in_cell_operation(
    filepath: str, sheet_name: str, cell: str, formula: str
) -> dict[str, Any]:
    if not validate_cell_reference(cell):
        return {"status": "error", "message": f"Invalid cell reference: {cell}"}
    from excel_mcp.core._io import workbook_session

    try:
        with workbook_session(filepath, mode="read") as wb:
            if sheet_name not in wb.sheetnames:
                return {"status": "error", "message": f"Sheet '{sheet_name}' not found"}
            result: tuple[bool, str] = validate_formula(formula)
            is_valid, message = result
            if not is_valid:
                return {
                    "status": "error",
                    "message": f"Invalid formula syntax: {message}",
                }
            sec_valid, sec_message = validate_formula_secure(formula)
            if not sec_valid:
                return {
                    "status": "error",
                    "message": f"Formula security validation failed: {sec_message}",
                }
            cell_refs = re.findall(r"[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?", formula)
            for ref in cell_refs:
                if ":" in ref:
                    start, end = ref.split(":")
                    if not (
                        validate_cell_reference(start) and validate_cell_reference(end)
                    ):
                        return {
                            "status": "error",
                            "message": f"Invalid cell range reference in formula: {ref}",
                        }
                elif not validate_cell_reference(ref):
                    return {
                        "status": "error",
                        "message": f"Invalid cell reference in formula: {ref}",
                    }
            return {
                "status": "success",
                "message": f"Formula '{formula}' is valid and ready to be applied to cell {cell}",
                "valid": True,
                "cell": cell,
                "formula": formula,
            }
    except Exception as e:
        return {"status": "error", "message": str(e)}


def validate_range_in_sheet_operation(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str | None = None,
) -> dict[str, Any]:
    if not validate_cell_reference(start_cell):
        return {
            "status": "error",
            "message": f"Invalid start cell reference: {start_cell}",
        }
    if end_cell and not validate_cell_reference(end_cell):
        return {"status": "error", "message": f"Invalid end cell reference: {end_cell}"}
    from excel_mcp.core._io import workbook_session

    try:
        with workbook_session(filepath, mode="read") as wb:
            if sheet_name not in wb.sheetnames:
                return {"status": "error", "message": f"Sheet '{sheet_name}' not found"}
            worksheet = wb[sheet_name]
            data_max_row = worksheet.max_row
            data_max_col = worksheet.max_column
            try:
                start_row, start_col, end_row, end_col = parse_cell_range(
                    start_cell, end_cell
                )
            except ValueError as e:
                return {"status": "error", "message": f"Invalid range: {str(e)}"}
            if end_row is None:
                end_row = start_row
            if end_col is None:
                end_col = start_col
            is_valid, message = validate_range_bounds(
                worksheet, start_row, start_col, end_row, end_col
            )
            if not is_valid:
                return {"status": "error", "message": message}
            range_str = (
                f"{start_cell}" if end_cell is None else f"{start_cell}:{end_cell}"
            )
            data_range_str = f"A1:{get_column_letter(data_max_col)}{data_max_row}"
            extends_beyond_data = end_row > data_max_row or end_col > data_max_col
            return {
                "status": "success",
                "message": (
                    f"Range '{range_str}' is valid. "
                    f"Sheet contains data in range '{data_range_str}'"
                ),
                "valid": True,
                "range": range_str,
                "data_range": data_range_str,
                "extends_beyond_data": extends_beyond_data,
                "data_dimensions": {
                    "max_row": data_max_row,
                    "max_col": data_max_col,
                    "max_col_letter": get_column_letter(data_max_col),
                },
            }
    except Exception as e:
        return {"status": "error", "message": str(e)}


def validate_range_bounds(
    worksheet: Worksheet,
    start_row: int,
    start_col: int,
    end_row: int | None = None,
    end_col: int | None = None,
) -> tuple[bool, str]:
    max_row = worksheet.max_row
    max_col = worksheet.max_column
    try:
        if start_row < 1 or start_row > max_row:
            return False, f"Start row {start_row} out of bounds (1-{max_row})"
        if start_col < 1 or start_col > max_col:
            return False, (
                f"Start column {get_column_letter(start_col)} "
                f"out of bounds (A-{get_column_letter(max_col)})"
            )
        if end_row is not None and end_col is not None:
            if end_row < start_row:
                return False, "End row cannot be before start row"
            if end_col < start_col:
                return False, "End column cannot be before start column"
            if end_row > max_row:
                return False, f"End row {end_row} out of bounds (1-{max_row})"
            if end_col > max_col:
                return False, (
                    f"End column {get_column_letter(end_col)} "
                    f"out of bounds (A-{get_column_letter(max_col)})"
                )
        return True, "Range is valid"
    except Exception as e:
        return False, f"Invalid range: {e!s}"
