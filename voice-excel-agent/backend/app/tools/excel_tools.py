"""
Excel tools exposed to the agent.

SECURITY NOTE
-------------
This is the ONLY interface the LLM has to the workbook. There is no
`run_python`, `eval`, or `exec` tool anywhere in this codebase, and
there must never be one added here. Every tool below is a narrow,
typed wrapper around a single WorkbookManager method. If a requested
action can't be expressed by one of these tools, the agent must say
so instead of improvising code execution.

Formula strings (compute_formula/apply_formula) are evaluated in the backend
through a whitelisted parser (app/excel/formula_engine.py — no eval/exec) and
only the computed ANSWER is stored/returned — the raw '=...' text is never
pasted into a cell. That evaluator is not an arbitrary-code surface.

OLLAMA COMPATIBILITY NOTE
-------------------------
Ollama's chat template crashes (500: `index $prop.Type`) on JSON
schemas that use union types (`anyOf`: Optional[X], X | None, Any).
Therefore every tool signature below uses ONLY plain `str`, `int`,
`float`, `bool`, `list`, `dict` — never Optional/Union/Any. Missing
strings are `""`, missing ints use a documented sentinel (e.g. 0 =
append), missing lists/dicts are empty. Coercion helpers below convert
these sentinels back to None before calling the WorkbookManager.
"""

from __future__ import annotations

from typing import Any

from langchain_core.tools import tool

from app.excel.workbook_manager import WorkbookError, workbook_manager as wm


def _ok(payload: Any) -> dict:
    return {"success": True, "result": payload}


def _err(exc: Exception) -> dict:
    return {"success": False, "error": str(exc)}


def _guard(fn, *args, **kwargs):
    try:
        return _ok(fn(*args, **kwargs))
    except WorkbookError as exc:
        return _err(exc)
    except Exception as exc:  # keep the agent alive on unexpected errors too
        return _err(exc)


def _opt(s: str) -> str | None:
    """Empty/blank string sentinel -> None (real sheet/column names are never blank)."""
    return s if s and s.strip() else None


def _coerce_value(v: Any) -> Any:
    """Turn LLM string values into int/float/bool/None where obvious, so numbers
    land in Excel as numbers (SUM-able) instead of text."""
    if v is None:
        return None
    if isinstance(v, (int, float, bool)):
        return v
    s = str(v).strip()
    if s == "" or s.lower() in ("none", "null"):
        return None
    if s.lower() in ("true", "yes"):
        return True
    if s.lower() in ("false", "no"):
        return False
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s)
    except ValueError:
        pass
    return str(v)


def _coerce_list(values: list) -> list:
    return [_coerce_value(v) if not isinstance(v, list) else _coerce_list(v) for v in values]


def _coerce_dict(values: dict) -> dict:
    return {str(k): _coerce_value(v) for k, v in values.items()}


def _parse_bool_flag(s: str) -> bool | None:
    s = (s or "").strip().lower()
    if not s or s in ("none", "null", "unset"):
        return None
    if s in ("true", "yes", "1", "bold", "italic"):
        return True
    if s in ("false", "no", "0"):
        return False
    return None


# ---------------------------------------------------------------------- #
# Inspection tools (read-only — the agent should call these liberally)
# ---------------------------------------------------------------------- #
@tool
def get_workbook_summary() -> dict:
    """Get the list of sheets, active sheet, and file path of the currently open workbook.
    Always call this first if you are unsure what workbook/sheets exist."""
    return _guard(wm.get_summary)


@tool
def preview_sheet(sheet: str = "", max_rows: int = 25, max_cols: int = 12) -> dict:
    """Preview the top-left block of a sheet (25 rows/page). Row 1 is ALWAYS the
    column-names (header) row — it is pinned on every page, so later pages still
    show which column is which. Formula cells show their computed answers.
    sheet defaults to the active sheet if omitted (pass empty string)."""
    return _guard(wm.sheet_preview, sheet=_opt(sheet), max_rows=max_rows, max_cols=max_cols)


@tool
def read_range(cell_range: str, sheet: str = "") -> dict:
    """Read the values in a cell range, e.g. 'A1:C10'. Formula cells are returned as
    their computed answers, not the raw '=...' text. sheet defaults to the active sheet
    if omitted (pass empty string)."""
    return _guard(wm.read_range, cell_range=cell_range, sheet=_opt(sheet))


@tool
def find_value(query: str, sheet: str = "") -> dict:
    """Search the sheet for any cell whose text contains the given query (case-insensitive).
    Returns matching cell coordinates and values. Use this to locate a person/row by name
    before updating or deleting it. sheet defaults to the active sheet if omitted."""
    return _guard(wm.find, query=query, sheet=_opt(sheet))


@tool
def find_row_by_column_value(column: str, value: str, sheet: str = "") -> dict:
    """Find the row number (excluding the header row) where `column` (a header name or
    column letter) equals `value`. Returns the 1-based row number, or null if not found.
    Use this before update_row/delete_row when the user refers to a record by name
    (e.g. 'set Arun's salary to 40000'). sheet defaults to the active sheet if omitted."""
    return _guard(wm.find_row_by_value, column=column, value=value, sheet=_opt(sheet))


@tool
def get_headers(sheet: str = "") -> dict:
    """Get the header row (column names) of a sheet. sheet defaults to the active sheet
    if omitted (pass empty string)."""
    return _guard(wm.get_headers, sheet=_opt(sheet))


@tool
def list_sheets() -> dict:
    """List all sheet names in the workbook."""
    return _guard(wm.list_sheets)


# ---------------------------------------------------------------------- #
# Cell / range writes
# ---------------------------------------------------------------------- #
@tool
def write_cell(cell: str, value: str, sheet: str = "") -> dict:
    """Write a single value into a cell, e.g. cell='B2', value='42'. Numbers given as
    text are stored as numbers. sheet defaults to the active sheet if omitted."""
    return _guard(wm.write_cell, cell=cell, value=_coerce_value(value), sheet=_opt(sheet))


@tool
def write_range(start_cell: str, values: list, sheet: str = "") -> dict:
    """Write a 2D block of values starting at start_cell (top-left corner).
    values is a list of rows, each row a list of cell values. sheet defaults to
    the active sheet if omitted (pass empty string)."""
    return _guard(wm.write_range, start_cell=start_cell, values=_coerce_list(values), sheet=_opt(sheet))


# ---------------------------------------------------------------------- #
# Row operations
# ---------------------------------------------------------------------- #
@tool
def add_row(values: list, sheet: str = "", at_row: int = 0) -> dict:
    """Append a new row of values to the end of a sheet (or insert at at_row if given
    a positive row number; 0 means append). values must be ordered to match the
    existing column headers — call get_headers or preview_sheet first to confirm
    the column order. sheet defaults to the active sheet if omitted."""
    return _guard(wm.add_row, values=_coerce_list(values), sheet=_opt(sheet),
                  at_row=at_row if at_row > 0 else None)


@tool
def update_row(row_number: int, values: dict, sheet: str = "") -> dict:
    """Update specific columns of an existing row. row_number is the 1-based sheet row
    (use find_row_by_column_value to look it up). values maps column header name (or
    column letter) to the new value, e.g. {"Salary": 40000}. sheet defaults to the
    active sheet if omitted (pass empty string)."""
    return _guard(wm.update_row, row_number=row_number, values=_coerce_dict(values), sheet=_opt(sheet))


@tool
def delete_row(row_number: int, sheet: str = "") -> dict:
    """Delete a row by its 1-based row number. This is destructive — only call it after
    the user has confirmed. sheet defaults to the active sheet if omitted."""
    return _guard(wm.delete_row, row_number=row_number, sheet=_opt(sheet))


# ---------------------------------------------------------------------- #
# Column operations
# ---------------------------------------------------------------------- #
@tool
def add_column(header: str, values: list = [], sheet: str = "") -> dict:  # noqa: B006
    """Add a new column at the end of the sheet with the given header and optional values
    (pass an empty list if none). sheet defaults to the active sheet if omitted."""
    return _guard(wm.add_column, header=header, values=_coerce_list(values) or None, sheet=_opt(sheet))


@tool
def rename_column(old_name: str, new_name: str, sheet: str = "") -> dict:
    """Rename a column header. sheet defaults to the active sheet if omitted."""
    return _guard(wm.rename_column, old_name=old_name, new_name=new_name, sheet=_opt(sheet))


@tool
def delete_column(column: str, sheet: str = "") -> dict:
    """Delete a column by header name or column letter (e.g. 'C'). Destructive — only
    call after user confirmation. sheet defaults to the active sheet if omitted."""
    return _guard(wm.delete_column, column=column, sheet=_opt(sheet))


# ---------------------------------------------------------------------- #
# Row/column insertion + moves (structural — confirmation-gated)
# ---------------------------------------------------------------------- #
@tool
def insert_column_before(reference_column: str, header: str, values: list = [], sheet: str = "") -> dict:  # noqa: B006
    """Insert a new column BEFORE reference_column (letter or header name), with the
    given header and optional values (empty list if none). Destructive/shifts data —
    only call after user confirmation. sheet defaults to the active sheet."""
    return _guard(wm.insert_column_before, reference_column=reference_column, header=header,
                  values=_coerce_list(values) or None, sheet=_opt(sheet))


@tool
def insert_column_after(reference_column: str, header: str, values: list = [], sheet: str = "") -> dict:  # noqa: B006
    """Insert a new column AFTER reference_column (letter or header name), with the
    given header and optional values (empty list if none). Destructive/shifts data —
    only call after user confirmation. sheet defaults to the active sheet."""
    return _guard(wm.insert_column_after, reference_column=reference_column, header=header,
                  values=_coerce_list(values) or None, sheet=_opt(sheet))


@tool
def insert_row_before(reference_row: int, values: list = [], sheet: str = "") -> dict:  # noqa: B006
    """Insert a new row BEFORE reference_row (1-based), with optional values ordered
    to match the columns (empty list for a blank row). Shifts data — only call after
    user confirmation. sheet defaults to the active sheet."""
    return _guard(wm.insert_row_before, reference_row=reference_row,
                  values=_coerce_list(values) or None, sheet=_opt(sheet))


@tool
def insert_row_after(reference_row: int, values: list = [], sheet: str = "") -> dict:  # noqa: B006
    """Insert a new row AFTER reference_row (1-based), with optional values ordered
    to match the columns (empty list for a blank row). Shifts data — only call after
    user confirmation. sheet defaults to the active sheet."""
    return _guard(wm.insert_row_after, reference_row=reference_row,
                  values=_coerce_list(values) or None, sheet=_opt(sheet))


@tool
def move_column(column: str, new_position: str, sheet: str = "") -> dict:
    """Move a column (letter or header name) to a new 1-based position (e.g. '1') or
    column letter (e.g. 'A'). Shifts data — only call after user confirmation.
    sheet defaults to the active sheet."""
    pos = new_position.strip()
    return _guard(wm.move_column, column=column,
                  new_position=int(pos) if pos.isdigit() else pos, sheet=_opt(sheet))


@tool
def move_row(row: int, new_position: int, sheet: str = "") -> dict:
    """Move a row (1-based) to a new 1-based position. Shifts data — only call after
    user confirmation. sheet defaults to the active sheet."""
    return _guard(wm.move_row, row=row, new_position=new_position, sheet=_opt(sheet))


# ---------------------------------------------------------------------- #
# Cell / range modifications
# ---------------------------------------------------------------------- #
@tool
def copy_range(source_range: str, dest_cell: str, sheet: str = "", source_sheet: str = "") -> dict:
    """Copy values and formatting from source_range (e.g. 'A1:C3') to dest_cell
    (top-left, e.g. 'E1'), within or across sheets. source_sheet defaults to sheet
    (pass empty string for the active sheet)."""
    return _guard(wm.copy_range, source_range=source_range, dest_cell=dest_cell,
                  sheet=_opt(sheet), source_sheet=_opt(source_sheet))


@tool
def clear_range(cell_range: str, sheet: str = "") -> dict:
    """Clear VALUES in a range without deleting the cells (formatting kept).
    Destructive for large ranges — confirm first when clearing many rows.
    sheet defaults to the active sheet."""
    return _guard(wm.clear_range, cell_range=cell_range, sheet=_opt(sheet))


@tool
def merge_cells(cell_range: str, sheet: str = "") -> dict:
    """Merge a cell range (e.g. 'A1:C1'). Structural — only call after user
    confirmation. sheet defaults to the active sheet."""
    return _guard(wm.merge_cells, cell_range=cell_range, sheet=_opt(sheet))


@tool
def unmerge_cells(cell_range: str, sheet: str = "") -> dict:
    """Unmerge a previously merged range. sheet defaults to the active sheet."""
    return _guard(wm.unmerge_cells, cell_range=cell_range, sheet=_opt(sheet))


@tool
def freeze_panes(cell: str, sheet: str = "") -> dict:
    """Freeze panes at a cell (e.g. 'B2' freezes row 1 + column A; 'A1' unfreezes).
    sheet defaults to the active sheet."""
    return _guard(wm.freeze_panes, cell=cell, sheet=_opt(sheet))


@tool
def set_column_width(column: str, width: float, sheet: str = "") -> dict:
    """Set a column's width (letter or header name; width in Excel units, ~8-60).
    sheet defaults to the active sheet."""
    return _guard(wm.set_column_width, column=column, width=width, sheet=_opt(sheet))


@tool
def set_row_height(row: int, height: float, sheet: str = "") -> dict:
    """Set a row's height in points. sheet defaults to the active sheet."""
    return _guard(wm.set_row_height, row=row, height=height, sheet=_opt(sheet))


@tool
def autofit_columns(sheet: str = "") -> dict:
    """Approximately autofit all column widths from content length (openpyxl has no
    true autofit; widths are estimated). sheet defaults to the active sheet."""
    return _guard(wm.autofit_columns, sheet=_opt(sheet))


@tool
def duplicate_sheet(source_name: str, new_name: str) -> dict:
    """Duplicate a sheet under a new name (values + formatting)."""
    return _guard(wm.duplicate_sheet, source_name=source_name, new_name=new_name)


@tool
def set_cell_comment(cell: str, text: str, sheet: str = "") -> dict:
    """Attach a comment/note to a cell. sheet defaults to the active sheet."""
    return _guard(wm.set_cell_comment, cell=cell, text=text, sheet=_opt(sheet))


@tool
def conditional_format(cell_range: str, rule_type: str, sheet: str = "", options: dict = {}) -> dict:  # noqa: B006
    """Apply conditional formatting. rule_type is one of: greater_than, less_than,
    between, color_scale, data_bars, duplicate_values. options carries the rule
    params: greater_than/less_than need {"threshold": 100}; between needs
    {"minimum": 0, "maximum": 100}; color_scale accepts {"min_color": hex,
    "max_color": hex}; data_bars accepts {"color": hex}; all cell rules accept
    {"fill_color": hex}. sheet defaults to the active sheet."""
    return _guard(wm.conditional_format, cell_range=cell_range, rule_type=rule_type,
                  sheet=_opt(sheet), **{str(k): v for k, v in (options or {}).items()})


@tool
def data_validation(cell_range: str, validation_type: str, sheet: str = "", params: dict = {}) -> dict:  # noqa: B006
    """Apply data validation. validation_type is one of: list (dropdown),
    whole_number, decimal, date. params examples: list -> {"options": ["A","B"]};
    numeric/date -> {"operator": "between", "minimum": 1, "maximum": 10}.
    sheet defaults to the active sheet."""
    return _guard(wm.data_validation, cell_range=cell_range, validation_type=validation_type,
                  sheet=_opt(sheet), **{str(k): v for k, v in (params or {}).items()})


@tool
def add_chart(chart_type: str, data_range: str, sheet: str = "", title: str = "", position: str = "") -> dict:
    """Add a bar/line/pie chart over data_range (e.g. 'A1:B6'). title and position
    (e.g. 'E5') are optional (pass empty string). sheet defaults to the active sheet."""
    return _guard(wm.add_chart, chart_type=chart_type, data_range=data_range,
                  sheet=_opt(sheet), title=_opt(title), position=_opt(position))


# ---------------------------------------------------------------------- #
# Sheet operations
# ---------------------------------------------------------------------- #
@tool
def create_sheet(name: str) -> dict:
    """Create a new, empty sheet with the given name."""
    return _guard(wm.create_sheet, name=name)


@tool
def rename_sheet(old_name: str, new_name: str) -> dict:
    """Rename an existing sheet."""
    return _guard(wm.rename_sheet, old_name=old_name, new_name=new_name)


@tool
def delete_sheet(name: str) -> dict:
    """Delete a sheet entirely. Destructive — only call after user confirmation."""
    return _guard(wm.delete_sheet, name=name)


@tool
def set_active_sheet(name: str) -> dict:
    """Switch the active sheet that subsequent operations default to."""
    return _guard(wm.set_active_sheet, name=name)


# ---------------------------------------------------------------------- #
# Sort / filter
# ---------------------------------------------------------------------- #
@tool
def sort_range(column: str, ascending: bool = True, sheet: str = "") -> dict:
    """Sort all data rows (below the header) by the given column, ascending or descending.
    sheet defaults to the active sheet if omitted (pass empty string)."""
    return _guard(wm.sort_range, column=column, ascending=ascending, sheet=_opt(sheet))


@tool
def filter_rows(column: str, condition: str, value: str, sheet: str = "") -> dict:
    """Filter rows where column meets condition against value.
    condition must be one of: eq, neq, gt, gte, lt, lte, contains.
    sheet defaults to the active sheet if omitted (pass empty string)."""
    return _guard(wm.filter_rows, column=column, condition=condition,
                  value=_coerce_value(value), sheet=_opt(sheet))


# ---------------------------------------------------------------------- #
# Formulas — computed in the backend, only the ANSWER is stored/shown
# ---------------------------------------------------------------------- #
@tool
def compute_formula(formula: str, sheet: str = "") -> dict:
    """Evaluate an Excel formula string in the backend and return the computed
    ANSWER without writing anything to the sheet. Use this when the user did NOT
    name a destination cell and just wants the number in chat (e.g. 'what is the
    total salary?'). sheet defaults to the active sheet if omitted."""
    return _guard(wm.compute_formula, formula=formula, sheet=_opt(sheet))


@tool
def apply_formula(cell: str, formula: str, sheet: str = "") -> dict:
    """Compute an Excel formula string in the backend and store the resulting
    ANSWER (static value) in `cell` — the raw '=...' text is NEVER pasted into
    the sheet. `cell` is the user-specified destination ('C5'; a request like
    'column C row 5' means C5 — resolve it first). The tool result returns the
    computed 'value': report it in chat. If the user named no destination cell,
    use compute_formula instead and just report the answer in chat.
    sheet defaults to the active sheet if omitted (pass empty string)."""
    return _guard(wm.apply_formula, cell=cell, formula=formula, sheet=_opt(sheet))


@tool
def apply_simple_formula(cell: str, formula_type: str, range_ref: str, sheet: str = "") -> dict:
    """Convenience wrapper for the common aggregate case. formula_type must be one of:
    SUM, AVERAGE, COUNT, MIN, MAX. range_ref is the range the formula applies to,
    e.g. 'C2:C50'. Computes in the backend and stores the ANSWER, never '=...' text.
    sheet defaults to the active sheet if omitted (pass empty string)."""
    return _guard(wm.apply_simple_formula, cell=cell, formula_type=formula_type,
                  range_ref=range_ref, sheet=_opt(sheet))


@tool
def suggest_formula(description: str, sheet: str = "") -> dict:
    """Given a plain-English calculation description (e.g. 'total sales where region
    is West and month is Jan'), inspect the sheet's headers/sample rows and return
    grounding context so you (the agent) can compose the correct live Excel formula
    (SUMIFS/VLOOKUP/INDEX-MATCH/IF/COUNTIFS/...) with your own reasoning, then write
    it via apply_formula. This tool does NOT compute anything — it only supplies the
    headers and sample data you need to write the formula. sheet defaults to the
    active sheet if omitted."""
    return _guard(wm.suggest_formula, description=description, sheet=_opt(sheet))


# ---------------------------------------------------------------------- #
# Formatting
# ---------------------------------------------------------------------- #
@tool
def format_cells(
    cell_range: str,
    sheet: str = "",
    bold: str = "",
    italic: str = "",
    font_color: str = "",
    fill_color: str = "",
    number_format: str = "",
    alignment: str = "",
    border: str = "",
) -> dict:
    """Apply formatting to a cell range. Colors are 6-digit hex strings without '#'
    (e.g. 'FF0000'); pass empty string for any option to leave it unchanged.
    bold/italic/border accept 'true'/'false' (empty = unchanged).
    number_format examples: '#,##0.00' for decimals, '0%' for percent.
    alignment is one of 'left', 'center', 'right' (empty = unchanged).
    sheet defaults to the active sheet if omitted (pass empty string)."""
    return _guard(
        wm.format_cells,
        cell_range=cell_range,
        sheet=_opt(sheet),
        bold=_parse_bool_flag(bold),
        italic=_parse_bool_flag(italic),
        font_color=_opt(font_color),
        fill_color=_opt(fill_color),
        number_format=_opt(number_format),
        alignment=_opt(alignment),
        border=_parse_bool_flag(border),
    )


# ---------------------------------------------------------------------- #
# Analysis
# ---------------------------------------------------------------------- #
@tool
def analyze_sheet(sheet: str = "") -> dict:
    """Run basic Pandas analysis on a sheet: row count, columns, and numeric summary
    statistics (mean, std, min, max, quartiles) for numeric columns.
    sheet defaults to the active sheet if omitted (pass empty string)."""
    return _guard(wm.analyze, sheet=_opt(sheet))


# ---------------------------------------------------------------------- #
# Save
# ---------------------------------------------------------------------- #
@tool
def save_workbook(save_as: str = "") -> dict:
    """Save the workbook to disk. If save_as is omitted (pass empty string),
    saves to the current path."""
    return _guard(wm.save_workbook, save_as=_opt(save_as))


ALL_TOOLS = [
    get_workbook_summary,
    preview_sheet,
    read_range,
    find_value,
    find_row_by_column_value,
    get_headers,
    list_sheets,
    write_cell,
    write_range,
    add_row,
    update_row,
    delete_row,
    add_column,
    rename_column,
    delete_column,
    insert_column_before,
    insert_column_after,
    insert_row_before,
    insert_row_after,
    move_column,
    move_row,
    copy_range,
    clear_range,
    merge_cells,
    unmerge_cells,
    freeze_panes,
    set_column_width,
    set_row_height,
    autofit_columns,
    duplicate_sheet,
    set_cell_comment,
    conditional_format,
    data_validation,
    add_chart,
    create_sheet,
    rename_sheet,
    delete_sheet,
    set_active_sheet,
    sort_range,
    filter_rows,
    compute_formula,
    apply_formula,
    apply_simple_formula,
    suggest_formula,
    format_cells,
    analyze_sheet,
    save_workbook,
]

# Tool names considered destructive — the agent layer requires an explicit
# confirmation flag before actually invoking these. Covers deletes, merges,
# clears, and insert/move operations that shift existing data.
DESTRUCTIVE_TOOLS = {
    "delete_row", "delete_column", "delete_sheet",
    "insert_column_before", "insert_column_after",
    "insert_row_before", "insert_row_after",
    "move_column", "move_row",
    "merge_cells", "clear_range",
}

# Threshold above which a delete_row / range clear is treated as "large"
# and requires confirmation even if DESTRUCTIVE_TOOLS confirmation is on
# a per-call basis (kept here for the agent/confirmation layer to reuse).
LARGE_RANGE_ROW_THRESHOLD = 10
