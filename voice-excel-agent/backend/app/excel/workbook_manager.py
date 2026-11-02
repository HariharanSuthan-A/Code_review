"""
Excela AI — WorkbookManager
===========================
Thin, stateful wrapper around openpyxl that:
  * Keeps exactly one "active" workbook in memory per process/session.
  * Snapshots the workbook (as bytes) before every mutating operation so
    a single `undo()` call can restore the previous state.
  * Exposes small, structured, JSON-serialisable read helpers that the
    agent/tools layer uses to "inspect" the sheet before deciding what
    to do (this is what lets the LLM ground its actions in reality
    instead of guessing cell coordinates).

This module never evaluates user-supplied code. Every operation is a
plain Python method with explicit, typed parameters. Excel formula
strings (see compute_formula/apply_formula) are EVALUATED in the backend
through a whitelisted parser (formula_engine.py — no eval/exec) and only
the computed ANSWER is stored in cells / returned for chat — the raw
'=...' text is never pasted into the sheet. Row 1 of every sheet is the
column-names (header) row: it is styled/bordered, frozen (A2), and pinned
on every 25-row preview page.

Transport note (voice-ready architecture): Excela AI is text-only (chat)
for now; voice STT/TTS is a planned Phase 2 addition. The agent and tool
layers make no text-only assumptions — only the transport (API routes /
frontend) changes when voice is re-attached via the /ws/agent
status-stream contract.
"""
from __future__ import annotations

import copy
import io
import os
import re
import threading
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart
from openpyxl.chart.reference import Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

from app.excel.formula_engine import try_evaluate_on_workbook


MAX_UNDO_STACK = 20


class WorkbookError(Exception):
    """Raised for any user-facing workbook operation failure."""


@dataclass
class HistoryEntry:
    timestamp: str
    action: str
    detail: str


@dataclass
class WorkbookSession:
    """Holds workbook state for one logical session (single-user MVP)."""

    path: str | None = None
    workbook: Workbook | None = None
    active_sheet: str | None = None
    undo_stack: list[bytes] = field(default_factory=list)
    history: list[HistoryEntry] = field(default_factory=list)
    lock: threading.RLock = field(default_factory=threading.RLock)


class WorkbookManager:
    """
    Manages a single active workbook for the running server instance.
    In a multi-user deployment you'd key a dict of these by session id;
    the MVP keeps one global session for simplicity.
    """

    def __init__(self, workbooks_dir: str = "workbooks") -> None:
        self.workbooks_dir = workbooks_dir
        os.makedirs(self.workbooks_dir, exist_ok=True)
        self.session = WorkbookSession()

    # ------------------------------------------------------------------ #
    # Lifecycle
    # ------------------------------------------------------------------ #
    def _safe_filename(self, filename: str) -> str:
        """Strip any directory components so callers can't escape workbooks_dir
        (e.g. '../../tmp/pwned.xlsx' -> 'pwned.xlsx'). Rejects empty names,
        '.'/'..', and names with path separators that survive basenaming."""
        base = os.path.basename(str(filename or "").strip())
        if not base or base in (".", ".."):
            raise WorkbookError(f"Invalid filename {filename!r}")
        return base

    def _contained_path(self, target: str) -> str:
        """Resolve target under workbooks_dir; raise if it escapes."""
        base_dir = os.path.abspath(self.workbooks_dir)
        resolved = os.path.abspath(target if os.path.isabs(target) else os.path.join(self.workbooks_dir, target))
        if resolved != base_dir and not resolved.startswith(base_dir + os.sep):
            raise WorkbookError(f"Refusing to write outside the workbooks directory: {target!r}")
        return resolved

    def new_workbook(self, filename: str = "untitled.xlsx") -> dict[str, Any]:
        with self.session.lock:
            filename = self._safe_filename(filename)
            wb = Workbook()
            self.session.workbook = wb
            self.session.path = os.path.join(self.workbooks_dir, filename)
            self.session.active_sheet = wb.active.title
            # NOTE: do NOT touch any cell here (even reading A1 marks the
            # fresh sheet non-empty and breaks the first ws.append). Row 1
            # becomes the styled column-names row once headers are written.
            wb.active.freeze_panes = "A2"
            self.session.undo_stack.clear()
            self._log("new_workbook", f"Created new workbook '{filename}'")
            return self.get_summary()

    def open_workbook(self, file_path: str) -> dict[str, Any]:
        with self.session.lock:
            if not os.path.exists(file_path):
                raise WorkbookError(f"File not found: {file_path}")
            try:
                wb = load_workbook(file_path)
            except WorkbookError:
                raise
            except Exception as exc:
                raise WorkbookError(f"Cannot open {os.path.basename(file_path)!r} as an Excel workbook: {exc}") from exc
            self.session.workbook = wb
            self.session.path = file_path
            self.session.active_sheet = wb.active.title
            for ws in wb.worksheets:
                # First row of every sheet is the column-names row.
                self._style_header_row(ws)
                if not ws.freeze_panes:
                    ws.freeze_panes = "A2"
            self.session.undo_stack.clear()
            self._log("open_workbook", f"Opened workbook '{os.path.basename(file_path)}'")
            return self.get_summary()

    def save_workbook(self, save_as: str | None = None) -> dict[str, Any]:
        wb = self._wb()
        target = save_as or self.session.path
        if not target:
            raise WorkbookError("No path to save to. Provide save_as.")
        if save_as:
            # Sanitize a user-supplied name, then guarantee containment.
            target = self._contained_path(self._safe_filename(target))
        elif not os.path.isabs(target):
            # new_workbook()/open_workbook() already store paths under
            # workbooks_dir — don't join it twice.
            norm = os.path.normpath(target)
            if norm != os.path.normpath(self.workbooks_dir) and not norm.startswith(
                os.path.normpath(self.workbooks_dir) + os.sep
            ):
                target = os.path.join(self.workbooks_dir, target)
        os.makedirs(os.path.dirname(os.path.abspath(target)), exist_ok=True)
        wb.save(target)
        self.session.path = target
        self._log("save_workbook", f"Saved to '{target}'")
        return {"saved_path": target}

    def is_loaded(self) -> bool:
        return self.session.workbook is not None

    # ------------------------------------------------------------------ #
    # Snapshot / undo
    # ------------------------------------------------------------------ #
    def _snapshot(self) -> None:
        """Push current workbook bytes onto the undo stack."""
        wb = self._wb()
        buf = io.BytesIO()
        wb.save(buf)
        self.session.undo_stack.append(buf.getvalue())
        if len(self.session.undo_stack) > MAX_UNDO_STACK:
            self.session.undo_stack.pop(0)

    def undo(self) -> dict[str, Any]:
        with self.session.lock:
            if not self.session.undo_stack:
                raise WorkbookError("Nothing to undo.")
            prev_bytes = self.session.undo_stack.pop()
            buf = io.BytesIO(prev_bytes)
            self.session.workbook = load_workbook(buf)
            if self.session.active_sheet not in self.session.workbook.sheetnames:
                self.session.active_sheet = self.session.workbook.sheetnames[0]
            self._log("undo", "Reverted last change")
            return self.get_summary()

    # ------------------------------------------------------------------ #
    # Introspection (read-only, used by the agent to "look before it acts")
    # ------------------------------------------------------------------ #
    def get_summary(self) -> dict[str, Any]:
        wb = self._wb()
        return {
            "path": self.session.path,
            "sheets": wb.sheetnames,
            "active_sheet": self.session.active_sheet,
        }

    def list_sheets(self) -> list[str]:
        return self._wb().sheetnames

    def _is_empty_sheet(self, ws) -> bool:
        """True when the sheet has no DATA values. Inspects the internal
        cell cache only — never touches a cell, because even READING A1
        marks a fresh sheet non-empty for openpyxl's append()."""
        try:
            cells = ws._cells
        except AttributeError:
            return ws.max_row == 1 and ws.max_column == 1
        for cell in cells.values():
            try:
                if cell.value is not None:
                    return False
            except AttributeError:
                continue
        return True

    def _style_header_row(self, ws) -> None:
        """Treat row 1 as the column-names row: bold + highlighted fill +
        thin borders, so it renders as a bordered header in Excel too.
        Idempotent — safe to call whenever the sheet already has headers.
        Never call on a truly empty sheet (see _is_empty_sheet)."""
        if self._is_empty_sheet(ws):
            return
        thin = Side(style="thin", color="2A2F35")
        header_fill = PatternFill(start_color="2A3440", end_color="2A3440", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        for cell in next(ws.iter_rows(min_row=1, max_row=1)):
            if cell.value is None:
                continue
            cell.font = header_font
            cell.fill = header_fill
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    def _display_value(self, ws, cell) -> Any:
        """What the user sees in Excel: the computed answer for formula
        cells, the raw value otherwise. The stored formula is untouched —
        only the *displayed* value is evaluated. Returns the raw formula
        string when the evaluator does not support it."""
        v = cell.value
        if isinstance(v, str) and v.startswith("="):
            computed = try_evaluate_on_workbook(v, self._wb(), ws.title)
            return computed if computed is not None else v
        return v

    def sheet_preview(self, sheet: str | None = None, max_rows: int = 25, max_cols: int = 12, start_row: int = 1) -> dict[str, Any]:
        """Paginated grid preview (25 rows/page). Row 1 is ALWAYS the
        column-names (header) row: it is returned in `headers` AND pinned as
        the first entry of `preview` on every page, so pages 2+ still show
        which column is which. `start_row` applies to DATA rows only when
        start_row > 1; `data_start_row` reports the first data row shown."""
        ws = self._ws(sheet)
        self._style_header_row(ws)
        # Empty sheet: return the empty shape WITHOUT touching any cell,
        # so a later add_row still lands on row 1 (row 1 = future headers).
        if self._is_empty_sheet(ws):
            return {
                "sheet": ws.title,
                "max_row": 1,
                "max_column": 1,
                "start_row": 1,
                "data_start_row": 1,
                "headers": [],
                "column_letters": ["A"],
                "header_row": [],
                "header_row_number": 1,
                "page_size": max_rows,
                "preview": [],
            }
        headers = self.get_headers(ws.title)
        ncols = min(ws.max_column, max_cols)
        col_letters = [get_column_letter(c) for c in range(1, ncols + 1)]

        def _row_vals(r: int) -> list[Any]:
            return [self._display_value(ws, ws.cell(row=r, column=c)) for c in range(1, ncols + 1)]

        header_vals = _row_vals(1)
        start_row = max(1, int(start_row or 1))
        if start_row <= 1:
            end_row = min(ws.max_row, max_rows)
            rows_out = [_row_vals(r) for r in range(1, end_row + 1)]
            data_start = 1
        else:
            end_row = min(ws.max_row, start_row + max_rows - 1)
            page = [_row_vals(r) for r in range(start_row, end_row + 1)] if start_row <= ws.max_row else []
            # Pin the header row on top of every later page.
            rows_out = [header_vals] + page
            data_start = start_row
        return {
            "sheet": ws.title,
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "start_row": start_row,
            "data_start_row": data_start,
            "headers": headers[:ncols],
            "column_letters": col_letters,
            "header_row": header_vals,
            "header_row_number": 1,
            "page_size": max_rows,
            "preview": rows_out,
        }

    def read_range(self, cell_range: str, sheet: str | None = None) -> list[list[Any]]:
        ws = self._ws(sheet)
        try:
            cells = ws[cell_range]
        except Exception as exc:  # openpyxl raises plain Exception/ValueError
            raise WorkbookError(f"Invalid range '{cell_range}': {exc}") from exc
        if isinstance(cells, tuple):
            if isinstance(cells[0], tuple):
                return [[self._display_value(ws, c) for c in row] for row in cells]
            return [[self._display_value(ws, c) for c in cells]]
        return [[self._display_value(ws, cells)]]

    def find(self, query: str, sheet: str | None = None) -> list[dict[str, Any]]:
        ws = self._ws(sheet)
        query_l = str(query).lower()
        matches = []
        for row in ws.iter_rows():
            for cell in row:
                shown = self._display_value(ws, cell)
                if shown is not None and query_l in str(shown).lower():
                    matches.append({"cell": cell.coordinate, "value": shown, "row": cell.row, "col": cell.column})
        return matches

    def get_headers(self, sheet: str | None = None) -> list[str]:
        ws = self._ws(sheet)
        try:
            first = next(ws.iter_rows(min_row=1, max_row=1))
        except StopIteration:
            return []
        return [c.value for c in first]

    # ------------------------------------------------------------------ #
    # Cell / range writes
    # ------------------------------------------------------------------ #
    def write_cell(self, cell: str, value: Any, sheet: str | None = None) -> dict[str, Any]:
        if not re.fullmatch(r"[A-Za-z]{1,3}\d+", str(cell or "").strip()):
            raise WorkbookError(f"Invalid cell {cell!r}. Use e.g. 'B2'.")
        ws = self._ws(sheet)
        self._snapshot()
        ws[cell] = value
        try:
            import re as _re
            if int(_re.sub(r"[^0-9]", "", str(cell)) or 0) == 1:
                self._style_header_row(ws)
        except Exception:
            pass
        self._log("write_cell", f"{ws.title}!{cell} = {value!r}")
        return {"cell": cell, "value": value, "sheet": ws.title}

    def write_range(self, start_cell: str, values: list[list[Any]], sheet: str | None = None) -> dict[str, Any]:
        m = re.fullmatch(r"([A-Za-z]{1,3})(\d+)", str(start_cell or "").strip())
        if not m:
            raise WorkbookError(f"Invalid start cell {start_cell!r}. Use e.g. 'A1'.")
        ws = self._ws(sheet)
        self._snapshot()
        start_col = column_index_from_string(m.group(1).upper())
        start_row = int(m.group(2))
        if start_row < 1:
            raise WorkbookError(f"Invalid start cell {start_cell!r}. Row must be >= 1.")
        for r_off, row in enumerate(values):
            for c_off, val in enumerate(row):
                ws.cell(row=start_row + r_off, column=start_col + c_off, value=val)
        if start_row == 1:
            self._style_header_row(ws)
        self._log("write_range", f"{ws.title} starting at {start_cell}, {len(values)} rows")
        return {"start_cell": start_cell, "rows_written": len(values), "sheet": ws.title}

    # ------------------------------------------------------------------ #
    # Row operations
    # ------------------------------------------------------------------ #
    def add_row(self, values: list[Any], sheet: str | None = None, at_row: int | None = None) -> dict[str, Any]:
        self._snapshot()
        ws = self._ws(sheet)
        if at_row is None:
            if self._is_empty_sheet(ws):
                # Write headers straight into row 1 (ws.append would work
                # here only if no cell was ever touched — don't rely on it).
                for c_off, val in enumerate(values):
                    ws.cell(row=1, column=c_off + 1, value=val)
                row_idx = 1
            else:
                ws.append(values)
                row_idx = ws.max_row
        else:
            ws.insert_rows(at_row)
            for c_off, val in enumerate(values):
                ws.cell(row=at_row, column=c_off + 1, value=val)
            row_idx = at_row
        if row_idx == 1:
            self._style_header_row(ws)
        self._log("add_row", f"{ws.title} row {row_idx}: {values}")
        return {"row": row_idx, "sheet": ws.title}

    def update_row(self, row_number: int, values: dict[str, Any], sheet: str | None = None) -> dict[str, Any]:
        """values: mapping of column letter OR header name -> new value."""
        ws = self._ws(sheet)
        if int(row_number) < 1 or int(row_number) > ws.max_row:
            raise WorkbookError(f"Row {row_number} does not exist (sheet has {ws.max_row} rows)")
        self._snapshot()
        headers = {str(c.value): c.column for c in next(ws.iter_rows(min_row=1, max_row=1)) if c.value is not None}
        for key, val in values.items():
            if key.upper() in [get_column_letter(i) for i in range(1, ws.max_column + 1)] and key.isalpha():
                col_idx = column_index_from_string(key.upper())
            elif key in headers:
                col_idx = headers[key]
            else:
                raise WorkbookError(f"Unknown column reference '{key}'")
            ws.cell(row=row_number, column=col_idx, value=val)
        self._log("update_row", f"{ws.title} row {row_number}: {values}")
        return {"row": row_number, "sheet": ws.title, "updated": values}

    def delete_row(self, row_number: int, sheet: str | None = None) -> dict[str, Any]:
        ws = self._ws(sheet)
        if int(row_number) < 1 or int(row_number) > ws.max_row:
            raise WorkbookError(f"Row {row_number} does not exist (sheet has {ws.max_row} rows)")
        self._snapshot()
        ws.delete_rows(row_number)
        self._log("delete_row", f"{ws.title} row {row_number}")
        return {"deleted_row": row_number, "sheet": ws.title}

    def find_row_by_value(self, column: str, value: Any, sheet: str | None = None) -> int | None:
        """Find the first row (excluding header) whose given column/header matches value."""
        ws = self._ws(sheet)
        headers = {str(c.value): c.column for c in next(ws.iter_rows(min_row=1, max_row=1)) if c.value is not None}
        col_idx = headers.get(column)
        if col_idx is None and column.isalpha():
            col_idx = column_index_from_string(column.upper())
        if col_idx is None:
            raise WorkbookError(f"Unknown column '{column}'")
        for row in ws.iter_rows(min_row=2):
            cell = row[col_idx - 1]
            if cell.value is not None and str(cell.value).lower() == str(value).lower():
                return cell.row
        return None

    # ------------------------------------------------------------------ #
    # Column operations
    # ------------------------------------------------------------------ #
    def add_column(self, header: str, values: list[Any] | None = None, sheet: str | None = None) -> dict[str, Any]:
        self._snapshot()
        ws = self._ws(sheet)
        col_idx = ws.max_column + 1
        ws.cell(row=1, column=col_idx, value=header)
        if values:
            for r_off, val in enumerate(values):
                ws.cell(row=2 + r_off, column=col_idx, value=val)
        self._log("add_column", f"{ws.title} column '{header}' at {get_column_letter(col_idx)}")
        return {"column": get_column_letter(col_idx), "header": header, "sheet": ws.title}

    def rename_column(self, old_name: str, new_name: str, sheet: str | None = None) -> dict[str, Any]:
        self._snapshot()
        ws = self._ws(sheet)
        for cell in next(ws.iter_rows(min_row=1, max_row=1)):
            if str(cell.value) == old_name:
                cell.value = new_name
                self._log("rename_column", f"{ws.title} '{old_name}' -> '{new_name}'")
                return {"old_name": old_name, "new_name": new_name, "sheet": ws.title}
        raise WorkbookError(f"Column header '{old_name}' not found")

    def delete_column(self, column: str, sheet: str | None = None) -> dict[str, Any]:
        ws = self._ws(sheet)
        col_idx = self._resolve_column_index(ws, column)
        if col_idx > ws.max_column:
            raise WorkbookError(f"Column '{column}' does not exist (sheet has {ws.max_column} columns)")
        self._snapshot()
        ws.delete_cols(col_idx)
        self._log("delete_column", f"{ws.title} column {column}")
        return {"deleted_column": column, "sheet": ws.title}

    # ------------------------------------------------------------------ #
    # Row/column insertion (not just append/delete) + moves
    # ------------------------------------------------------------------ #
    def _resolve_column_index(self, ws: Worksheet, column: str | int) -> int:
        if isinstance(column, int):
            if not 1 <= column <= ws.max_column:
                raise WorkbookError(f"Column index {column} out of range (1..{ws.max_column})")
            return column
        col = str(column).strip()
        if re.fullmatch(r"[A-Za-z]{1,3}", col):
            try:
                return column_index_from_string(col.upper())
            except ValueError:
                pass
        headers = {str(c.value): c.column for c in next(ws.iter_rows(min_row=1, max_row=1)) if c.value is not None}
        if col in headers:
            return headers[col]
        raise WorkbookError(f"Unknown column '{column}'")

    def insert_column_before(
        self, reference_column: str, header: str, values: list[Any] | None = None, sheet: str | None = None
    ) -> dict[str, Any]:
        self._snapshot()
        ws = self._ws(sheet)
        idx = self._resolve_column_index(ws, reference_column)
        ws.insert_cols(idx)
        ws.cell(row=1, column=idx, value=header)
        if values:
            for r_off, val in enumerate(values):
                ws.cell(row=2 + r_off, column=idx, value=val)
        self._log("insert_column_before", f"{ws.title} '{header}' before {reference_column}")
        return {"column": get_column_letter(idx), "header": header, "sheet": ws.title}

    def insert_column_after(
        self, reference_column: str, header: str, values: list[Any] | None = None, sheet: str | None = None
    ) -> dict[str, Any]:
        self._snapshot()
        ws = self._ws(sheet)
        idx = self._resolve_column_index(ws, reference_column) + 1
        ws.insert_cols(idx)
        ws.cell(row=1, column=idx, value=header)
        if values:
            for r_off, val in enumerate(values):
                ws.cell(row=2 + r_off, column=idx, value=val)
        self._log("insert_column_after", f"{ws.title} '{header}' after {reference_column}")
        return {"column": get_column_letter(idx), "header": header, "sheet": ws.title}

    def insert_row_before(
        self, reference_row: int, values: list[Any] | None = None, sheet: str | None = None
    ) -> dict[str, Any]:
        if int(reference_row) < 1:
            raise WorkbookError(f"Invalid row {reference_row}")
        self._snapshot()
        ws = self._ws(sheet)
        ws.insert_rows(int(reference_row))
        if values:
            for c_off, val in enumerate(values):
                ws.cell(row=int(reference_row), column=c_off + 1, value=val)
        self._log("insert_row_before", f"{ws.title} row {reference_row}: {values}")
        return {"row": int(reference_row), "sheet": ws.title}

    def insert_row_after(
        self, reference_row: int, values: list[Any] | None = None, sheet: str | None = None
    ) -> dict[str, Any]:
        if int(reference_row) < 1:
            raise WorkbookError(f"Invalid row {reference_row}")
        self._snapshot()
        ws = self._ws(sheet)
        new_row = int(reference_row) + 1
        ws.insert_rows(new_row)
        if values:
            for c_off, val in enumerate(values):
                ws.cell(row=new_row, column=c_off + 1, value=val)
        self._log("insert_row_after", f"{ws.title} row {new_row}: {values}")
        return {"row": new_row, "sheet": ws.title}

    def move_column(self, column: str | int, new_position: int | str, sheet: str | None = None) -> dict[str, Any]:
        self._snapshot()
        ws = self._ws(sheet)
        src = self._resolve_column_index(ws, column)
        if isinstance(new_position, str) and re.fullmatch(r"[A-Za-z]{1,3}", new_position.strip()):
            dst = column_index_from_string(new_position.strip().upper())
        else:
            dst = int(new_position)
        if not 1 <= dst <= ws.max_column:
            raise WorkbookError(f"New position {new_position} out of range")
        if src == dst:
            return {"column": get_column_letter(src), "new_position": dst, "sheet": ws.title, "moved": False}
        col_cells = [ws.cell(row=r, column=src) for r in range(1, ws.max_row + 1)]
        saved = [(c.value, c.font.copy(), c.fill.copy(), c.border.copy(),
                  c.alignment.copy(), c.number_format, c.protection.copy()) for c in col_cells]
        ws.delete_cols(src)
        ws.insert_cols(dst)
        for r, (val, font, fill, border, alignment, num_fmt, prot) in enumerate(saved, start=1):
            c = ws.cell(row=r, column=dst)
            c.value = val
            c.font, c.fill, c.border, c.alignment = font, fill, border, alignment
            c.number_format, c.protection = num_fmt, prot
        self._log("move_column", f"{ws.title} column {column} -> position {dst}")
        return {"column": get_column_letter(dst), "new_position": dst, "sheet": ws.title}

    def move_row(self, row: int, new_position: int, sheet: str | None = None) -> dict[str, Any]:
        row, new_position = int(row), int(new_position)
        if row < 1 or new_position < 1:
            raise WorkbookError("Row numbers must be >= 1")
        ws = self._ws(sheet)
        if row > ws.max_row:
            raise WorkbookError(f"Row {row} does not exist (sheet has {ws.max_row} rows)")
        self._snapshot()
        if row == new_position:
            return {"row": row, "new_position": new_position, "sheet": ws.title, "moved": False}
        row_cells = [ws.cell(row=row, column=c) for c in range(1, ws.max_column + 1)]
        saved = [(c.value, c.font.copy(), c.fill.copy(), c.border.copy(),
                  c.alignment.copy(), c.number_format, c.protection.copy()) for c in row_cells]
        ws.delete_rows(row)
        dst = new_position
        ws.insert_rows(dst)
        for c_off, (val, font, fill, border, alignment, num_fmt, prot) in enumerate(saved, start=1):
            c = ws.cell(row=dst, column=c_off)
            c.value = val
            c.font, c.fill, c.border, c.alignment = font, fill, border, alignment
            c.number_format, c.protection = num_fmt, prot
        self._log("move_row", f"{ws.title} row {row} -> position {dst}")
        return {"row": dst, "new_position": dst, "sheet": ws.title}

    # ------------------------------------------------------------------ #
    # Cell / range modifications
    # ------------------------------------------------------------------ #
    def copy_range(
        self, source_range: str, dest_cell: str, sheet: str | None = None, source_sheet: str | None = None
    ) -> dict[str, Any]:
        self._snapshot()
        src_ws = self._ws(source_sheet or sheet)
        dst_ws = self._ws(sheet)
        try:
            min_c, min_r, max_c, max_r = range_boundaries(source_range)
        except Exception as exc:
            raise WorkbookError(f"Invalid source range '{source_range}': {exc}") from exc
        m = re.fullmatch(r"([A-Za-z]{1,3})(\d+)", dest_cell.strip())
        if not m:
            raise WorkbookError(f"Invalid destination cell '{dest_cell}'")
        dst_c = column_index_from_string(m.group(1).upper())
        dst_r = int(m.group(2))
        for r_off in range(max_r - min_r + 1):
            for c_off in range(max_c - min_c + 1):
                src = src_ws.cell(row=min_r + r_off, column=min_c + c_off)
                dst = dst_ws.cell(row=dst_r + r_off, column=dst_c + c_off)
                dst.value = src.value
                dst.font = src.font.copy()
                dst.fill = src.fill.copy()
                dst.border = src.border.copy()
                dst.alignment = src.alignment.copy()
                dst.number_format = src.number_format
                dst.protection = src.protection.copy()
        self._log("copy_range", f"{src_ws.title}!{source_range} -> {dst_ws.title}!{dest_cell}")
        return {"source_range": source_range, "dest_cell": dest_cell, "sheet": dst_ws.title}

    def clear_range(self, cell_range: str, sheet: str | None = None) -> dict[str, Any]:
        """Clear values without deleting the cells (formatting preserved)."""
        self._snapshot()
        ws = self._ws(sheet)
        try:
            cells = ws[cell_range]
        except Exception as exc:
            raise WorkbookError(f"Invalid range '{cell_range}': {exc}") from exc
        flat = []
        if isinstance(cells, tuple):
            for item in cells:
                flat.extend(item) if isinstance(item, tuple) else flat.append(item)
        else:
            flat = [cells]
        for c in flat:
            c.value = None
        self._log("clear_range", f"{ws.title}!{cell_range} ({len(flat)} cells cleared)")
        return {"range": cell_range, "cells_cleared": len(flat), "sheet": ws.title}

    def merge_cells(self, cell_range: str, sheet: str | None = None) -> dict[str, Any]:
        self._snapshot()
        ws = self._ws(sheet)
        try:
            ws.merge_cells(cell_range)
        except Exception as exc:
            raise WorkbookError(f"Cannot merge '{cell_range}': {exc}") from exc
        self._log("merge_cells", f"{ws.title}!{cell_range}")
        return {"range": cell_range, "sheet": ws.title}

    def unmerge_cells(self, cell_range: str, sheet: str | None = None) -> dict[str, Any]:
        self._snapshot()
        ws = self._ws(sheet)
        try:
            ws.unmerge_cells(cell_range)
        except Exception as exc:
            raise WorkbookError(f"Cannot unmerge '{cell_range}': {exc}") from exc
        self._log("unmerge_cells", f"{ws.title}!{cell_range}")
        return {"range": cell_range, "sheet": ws.title}

    def freeze_panes(self, cell: str, sheet: str | None = None) -> dict[str, Any]:
        self._snapshot()
        ws = self._ws(sheet)
        cell = (cell or "").strip() or "A1"
        ws.freeze_panes = None if cell.upper() == "A1" else cell
        self._log("freeze_panes", f"{ws.title} freeze at {cell}")
        return {"cell": cell, "sheet": ws.title}

    def set_column_width(self, column: str, width: float, sheet: str | None = None) -> dict[str, Any]:
        self._snapshot()
        ws = self._ws(sheet)
        idx = self._resolve_column_index(ws, column)
        ws.column_dimensions[get_column_letter(idx)].width = float(width)
        self._log("set_column_width", f"{ws.title} col {get_column_letter(idx)} width {width}")
        return {"column": get_column_letter(idx), "width": float(width), "sheet": ws.title}

    def set_row_height(self, row: int, height: float, sheet: str | None = None) -> dict[str, Any]:
        if int(row) < 1:
            raise WorkbookError(f"Invalid row {row}. Row numbers must be >= 1.")
        self._snapshot()
        ws = self._ws(sheet)
        ws.row_dimensions[int(row)].height = float(height)
        self._log("set_row_height", f"{ws.title} row {row} height {height}")
        return {"row": int(row), "height": float(height), "sheet": ws.title}

    def autofit_columns(self, sheet: str | None = None) -> dict[str, Any]:
        """Approximate autofit — openpyxl has no true autofit; width is
        estimated from the longest stringified cell content per column."""
        self._snapshot()
        ws = self._ws(sheet)
        widths = {}
        for col in ws.iter_cols():
            letter = get_column_letter(col[0].column)
            longest = 0
            for c in col:
                if c.value is not None:
                    longest = max(longest, len(str(c.value)))
            widths[letter] = min(60.0, max(8.0, longest + 2))
            ws.column_dimensions[letter].width = widths[letter]
        self._log("autofit_columns", f"{ws.title}: {widths}")
        return {"sheet": ws.title, "widths": widths}

    def duplicate_sheet(self, source_name: str, new_name: str) -> dict[str, Any]:
        self._snapshot()
        wb = self._wb()
        if source_name not in wb.sheetnames:
            raise WorkbookError(f"Sheet '{source_name}' not found")
        if new_name in wb.sheetnames:
            raise WorkbookError(f"Sheet '{new_name}' already exists")
        ws_copy = wb.copy_worksheet(wb[source_name])
        ws_copy.title = new_name
        self._log("duplicate_sheet", f"'{source_name}' -> '{new_name}'")
        return {"source": source_name, "new_sheet": new_name, "sheets": wb.sheetnames}

    def set_cell_comment(self, cell: str, text: str, sheet: str | None = None) -> dict[str, Any]:
        self._snapshot()
        ws = self._ws(sheet)
        ws[cell].comment = Comment(text, "Excela AI")
        self._log("set_cell_comment", f"{ws.title}!{cell}")
        return {"cell": cell, "sheet": ws.title}

    def conditional_format(
        self, cell_range: str, rule_type: str, sheet: str | None = None, **rule_params: Any
    ) -> dict[str, Any]:
        """rule_type: greater_than | less_than | between | color_scale |
        data_bars | duplicate_values. Extra params via rule_params, e.g.
        threshold=100, minimum=0/maximum=100, colors, fill_color."""
        self._snapshot()
        ws = self._ws(sheet)
        rt = rule_type.strip().lower()

        def _fill(color: str | None) -> PatternFill | None:
            if not color:
                return None
            hex_color = color.lstrip("#")
            if len(hex_color) == 6:
                hex_color = "FF" + hex_color
            return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

        try:
            if rt == "greater_than":
                threshold = rule_params.get("threshold", rule_params.get("value"))
                if threshold is None:
                    raise WorkbookError("greater_than requires threshold=value")
                fill = _fill(rule_params.get("fill_color", "FFC7CE"))
                ws.conditional_formatting.add(cell_range, CellIsRule(
                    operator="greaterThan", formula=[str(threshold)], fill=fill))
            elif rt == "less_than":
                threshold = rule_params.get("threshold", rule_params.get("value"))
                if threshold is None:
                    raise WorkbookError("less_than requires threshold=value")
                fill = _fill(rule_params.get("fill_color", "FFC7CE"))
                ws.conditional_formatting.add(cell_range, CellIsRule(
                    operator="lessThan", formula=[str(threshold)], fill=fill))
            elif rt == "between":
                lo = rule_params.get("minimum", rule_params.get("formula"))
                hi = rule_params.get("maximum", rule_params.get("formula2"))
                if lo is None or hi is None:
                    raise WorkbookError("between requires minimum= and maximum=")
                fill = _fill(rule_params.get("fill_color", "C6EFCE"))
                ws.conditional_formatting.add(cell_range, CellIsRule(
                    operator="between", formula=[str(lo), str(hi)], fill=fill))
            elif rt == "color_scale":
                ws.conditional_formatting.add(cell_range, ColorScaleRule(
                    start_type="min", start_color=rule_params.get("min_color", "F8696B"),
                    end_type="max", end_color=rule_params.get("max_color", "63BE7B")))
            elif rt == "data_bars":
                ws.conditional_formatting.add(cell_range, DataBarRule(
                    start_type="min", end_type="max", color=rule_params.get("color", "638EC6")))
            elif rt == "duplicate_values":
                fill = _fill(rule_params.get("fill_color", "FFC7CE"))
                col_letter = re.sub(r"[^A-Za-z]", "", cell_range.split(":")[0])
                first_row = re.sub(r"[^0-9]", "", cell_range.split(":")[0])
                ws.conditional_formatting.add(cell_range, FormulaRule(
                    formula=[f'COUNTIF(${col_letter}$1:${col_letter}$1048576,{col_letter}{first_row})>1'],
                    fill=fill))
            else:
                raise WorkbookError(
                    f"Unsupported rule_type '{rule_type}'. Use: greater_than, less_than, "
                    "between, color_scale, data_bars, duplicate_values")
        except WorkbookError:
            raise
        except Exception as exc:
            raise WorkbookError(f"Cannot apply conditional format: {exc}") from exc
        self._log("conditional_format", f"{ws.title}!{cell_range} {rule_type}")
        return {"range": cell_range, "rule_type": rt, "sheet": ws.title}

    def data_validation(
        self, cell_range: str, validation_type: str, sheet: str | None = None, **params: Any
    ) -> dict[str, Any]:
        """validation_type: list (dropdown) | whole_number | decimal | date.
        Params: for list: options=[...] or formula1; for numeric/date:
        operator=, minimum=/formula1, maximum=/formula2, allow_blank=,
        error_title=/error=, prompt_title=/prompt=."""
        self._snapshot()
        ws = self._ws(sheet)
        vt = validation_type.strip().lower()
        allow_blank = bool(params.get("allow_blank", True))
        try:
            if vt == "list":
                options = params.get("options", params.get("values"))
                if options is not None and not isinstance(options, str):
                    formula1 = '"' + ",".join(str(o) for o in options) + '"'
                else:
                    formula1 = params.get("formula1", str(options or ""))
                    if not formula1:
                        raise WorkbookError("list validation requires options=[...] or formula1=")
                dv = DataValidation(type="list", formula1=formula1, allow_blank=allow_blank)
            elif vt in ("whole_number", "whole"):
                dv = DataValidation(
                    type="whole",
                    operator=params.get("operator", "between"),
                    formula1=str(params.get("minimum", params.get("formula1", ""))),
                    formula2=str(params.get("maximum", params.get("formula2", ""))),
                    allow_blank=allow_blank)
            elif vt == "decimal":
                dv = DataValidation(
                    type="decimal",
                    operator=params.get("operator", "between"),
                    formula1=str(params.get("minimum", params.get("formula1", ""))),
                    formula2=str(params.get("maximum", params.get("formula2", ""))),
                    allow_blank=allow_blank)
            elif vt == "date":
                dv = DataValidation(
                    type="date",
                    operator=params.get("operator", "between"),
                    formula1=str(params.get("minimum", params.get("formula1", ""))),
                    formula2=str(params.get("maximum", params.get("formula2", ""))),
                    allow_blank=allow_blank)
            else:
                raise WorkbookError(
                    f"Unsupported validation_type '{validation_type}'. Use: list, whole_number, decimal, date")
            dv.errorTitle = params.get("error_title", "")
            dv.error = params.get("error", "")
            dv.promptTitle = params.get("prompt_title", "")
            dv.prompt = params.get("prompt", "")
            dv.showErrorMessage = bool(dv.error)
            dv.showDropDown = False if vt == "list" else True
            ws.add_data_validation(dv)
            dv.add(cell_range)
        except WorkbookError:
            raise
        except Exception as exc:
            raise WorkbookError(f"Cannot apply data validation: {exc}") from exc
        self._log("data_validation", f"{ws.title}!{cell_range} {validation_type}")
        return {"range": cell_range, "validation_type": vt, "sheet": ws.title}

    def add_chart(
        self, chart_type: str, data_range: str, sheet: str | None = None,
        title: str | None = None, position: str | None = None,
    ) -> dict[str, Any]:
        ct = chart_type.strip().lower()
        if ct not in ("bar", "line", "pie"):
            raise WorkbookError(f"Unsupported chart_type '{chart_type}'. Use: bar, line, pie")
        self._snapshot()
        ws = self._ws(sheet)
        try:
            min_c, min_r, max_c, max_r = range_boundaries(data_range)
        except Exception as exc:
            raise WorkbookError(f"Invalid data_range '{data_range}': {exc}") from exc
        data = Reference(ws, min_col=min_c, min_row=min_r, max_col=max_c, max_row=max_r)
        cats = Reference(ws, min_col=min_c, min_row=min_r + 1, max_row=max_r) if max_c >= min_c else None
        chart = {"bar": BarChart, "line": LineChart, "pie": PieChart}[ct]()
        chart.add_data(data, titles_from_data=True)
        if cats is not None and ct in ("bar", "line"):
            chart.set_categories(cats)
        if title:
            chart.title = title
        anchor = position or "E5"
        ws.add_chart(chart, anchor)
        self._log("add_chart", f"{ws.title} {chart_type} over {data_range} at {anchor}")
        return {"chart_type": ct, "data_range": data_range, "sheet": ws.title, "position": anchor}

    # ------------------------------------------------------------------ #
    # Sheet operations
    # ------------------------------------------------------------------ #
    def create_sheet(self, name: str, index: int | None = None) -> dict[str, Any]:
        self._snapshot()
        wb = self._wb()
        if name in wb.sheetnames:
            raise WorkbookError(f"Sheet '{name}' already exists")
        wb.create_sheet(title=name, index=index)
        # Fresh sheet is empty — pin row 1 but don't touch any cell yet.
        wb[name].freeze_panes = "A2"
        self._log("create_sheet", f"Created sheet '{name}'")
        return {"created_sheet": name, "sheets": wb.sheetnames}

    def rename_sheet(self, old_name: str, new_name: str) -> dict[str, Any]:
        self._snapshot()
        wb = self._wb()
        if old_name not in wb.sheetnames:
            raise WorkbookError(f"Sheet '{old_name}' not found")
        wb[old_name].title = new_name
        if self.session.active_sheet == old_name:
            self.session.active_sheet = new_name
        self._log("rename_sheet", f"'{old_name}' -> '{new_name}'")
        return {"old_name": old_name, "new_name": new_name, "sheets": wb.sheetnames}

    def delete_sheet(self, name: str) -> dict[str, Any]:
        self._snapshot()
        wb = self._wb()
        if name not in wb.sheetnames:
            raise WorkbookError(f"Sheet '{name}' not found")
        if len(wb.sheetnames) == 1:
            raise WorkbookError("Cannot delete the only remaining sheet")
        del wb[name]
        if self.session.active_sheet == name:
            self.session.active_sheet = wb.sheetnames[0]
        self._log("delete_sheet", f"Deleted sheet '{name}'")
        return {"deleted_sheet": name, "sheets": wb.sheetnames}

    def set_active_sheet(self, name: str) -> dict[str, Any]:
        wb = self._wb()
        if name not in wb.sheetnames:
            raise WorkbookError(f"Sheet '{name}' not found")
        self.session.active_sheet = name
        return {"active_sheet": name}

    # ------------------------------------------------------------------ #
    # Sort / filter
    # ------------------------------------------------------------------ #
    def sort_range(self, column: str, ascending: bool = True, sheet: str | None = None, has_header: bool = True) -> dict[str, Any]:
        self._snapshot()
        ws = self._ws(sheet)
        start_row = 2 if has_header else 1
        headers = {str(c.value): c.column for c in next(ws.iter_rows(min_row=1, max_row=1))} if has_header else {}
        col_idx = column_index_from_string(column.upper()) if column.isalpha() and len(column) <= 3 and column.upper() in [get_column_letter(i) for i in range(1, ws.max_column + 1)] else headers.get(column)
        if col_idx is None:
            raise WorkbookError(f"Unknown column '{column}'")

        data = [[cell.value for cell in row] for row in ws.iter_rows(min_row=start_row)]
        data.sort(key=lambda r: (r[col_idx - 1] is None, r[col_idx - 1]), reverse=not ascending)
        for r_off, row_vals in enumerate(data):
            for c_off, val in enumerate(row_vals):
                ws.cell(row=start_row + r_off, column=c_off + 1, value=val)
        self._log("sort_range", f"{ws.title} sorted by '{column}' ({'asc' if ascending else 'desc'})")
        return {"sorted_by": column, "ascending": ascending, "sheet": ws.title}

    def filter_rows(self, column: str, condition: str, value: Any, sheet: str | None = None) -> list[dict[str, Any]]:
        """condition: one of eq, neq, gt, gte, lt, lte, contains"""
        ws = self._ws(sheet)
        headers = {str(c.value): c.column for c in next(ws.iter_rows(min_row=1, max_row=1))}
        col_idx = headers.get(column)
        if col_idx is None and column.isalpha():
            col_idx = column_index_from_string(column.upper())
        if col_idx is None:
            raise WorkbookError(f"Unknown column '{column}'")

        ops = {
            "eq": lambda a, b: a == b,
            "neq": lambda a, b: a != b,
            "gt": lambda a, b: a is not None and a > b,
            "gte": lambda a, b: a is not None and a >= b,
            "lt": lambda a, b: a is not None and a < b,
            "lte": lambda a, b: a is not None and a <= b,
            "contains": lambda a, b: a is not None and str(b).lower() in str(a).lower(),
        }
        if condition not in ops:
            raise WorkbookError(f"Unknown condition '{condition}'")

        results = []
        for row in ws.iter_rows(min_row=2):
            cell_val = row[col_idx - 1].value
            try:
                cmp_val = type(cell_val)(value) if condition in ("gt", "gte", "lt", "lte") and cell_val is not None else value
            except (TypeError, ValueError):
                cmp_val = value
            if ops[condition](cell_val, cmp_val):
                results.append({"row": row[0].row, "values": [c.value for c in row]})
        return results

    # ------------------------------------------------------------------ #
    # Formulas — computed in the backend, only the ANSWER is stored/shown
    # ------------------------------------------------------------------ #
    def compute_formula(self, formula: str, sheet: str | None = None) -> dict[str, Any]:
        """Evaluate an Excel formula string in the backend (whitelisted
        parser — no eval/exec) and return the computed answer WITHOUT
        writing anything to the sheet. Use when the user just asks for a
        number in chat ('what is the total salary?')."""
        if not isinstance(formula, str) or not formula.startswith("=") or len(formula) < 2:
            raise WorkbookError("Formula must be a string starting with '=', e.g. '=SUM(C2:C50)'.")
        if formula.count("(") != formula.count(")"):
            raise WorkbookError(f"Formula has unbalanced parentheses: {formula!r}")
        if "\n" in formula or "\r" in formula:
            raise WorkbookError("Formula must be a single-line Excel formula string.")
        ws = self._ws(sheet)
        computed = try_evaluate_on_workbook(formula, self._wb(), ws.title)
        if computed is None:
            raise WorkbookError(
                f"Could not compute {formula!r} in the backend. "
                "It may use a function Excel supports but this backend evaluator does not — "
                "rephrase with SUM/AVERAGE/COUNT/MIN/MAX/SUMIFS/COUNTIFS/IF/VLOOKUP/INDEX-MATCH/TEXTJOIN/CONCATENATE."
            )
        self._log("compute_formula", f"{ws.title} {formula} => {computed!r}")
        return {"formula": formula, "value": computed, "sheet": ws.title}

    def apply_formula(self, cell: str, formula: str, sheet: str | None = None) -> dict[str, Any]:
        """Compute an Excel formula string in the backend and write the
        resulting ANSWER (a static value) into `cell` — the raw '=...'
        text is NEVER pasted into the sheet.

        `cell` is the user-specified destination ('C5', or 'Salary by row 5'
        resolved to a coordinate first). Report the returned 'value' in
        chat. If no destination was given, call compute_formula instead and
        just report its answer in chat.
        """
        computed = self.compute_formula(formula, sheet=sheet)["value"]
        self._snapshot()
        ws = self._ws(sheet)
        ws[cell] = computed
        self._log("apply_formula", f"{ws.title}!{cell} = {computed!r} (from {formula})")
        return {"cell": cell, "formula": formula, "value": computed, "sheet": ws.title}

    def apply_simple_formula(
        self, cell: str, formula_type: str, range_ref: str, sheet: str | None = None
    ) -> dict[str, Any]:
        """Convenience wrapper for the common aggregate case (SUM/AVERAGE/
        COUNT/MIN/MAX). Computes in the backend and stores the ANSWER, never
        the '=...' text. Kept so 'sum column C' stays a one-shot call."""
        valid = {"SUM", "AVERAGE", "COUNT", "MIN", "MAX"}
        formula_type = formula_type.upper()
        if formula_type not in valid:
            raise WorkbookError(f"Unsupported formula '{formula_type}'. Use one of {valid}")
        return self.apply_formula(cell, f"={formula_type}({range_ref})", sheet=sheet)

    def suggest_formula(self, description: str, sheet: str | None = None) -> dict[str, Any]:
        """Natural-language-to-formula helper (no ML model here).

        Returns the sheet's headers + a data sample alongside the plain-
        English description so the AGENT can compose the correct Excel
        formula itself (SUMIFS/VLOOKUP/INDEX-MATCH/IF/COUNTIFS/...) using
        its own reasoning, then write it via apply_formula. The agent —
        not this method — does the reasoning; this method only supplies
        the grounding context.
        """
        ws = self._ws(sheet)
        headers = self.get_headers(ws.title)
        preview = self.sheet_preview(ws.title, max_rows=6, max_cols=12)
        return {
            "description": description,
            "sheet": ws.title,
            "headers": headers,
            "sample": preview["preview"],
            "guidance": (
                "Compose a standard Excel formula string (e.g. SUMIFS/COUNTIFS/"
                "AVERAGEIFS/VLOOKUP/XLOOKUP/INDEX-MATCH/IF), then evaluate it via "
                "compute_formula (chat-only answer) or apply_formula (compute in "
                "backend and store the ANSWER in the user-specified cell, e.g. "
                "column C row 5 -> C5). Never paste raw '=...' text into a cell."
            ),
        }

    # ------------------------------------------------------------------ #
    # Formatting
    # ------------------------------------------------------------------ #
    def format_cells(
        self,
        cell_range: str,
        sheet: str | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        font_color: str | None = None,
        fill_color: str | None = None,
        number_format: str | None = None,
        alignment: str | None = None,
        border: bool | None = None,
    ) -> dict[str, Any]:
        self._snapshot()
        ws = self._ws(sheet)
        cells = ws[cell_range]
        flat_cells = []
        if isinstance(cells, tuple):
            for item in cells:
                if isinstance(item, tuple):
                    flat_cells.extend(item)
                else:
                    flat_cells.append(item)
        else:
            flat_cells = [cells]

        thin = Side(style="thin")
        for cell in flat_cells:
            if bold is not None or italic is not None or font_color is not None:
                cell.font = Font(
                    bold=bold if bold is not None else cell.font.bold,
                    italic=italic if italic is not None else cell.font.italic,
                    color=font_color if font_color is not None else cell.font.color,
                )
            if fill_color is not None:
                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            if number_format is not None:
                cell.number_format = number_format
            if alignment is not None:
                cell.alignment = Alignment(horizontal=alignment)
            if border:
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

        self._log("format_cells", f"{ws.title}!{cell_range}")
        return {"range": cell_range, "sheet": ws.title}

    # ------------------------------------------------------------------ #
    # Pandas-backed analysis
    # ------------------------------------------------------------------ #
    def analyze(self, sheet: str | None = None) -> dict[str, Any]:
        import pandas as pd

        ws = self._ws(sheet)
        # Use displayed (computed) values so formula cells contribute
        # their answers, not their "=..." strings, to the statistics.
        data = [
            [self._display_value(ws, c) for c in row]
            for row in ws.iter_rows()
        ]
        if not data:
            return {"error": "Sheet is empty"}
        headers, *rows = data
        # Normalize blank/duplicate headers so a stray value written into
        # row 1 (e.g. a formula answer in F1 with no header) doesn't leak
        # into the analysis as a bogus column name.
        seen: dict[str, int] = {}
        clean_headers = []
        for i, h in enumerate(headers):
            name = str(h).strip() if h is not None and str(h).strip() else f"Column_{i + 1}"
            seen[name] = seen.get(name, 0) + 1
            if seen[name] > 1:
                name = f"{name}_{seen[name]}"
            clean_headers.append(name)
        df = pd.DataFrame(rows, columns=clean_headers)
        numeric_df = df.select_dtypes(include="number")
        summary = {
            "row_count": len(df),
            "columns": list(df.columns),
            "numeric_summary": numeric_df.describe().to_dict() if not numeric_df.empty else {},
        }
        return summary

    # ------------------------------------------------------------------ #
    # History
    # ------------------------------------------------------------------ #
    def get_history(self, limit: int = 50) -> list[dict[str, Any]]:
        return [h.__dict__ for h in self.session.history[-limit:]]

    def _log(self, action: str, detail: str) -> None:
        self.session.history.append(HistoryEntry(timestamp=datetime.now(timezone.utc).isoformat(), action=action, detail=detail))

    # ------------------------------------------------------------------ #
    # Internal helpers
    # ------------------------------------------------------------------ #
    def _wb(self) -> Workbook:
        if self.session.workbook is None:
            raise WorkbookError("No workbook is open. Upload or create one first.")
        return self.session.workbook

    def _ws(self, sheet: str | None) -> Worksheet:
        wb = self._wb()
        name = sheet or self.session.active_sheet
        if name not in wb.sheetnames:
            raise WorkbookError(f"Sheet '{name}' not found. Available: {wb.sheetnames}")
        return wb[name]


# Single shared instance used by the whole app (MVP: one session)
workbook_manager = WorkbookManager()
