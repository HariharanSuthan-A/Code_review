"""
Safe Excel formula evaluator (display layer only).

Purpose: openpyxl stores ``=...`` strings but never computes them, so
previews/read_range used to show the raw formula text
(e.g. ``=CONCATENATE("Excellent: ",COUNTIF(N2:N51,"Excellent"),...)``).
Excel itself shows the *answer*. This module evaluates common formulas
in Python so reads/previews can show the same answer the user sees in
Excel, while the saved .xlsx keeps the live formula untouched.

Security: whitelisted expression parser — no eval/exec, no attribute
access, no I/O. Only the functions in _FUNCTIONS plus standard
operators are understood. Anything else raises FormulaEvalError and
callers fall back to showing the raw formula string.
"""
from __future__ import annotations

import re
from datetime import date, datetime


class FormulaEvalError(Exception):
    pass


# ---------------------------------------------------------------- value helpers

def _is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _to_num(v):
    if isinstance(v, bool):
        return 1 if v else 0
    if _is_num(v):
        return v
    if v is None or v == "":
        return 0
    try:
        s = str(v).strip()
        return int(s) if re.fullmatch(r"[+-]?\d+", s) else float(s)
    except (ValueError, TypeError):
        raise FormulaEvalError(f"Expected a number, got {v!r}")


def _to_bool(v) -> bool:
    if isinstance(v, bool):
        return v
    if v is None or v == "":
        return False
    if _is_num(v):
        return v != 0
    s = str(v).strip().lower()
    if s in ("true", "yes"):
        return True
    if s in ("false", "no"):
        return False
    return bool(s)


def _display_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _flatten(args) -> list:
    out = []
    for a in args:
        if isinstance(a, list):
            for row in a:
                if isinstance(row, list):
                    out.extend(row)
                else:
                    out.append(row)
        else:
            out.append(a)
    return out


def _match_criteria(value, criteria) -> bool:
    """Excel COUNTIF/SUMIF criteria matching: numbers, plain strings
    (case-insensitive), comparison prefixes (>,<,>=,<=,<>,=) and
    wildcards * and ?."""
    if criteria is None:
        return value is None or value == ""
    if _is_num(criteria) and _is_num(value):
        return value == criteria
    s = str(criteria)
    m = re.match(r"^(>=|<=|<>|>|<|=)(.*)$", s, re.DOTALL)
    if m:
        op, operand_s = m.group(1), m.group(2)
        try:
            operand = float(operand_s) if operand_s != "" else 0
            v = float(value) if value not in (None, "") else None
            if v is None:
                return False
            return {
                ">": v > operand, "<": v < operand, ">=": v >= operand,
                "<=": v <= operand, "=": v == operand, "<>": v != operand,
            }[op]
        except (ValueError, TypeError):
            vo, oo = str(value).lower(), operand_s.lower()
            if op == "=":
                return vo == oo
            if op == "<>":
                return vo != oo
            return {">": vo > oo, "<": vo < oo, ">=": vo >= oo, "<=": vo <= oo}[op]
    if "*" in s or "?" in s:
        pat = "".join(
            ".*" if ch == "*" else "." if ch == "?" else re.escape(ch) for ch in s
        )
        return re.fullmatch(pat, "" if value is None else str(value), re.IGNORECASE) is not None
    if value is None:
        return s == ""
    if _is_num(value):
        try:
            return float(value) == float(s)
        except ValueError:
            return str(value).lower() == s.lower()
    return str(value).lower() == s.lower()


# ---------------------------------------------------------------- tokenizer / parser

_TOKEN = re.compile(r"""
    (?P<dq>"(?:[^"]|"")*")
  | (?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?|\.\d+)
  | (?P<cell>\$?[A-Za-z]{1,3}\$?\d+)
  | (?P<word>[A-Za-z_][A-Za-z0-9_.]*)
  | (?P<op><>|<=|>=|[+\-*/^&%(),:!<>=])
  | (?P<ws>\s+)
  | (?P<sq>'(?:[^']|'')*')
""", re.VERBOSE)


class _Ctx:
    """Resolution callbacks — lets the parser stay workbook-agnostic."""

    def __init__(self, cell_value, range_values):
        self.cell_value = cell_value      # (sheet|None, row, col) -> evaluated scalar
        self.range_values = range_values  # (sheet|None, min_c, min_r, max_c, max_r) -> list[list]


class _Parser:
    def __init__(self, text, ctx: _Ctx):
        self.ctx = ctx
        self.toks: list[tuple[str, str]] = []
        pos = 0
        while pos < len(text):
            m = _TOKEN.match(text, pos)
            if not m:
                raise FormulaEvalError(f"Cannot parse near {text[pos:pos + 12]!r}")
            pos = m.end()
            kind = m.lastgroup
            if kind == "ws":
                continue
            self.toks.append((kind, m.group(0)))
        self.i = 0

    # -- token stream --
    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def next(self):
        t = self.peek()
        if t[0] is None:
            raise FormulaEvalError("Unexpected end of formula")
        self.i += 1
        return t

    # -- grammar --
    def parse(self):
        v = self.parse_concat()
        if self.i != len(self.toks):
            raise FormulaEvalError(f"Unexpected {self.toks[self.i][1]!r}")
        return v

    def parse_concat(self):
        v = self.parse_comparison()
        while self.peek()[1] == "&":
            self.next()
            rhs = self.parse_comparison()
            v = _display_str(v) + _display_str(rhs)
        return v

    def parse_comparison(self):
        v = self.parse_add()
        while True:
            _, op = self.peek()
            if op in ("=", "<>", "<", ">", "<=", ">="):
                self.next()
                rhs = self.parse_add()
                v = _compare(v, op, rhs)
            else:
                return v

    def parse_add(self):
        v = self.parse_mul()
        while self.peek()[1] in ("+", "-"):
            _, op = self.next()
            rhs = self.parse_mul()
            v = _to_num(v) + _to_num(rhs) if op == "+" else _to_num(v) - _to_num(rhs)
        return v

    def parse_mul(self):
        v = self.parse_power()
        while self.peek()[1] in ("*", "/"):
            _, op = self.next()
            rhs = self.parse_power()
            if op == "*":
                v = _to_num(v) * _to_num(rhs)
            else:
                d = _to_num(rhs)
                if d == 0:
                    raise FormulaEvalError("Division by zero")
                v = _to_num(v) / d
        return v

    def parse_power(self):
        v = self.parse_unary()
        if self.peek()[1] == "^":
            self.next()
            rhs = self.parse_unary()
            v = _to_num(v) ** _to_num(rhs)
        return v

    def parse_unary(self):
        _, op = self.peek()
        if op == "-":
            self.next()
            return -_to_num(self.parse_unary())
        if op == "+":
            self.next()
            return _to_num(self.parse_unary())
        return self.parse_postfix()

    def parse_postfix(self):
        v = self.parse_primary()
        while self.peek()[1] == "%":
            self.next()
            v = _to_num(v) / 100
        return v

    # -- primary: literals, refs, calls, parens --
    def parse_primary(self):
        kind, val = self.next()
        if kind == "dq":
            return val[1:-1].replace('""', '"')
        if kind == "num":
            return int(val) if re.fullmatch(r"[+-]?\d+", val) else float(val)
        if kind == "op" and val == "(":
            v = self.parse_concat()
            k2, v2 = self.next()
            if v2 != ")":
                raise FormulaEvalError(f"Expected ')', got {v2!r}")
            return v
        if kind == "sq":
            # quoted sheet name: 'My Sheet'!A1...  (must be followed by !)
            sheet = val[1:-1].replace("''", "'")
            k2, v2 = self.peek()
            if v2 != "!":
                raise FormulaEvalError(f"Unexpected {val!r}")
            self.next()  # consume !
            return self.parse_ref_with_sheet(sheet)
        if kind in ("cell", "word"):
            return self.parse_ref_or_call(kind, val)
        raise FormulaEvalError(f"Unexpected {val!r}")

    def parse_ref_or_call(self, kind, val):
        # function call?
        if self.peek()[1] == "(":
            if kind != "word":
                raise FormulaEvalError(f"Unexpected {val!r}")
            return self.parse_call(val)
        # sheet-qualified?  word ! ref
        if self.peek()[1] == "!" and kind == "word":
            self.next()
            return self.parse_ref_with_sheet(val)
        text = val
        # range?  A1:B2  (also A:A / 1:1 whole col/row — unsupported, raise)
        if self.peek()[1] == ":":
            self.next()
            k2, v2 = self.next()
            if k2 not in ("cell", "word"):
                raise FormulaEvalError(f"Bad range end {v2!r}")
            text = text + ":" + v2
            return self.ctx.range_values(None, *self._bounds(text))
        if kind == "cell":
            return self.ctx.cell_value(None, *self._cell(text))
        up = val.upper()
        if up == "TRUE":
            return True
        if up == "FALSE":
            return False
        # bare column (A) or unknown name
        raise FormulaEvalError(f"Unknown name {val!r}")

    def parse_ref_with_sheet(self, sheet):
        kind, val = self.next()
        if kind not in ("cell", "word"):
            raise FormulaEvalError(f"Bad reference after {sheet}! ")
        text = val
        if self.peek()[1] == ":":
            self.next()
            k2, v2 = self.next()
            if k2 not in ("cell", "word"):
                raise FormulaEvalError(f"Bad range end {v2!r}")
            text = text + ":" + v2
            return self.ctx.range_values(sheet, *self._bounds(text))
        return self.ctx.cell_value(sheet, *self._cell(text))

    def parse_call(self, name):
        self.next()  # consume (
        args = []
        if self.peek()[1] == ")":
            self.next()
            return _call(name, args, self.ctx)
        while True:
            # a range arg must stay a 2D list, not collapse to scalar:
            args.append(self.parse_range_or_value())
            _, op = self.peek()
            if op == ",":
                self.next()
                continue
            if op == ")":
                self.next()
                return _call(name, args, self.ctx)
            raise FormulaEvalError(f"Expected ',' or ')' in {name}()")

    def parse_range_or_value(self):
        # lookahead: [sheet!]cell ":" cell  -> keep as range grid
        save = self.i
        try:
            kind, val = self.peek()
            sheet = None
            if kind == "sq":
                self.next()
                if self.peek()[1] != "!":
                    raise FormulaEvalError("bad ref")
                self.next()
                sheet = val[1:-1].replace("''", "'")
                kind, val = self.peek()
            elif kind == "word":
                # could be Sheet!A1 — check two-ahead
                if self.i + 1 < len(self.toks) and self.toks[self.i + 1][1] == "!":
                    self.next()
                    self.next()
                    sheet = val
                    kind, val = self.peek()
            if kind in ("cell", "word") and self.i + 1 < len(self.toks) \
                    and self.toks[self.i + 1][1] == ":":
                start = val
                self.next()
                self.next()  # consume :
                k2, v2 = self.peek()
                if k2 in ("cell", "word"):
                    self.next()
                    return self.ctx.range_values(sheet, *self._bounds(start + ":" + v2))
        except FormulaEvalError:
            pass
        self.i = save
        return self.parse_concat()

    @staticmethod
    def _cell(text):
        m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", text)
        if not m:
            raise FormulaEvalError(f"Bad cell {text!r}")
        from openpyxl.utils import column_index_from_string
        return int(m.group(2)), column_index_from_string(m.group(1).upper())

    @staticmethod
    def _bounds(text):
        from openpyxl.utils import column_index_from_string
        parts = text.split(":")
        if len(parts) != 2:
            raise FormulaEvalError(f"Bad range {text!r}")
        # Whole-column refs like A:C / $A:$C — rows resolved by the
        # caller (None = 1..ws.max_row).
        cols = []
        for p in parts:
            m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?", p.strip())
            if m:
                cols.append(column_index_from_string(m.group(1).upper()))
        if len(cols) == 2:
            return min(cols), 1, max(cols), None
        bounds = []
        for p in parts:
            m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", p.strip())
            if not m:
                raise FormulaEvalError(f"Unsupported range {text!r}")
            bounds.append((column_index_from_string(m.group(1).upper()), int(m.group(2))))
        (c1, r1), (c2, r2) = bounds
        return min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)


def _compare(l, op, r):
    if _is_num(l) and _is_num(r) and not isinstance(l, bool) and not isinstance(r, bool):
        ok = {"=": l == r, "<>": l != r, "<": l < r, ">": l > r, "<=": l <= r, ">=": l >= r}[op]
        return ok
    ls, rs = _display_str(l).lower(), _display_str(r).lower()
    try:
        ln, rn = float(ls), float(rs)
        ok = {"=": ln == rn, "<>": ln != rn, "<": ln < rn, ">": ln > rn, "<=": ln <= rn, ">=": ln >= rn}[op]
        return ok
    except ValueError:
        pass
    return {"=": ls == rs, "<>": ls != rs, "<": ls < rs, ">": ls > rs, "<=": ls <= rs, ">=": ls >= rs}[op]


# ---------------------------------------------------------------- functions

def _fn_sum(*args):
    total = 0
    for v in _flatten(args):
        if _is_num(v):
            total += v
    return total


def _fn_average(*args):
    nums = [v for v in _flatten(args) if _is_num(v)]
    if not nums:
        raise FormulaEvalError("AVERAGE of no numbers")
    return sum(nums) / len(nums)


def _fn_count(*args):
    return sum(1 for v in _flatten(args) if _is_num(v))


def _fn_counta(*args):
    return sum(1 for v in _flatten(args) if v not in (None, ""))


def _fn_countblank(arg):
    if not isinstance(arg, list):
        arg = [[arg]]
    grid = arg if isinstance(arg[0], list) else [arg]
    return sum(1 for row in grid for v in row if v in (None, ""))


def _fn_countif(rng, criteria):
    grid = rng if isinstance(rng, list) else [[rng]]
    rows = grid if grid and isinstance(grid[0], list) else [grid]
    return sum(1 for row in rows for v in row if _match_criteria(_resolve(v), criteria))


def _fn_countifs(*args):
    if len(args) % 2 != 0 or not args:
        raise FormulaEvalError("COUNTIFS needs range/criteria pairs")
    pairs = [(args[i], args[i + 1]) for i in range(0, len(args), 2)]
    grids = []
    for rng, _ in pairs:
        g = rng if isinstance(rng, list) else [[rng]]
        grids.append(g if g and isinstance(g[0], list) else [g])
    n = max(len(g) for g in grids)
    total = 0
    for r in range(n):
        for c in range(max(len(g[r]) if r < len(g) else 0 for g in grids)):
            ok = True
            for g, (_, crit) in zip(grids, pairs):
                v = g[r][c] if r < len(g) and c < len(g[r]) else None
                if not _match_criteria(_resolve(v), crit):
                    ok = False
                    break
            if ok:
                total += 1
    return total


def _resolve(v):
    return v  # nested formulas are already evaluated by ctx callbacks


def _sumif(rng, criteria, sum_range=None):
    vals = rng if isinstance(rng, list) else [[rng]]
    rows = vals if vals and isinstance(vals[0], list) else [vals]
    if sum_range is None:
        srows = rows
    else:
        s = sum_range if isinstance(sum_range, list) else [[sum_range]]
        srows = s if s and isinstance(s[0], list) else [s]
    total = 0
    for r in range(len(rows)):
        for c in range(len(rows[r])):
            if _match_criteria(rows[r][c], criteria):
                sv = srows[r][c] if r < len(srows) and c < len(srows[r]) else 0
                if _is_num(sv):
                    total += sv
    return total


def _fn_sumifs(sum_range, *args):
    if len(args) % 2 != 0:
        raise FormulaEvalError("SUMIFS needs range/criteria pairs")
    s = sum_range if isinstance(sum_range, list) else [[sum_range]]
    srows = s if s and isinstance(s[0], list) else [s]
    pairs = [(args[i], args[i + 1]) for i in range(0, len(args), 2)]
    grids = []
    for rng, _ in pairs:
        g = rng if isinstance(rng, list) else [[rng]]
        grids.append(g if g and isinstance(g[0], list) else [g])
    total = 0
    n = max([len(srows)] + [len(g) for g in grids])
    for r in range(n):
        width = max([len(srows[r]) if r < len(srows) else 0] +
                    [len(g[r]) if r < len(g) else 0 for g in grids])
        for c in range(width):
            if all(_match_criteria(g[r][c] if r < len(g) and c < len(g[r]) else None, crit)
                   for g, (_, crit) in zip(grids, pairs)):
                sv = srows[r][c] if r < len(srows) and c < len(srows[r]) else 0
                if _is_num(sv):
                    total += sv
    return total


def _fn_averageif(rng, criteria, avg_range=None):
    vals = rng if isinstance(rng, list) else [[rng]]
    rows = vals if vals and isinstance(vals[0], list) else [vals]
    a = avg_range if avg_range is not None else rng
    ag = a if isinstance(a, list) else [[a]]
    arows = ag if ag and isinstance(ag[0], list) else [ag]
    nums = []
    for r in range(len(rows)):
        for c in range(len(rows[r])):
            if _match_criteria(rows[r][c], criteria):
                sv = arows[r][c] if r < len(arows) and c < len(arows[r]) else None
                if _is_num(sv):
                    nums.append(sv)
    if not nums:
        raise FormulaEvalError("AVERAGEIF of no numbers")
    return sum(nums) / len(nums)


def _fn_if(cond, t, f=None):
    return t if _to_bool(cond) else (f if f is not None else False)


def _fn_vlookup(key, rng, col, approx=None):
    if not isinstance(rng, list) or not rng or not isinstance(rng[0], list):
        raise FormulaEvalError("VLOOKUP needs a table range")
    ci = int(_to_num(col)) - 1
    if ci < 0 or ci >= len(rng[0]):
        raise FormulaEvalError("VLOOKUP column out of range")
    approx = _to_bool(approx) if approx is not None else False
    if approx:
        best = None
        for row in rng:
            try:
                if float(row[0]) <= float(key):
                    best = row
            except (ValueError, TypeError):
                if str(row[0]).lower() <= str(key).lower():
                    best = row
        if best is None:
            raise FormulaEvalError("VLOOKUP: no match")
        return best[ci]
    for row in rng:
        if _match_criteria(row[0], key):
            return row[ci]
    raise FormulaEvalError("VLOOKUP: no exact match")


def _fn_index(rng, r, c=None):
    g = rng if isinstance(rng, list) else [[rng]]
    rows = g if g and isinstance(g[0], list) else [g]
    ri = int(_to_num(r)) - 1
    ci = int(_to_num(c)) - 1 if c is not None else 0
    try:
        return rows[ri][ci]
    except IndexError:
        raise FormulaEvalError("INDEX out of range") from None


def _fn_match(key, rng, mtype=1):
    # single row or column vector
    g = rng if isinstance(rng, list) else [[rng]]
    rows = g if g and isinstance(g[0], list) else [g]
    vec = rows[0] if len(rows) == 1 else [row[0] for row in rows]
    mtype = int(_to_num(mtype))
    if mtype == 0:
        for i, v in enumerate(vec):
            if _match_criteria(v, key):
                return i + 1
        raise FormulaEvalError("MATCH: no exact match")
    raise FormulaEvalError("MATCH only supports exact (0) mode")


def _fn_textjoin(delim, ignore_empty, *args):
    parts = [_display_str(v) for v in _flatten(args)]
    if _to_bool(ignore_empty):
        parts = [p for p in parts if p != ""]
    return _display_str(delim).join(parts)


def _fn_xlookup(key, lookup, ret, not_found=None):
    lg = lookup if isinstance(lookup, list) else [[lookup]]
    lrows = lg if lg and isinstance(lg[0], list) else [lg]
    rg = ret if isinstance(ret, list) else [[ret]]
    rrows = rg if rg and isinstance(rg[0], list) else [rg]
    # single row or single column vectors (exact match only)
    if len(lrows) == 1:
        vec = lrows[0]
        rvec = rrows[0] if len(rrows) == 1 else [row[0] for row in rrows]
        for i, v in enumerate(vec):
            if _match_criteria(v, key):
                try:
                    return rvec[i]
                except IndexError:
                    raise FormulaEvalError("XLOOKUP: return range too short") from None
    elif len(lrows[0]) == 1:
        vec = [row[0] for row in lrows]
        rvec = [row[0] for row in rrows] if len(rrows[0]) == 1 else rrows[0]
        for i, v in enumerate(vec):
            if _match_criteria(v, key):
                try:
                    return rvec[i]
                except IndexError:
                    raise FormulaEvalError("XLOOKUP: return range too short") from None
    else:
        raise FormulaEvalError("XLOOKUP needs a single row/column lookup array")
    if not_found is not None:
        return not_found
    raise FormulaEvalError("XLOOKUP: no match")


_FUNCTIONS = {
    "SUM": lambda *a: _fn_sum(*a),
    "AVERAGE": lambda *a: _fn_average(*a),
    "AVG": lambda *a: _fn_average(*a),
    "COUNT": lambda *a: _fn_count(*a),
    "COUNTA": lambda *a: _fn_counta(*a),
    "COUNTBLANK": lambda a: _fn_countblank(a),
    "COUNTIF": lambda r, c: _fn_countif(r, c),
    "COUNTIFS": lambda *a: _fn_countifs(*a),
    "SUMIF": lambda r, c, s=None: _sumif(r, c, s),
    "SUMIFS": lambda s, *a: _fn_sumifs(s, *a),
    "AVERAGEIF": lambda r, c, a=None: _fn_averageif(r, c, a),
    "AVERAGEIFS": lambda a, *rest: _fn_averageif(rest[0], rest[1], a) if len(rest) == 2 else (_ for _ in ()).throw(FormulaEvalError("AVERAGEIFS needs one range/criteria pair")),
    "MIN": lambda *a: min([v for v in _flatten(a) if _is_num(v)], default=0),
    "MAX": lambda *a: max([v for v in _flatten(a) if _is_num(v)], default=0),
    "PRODUCT": lambda *a: _prod(_flatten(a)),
    "ABS": lambda x: abs(_to_num(x)),
    "ROUND": lambda x, n=0: round(float(_to_num(x)), int(_to_num(n))),
    "INT": lambda x: int(float(_to_num(x)) // 1),
    "MOD": lambda x, y: _to_num(x) % _to_num(y),
    "POWER": lambda x, y: _to_num(x) ** _to_num(y),
    "SQRT": lambda x: _to_num(x) ** 0.5,
    "IF": lambda c, t, f=None: _fn_if(c, t, f),
    "IFS": lambda *a: _fn_ifs(a),
    "AND": lambda *a: all(_to_bool(v) for v in _flatten(a)),
    "OR": lambda *a: any(_to_bool(v) for v in _flatten(a)),
    "NOT": lambda x: not _to_bool(x),
    "CONCATENATE": lambda *a: "".join(_display_str(v) for v in _flatten(a)),
    "CONCAT": lambda *a: "".join(_display_str(v) for v in _flatten(a)),
    "TEXTJOIN": lambda d, i, *a: _fn_textjoin(d, i, *a),
    "UPPER": lambda x: _display_str(x).upper(),
    "LOWER": lambda x: _display_str(x).lower(),
    "TRIM": lambda x: _display_str(x).strip(),
    "LEN": lambda x: len(_display_str(x)),
    "LEFT": lambda x, n=1: _display_str(x)[:int(_to_num(n))],
    "RIGHT": lambda x, n=1: _display_str(x)[-int(_to_num(n)):] if int(_to_num(n)) else "",
    "MID": lambda x, s, n: _display_str(x)[int(_to_num(s)) - 1:int(_to_num(s)) - 1 + int(_to_num(n))],
    "VALUE": lambda x: _to_num(x),
    "TEXT": lambda x, fmt=None: _display_str(x),
    "VLOOKUP": lambda k, r, c, a=None: _fn_vlookup(k, r, c, a),
    "XLOOKUP": lambda k, l, r, n=None: _fn_xlookup(k, l, r, n),
    "INDEX": lambda r, i, c=None: _fn_index(r, i, c),
    "MATCH": lambda k, r, t=0: _fn_match(k, r, t),
    "TODAY": lambda: date.today(),
    "NOW": lambda: datetime.now(),
    "IFERROR": lambda v, fb: fb if isinstance(v, FormulaEvalError) else v,
}


def _prod(vals):
    p = 1
    for v in vals:
        if _is_num(v):
            p *= v
    return p


def _fn_ifs(args):
    if len(args) % 2 != 0:
        raise FormulaEvalError("IFS needs condition/value pairs")
    for i in range(0, len(args), 2):
        if _to_bool(args[i]):
            return args[i + 1]
    raise FormulaEvalError("IFS: no condition true")


def _call(name, args, ctx):
    fn = _FUNCTIONS.get(str(name).upper())
    if fn is None:
        raise FormulaEvalError(f"Unsupported function {name}")
    # IFERROR needs lazy eval — approximate: caller already evaluated, so
    # plain passthrough is fine (errors raise before reaching here).
    try:
        return fn(*args)
    except FormulaEvalError:
        raise
    except Exception as exc:
        raise FormulaEvalError(f"{name}: {exc}") from exc


# ---------------------------------------------------------------- public API

def evaluate(formula: str, cell_value_cb, range_values_cb):
    """Evaluate an Excel formula string (with or without leading '=').

    cell_value_cb(sheet|None, row, col) -> scalar (already evaluated).
    range_values_cb(sheet|None, min_col, min_row, max_col, max_row) -> grid.
    """
    text = formula.strip()
    if text.startswith("="):
        text = text[1:]
    if not text:
        raise FormulaEvalError("Empty formula")
    parser = _Parser(text, _Ctx(cell_value_cb, range_values_cb))
    return parser.parse()


def evaluate_on_workbook(formula: str, workbook, current_sheet: str, _seen: frozenset | None = None):
    """Evaluate `formula` against an openpyxl Workbook, recursively
    resolving referenced cells (including cells that themselves contain
    formulas). `_seen` guards against circular references."""
    from openpyxl.utils import column_index_from_string

    _seen = _seen or frozenset()

    def ws_for(name):
        sname = name or current_sheet
        if sname not in workbook.sheetnames:
            raise FormulaEvalError(f"Unknown sheet {sname!r}")
        return workbook[sname], sname

    def cell_cb(sheet, row, col):
        ws, sname = ws_for(sheet)
        key = (sname, row, col)
        if key in _seen:
            raise FormulaEvalError("Circular reference")
        cell = ws.cell(row=row, column=col)
        v = cell.value
        if isinstance(v, str) and v.startswith("="):
            return evaluate_on_workbook(v, workbook, sname, _seen | {key})
        return v

    def range_cb(sheet, min_c, min_r, max_c, max_r):
        ws, _ = ws_for(sheet)
        if max_r is None:  # whole-column ref like A:C
            max_r = max(ws.max_row, min_r)
        grid = []
        for r in range(min_r, max_r + 1):
            row = []
            for c in range(min_c, max_c + 1):
                row.append(cell_cb(ws.title, r, c))
            grid.append(row)
        return grid

    return evaluate(formula, cell_cb, range_cb)


def try_evaluate_on_workbook(formula: str, workbook, current_sheet: str):
    """Return the computed value, or None if the formula uses something
    the evaluator does not support."""
    try:
        return evaluate_on_workbook(formula, workbook, current_sheet)
    except FormulaEvalError:
        return None
    except Exception:
        return None
