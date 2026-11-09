#!/usr/bin/env python3
"""
excel_to_markdown.py

Turn an .xlsx workbook into markdown that keeps Excel's cell coordinates, so an
LLM can read the data AND cite exact cells (e.g. 'Q3 Report'!B7).

How it works (no guessing about table positions):
  1. Read every sheet (visible, hidden and "very hidden") with openpyxl.
  2. Find "regions": groups of filled cells that touch each other. Each region
     is one table, one stray cell, or (if very sparse) a list of cells.
  3. Print each table with column letters across the top and row numbers down
     the left side, exactly like Excel. (A table's own header, e.g. "Due Date",
     is just an ordinary cell, so it appears as the first row under the letters.)
  4. List merged cells under the table, and (only if asked) comments, formulas,
     and dates that Excel stores as plain numbers (General format).

Command line (run from the repo root):
    python src/excel_to_markdown.py book.xlsx -o book.md
    python src/excel_to_markdown.py book.xlsx --comments --formulas --detect-general-dates
    python src/excel_to_markdown.py --help          # every option

As a library:
    from excel_to_markdown import parse_workbook
    result = parse_workbook("book.xlsx")      # a path, raw bytes, or an open file
    print(result.markdown)      # the whole workbook as one text
    print(result.sheets)        # one dict per sheet: name, state, regions, markdown
    print(result.warnings)      # anything that went wrong (never raises for a bad sheet)

Full documentation: docs/excel_to_markdown.md
Limits: .xlsx / .xlsm only (not the old .xls). Charts, images and shapes are ignored.
"""
from __future__ import annotations

import argparse
import datetime as dt
import io
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Optional

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import WINDOWS_EPOCH, from_excel, to_excel

MAX_MERGE_FILL = 10_000  # don't expand giant merged areas cell by cell


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def a1(r: int, c: int) -> str:
    return f"{get_column_letter(c)}{r}"


def span(r1: int, c1: int, r2: int, c2: int) -> str:
    return a1(r1, c1) if (r1, c1) == (r2, c2) else f"{a1(r1, c1)}:{a1(r2, c2)}"


def quote_sheet(name: str) -> str:
    """Sheet names with spaces or symbols need quotes in Excel references."""
    safe = bool(name) and name.replace("_", "").isalnum() and not name[0].isdigit()
    return name if safe else "'" + name.replace("'", "''") + "'"


_DATE_LETTERS = re.compile(r"[yd]", re.I)


def has_date_part(number_format: str) -> bool:
    """True if an Excel number format shows a day or year (so it is a date, not only a time)."""
    stripped = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", number_format or "")
    return bool(_DATE_LETTERS.search(stripped))


def is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def fmt(v, max_chars: int = 1000) -> str:
    """Turn a cell value into clean text."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        if v != v or v in (float("inf"), float("-inf")):
            return str(v)
        if v.is_integer() and abs(v) < 1e15:
            return str(int(v))
        return format(v, ".15g")  # 0.30000000000000004 -> 0.3
    if isinstance(v, dt.datetime):
        return v.date().isoformat() if v.time() == dt.time(0) else v.isoformat(" ")
    if isinstance(v, (dt.date, dt.time)):
        return v.isoformat()
    if isinstance(v, dt.timedelta):            # elapsed time, e.g. format [h]:mm
        total = int(round(v.total_seconds()))
        hours, rest = divmod(abs(total), 3600)
        return f"{'-' if total < 0 else ''}{hours}:{rest // 60:02d}:{rest % 60:02d}"
    s = str(v).replace("\r\n", "\n").replace("\r", "\n").strip()
    if len(s) > max_chars:
        s = s[:max_chars] + f"… [cut: {len(s)} characters in total]"
    return s


def clean_comment(text: str) -> str:
    """Tidy a cell comment. Newer Excel 'threaded comments' are stored with a long
    boilerplate warning in front of the real text; drop it and keep Comment/Reply lines."""
    text = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    if text.startswith("[Threaded comment]"):
        i = text.find("Comment:")
        text = text[i:] if i >= 0 else text
        text = re.sub(r"(Comment|Reply):[ \t]*\n[ \t]*", r"\1: ", text)
    return " / ".join(line.strip() for line in text.split("\n") if line.strip())


def short(v, n: int = 40) -> str:
    """Short preview of a value for use inside a sentence."""
    t = fmt(v, 10_000).replace("\n", " ")
    return t if len(t) <= n else t[: n - 1] + "…"


def esc(s: str) -> str:
    """Make text safe inside a markdown table cell."""
    return s.replace("|", "\\|").replace("\n", "<br>")


# --------------------------------------------------------------------------
# Data containers
# --------------------------------------------------------------------------
@dataclass
class Options:
    include_formulas: bool = False          # list formulas under each table
    include_comments: bool = False          # list cell comments (off unless asked for)
    include_hidden_sheets: bool = True
    include_hidden_rows_cols: bool = True
    detect_general_dates: bool = False      # show date-like numbers in General cells as dates (see detect_general_dates)
    date_keywords: Optional[tuple] = None   # header words that mean "date"; None = DEFAULT_DATE_KEYWORDS
    date_min_year: int = 1950               # numbers that would be dates before this year are left alone
    date_max_year: int = 2100               # ... or after this year
    gap: int = 0                            # blank rows/cols allowed inside one table
    max_rows_per_chunk: int = 200           # long tables are split into parts of this many rows
    max_cell_chars: int = 1000
    row_limit: Optional[int] = None         # only read the first N rows of each sheet (None = all)
    column_limit: Optional[int] = None      # only read the first N columns of each sheet (None = all)


@dataclass
class Grid:
    cells: dict = field(default_factory=dict)      # (row, col) -> value
    occupied: set = field(default_factory=set)     # filled cells + merged areas
    merged: list = field(default_factory=list)     # (r1, c1, r2, c2)
    comments: dict = field(default_factory=dict)   # (row, col) -> (author, text)
    formulas: dict = field(default_factory=dict)   # (row, col) -> formula text
    hidden_rows: set = field(default_factory=set)
    hidden_cols: set = field(default_factory=set)
    general_numbers: set = field(default_factory=set)   # numbers in General-format cells (only filled if detection is on)
    date_notes: list = field(default_factory=list)      # (column, first_row, header, count) for converted columns
    epoch: object = None                                # workbook date system (1900 or 1904)
    cut_rows: bool = False                              # data existed beyond row_limit
    cut_cols: bool = False                              # data existed beyond column_limit


@dataclass
class Region:
    members: list
    r1: int
    c1: int
    r2: int
    c2: int


@dataclass
class ParseResult:
    markdown: str      # the whole workbook as one text
    sheets: list       # one dict per sheet: name, state, regions, and "markdown" (that sheet's own text)
    warnings: list     # anything that went wrong (the program does not crash on bad sheets)
    legend: str = ""   # the "how to read this" note (also at the top of `markdown`)


# --------------------------------------------------------------------------
# Reading a sheet into a grid
# --------------------------------------------------------------------------
def stored_cells(ws):
    """Only the cells that really exist in the file.

    Walking the whole used range (ws.iter_rows) can create billions of empty cells when a
    sheet has formatting far from its data, and runs out of memory. ws._cells holds just
    the cells that were stored, so it stays small.
    """
    cells = getattr(ws, "_cells", None)
    if cells is not None:
        return list(cells.values())
    return [c for row in ws.iter_rows() for c in row]  # fallback if openpyxl changes


def build_grid(ws, ws_formulas, o: Options) -> Grid:
    g = Grid()
    g.epoch = getattr(getattr(ws, "parent", None), "epoch", None)

    def within_limits(r: int, c: int) -> bool:
        return (o.row_limit is None or r <= o.row_limit) and (o.column_limit is None or c <= o.column_limit)

    for cell in stored_cells(ws):
        v = cell.value
        r, c = cell.row, cell.column
        ok = within_limits(r, c)
        if not is_blank(v):
            if not ok:
                g.cut_rows = g.cut_rows or (o.row_limit is not None and r > o.row_limit)
                g.cut_cols = g.cut_cols or (o.column_limit is not None and c > o.column_limit)
            else:
                if v == dt.time(0, 0) and type(v) is dt.time and has_date_part(getattr(cell, "number_format", "")):
                    v = "1900-01-00"   # serial 0 shown with a date format (e.g. formula pointing at a blank cell)
                g.cells[(r, c)] = v
                if (o.detect_general_dates and isinstance(v, (int, float)) and not isinstance(v, bool)
                        and str(getattr(cell, "number_format", "General")).lower() == "general"):
                    g.general_numbers.add((r, c))
        if o.include_comments and ok:
            cm = getattr(cell, "comment", None)
            if cm is not None and getattr(cm, "text", None) and cm.text.strip():
                author = getattr(cm, "author", None) or ""
                if author.startswith("tc=") or author.strip().lower() in ("unknown author", "none"):
                    author = ""                   # internal ids / placeholders carry no information
                g.comments[(r, c)] = (author.strip(), clean_comment(cm.text))

    # Real extent of the data (not of the formatting): used to cap merged areas and column ranges.
    max_r = max((r for r, _ in g.cells), default=1)
    max_c = max((c for _, c in g.cells), default=1)

    g.occupied = set(g.cells)

    for mr in ws.merged_cells.ranges:
        rng = (mr.min_row, mr.min_col, mr.max_row, mr.max_col)
        g.merged.append(rng)
        if (rng[0], rng[1]) in g.cells:  # only merges that hold a value
            r_hi, c_hi = min(rng[2], max_r), min(rng[3], max_c)
            if (r_hi - rng[0] + 1) * (c_hi - rng[1] + 1) <= MAX_MERGE_FILL:
                for r in range(rng[0], r_hi + 1):
                    for c in range(rng[1], c_hi + 1):
                        g.occupied.add((r, c))

    g.hidden_rows = {r for r, d in ws.row_dimensions.items() if d.hidden}
    for key, d in ws.column_dimensions.items():
        if d.hidden:
            lo = d.min or column_index_from_string(key)
            hi = min(d.max or lo, max_c)
            g.hidden_cols.update(range(lo, hi + 1))

    if ws_formulas is not None:
        for cell in stored_cells(ws_formulas):
            v = cell.value
            text = getattr(v, "text", v)  # array formulas keep text in .text
            if isinstance(text, str) and text.startswith("=") and within_limits(cell.row, cell.column):
                g.formulas[(cell.row, cell.column)] = text
    return g


# --------------------------------------------------------------------------
# Opt-in: dates stored as plain numbers (General format)
# --------------------------------------------------------------------------
# A cell formatted as General that holds 46300 is a number as far as the file is concerned, even if
# the business meaning is "5 October 2026". Detection is deliberately cautious. A number is shown
# as a date only if ALL of these hold:
#   1. the cell's format is General (cells with a real date format are already handled),
#   2. the column's header looks like a date header (see DEFAULT_DATE_KEYWORDS) and does not look
#      like an amount, count, id, etc. (NOT_DATE_WORDS). The "column header" is the unbroken stack of
#      text cells directly above the first number in that column (so a two-row header works), plus
#      a merged group header sitting directly over that stack. In the markdown output these are
#      ordinary rows under the A B C letters, so the check looks at the sheet's cells, not at the
#      letters row. Numbers with no text above them are never touched,
#   3. the number falls in a plausible date range (date_min_year .. date_max_year),
#   4. at least half of the column's General numbers are in that range.
# The original number is always kept: `2026-10-05 (serial 46300)`.
DEFAULT_DATE_KEYWORDS = (
    "date", "dates", "datetime", "timestamp", "dt", "dob", "created", "modified", "updated", "due",
    "deadline", "issued", "expiry", "expiration", "expires", "effective", "posted", "opened", "closed", "received",
)
NOT_DATE_WORDS = (
    "amount", "balance", "total", "count", "qty", "quantity", "price", "cost", "rate",
    "days", "age", "duration", "hours", "minutes", "id", "number",
)


def header_words(text: str) -> list:
    """'Due Date (UTC)' -> ['due', 'date', 'utc'];  'DueDate' -> ['due', 'date']"""
    spaced = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", str(text))
    return [w for w in re.split(r"[^A-Za-z0-9]+", spaced.lower()) if w]


def looks_like_date_header(texts: list, keywords) -> bool:
    words = [w for t in texts for w in header_words(t)]
    if any(w in NOT_DATE_WORDS for w in words):
        return False
    return any(w in keywords for w in words)


def detect_general_dates(grid: Grid, reg: "Region", o: Options) -> None:
    """Replace date-like General numbers in this table with `date (serial N)` text. Changes grid.cells."""
    keywords = tuple(k.lower() for k in (o.date_keywords or DEFAULT_DATE_KEYWORDS))
    epoch = grid.epoch or WINDOWS_EPOCH
    lo = to_excel(dt.datetime(o.date_min_year, 1, 1), epoch)
    hi = to_excel(dt.datetime(o.date_max_year, 12, 31), epoch)

    rows_by_col: dict = {}
    for (r, c) in reg.members:
        if (r, c) in grid.cells:
            rows_by_col.setdefault(c, []).append(r)

    for c, rows in rows_by_col.items():
        rows.sort()
        numbers = [r for r in rows if (r, c) in grid.general_numbers]
        if not numbers:
            continue
        first = numbers[0]

        # The column header = the unbroken stack of text cells directly above the first number
        # (so a two-row header works, and so does a text value such as "TBD" between header and
        # numbers), plus a merged group header sitting directly over that stack.
        r = first - 1
        while r >= reg.r1 and (r, c) not in grid.cells:      # step over blank cells (tables joined with gap > 0)
            r -= 1
        stack = []
        while r >= reg.r1 and isinstance(grid.cells.get((r, c)), str):
            stack.append(r)
            r -= 1
        if not stack:
            continue          # no header above the numbers: never guess
        texts = [grid.cells[(x, c)] for x in stack]
        for (r1, c1, r2, c2) in grid.merged:
            anchor = grid.cells.get((r1, c1))
            if c1 <= c <= c2 and r2 == stack[-1] - 1 and r1 >= reg.r1 and isinstance(anchor, str):
                texts.append(anchor)
        if not looks_like_date_header(texts, keywords):
            continue
        matched = next(t for t in texts if any(w in keywords for w in header_words(t)))   # the header text that qualified

        plausible = [r for r in numbers if lo <= grid.cells[(r, c)] <= hi]
        if len(plausible) * 2 < len(numbers):
            continue          # most numbers are not date-sized: this is not a date column

        for r in plausible:
            v = grid.cells[(r, c)]
            try:
                grid.cells[(r, c)] = f"{fmt(from_excel(v, epoch))} (serial {fmt(v)})"
            except (ValueError, OverflowError, TypeError):
                continue
        grid.date_notes.append((c, stack[0], str(matched).strip(), len(plausible)))


# --------------------------------------------------------------------------
# Finding regions (tables) on a sheet
# --------------------------------------------------------------------------
def _make_region(members: list) -> Region:
    rs = [p[0] for p in members]
    cs = [p[1] for p in members]
    return Region(members, min(rs), min(cs), max(rs), max(cs))


def _overlap(a: Region, b: Region) -> bool:
    return not (a.r2 < b.r1 or b.r2 < a.r1 or a.c2 < b.c1 or b.c2 < a.c1)


def find_regions(occupied: set, gap: int = 0) -> list:
    """Group filled cells that touch (including diagonally) into regions."""
    parent = {p: p for p in occupied}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    reach = gap + 1
    for (r, c) in occupied:
        for dr in range(0, reach + 1):
            for dc in range(-reach, reach + 1):
                if dr == 0 and dc <= 0:
                    continue
                n = (r + dr, c + dc)
                if n in occupied:
                    a, b = find((r, c)), find(n)
                    if a != b:
                        parent[a] = b

    groups: dict = {}
    for p in occupied:
        groups.setdefault(find(p), []).append(p)
    regions = [_make_region(m) for m in groups.values()]

    # Regions whose outlines overlap are really one table: merge them.
    changed = True
    while changed:
        changed = False
        regions.sort(key=lambda g: (g.r1, g.c1))
        out: list = []
        for g in regions:
            for o in out:
                if _overlap(o, g):
                    o.members += g.members
                    o.r1, o.c1 = min(o.r1, g.r1), min(o.c1, g.c1)
                    o.r2, o.c2 = max(o.r2, g.r2), max(o.c2, g.c2)
                    changed = True
                    break
            else:
                out.append(g)
        regions = out
    return sorted(regions, key=lambda g: (g.r1, g.c1))


# --------------------------------------------------------------------------
# Rendering
# --------------------------------------------------------------------------
def comment_line(r: int, c: int, author: str, text: str) -> str:
    who = f" (by {author})" if author else ""
    return f"Comment on {a1(r, c)}{who}: {esc(fmt(text, 500))}"


def _footers(reg: Region, grid: Grid, o: Options) -> list:
    lines = []
    inside = lambda r, c: reg.r1 <= r <= reg.r2 and reg.c1 <= c <= reg.c2

    merges = sorted(m for m in grid.merged if inside(m[0], m[1]) and (m[0], m[1]) in grid.cells)
    if merges:
        shown = [f'{span(*m)} ("{esc(short(grid.cells[(m[0], m[1])]))}")' for m in merges[:40]]
        more = f" … and {len(merges) - 40} more" if len(merges) > 40 else ""
        lines.append("Merged cells (value sits in the first cell of the range): " + "; ".join(shown) + more)

    for (r, c), (author, text) in sorted(grid.comments.items()):
        if inside(r, c):
            lines.append(comment_line(r, c, author, text))

    for (c, first_row, header, n) in sorted(grid.date_notes):
        if inside(first_row, c):
            lines.append(f'Date detection (opt-in): column {get_column_letter(c)} ("{esc(short(header, 40))}"): '
                         f'{n} number(s) stored in General format are shown as dates, with the original number in brackets')

    if o.include_formulas:
        items = sorted((p, f) for p, f in grid.formulas.items() if inside(*p))
        for (r, c), f in items[:60]:
            note = "" if (r, c) in grid.cells else "  [no saved value in file]"
            lines.append(f"Formula in {a1(r, c)}: `{f}`{note}")
        if len(items) > 60:
            lines.append(f"… and {len(items) - 60} more formulas")
    return lines


def render_table(sname: str, idx: int, reg: Region, grid: Grid, o: Options) -> list:
    filled = [p for p in reg.members if p in grid.cells]
    rows = sorted({r for r, _ in filled})
    cols = sorted({c for _, c in filled})
    if not o.include_hidden_rows_cols:
        rows = [r for r in rows if r not in grid.hidden_rows]
        cols = [c for c in cols if c not in grid.hidden_cols]
    if not rows or not cols:
        return []

    col_labels = [get_column_letter(c) + (" (hidden)" if c in grid.hidden_cols else "") for c in cols]

    def row_line(r: int, suffix: str = "") -> str:
        label = str(r) + (" (hidden)" if r in grid.hidden_rows else "") + suffix
        vals = [esc(fmt(grid.cells.get((r, c)), o.max_cell_chars)) for c in cols]
        return "| " + label + " | " + " | ".join(vals) + " |"

    n = max(1, o.max_rows_per_chunk)
    chunks = [rows[i:i + n] for i in range(0, len(rows), n)]
    out: list = []
    for k, chunk in enumerate(chunks, 1):
        title = f"### Table {idx} — {quote_sheet(sname)}!{span(reg.r1, reg.c1, reg.r2, reg.c2)}"
        if len(chunks) > 1:
            title += f" (part {k} of {len(chunks)}, rows {chunk[0]}–{chunk[-1]})"
        out += [title, "", "| | " + " | ".join(col_labels) + " |", "|---|" + "---|" * len(cols)]
        if k > 1:
            out.append(row_line(rows[0], " (first row repeated for context)"))
        out += [row_line(r) for r in chunk]
        out.append("")
    foot = _footers(reg, grid, o)
    if foot:
        out += [f"- {x}" for x in foot] + [""]
    return out


def render_cell_list(sname: str, idx: int, reg: Region, grid: Grid, o: Options, kind: str) -> list:
    """One line per cell: used for single cells and very scattered cells."""
    cells = sorted(p for p in reg.members if p in grid.cells)
    if not o.include_hidden_rows_cols:
        cells = [p for p in cells if p[0] not in grid.hidden_rows and p[1] not in grid.hidden_cols]
    if not cells:
        return []
    if len(cells) == 1 and kind == "cell":
        r, c = cells[0]
        m = next((m for m in grid.merged if (m[0], m[1]) == (r, c) and m[:2] != m[2:]), None)
        label = f"merged cell {span(*m)}" if m else "single cell"
        out = [f"**{a1(r, c)}** ({label}): {esc(fmt(grid.cells[(r, c)], o.max_cell_chars))}"]
    else:
        out = [f"### Scattered cells {idx} — {quote_sheet(sname)}!{span(reg.r1, reg.c1, reg.r2, reg.c2)}", ""]
        out += [f"- **{a1(r, c)}**: {esc(fmt(grid.cells[(r, c)], o.max_cell_chars))}" for r, c in cells]
    foot = _footers(reg, grid, o)
    if foot:
        out += [""] + [f"- {x}" for x in foot]
    return out + [""]


def fallback_cell_dump(sname: str, grid: Grid, o: Options) -> list:
    """Last resort if a region fails to render: dump its cells one per line."""
    out = [f"### Cell-by-cell dump — {quote_sheet(sname)}", ""]
    for (r, c), v in sorted(grid.cells.items()):
        out.append(f"- **{a1(r, c)}**: {esc(fmt(v, o.max_cell_chars))}")
    return out + [""]


# --------------------------------------------------------------------------
# Sheet and workbook level
# --------------------------------------------------------------------------
STATE_LABEL = {"visible": "visible", "hidden": "HIDDEN SHEET", "veryHidden": "VERY HIDDEN SHEET"}


def process_sheet(ws, ws_formulas, o: Options, warnings: list):
    name = ws.title
    state = getattr(ws, "sheet_state", "visible")
    meta = {"name": name, "state": state, "regions": []}
    lines = [f"## Sheet {quote_sheet(name)} — {STATE_LABEL.get(state, state)}", ""]

    grid = build_grid(ws, ws_formulas, o)
    if not grid.cells and not grid.comments:
        lines += ["(empty sheet)", ""]
        meta["markdown"] = "\n".join(lines).rstrip() + "\n"
        return lines, meta

    regions = find_regions(grid.occupied, o.gap)
    plan = []
    for reg in regions:
        filled = [p for p in reg.members if p in grid.cells]
        if not filled:
            continue
        area = (reg.r2 - reg.r1 + 1) * (reg.c2 - reg.c1 + 1)
        if len(filled) == 1:
            kind = "cell"
        elif area >= 60 and len(filled) / area < 0.12:
            kind = "scattered"
        else:
            kind = "table"
        plan.append((reg, filled, kind))

    if o.detect_general_dates:
        for reg, _, kind in plan:
            if kind == "table":
                try:
                    detect_general_dates(grid, reg, o)
                except Exception as e:  # detection is a bonus: never let it stop the table
                    warnings.append(f"Sheet '{name}' region {span(reg.r1, reg.c1, reg.r2, reg.c2)}: date detection skipped ({type(e).__name__}: {e})")

    body: list = []
    n_table = n_other = 0
    for reg, filled, kind in plan:
        rng = span(reg.r1, reg.c1, reg.r2, reg.c2)
        try:
            if kind == "table":
                n_table += 1
                block = render_table(name, n_table, reg, grid, o)
            else:
                n_other += 1
                block = render_cell_list(name, n_other, reg, grid, o, kind)
        except Exception as e:  # one bad region must not stop the file
            warnings.append(f"Sheet '{name}' region {rng}: {type(e).__name__}: {e}; used cell dump")
            block = fallback_cell_dump(name, Grid(cells={p: grid.cells[p] for p in filled}), o)
            kind = "fallback"
        if block:
            body += block
            meta["regions"].append({"kind": kind, "range": rng, "filled_cells": len(filled)})

    # Comments sitting on cells outside every block (e.g. on an empty cell far from the data)
    # would otherwise be lost, so list them in their own block.
    def in_any_region(r, c):
        return any(g.r1 <= r <= g.r2 and g.c1 <= c <= g.c2 for g in regions)

    stray = [(p, v) for p, v in sorted(grid.comments.items()) if not in_any_region(*p)]
    if stray:
        body += [f"### Comments on cells outside any table — {quote_sheet(name)}", ""]
        body += [f"- {comment_line(r, c, *v)}" for (r, c), v in stray]
        body += [""]
        meta["regions"].append({"kind": "comments on", "range": ", ".join(a1(*p) for p, _ in stray),
                                "filled_cells": 0})

    if grid.cut_rows or grid.cut_cols:
        parts = []
        if grid.cut_rows:
            parts.append(f"the first {o.row_limit} rows")
        if grid.cut_cols:
            parts.append(f"the first {o.column_limit} columns")
        lines += [f"⚠ Only {' and '.join(parts)} of this sheet were read (size limit); anything beyond is not included.", ""]
        warnings.append(f"Sheet '{name}': data beyond the size limit was not read")

    if o.include_hidden_rows_cols and (grid.hidden_rows or grid.hidden_cols):
        hr = sorted(grid.hidden_rows)
        hc = [get_column_letter(c) for c in sorted(grid.hidden_cols)]
        lines.append(f"Hidden rows: {hr or 'none'}; hidden columns: {hc or 'none'} (shown below and marked)")
        lines.append("")
    if len(meta["regions"]) > 1:
        lines.append("Regions on this sheet: " + ", ".join(f"{r['kind']} {r['range']}" for r in meta["regions"]))
        lines.append("")
    lines = lines + body
    meta["markdown"] = "\n".join(lines).rstrip() + "\n"
    return lines, meta


def make_legend(o: Options) -> str:
    text = (
        "> **How to read this file.** Each sheet is split into blocks. In every table, the "
        "column letters are across the top and the row numbers are down the left side, exactly "
        "as in Excel. To cite a value, give the sheet name and the cell, for example "
        "`'Q3 Report'!B7`. Empty cells are blank. Rows and columns that are completely empty "
        "are left out, so row numbers can skip. Merged cells show their value only in the "
        "first cell of the range (listed under the table)."
    )
    if o.include_comments:
        text += " Cell comments are listed under the table as `Comment on B2 (by author): text`."
    if o.detect_general_dates:
        text += (" Numbers in date-like columns that Excel stores as plain numbers are shown as "
                 "`2026-10-05 (serial 46300)`: the date, then the original number.")
    return text + " Hidden sheets, rows and columns are marked."


LEGEND = make_legend(Options(include_comments=True))   # kept for backwards compatibility


def _load(source, data_only: bool):
    """Open a workbook from a path, raw bytes, or a file-like object (each call starts from the beginning)."""
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(bytes(source))
    elif hasattr(source, "seek"):
        source.seek(0)
    elif isinstance(source, os.PathLike):
        source = os.fspath(source)
    return load_workbook(source, data_only=data_only)


def parse_workbook(source, *, name: Optional[str] = None, legend: bool = True, **kwargs) -> ParseResult:
    """Convert a workbook (path, bytes or file object) to coordinate-aware markdown.

    Options (keyword arguments) are the fields of `Options`, e.g. include_comments=True,
    detect_general_dates=True, include_formulas=True, include_hidden_sheets=False, row_limit=2000.
    """
    o = Options(**kwargs)
    warnings: list = []
    sheets_meta: list = []
    out: list = []

    wb = _load(source, data_only=True)
    wbf = _load(source, data_only=False) if o.include_formulas else None

    if name is None:
        name = os.path.basename(os.fspath(source).replace("\\", "/")) if isinstance(source, (str, os.PathLike)) else "workbook"
    legend_text = make_legend(o)

    out.append(f"# Workbook: {name}")
    out.append("")
    if legend:
        out += [legend_text, ""]
    states = ", ".join(f"{quote_sheet(ws.title)} ({ws.sheet_state})" for ws in wb.worksheets)
    out += [f"Sheets in order: {states}", ""]
    if getattr(wb, "chartsheets", None):
        out += [f"Chart-only sheets (no cell data, skipped): {', '.join(c.title for c in wb.chartsheets)}", ""]

    for ws in wb.worksheets:
        if ws.sheet_state != "visible" and not o.include_hidden_sheets:
            sheets_meta.append({"name": ws.title, "state": ws.sheet_state, "regions": [], "skipped": True})
            continue
        try:
            lines, meta = process_sheet(ws, wbf[ws.title] if wbf else None, o, warnings)
        except Exception as e:  # one bad sheet must not stop the file
            warnings.append(f"Sheet '{ws.title}' failed: {type(e).__name__}: {e}")
            lines = [f"## Sheet {quote_sheet(ws.title)}", "", f"⚠ This sheet could not be read ({type(e).__name__}).", ""]
            meta = {"name": ws.title, "state": ws.sheet_state, "regions": [], "error": str(e),
                    "markdown": "\n".join(lines).rstrip() + "\n"}
        out += lines
        sheets_meta.append(meta)

    return ParseResult("\n".join(out).rstrip() + "\n", sheets_meta, warnings, legend_text)


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="Convert an Excel workbook to coordinate-aware markdown.")
    ap.add_argument("input")
    ap.add_argument("-o", "--output", help="write markdown here (default: print to screen)")
    ap.add_argument("--formulas", action="store_true", help="also list formulas under each table")
    ap.add_argument("--comments", action="store_true", help="also list cell comments (off by default)")
    ap.add_argument("--detect-general-dates", action="store_true",
                    help="show date-like numbers in General-format cells as dates, under date-like headers (off by default)")
    ap.add_argument("--date-keywords", help="comma-separated header words that mean 'date' (replaces the built-in list)")
    ap.add_argument("--date-min-year", type=int, default=1950)
    ap.add_argument("--date-max-year", type=int, default=2100)
    ap.add_argument("--skip-hidden-sheets", action="store_true")
    ap.add_argument("--skip-hidden-rows-cols", action="store_true")
    ap.add_argument("--gap", type=int, default=0, help="blank rows/cols allowed inside one table (default 0)")
    ap.add_argument("--chunk-rows", type=int, default=200, help="split tables longer than this into parts (default 200)")
    ap.add_argument("--row-limit", type=int, help="only read the first N rows of each sheet")
    ap.add_argument("--column-limit", type=int, help="only read the first N columns of each sheet")
    ap.add_argument("--max-cell-chars", type=int, default=1000)
    return ap


def main(argv=None) -> int:
    a = build_parser().parse_args(argv)

    res = parse_workbook(
        a.input,
        include_formulas=a.formulas,
        include_comments=a.comments,
        detect_general_dates=a.detect_general_dates,
        date_keywords=tuple(k.strip() for k in a.date_keywords.split(",") if k.strip()) if a.date_keywords else None,
        date_min_year=a.date_min_year,
        date_max_year=a.date_max_year,
        include_hidden_sheets=not a.skip_hidden_sheets,
        include_hidden_rows_cols=not a.skip_hidden_rows_cols,
        gap=a.gap,
        max_rows_per_chunk=a.chunk_rows,
        row_limit=a.row_limit,
        column_limit=a.column_limit,
        max_cell_chars=a.max_cell_chars,
    )
    if a.output:
        with open(a.output, "w", encoding="utf-8") as f:
            f.write(res.markdown)
    else:
        sys.stdout.write(res.markdown)
    for w in res.warnings:
        print("WARNING:", w, file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
