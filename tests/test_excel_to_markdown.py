#!/usr/bin/env python3
"""
Tests for src/excel_to_markdown.py (the converter on its own: layouts, merged cells, comments,
dates, hidden content, robustness). The tests for the extractor/app wiring are in
test_excel_openpyxl_method.py.

Run:   pytest tests/test_excel_to_markdown.py
   or: python tests/test_excel_to_markdown.py        (plain Python, no extra tools)

Most tests read samples/excel_edge_cases.xlsx (14 tabs, see docs/excel_to_markdown.md). A few build tiny
workbooks on the fly (comment authors, far-away formatting) because they need
things the saved test book cannot hold.
"""
import os
import re
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "src"))
import excel_to_markdown as m  # noqa: E402

BOOK = os.path.join(os.path.dirname(HERE), "samples", "excel_edge_cases.xlsx")


# ---------------------------------------------------------------- helpers
def block(md: str, sheet: str) -> str:
    """The part of the output that belongs to one sheet."""
    head = f"## Sheet {m.quote_sheet(sheet)} — "
    assert head in md, f"sheet {sheet!r} not in output"
    return md.split(head, 1)[1].split("\n## Sheet ", 1)[0]


def table_row(text: str, row_label: str) -> list:
    """Cells of one table row, e.g. table_row(t, '3') -> ['Region', 'Country', ...]."""
    line = next(l for l in text.splitlines() if l.startswith(f"| {row_label} |"))
    return [c.strip() for c in line.strip().strip("|").split("|")][1:]


def regions(res, sheet: str) -> list:
    return [r["range"] for s in res.sheets if s["name"] == sheet for r in s["regions"]]


def tmp_book():
    from openpyxl import Workbook
    return Workbook(), tempfile.mktemp(suffix=".xlsx")


RES = m.parse_workbook(BOOK)                                              # defaults
RES_F = m.parse_workbook(BOOK, include_formulas=True)
RES_C = m.parse_workbook(BOOK, include_comments=True)                      # comments switched on
RES_D = m.parse_workbook(BOOK, detect_general_dates=True)                  # General-date detection switched on
MD = RES.markdown
MD_C = RES_C.markdown


# ------------------------------------------------- 1. whole workbook
def test_every_sheet_is_read_with_its_visibility():
    states = {s["name"]: s["state"] for s in RES.sheets}
    assert len(states) == 14
    assert states["Summary"] == "visible"
    assert states["Hidden Rates"] == "hidden"
    assert states["Config Secret"] == "veryHidden"
    assert not RES.warnings, RES.warnings


# ------------------------------------------------- 2. layouts
def test_simple_table_and_number_cleanup():
    t = block(MD, "Summary")
    assert "### Table 1 — Summary!A1:C5" in t
    assert table_row(t, "1") == ["Product", "Units", "Price"]
    assert table_row(t, "4") == ["Gizmo", "30", "0.3"]          # not 0.30000000000000004


def test_title_notes_and_table_stay_separate():
    assert regions(RES, "Q3 Report") == ["A1:E3", "A5:E11", "A13:E14"]
    t = block(MD, "Q3 Report")
    assert "Prepared by: Finance Team" in t
    assert "unaudited \\| source" in t and "<br>Second line" in t   # pipe escaped, line break kept


def test_several_tables_on_one_sheet_and_a_table_without_header():
    assert regions(RES, "Multi Tables") == ["B2:D6", "H3:I5", "B11:D14", "K20", "B25:C26"]
    t = block(MD, "Multi Tables")
    assert table_row(t.split("### Table 3")[1], "11") == ["1", "2", "3"]   # numbers only, no header
    assert "**K20** (single cell): See appendix B" in t
    assert "2026-09-30 14:30:00" in t                                    # date with time


def test_numbers_dates_booleans_and_gaps():
    t = block(MD, "No Header Data")
    assert table_row(t, "3") == ["1", "2.5", "3", "TRUE"]
    assert table_row(t, "5") == ["7", "", "2026-01-15", "n/a"]


def test_empty_sheet():
    assert "(empty sheet)" in block(MD, "Empty")


# ------------------------------------------------- 3. formulas
def test_saved_values_by_default_and_formulas_on_request():
    assert table_row(block(MD, "Q3 Report"), "7") == ["East", "1200", "1350", "1100", "3650"]
    assert "Formula in" not in block(MD, "Q3 Report")
    assert "Formula in E7: `=SUM(B7:D7)`" in block(RES_F.markdown, "Q3 Report")


# ------------------------------------------------- 3b. dates and times
def test_dates_times_and_durations_become_readable_text():
    t = block(MD, "Dates")
    got = {r: table_row(t, str(r))[1] for r in range(2, 11)}
    assert got == {
        2: "2026-10-05",             # dd/mm/yyyy in Excel -> unambiguous year-month-day
        3: "2026-10-05",             # shown as "Oct 2026" in Excel; the real stored date is kept
        4: "2026-10-05 14:30:00",    # date and time
        5: "09:15:30",               # time of day
        6: "36:30:00",               # elapsed time [h]:mm  (not "1 day, 12:30:00")
        7: "2026-11-04",             # date calculated by a formula
        8: "1900-01-00",             # zero date, as Excel displays it (not "00:00:00")
        9: "2026-10-05",             # text that looks like a date stays as it is
        10: "46300",                 # a number in General format stays a number
    }


def test_date_edge_cases_do_not_break():
    import datetime as dt
    import warnings
    from openpyxl.utils.datetime import CALENDAR_MAC_1904
    # Mac-style workbook that counts dates from 1904
    wb, path = tmp_book()
    wb.epoch = CALENDAR_MAC_1904
    wb.active["A1"] = dt.datetime(2026, 10, 5, 14, 30)
    wb.active["A1"].number_format = "yyyy-mm-dd hh:mm"
    wb.save(path)
    assert "2026-10-05 14:30:00" in m.parse_workbook(path).markdown
    os.remove(path)
    # a date-formatted number far outside Excel's calendar must not crash the file
    wb, path = tmp_book()
    wb.active["A1"] = "ok"; wb.active["B1"] = 99999999
    wb.active["B1"].number_format = "yyyy-mm-dd"
    wb.save(path)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        r = m.parse_workbook(path)
    assert "ok" in r.markdown and "#VALUE!" in r.markdown
    os.remove(path)


# ------------------------------------------------- 4. hidden content
def test_hidden_and_very_hidden_sheets_are_read():
    assert "HIDDEN SHEET" in MD and "VERY HIDDEN SHEET" in MD
    assert "| 2 | EUR | 1.08 |" in block(MD, "Hidden Rates")
    assert "| 2 | tax_rate | 0.21 |" in block(MD, "Config Secret")


def test_hidden_sheets_can_be_skipped():
    r = m.parse_workbook(BOOK, include_hidden_sheets=False)
    assert "EUR" not in r.markdown and "tax_rate" not in r.markdown
    assert "Widget" in r.markdown


def test_hidden_rows_and_columns_are_marked_or_skipped():
    t = block(MD, "Hidden Rows Cols")
    assert "| 3 (hidden) | 2 | Ben |" in t and "C (hidden)" in t
    t2 = block(m.parse_workbook(BOOK, include_hidden_rows_cols=False).markdown, "Hidden Rows Cols")
    assert "Ben" not in t2 and "Salary band" not in t2 and "Ana" in t2


# ------------------------------------------------- 5. big and messy
def test_long_table_is_chunked_without_losing_or_duplicating_rows():
    t = block(MD, "Big Log")
    assert "part 1 of 3" in t and "part 2 of 3" in t and "part 3 of 3" in t
    assert t.count("first row repeated for context") == 2
    rows = [int(x) for x in re.findall(r"^\| (\d+) \|", t, re.M)]
    assert rows == list(range(1, 452))                      # every row exactly once, in order


def test_messy_sheet():
    t = block(MD, "Messy")
    assert "B20" not in t                                        # whitespace-only cell ignored
    assert "Line one<br>Line two \\| with pipe" in t             # line break + pipe handled
    assert "#DIV/0!" in t                                        # error value kept as text
    assert "### Table 1 — Messy!J30:K31" in t                    # diagonal neighbours grouped
    assert "[cut: 1500 characters in total]" in t                # very long cell shortened
    assert "**Z50** (single cell): Far away cell" in t


# ------------------------------------------------- 6. merged cells
def _merged():
    t = block(MD, "Merged Cells")
    return t, t.split("### Table 1")[1].split("### Table 2")[0], t.split("### Table 2")[1]


def test_merged_horizontal():
    t, t1, _ = _merged()
    assert 'A1:F1 ("Sales by Region and Half-year")' in t
    assert table_row(t1, "3")[2:6] == ["H1", "", "H2", ""]       # value only in first cell
    assert 'C3:D3 ("H1")' in t1 and 'E3:F3 ("H2")' in t1


def test_merged_vertical():
    _, t1, _ = _merged()
    assert table_row(t1, "3")[0] == "Region" and table_row(t1, "4")[0] == ""
    assert [table_row(t1, str(r))[0] for r in (5, 6, 7)] == ["Europe", "", ""]
    assert [table_row(t1, str(r))[0] for r in (8, 9)] == ["Asia", ""]
    assert 'A5:A7 ("Europe")' in t1 and 'A8:A9 ("Asia")' in t1 and 'B3:B4 ("Country")' in t1


def test_merged_combined_rows_and_columns():
    _, _, t2 = _merged()
    assert table_row(t2, "3")[:2] == ["Target band", ""] and table_row(t2, "4")[:2] == ["", ""]
    assert "Low" in table_row(t2, "3") and "Total" in table_row(t2, "5")   # neighbours kept
    assert 'H3:I4 ("Target band")' in t2


def test_merges_do_not_split_tables_and_empty_merges_are_ignored():
    t, t1, _ = _merged()
    assert "'Merged Cells'!A3:F9" in t1
    assert "A12" not in t and "**A13** (single cell): After the empty merge" in t
    assert "**M5** (merged cell M5:R12)" in block(MD, "Messy")


def test_every_merge_that_holds_a_value_is_reported():
    from openpyxl import load_workbook
    missing = [(ws.title, mr.coord) for ws in load_workbook(BOOK).worksheets for mr in ws.merged_cells.ranges
               if ws.cell(mr.min_row, mr.min_col).value not in (None, "") and mr.coord not in MD]
    assert not missing, missing


# ------------------------------------------------- 7. comments
def test_comments_inside_a_table_are_listed_under_it():
    t = block(MD_C, "Comments")
    t1 = t.split("### Table 1")[1].split("### Comments on cells outside")[0]
    assert "- Comment on B1: Cost in USD, excluding tax." in t1
    assert "- Comment on B3: Quote expired. / Ask the vendor for a new one." in t1    # line break -> " / "
    assert "- Comment on C3: Waiting for sign-off" in t1                              # comment on an EMPTY cell


def test_threaded_comment_boilerplate_is_removed():
    t = block(MD_C, "Comments")
    assert "Comment on A4: Comment: Check this figure with Ops / Reply: Confirmed, figure is correct" in t
    assert "Your version of Excel" not in t


def test_comments_outside_tables_are_not_lost():
    t = block(MD_C, "Comments")
    assert "**A7** (single cell): Legend: Yes = approved" in t and "Comment on A7: Added by audit" in t
    tail = t.split("### Comments on cells outside any table")[1]
    assert "Comment on E10: Orphan note: ignore the Lamp row" in tail                # empty cell, no table nearby


def test_comment_authors_are_shown_when_real():
    from openpyxl.comments import Comment
    wb, path = tmp_book()
    ws = wb.active
    ws.title = "Authors"
    ws.append(["Item", "Cost"]); ws.append(["Pen", 2])
    ws["B1"].comment = Comment("Excluding tax", "Finance")
    ws["B2"].comment = Comment("Plain note", "None")                               # placeholder author: ignored
    ws["A2"].comment = Comment("[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment.\n\n"
                               "Comment:\n    Is this right?", "tc={ABC-123}")
    wb.save(path)
    t = m.parse_workbook(path, include_comments=True).markdown
    assert "Comment on B1 (by Finance): Excluding tax" in t
    assert "Comment on B2: Plain note" in t and "(by None)" not in t
    assert "Comment on A2: Comment: Is this right?" in t and "tc=" not in t         # internal id hidden
    os.remove(path)


# ------------------------------------------------- 7b. comments are opt-in
def test_comments_are_off_by_default_and_on_request():
    assert "Comment on" not in MD and "Comments on cells outside" not in MD
    assert "Cell comments are listed" not in MD                      # the note for the LLM does not mention them either
    assert "Comment on B1: Cost in USD, excluding tax." in MD_C
    assert "Cell comments are listed" in MD_C


# ------------------------------------------------- 7c. General-format dates are opt-in
def _gd(res, table_no):
    t = block(res.markdown, "General Dates")
    parts = t.split("### Table ")
    return next(p for p in parts if p.startswith(f"{table_no} "))


def test_general_dates_are_left_alone_by_default():
    t = block(MD, "General Dates")
    assert "serial" not in t and "Date detection" not in t
    assert table_row(t, "2")[1] == "46300"                            # still a plain number
    assert "serial" not in RES.legend


def test_general_dates_are_shown_with_the_original_number_when_asked():
    t = _gd(RES_D, 1)
    assert table_row(t, "2")[1] == "2026-10-05 (serial 46300)"
    assert table_row(t, "4")[1] == "2026-12-05 (serial 46361)"
    assert table_row(t, "2")[2] == "2026-10-05 14:30:00 (serial 46300.6041666667)"   # date and time
    assert table_row(t, "3")[2] == "2026-10-06 12:00:00 (serial 46301.5)"
    assert 'Date detection (opt-in): column B ("Due Date")' in t and 'column C ("Created")' in t
    assert "(serial 46300)" in RES_D.legend                           # the LLM is told how to read it


def test_general_dates_never_touch_ids_amounts_or_small_numbers():
    t = _gd(RES_D, 1)
    assert [table_row(t, str(r))[0] for r in (2, 3, 4)] == ["46300", "46301", "46302"]    # Order ID
    assert [table_row(t, str(r))[3] for r in (2, 3, 4)] == ["46300", "46301", "46302"]    # Amount Due
    assert [table_row(t, str(r))[4] for r in (2, 3, 4)] == ["12", "15", "7"]              # Date Code (not date-sized)
    for col in ("column A", "column D", "column E"):
        assert col not in t


def test_general_dates_skip_text_and_cells_that_are_already_dates():
    t = _gd(RES_D, 1)
    assert table_row(t, "3")[5] == "n/a"                              # text in a date column stays text
    assert table_row(t, "2")[5] == "2026-10-05 (serial 46300)"        # numbers in the same column convert
    assert table_row(t, "2")[6] == "2026-10-05" and "serial" not in table_row(t, "2")[6]   # real date format: as before


def test_general_dates_use_merged_group_headers_and_never_guess_without_a_header():
    t2 = _gd(RES_D, 2)
    assert table_row(t2, "3") == ["2026-10-05 (serial 46300)", "2027-10-05 (serial 46665)"]   # "Contract dates" over Start / End
    t3 = _gd(RES_D, 3)
    assert table_row(t3, "7") == ["46300", "46301"] and "serial" not in t3                    # no header: untouched


def test_the_column_header_is_the_stack_of_text_above_the_first_number():
    # "TBD" sits between the header and the first number: still the same header stack
    t4 = _gd(RES_D, 4)
    assert table_row(t4, "12") == ["A", "TBD"]
    assert table_row(t4, "13") == ["B", "2026-10-05 (serial 46300)"]
    assert 'column B ("Due Date")' in t4
    # a two-row header that is not merged ("Dates" above "Start"): both rows count; the note names the word that matched
    t5 = _gd(RES_D, 5)
    assert table_row(t5, "13") == ["B", "2026-10-05 (serial 46300)"]
    assert 'column F ("Dates")' in t5


def test_a_merged_header_counts_only_when_it_sits_directly_over_the_header():
    wb, path = tmp_book()
    ws = wb.active
    ws["B1"] = "Report date"; ws.merge_cells("B1:C1")        # merged cell, but NOT directly over the header row
    ws["A2"] = "note"                                        # (keeps everything one table)
    ws["B3"] = "Claim ref"; ws["C3"] = "Amount"
    ws["B4"] = 46300; ws["C4"] = 46300
    wb.save(path)
    assert "(serial" not in m.parse_workbook(path, legend=False, detect_general_dates=True).markdown
    os.remove(path)
    # ... but the same merged cell directly over the header does count (here over a "Start" header)
    wb, path = tmp_book()
    ws = wb.active
    ws["B1"] = "Contract dates"; ws.merge_cells("B1:C1")
    ws["B2"] = "Start"; ws["C2"] = "End"
    ws["B3"] = 46300; ws["C3"] = 46301
    wb.save(path)
    md = m.parse_workbook(path, legend=False, detect_general_dates=True).markdown
    assert "2026-10-05 (serial 46300)" in md and "2026-10-06 (serial 46301)" in md
    os.remove(path)


def test_a_title_separated_by_a_blank_row_is_not_a_header():
    wb, path = tmp_book()
    ws = wb.active
    ws["A1"] = "Report date"                                 # a title, then a blank row ...
    ws["A3"] = 46300; ws["A4"] = 46301                       # ... then numbers with no header of their own
    wb.save(path)
    assert "(serial" not in m.parse_workbook(path, legend=False, detect_general_dates=True).markdown
    os.remove(path)


def test_general_dates_custom_keywords_and_year_range():
    wb, path = tmp_book()
    ws = wb.active
    ws.append(["Ref", "Start", "Notes"]); ws.append([1, 46300, "x"]); ws.append([2, 46301, "y"])
    wb.save(path)
    # (legend=False: the "how to read this" note itself mentions "serial" when detection is on)
    assert "serial" not in m.parse_workbook(path, legend=False, detect_general_dates=True).markdown   # "Start" is not a built-in word
    r = m.parse_workbook(path, legend=False, detect_general_dates=True, date_keywords=("start",))
    assert "2026-10-05 (serial 46300)" in r.markdown                                                  # ... but you can name it
    r = m.parse_workbook(path, legend=False, detect_general_dates=True, date_keywords=("start",), date_max_year=2020)
    assert "serial" not in r.markdown                                                                 # outside the allowed years
    os.remove(path)


def test_general_dates_in_a_1904_workbook():
    from openpyxl.utils.datetime import CALENDAR_MAC_1904
    wb, path = tmp_book()
    wb.epoch = CALENDAR_MAC_1904
    wb.active.append(["Due Date"]); wb.active.append([44838]); wb.active.append([44839])
    wb.save(path)
    assert "2026-10-05 (serial 44838)" in m.parse_workbook(path, detect_general_dates=True).markdown
    os.remove(path)


def test_header_rules():
    kw = m.DEFAULT_DATE_KEYWORDS
    yes = [["Due Date"], ["Created"], ["DueDate"], ["Contract dates", "Start"], ["Updated (UTC)"], ["date_of_loss"]]
    no = [["Order ID"], ["Amount Due"], ["Claim ref"], ["Days to due date"], ["Update"], ["Date Count"], ["Total"]]
    assert all(m.looks_like_date_header(h, kw) for h in yes)
    assert not any(m.looks_like_date_header(h, kw) for h in no)


# ------------------------------------------------- 7d. input and output forms
def test_workbook_can_be_read_from_a_path_bytes_or_file_object():
    import io
    data = open(BOOK, "rb").read()
    from_path = m.parse_workbook(BOOK).markdown
    assert m.parse_workbook(data, name="excel_edge_cases.xlsx").markdown == from_path
    assert m.parse_workbook(io.BytesIO(data), name="excel_edge_cases.xlsx").markdown == from_path
    assert "Formula in E7: `=SUM(B7:D7)`" in m.parse_workbook(data, include_formulas=True).markdown   # needs a second read


def test_each_sheet_also_comes_with_its_own_text():
    sheets = {s["name"]: s["markdown"] for s in RES.sheets}
    assert sheets["Summary"].startswith("## Sheet Summary") and "Widget" in sheets["Summary"]
    assert "Widget" not in sheets["Q3 Report"]
    assert all(text in RES.markdown for text in sheets.values())     # the whole-workbook text is made of them


def test_row_and_column_limits():
    wb, path = tmp_book()
    ws = wb.active
    for r in range(1, 31):
        ws.append([f"r{r}c{c}" for c in range(1, 11)])
    wb.save(path)
    res = m.parse_workbook(path, row_limit=10, column_limit=4)
    assert "r10c4" in res.markdown and "r11c1" not in res.markdown and "r1c5" not in res.markdown
    assert "Only the first 10 rows and the first 4 columns of this sheet were read" in res.markdown
    assert any("size limit" in w for w in res.warnings)
    assert "were read" not in m.parse_workbook(path).markdown        # no limit, no notice
    os.remove(path)


def test_command_line_flags():
    out = os.path.join(tempfile.gettempdir(), "excel_to_markdown_cli_test.md")
    m.main([BOOK, "-o", out, "--comments", "--detect-general-dates", "--formulas"])
    text = open(out, encoding="utf-8").read()
    assert "Comment on B1" in text and "(serial 46300)" in text and "Formula in E7" in text
    m.main([BOOK, "-o", out])                                          # defaults: all three off
    text = open(out, encoding="utf-8").read()
    assert "Comment on" not in text and "(serial 46300)" not in text and "Formula in" not in text
    os.remove(out)


# ------------------------------------------------- 8. options
def test_gap_option_joins_tables_separated_by_one_blank_row():
    assert regions(m.parse_workbook(BOOK, gap=1), "Q3 Report") == ["A1:E14"]


# ------------------------------------------------- 9. robustness
def test_formatting_far_from_the_data_does_not_hang():
    from openpyxl.styles import PatternFill
    wb, path = tmp_book()
    ws = wb.active
    ws["A1"] = "Name"; ws["B1"] = "Qty"; ws["A2"] = "Nut"; ws["B2"] = 5
    ws["XFD200000"].fill = PatternFill("solid", fgColor="FFFF00")     # empty, formatted, very far away
    wb.save(path)
    start = time.time()
    r = m.parse_workbook(path)
    assert time.time() - start < 10 and "| 2 | Nut | 5 |" in r.markdown
    os.remove(path)


def test_one_bad_table_does_not_stop_the_file():
    original = m.render_table
    def boom(sname, idx, reg, grid, o):
        if sname == "Summary":
            raise ValueError("simulated bug")
        return original(sname, idx, reg, grid, o)
    m.render_table = boom
    try:
        r = m.parse_workbook(BOOK)
    finally:
        m.render_table = original
    assert any("Summary" in w and "simulated bug" in w for w in r.warnings)
    assert "Widget" in r.markdown and "Diagonal A" in r.markdown     # data kept (cell-by-cell) + other sheets fine


def test_one_bad_sheet_does_not_stop_the_file():
    original = m.process_sheet
    def boom(ws, wf, o, w):
        if ws.title == "Q3 Report":
            raise RuntimeError("simulated sheet failure")
        return original(ws, wf, o, w)
    m.process_sheet = boom
    try:
        r = m.parse_workbook(BOOK)
    finally:
        m.process_sheet = original
    assert any("Q3 Report" in w for w in r.warnings)
    assert "could not be read" in r.markdown and "Diagonal A" in r.markdown


def test_unreadable_files_give_clear_errors():
    for name, exc in [("does_not_exist.xlsx", FileNotFoundError), ("corrupt.xlsx", Exception)]:
        path = os.path.join(tempfile.gettempdir(), name)
        if name == "corrupt.xlsx":
            open(path, "wb").write(b"not an excel file")
        try:
            m.parse_workbook(path)
        except exc:
            pass
        else:
            raise AssertionError(f"{name} should have raised")


# ------------------------------------------------- runner
if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items(), key=lambda kv: kv[1].__code__.co_firstlineno
                                       if callable(kv[1]) and hasattr(kv[1], "__code__") else 0)
             if n.startswith("test_") and callable(f)]
    failed = 0
    for name, fn in tests:
        try:
            fn()
            print("PASS ", name)
        except Exception as e:  # noqa: BLE001
            failed += 1
            print("FAIL ", name, "->", type(e).__name__, str(e)[:200])
    print(f"\n{len(tests) - failed} of {len(tests)} tests passed")
    sys.exit(1 if failed else 0)
