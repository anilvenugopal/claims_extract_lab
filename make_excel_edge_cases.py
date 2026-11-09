#!/usr/bin/env python3
"""
Builds samples/excel_edge_cases.xlsx: a workbook with 14 tabs of different difficulty, used by
the tests (tests/test_excel_to_markdown.py) and the notebook (notebooks/02_excel_extraction_lab.ipynb).

You only need this if you want to regenerate or extend the workbook; the finished file is already
in samples/. (make_samples.py does not rebuild it, because this script needs LibreOffice.)

Steps this script performs:
  1. Build the workbook with openpyxl.
  2. Open it in LibreOffice (headless) so formulas get calculated and saved values.
     (needs `soffice` on your PATH; without it the file is still written, but
     formula cells will have no saved values)
  3. Put the "very hidden" state back on one tab (LibreOffice turns it into plain "hidden").

Run (from the repo root):
    python make_excel_edge_cases.py
    python src/excel_to_markdown.py samples/excel_edge_cases.xlsx -o docs/excel_edge_cases_output.md --comments --detect-general-dates
The second command refreshes the saved output that docs/excel_to_markdown.md quotes from.
"""
import datetime as dt
import os
import re
import shutil
import subprocess
import tempfile
import zipfile

from openpyxl import Workbook
from openpyxl.comments import Comment

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "samples", "excel_edge_cases.xlsx")

THREADED_COMMENT = (
    "[Threaded comment]\n\n"
    "Your version of Excel allows you to read this threaded comment; however, any edits to it "
    "will get removed if the file is opened in a newer version of Excel. "
    "Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\n"
    "Comment:\n    Check this figure with Ops\nReply:\n    Confirmed, figure is correct"
)


def build(path: str) -> None:
    wb = Workbook()

    # 1. Summary: simple, clean, header in row 1 --------------------------------
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Product", "Units", "Price"])
    for row in [("Widget", 120, 9.99), ("Gadget", 75, 24.5), ("Gizmo", 30, 0.1 + 0.2), ("Doohickey", 200, 4)]:
        ws.append(row)

    # 2. Q3 Report: merged title, notes above, 2-level merged header, formulas --
    ws = wb.create_sheet("Q3 Report")
    ws["A1"] = "Q3 Regional Sales Report"; ws.merge_cells("A1:E1")
    ws["A2"] = "Prepared by: Finance Team"
    ws["A3"] = "Date:"; ws["B3"] = dt.date(2026, 10, 1)
    ws["A5"] = "Region"; ws.merge_cells("A5:A6")
    ws["B5"] = "Q3 Revenue ($)"; ws.merge_cells("B5:D5")
    ws["E5"] = "Total"; ws.merge_cells("E5:E6")
    for col, m in zip("BCD", ["Jul", "Aug", "Sep"]):
        ws[f"{col}6"] = m
    data = [("East", 1200, 1350, 1100), ("West", 900, 950, 1010), ("North", 400, 420, 380), ("South", 760, 800, 845)]
    for i, (reg, j, a, s) in enumerate(data, start=7):
        ws[f"A{i}"] = reg; ws[f"B{i}"] = j; ws[f"C{i}"] = a; ws[f"D{i}"] = s
        ws[f"E{i}"] = f"=SUM(B{i}:D{i})"
    ws["A11"] = "Total"
    for col in "BCDE":
        ws[f"{col}11"] = f"=SUM({col}7:{col}10)"
    ws["A13"] = "Note: figures are unaudited | source: ERP export.\nSecond line of the note."
    ws["A14"] = "Footnote spanning the table width, merged across A to E."; ws.merge_cells("A14:E14")

    # 3. Multi Tables: several tables, offsets, one without headers, stray note --
    ws = wb.create_sheet("Multi Tables")
    ws["B2"] = "Headcount by department"
    for j, h in enumerate(["Dept", "Heads", "Open roles"]):
        ws.cell(3, 2 + j, h)
    for i, row in enumerate([("Eng", 40, 5), ("Sales", 25, 2), ("Ops", 12, 0)], start=4):
        for j, v in enumerate(row):
            ws.cell(i, 2 + j, v)
    ws["H3"] = "Dept"; ws["I3"] = "Budget"
    for i, row in enumerate([("Eng", 4_000_000), ("Sales", 1_500_000)], start=4):
        ws.cell(i, 8, row[0]); ws.cell(i, 9, row[1])
    for i, row in enumerate([(1, 2, 3), (4, 5, 6), (7, 8, 9), (10, 11, 12)], start=11):  # no header
        for j, v in enumerate(row):
            ws.cell(i, 2 + j, v)
    ws["K20"] = "See appendix B"
    ws["B25"] = "Owner"; ws["C25"] = "J. Rivera"
    ws["B26"] = "Reviewed"; ws["C26"] = dt.datetime(2026, 9, 30, 14, 30)

    # 4. No Header Data: numbers only, blanks, bool, text numbers ---------------
    ws = wb.create_sheet("No Header Data")
    rows = [(1, 2.5, "3", True), (None, 4, 5.25, False), (7, None, dt.date(2026, 1, 15), "n/a"), (10, 11, 12, 13)]
    for i, row in enumerate(rows, start=3):
        for j, v in enumerate(row):
            if v is not None:
                ws.cell(i, 3 + j, v)

    # 5. Hidden Rates: HIDDEN sheet ---------------------------------------------
    ws = wb.create_sheet("Hidden Rates")
    ws.append(["Currency", "Rate to USD"])
    for r in [("EUR", 1.08), ("GBP", 1.27), ("JPY", 0.0067)]:
        ws.append(r)
    ws.sheet_state = "hidden"

    # 6. Config Secret: VERY HIDDEN sheet (state restored after recalculation) --
    ws = wb.create_sheet("Config Secret")
    ws["A1"] = "Setting"; ws["B1"] = "Value"
    ws["A2"] = "tax_rate"; ws["B2"] = 0.21
    ws["A3"] = "fiscal_year_start"; ws["B3"] = "April"
    ws.sheet_state = "veryHidden"

    # 7. Hidden Rows Cols: visible sheet with hidden rows, a hidden column ------
    ws = wb.create_sheet("Hidden Rows Cols")
    ws.append(["ID", "Name", "Salary band", "Status"])
    for r in [(1, "Ana", "B2", "Active"), (2, "Ben", "B3", "Left"), (3, "Cy", "B1", "Active"), (4, "Dee", "B4", "Left")]:
        ws.append(r)
    ws.row_dimensions[3].hidden = True
    ws.row_dimensions[5].hidden = True
    ws.column_dimensions["C"].hidden = True

    # 8. Empty -------------------------------------------------------------------
    wb.create_sheet("Empty")

    # 9. Big Log: 450 rows to trigger chunking -----------------------------------
    ws = wb.create_sheet("Big Log")
    ws.append(["Row ID", "Event", "Value", "Flag"])
    for i in range(1, 451):
        ws.append([i, f"event_{i}", i * 1.5, i % 7 == 0])

    # 10. Messy: scattered cells, error, long text, blanks, pipe, touching cells -
    ws = wb.create_sheet("Messy")
    ws["A1"] = "Top-left note"
    ws["Z50"] = "Far away cell"
    ws["D10"] = "Line one\nLine two | with pipe"
    ws["F2"] = "x" * 1500
    ws["B20"] = "   "                      # whitespace only: counts as empty
    ws["H8"] = "=1/0"
    ws["J30"] = "Diagonal A"; ws["K31"] = "Diagonal B"   # touch only at a corner
    ws["M5"] = "Big merge value"; ws.merge_cells("M5:R12")

    # 11. Merged Cells: horizontal, vertical, combined, empty merge ---------------
    ws = wb.create_sheet("Merged Cells")
    ws["A1"] = "Sales by Region and Half-year"; ws.merge_cells("A1:F1")      # horizontal
    ws["A3"] = "Region"; ws.merge_cells("A3:A4")                             # vertical
    ws["B3"] = "Country"; ws.merge_cells("B3:B4")                            # vertical
    ws["C3"] = "H1"; ws.merge_cells("C3:D3")                                 # horizontal
    ws["E3"] = "H2"; ws.merge_cells("E3:F3")                                 # horizontal
    for col, q in zip("CDEF", ["Q1", "Q2", "Q3", "Q4"]):
        ws[f"{col}4"] = q
    rows = [("France", 10, 12, 14, 16), ("Spain", 8, 9, 11, 13), ("Italy", 7, 7, 9, 10),
            ("Japan", 20, 22, 21, 25), ("India", 15, 18, 19, 24)]
    for i, (country, *q) in enumerate(rows, start=5):
        ws[f"B{i}"] = country
        for col, v in zip("CDEF", q):
            ws[f"{col}{i}"] = v
    ws["A5"] = "Europe"; ws.merge_cells("A5:A7")                             # vertical group label
    ws["A8"] = "Asia"; ws.merge_cells("A8:A9")                               # vertical group label
    ws["H3"] = "Target band"; ws.merge_cells("H3:I4")                        # combined: 2 rows x 2 cols
    ws["J3"] = "Low"; ws["K3"] = "High"; ws["J4"] = 1; ws["K4"] = 2
    ws["H5"] = "Europe"; ws["I5"] = "Total"; ws["J5"] = 3; ws["K5"] = 4
    ws.merge_cells("A12:C12")                                                # merge with NO value
    ws["A13"] = "After the empty merge"

    # 12. Comments: on headers, data, empty cells, stray cells, threaded ----------
    ws = wb.create_sheet("Comments")
    ws.append(["Item", "Cost", "Approved"])
    ws.append(["Pen", 2.5, "Yes"])
    ws.append(["Desk", 180, None])
    ws.append(["Lamp", 35, "Yes"])
    ws["B1"].comment = Comment("Cost in USD, excluding tax.", "Finance")
    ws["B3"].comment = Comment("Quote expired.\nAsk the vendor for a new one.", "Ana")
    ws["C3"].comment = Comment("Waiting for sign-off", "Ben")               # on an EMPTY cell inside the table
    ws["A4"].comment = Comment(THREADED_COMMENT, "tc={9F2D1C3A-0001}")       # newer Excel 'threaded' comment
    ws["A7"] = "Legend: Yes = approved"                                      # stray cell with a comment
    ws["A7"].comment = Comment("Added by audit", "Audit")
    ws["E10"].comment = Comment("Orphan note: ignore the Lamp row", "Cy")    # on an EMPTY cell far from any table

    # 13. Dates: how dates, times and durations come through ----------------------
    ws = wb.create_sheet("Dates")
    ws.append(["Case", "Stored value", "Excel displays"])
    base = dt.date(2026, 10, 5)
    ws.append(["Date (dd/mm/yyyy)", base, "05/10/2026"]);                         ws["B2"].number_format = "dd/mm/yyyy"
    ws.append(["Date (month and year only)", base, "Oct 2026"]);                 ws["B3"].number_format = "mmm yyyy"
    ws.append(["Date and time", dt.datetime(2026, 10, 5, 14, 30), "2026-10-05 14:30"]); ws["B4"].number_format = "yyyy-mm-dd hh:mm"
    ws.append(["Time of day", dt.time(9, 15, 30), "09:15:30"]);                  ws["B5"].number_format = "hh:mm:ss"
    ws.append(["Elapsed time (hours)", 1.5208333333333333, "36:30"]);            ws["B6"].number_format = "[h]:mm"
    ws.append(["Date from a formula", "=B2+30", "04/11/2026"]);                  ws["B7"].number_format = "dd/mm/yyyy"
    ws.append(["Empty date (zero)", "=B2-B2", "00/01/1900"]);                    ws["B8"].number_format = "dd/mm/yyyy"
    ws.append(["Date stored as text", "2026-10-05", "2026-10-05"])
    ws.append(["Number, not a date", 46300, "46300"]);                           ws["B10"].number_format = "General"

    # 14. General Dates: dates stored as plain numbers (opt-in detection) ----------
    ws = wb.create_sheet("General Dates")
    ws.append(["Order ID", "Due Date", "Created", "Amount Due", "Date Code", "Updated", "Posted"])
    ws.append([46300, 46300, 46300.604166666664, 46300, 12, 46300, dt.date(2026, 10, 5)])
    ws.append([46301, 46330, 46301.5, 46301, 15, "n/a", dt.date(2026, 11, 4)])
    ws.append([46302, 46361, 46302.25, 46302, 7, 46303, dt.date(2026, 12, 5)])
    for r in (2, 3, 4):
        ws.cell(r, 7).number_format = "dd/mm/yyyy"          # a real date format: not touched by detection
    ws["I1"] = "Contract dates"; ws.merge_cells("I1:J1")    # merged group header over two columns
    ws["I2"] = "Start"; ws["J2"] = "End"
    ws["I3"] = 46300; ws["J3"] = 46665
    ws["I4"] = 46301; ws["J4"] = 46666
    ws["A7"] = 46300; ws["B7"] = 46301                      # numbers with no header: never guessed
    ws["A8"] = 46302; ws["B8"] = 46303
    # the header is the unbroken stack of text directly above the first number:
    ws["A11"] = "Claim"; ws["B11"] = "Due Date"                # a text value ("TBD") between header and numbers
    ws["A12"] = "A"; ws["B12"] = "TBD"
    ws["A13"] = "B"; ws["B13"] = 46300
    ws["A14"] = "C"; ws["B14"] = 46301
    ws["E11"] = "Claim"; ws["F11"] = "Dates"                   # a two-row header that is NOT merged
    ws["F12"] = "Start"
    ws["E13"] = "B"; ws["F13"] = 46300
    ws["E14"] = "C"; ws["F14"] = 46301

    wb.save(path)


def recalculate(path: str) -> None:
    """Open in LibreOffice and save, so formulas have stored values."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("WARNING: LibreOffice not found; formula cells will have no saved values.")
        return
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--convert-to", "xlsx", "--outdir", tmp, path],
                       check=True, capture_output=True, timeout=180)
        shutil.move(os.path.join(tmp, os.path.basename(path)), path)


def restore_very_hidden(path: str, sheet_name: str = "Config Secret") -> None:
    """LibreOffice turns 'very hidden' into 'hidden'; patch the file's XML to put it back."""
    tmp = path + ".tmp"
    pattern = re.compile(r'(<sheet [^>]*name="%s"[^>]*?)state="hidden"' % re.escape(sheet_name))
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = pattern.sub(r'\1state="veryHidden"', data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path)


if __name__ == "__main__":
    build(OUT)
    recalculate(OUT)
    restore_very_hidden(OUT)
    print("wrote", OUT)
