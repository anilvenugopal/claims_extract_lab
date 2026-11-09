# Excel extraction, option 2: coordinate-aware markdown (openpyxl)

`excelExtractText` can read Excel files in two ways. You choose with the `method` field:

| `method` | What it does | Good for |
|---|---|---|
| `"pandas"` (default) | The original extractor: one markdown table per sheet, generated column numbers (`0 1 2 ...`) as the header. | Keeping today's behaviour exactly as it is. |
| `"openpyxl"` | Reads the workbook cell by cell and writes markdown that keeps **Excel's own column letters and row numbers**, so an LLM can pull out fields **and cite the exact cell** (`'Loss Run'!G7`). | Messy real-world files: merged cells, several tables on a sheet, notes around tables, hidden sheets, comments. |

Nothing changes unless you ask for `method="openpyxl"`. No new libraries are needed (`openpyxl` was already a requirement).

---

## 1. Files

```text
src/
├── excel_to_markdown.py            the converter (also usable as a command-line tool)
├── text_extract.py                 class OpenpyxlExcelExtractor wraps the converter
└── app.py                          excelExtractText takes ExcelPayload (method + options)
local_harness.py                    excelExtractText uses ExcelPayload
make_samples.py                     copies the edge-case workbook along with the other samples
make_excel_edge_cases.py            rebuilds the edge-case workbook (needs LibreOffice)
samples/
└── excel_edge_cases.xlsx           the complex test workbook: 14 tabs, each a different kind of mess
docs/
├── excel_to_markdown.md            this file
└── excel_edge_cases_output.md      the converter's output for that workbook (comments and date detection on)
notebooks/
├── 01_claims_extraction_lab.ipynb  all file types (its Excel part points to notebook 02)
└── 02_excel_extraction_lab.ipynb   Excel only: pandas and openpyxl side by side, tab by tab
tests/
├── test_excel_to_markdown.py       the converter itself
├── test_excel_openpyxl_method.py   the method switch and options (excelExtractText)
└── test_docs_and_notebooks.py      checks that this document and the notebooks match the code
```

---

## 2. How to use it

### Through `excelExtractText` (the same call as before, plus options)

```python
import local_harness as h

# default: pandas, exactly as before
h.run("excelExtractText", "samples/loss_run.xlsx")

# openpyxl method
h.run("excelExtractText", "samples/loss_run.xlsx", method="openpyxl")

# with options
h.run("excelExtractText", "samples/loss_run.xlsx", method="openpyxl",
      include_comments=True, detect_general_dates=True)
```

The result has the same shape as before: one `{"sheet_name": ..., "text": ...}` per sheet.

| Payload field | Default | What it does |
|---|---|---|
| `method` | `"pandas"` | `"pandas"` or `"openpyxl"`. Anything else is an error. |
| `include_comments` | off | List cell comments under the table. |
| `include_formulas` | off | List formulas under the table. |
| `include_hidden_sheets` | on | Set to `False` to leave out hidden and very hidden sheets. |
| `detect_general_dates` | off | Show date-like numbers in General-format cells as dates (section 5). |
| `date_keywords` | built-in list | Header words that mean "date". Replaces the built-in list. |
| `include_legend` | on | Put the short "how to read this" note at the top of each sheet's text. |
| `fallback_to_pandas` | on | `.xls` and `.xlsb` cannot be read with openpyxl. On: pandas is used for them. Off: an error. |

The options only work with `method="openpyxl"`. Asking for one while using the pandas method is an error, so an option is never silently ignored.

Like the pandas version, only the first 2000 rows and 200 columns of a sheet are read. When something is cut off, the text says so.

### On its own (command line, from the repo root)

```bash
python src/excel_to_markdown.py report.xlsx -o report.md
python src/excel_to_markdown.py report.xlsx --comments --detect-general-dates --formulas
python src/excel_to_markdown.py --help
```

### On its own (Python)

```python
from excel_to_markdown import parse_workbook

result = parse_workbook("report.xlsx")          # a path, raw bytes, or an open file
result.markdown    # the whole workbook as one text
result.sheets      # per sheet: name, state, regions found, and "markdown" (that sheet's own text)
result.warnings    # anything that went wrong (the converter does not crash on bad sheets)
```

### Converter options

| Option (Python / command line) | Default | What it does |
|---|---|---|
| `include_comments` / `--comments` | **off** | List cell comments. |
| `include_formulas` / `--formulas` | off | List formulas. |
| `detect_general_dates` / `--detect-general-dates` | **off** | Section 5. |
| `date_keywords` / `--date-keywords a,b,c` | built-in list | Header words that mean "date". |
| `date_min_year`, `date_max_year` / `--date-min-year`, `--date-max-year` | 1950, 2100 | Numbers that would be dates outside these years are left alone. |
| `include_hidden_sheets` / `--skip-hidden-sheets` | hidden sheets **included** | Leave out hidden and very hidden sheets. |
| `include_hidden_rows_cols` / `--skip-hidden-rows-cols` | hidden rows and columns **included** | Leave them out. |
| `gap` / `--gap N` | 0 | Allow N blank rows or columns *inside* one table. |
| `max_rows_per_chunk` / `--chunk-rows N` | 200 | Split tables longer than this into parts. |
| `row_limit`, `column_limit` / `--row-limit`, `--column-limit` | none | Only read the first N rows or columns of each sheet. |
| `max_cell_chars` / `--max-cell-chars N` | 1000 | Shorten any single cell longer than this. |

---

## 3. What changes compared with the pandas method

All of this is checked by `tests/test_excel_openpyxl_method.py` against the lab's own sample workbooks.

| | `pandas` (default) | `openpyxl` |
|---|---|---|
| Header row | Generated numbers `0 1 2 ...`; the real headers become a data row. | Excel's column letters on top. The real headers are still an ordinary row (section 5 explains why that matters), now with Excel's row numbers beside them. |
| Several tables on one sheet | Flattened into one grid. A side-by-side table shares rows with the table next to it. | Each table is its own block with its cell range, for example `'Loss Run'!K4:M8`. |
| Merged cells | The value sits in the first row only; nothing says it was merged. | The value sits in the first cell, and the merged range is listed under the table. |
| Numbers | Rounded to 2 decimals (`0.6234` becomes `0.62`). | Full precision (`0.6234`), with floating-point noise removed (`0.30000000000000004` becomes `0.3`). |
| Dates | `2026-04-30 00:00:00` | `2026-04-30`. Date and time, time of day, elapsed hours and zero dates are shown as Excel shows them (section 4). |
| Formulas | Saved results only. | Saved results; formula text on request. |
| Comments | Ignored. | On request. |
| Hidden sheets | Read, but not marked. | Read and labelled `HIDDEN SHEET` / `VERY HIDDEN SHEET`. Hidden rows and columns are marked too. |
| Dates stored as plain numbers (General format) | Plain numbers. | Plain numbers, unless `detect_general_dates` is on (section 5). |
| Citations | Not possible. | Every table has Excel's coordinates. |
| File types | `.xls` `.xlsm` `.xlsb` `.xlsx` | `.xlsx` `.xlsm`. Others fall back to pandas. |
| `extract()` on a multi-sheet file | Error unless exactly one sheet. | Returns one text per sheet. (`extractText` still uses pandas.) |
| Size of the text | Compact for ordinary tables. For sparse sheets it prints a grid that is mostly empty cells. | Larger for ordinary tables (row numbers, the legend and the notes under tables). Smaller for sparse sheets, because it lists only the blocks that exist. |

What both methods still ignore: charts, images and text boxes.

---

## 4. What the output looks like

Every sheet's text starts with a short "how to read this" note for the LLM (turn it off with `include_legend=False`), then the sheet is split into blocks. All examples below are real output from [`excel_edge_cases_output.md`](excel_edge_cases_output.md); a test checks that they still match.

### Tables and notes
A sheet with a title, notes above the table, the table itself and a footnote becomes separate blocks (`'Q3 Report'`, ranges `A1:E3`, `A5:E11`, `A13:E14`):

```
### Table 2 — 'Q3 Report'!A5:E11

| | A | B | C | D | E |
|---|---|---|---|---|---|
| 5 | Region | Q3 Revenue ($) |  |  | Total |
| 6 |  | Jul | Aug | Sep |  |
| 7 | East | 1200 | 1350 | 1100 | 3650 |
| 8 | West | 900 | 950 | 1010 | 2860 |
```

A single cell on its own is one line, for example `**K20** (single cell): See appendix B`. Very scattered cells are listed one per line.

### Merged cells
The value is written once, in the first cell. The other cells of the merge are blank. The merged ranges are listed under the table. Horizontal (`C3:D3`), vertical (`A5:A7`) and combined (`H3:I4`, two rows and two columns at once) merges all work the same way:

```
| | A | B | C | D | E | F |
|---|---|---|---|---|---|---|
| 3 | Region | Country | H1 |  | H2 |  |
| 4 |  |  | Q1 | Q2 | Q3 | Q4 |
| 5 | Europe | France | 10 | 12 | 14 | 16 |
| 6 |  | Spain | 8 | 9 | 11 | 13 |
| 7 |  | Italy | 7 | 7 | 9 | 10 |
| 8 | Asia | Japan | 20 | 22 | 21 | 25 |

- Merged cells (value sits in the first cell of the range): A3:A4 ("Region"); B3:B4 ("Country"); C3:D3 ("H1"); E3:F3 ("H2"); A5:A7 ("Europe"); A8:A9 ("Asia")
```

```
| | H | I | J | K |
|---|---|---|---|---|
| 3 | Target band |  | Low | High |
| 4 |  |  | 1 | 2 |

- Merged cells (value sits in the first cell of the range): H3:I4 ("Target band")
```

A merge with no value (formatting only) is ignored.

### Comments (only with `include_comments=True`)
Comments are listed under their table as `Comment on <cell>: <text>`. A real author is added: `Comment on B1 (by Finance): ...`. Placeholder authors such as "None" or "Unknown Author" are left out. Line breaks become ` / `. Newer "threaded" comments are cleaned of Excel's long warning text.

Comments on **empty cells** are kept. A comment on a cell outside every table goes into its own block at the end of the sheet, so it is never lost:

```
### Table 1 — Comments!A1:C4

| | A | B | C |
|---|---|---|---|
| 1 | Item | Cost | Approved |
| 2 | Pen | 2.5 | Yes |
| 3 | Desk | 180 |  |
| 4 | Lamp | 35 | Yes |

- Comment on B1: Cost in USD, excluding tax.
- Comment on B3: Quote expired. / Ask the vendor for a new one.
- Comment on C3: Waiting for sign-off
- Comment on A4: Comment: Check this figure with Ops / Reply: Confirmed, figure is correct

**A7** (single cell): Legend: Yes = approved

- Comment on A7: Added by audit

### Comments on cells outside any table — Comments

- Comment on E10: Orphan note: ignore the Lamp row
```

(`C3` is an empty cell inside the table. `E10` is an empty cell far from any table.)

### Hidden sheets, rows and columns
Sheet headings say `HIDDEN SHEET` or `VERY HIDDEN SHEET` (the kind Excel's "Unhide" menu cannot show). Hidden rows and columns are marked:

```
Hidden rows: [3, 5]; hidden columns: ['C'] (shown below and marked)

| | A | B | C (hidden) | D |
|---|---|---|---|---|
| 1 | ID | Name | Salary band | Status |
| 2 | 1 | Ana | B2 | Active |
| 3 (hidden) | 2 | Ben | B3 | Left |
```

### Dates that Excel has formatted as dates
Excel stores a date as a plain number plus a display format. The reader turns it back into a real date, and the converter writes it in one fixed form. From the `Dates` tab (the right-hand column is what Excel displays):

```
| | A | B | C |
|---|---|---|---|
| 1 | Case | Stored value | Excel displays |
| 2 | Date (dd/mm/yyyy) | 2026-10-05 | 05/10/2026 |
| 3 | Date (month and year only) | 2026-10-05 | Oct 2026 |
| 4 | Date and time | 2026-10-05 14:30:00 | 2026-10-05 14:30 |
| 5 | Time of day | 09:15:30 | 09:15:30 |
| 6 | Elapsed time (hours) | 36:30:00 | 36:30 |
| 7 | Date from a formula | 2026-11-04 | 04/11/2026 |
| 8 | Empty date (zero) | 1900-01-00 | 00/01/1900 |
| 9 | Date stored as text | 2026-10-05 | 2026-10-05 |
| 10 | Number, not a date | 46300 | 46300 |
```

- Dates become `2026-10-05`, date and time `2026-10-05 14:30:00`, times `09:15:30`. Year-month-day cannot be misread as day/month or month/day.
- Elapsed times (Excel format `[h]:mm`) become total hours, for example `36:30:00`.
- A "zero date" (a formula pointing at a blank cell, formatted as a date, which Excel shows as `00/01/1900`) is written as `1900-01-00`.
- Files that count dates from 1904 (older Mac files) are converted correctly.
- A date-formatted number outside Excel's calendar is read as the error `#VALUE!`. It does not stop the file.
- Display formats are not applied: `Oct 2026` is written as the real stored date (`2026-10-05`).

### Formulas (only with `include_formulas=True`)
Values are always the saved results. With the option on, formulas are listed under the table: ``Formula in E7: `=SUM(B7:D7)` ``. If a file was written by a script and never opened in Excel, formula cells have no saved value, and the output says `[no saved value in file]`.

### Long tables
A table over 200 rows is split into parts. The header row is repeated and keeps its real row number:

```
### Table 1 — 'Big Log'!A1:D451 (part 2 of 3, rows 201–400)

| | A | B | C | D |
|---|---|---|---|---|
| 1 (first row repeated for context) | Row ID | Event | Value | Flag |
| 201 | 200 | event_200 | 300 | FALSE |
```

---

## 5. Dates stored as plain numbers (General format) — opt-in

A date in a cell whose format is **General** is stored as a plain number (for example `46300`), and Excel itself shows `46300`. The file has no record that it is a date, so by default the converter writes the number as it is. A model reading `46300` under a "Date" header may not know it is a date, and models are poor at converting these numbers by themselves.

With `detect_general_dates=True` the converter shows the date **and keeps the original number**: `2026-10-05 (serial 46300)`. This exists only in the openpyxl method. The pandas method has no date detection.

### What "column header" means here

In the markdown tables (both methods), the **top row is not your header**. It is only Excel's column letters (`A B C`), or `0 1 2` in the pandas method. Your table's own header, the words such as "Due Date", is an ordinary cell in the sheet. In the output it therefore shows up as the **first data row**, directly under the letters.

Date detection looks at the **cells of the sheet**, never at the letters row. For each column it takes as the header:

1. **the unbroken stack of text cells directly above the first General-format number in that column.** One row is the usual case. Several rows work too (a two-row header). A text value such as `TBD` between the header and the first number is part of the stack, so it does not hide the header.
2. **a merged cell sitting directly above that stack**, such as a group header "Contract dates" over "Start" and "End".

The header can be on any row of the sheet: a table that starts on row 11 under other content works the same way. All the examples below are from the `General Dates` tab with detection on.

**Header on the first row of the table** (`Due Date` and `Created` are converted; `Order ID` and `Amount Due` hold the same numbers and are not, because their headers do not look like dates):

```
| | A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|---|
| 1 | Order ID | Due Date | Created | Amount Due | Date Code | Updated | Posted |
| 2 | 46300 | 2026-10-05 (serial 46300) | 2026-10-05 14:30:00 (serial 46300.6041666667) | 46300 | 12 | 2026-10-05 (serial 46300) | 2026-10-05 |
| 3 | 46301 | 2026-11-04 (serial 46330) | 2026-10-06 12:00:00 (serial 46301.5) | 46301 | 15 | n/a | 2026-11-04 |
| 4 | 46302 | 2026-12-05 (serial 46361) | 2026-10-07 06:00:00 (serial 46302.25) | 46302 | 7 | 2026-10-08 (serial 46303) | 2026-12-05 |

- Date detection (opt-in): column B ("Due Date"): 3 number(s) stored in General format are shown as dates, with the original number in brackets
- Date detection (opt-in): column C ("Created"): 3 number(s) stored in General format are shown as dates, with the original number in brackets
- Date detection (opt-in): column F ("Updated"): 2 number(s) stored in General format are shown as dates, with the original number in brackets
```

The quoted text in each note is the header text that qualified. `Posted` has a real date format, so it needed no help. `Date Code` is date-like by name but not by size (see rule 4 below), and the `n/a` in `Updated` stays text.

**Header with a text value between it and the numbers** (`TBD` is part of the stack):

```
| | A | B |
|---|---|---|
| 11 | Claim | Due Date |
| 12 | A | TBD |
| 13 | B | 2026-10-05 (serial 46300) |
| 14 | C | 2026-10-06 (serial 46301) |
```

**Two header rows that are not merged** (the stack is `Start` and `Dates`; the note names `Dates`, the word that matched):

```
| | E | F |
|---|---|---|
| 11 | Claim | Dates |
| 12 |  | Start |
| 13 | B | 2026-10-05 (serial 46300) |
| 14 | C | 2026-10-06 (serial 46301) |

- Date detection (opt-in): column F ("Dates"): 2 number(s) stored in General format are shown as dates, with the original number in brackets
```

**A merged group header directly above the header** ("Contract dates" covers `Start` and `End`):

```
| | I | J |
|---|---|---|
| 1 | Contract dates |  |
| 2 | Start | End |
| 3 | 2026-10-05 (serial 46300) | 2027-10-05 (serial 46665) |
| 4 | 2026-10-06 (serial 46301) | 2027-10-06 (serial 46666) |
```

**What is not a header:**
- The letters row (`A B C`, or `0 1 2`).
- Text further up the column that is separated from the numbers by a blank cell, such as a title. A title with a blank row under it is a separate block anyway. A title **directly** above the header (no blank row) is part of the stack, so a date-like word there counts. The note under the table shows what matched.
- A merged cell that is not directly above the header stack.
- Text **below** the first number. That is data.
- Nothing at all: **numbers with no text above them are never touched.** The block at `A7:B8` of the `General Dates` tab has two columns of 46300-style numbers and no header, and stays as plain numbers.

### All the rules

A General-format number is converted only if **all** of these are true:

1. The cell's format is General (cells with a real date format are already handled).
2. The column header (as defined above) contains a word from the built-in list: `date`, `dates`, `datetime`, `timestamp`, `dt`, `dob`, `created`, `modified`, `updated`, `due`, `deadline`, `issued`, `expiry`, `expiration`, `expires`, `effective`, `posted`, `opened`, `closed`, `received`. Words are split on spaces, punctuation and capital letters, so `DueDate` and `due_date` both match.
3. The header does **not** also contain a word that suggests an amount, count or id: `amount`, `balance`, `total`, `count`, `qty`, `quantity`, `price`, `cost`, `rate`, `days`, `age`, `duration`, `hours`, `minutes`, `id`, `number`. So "Amount Due" and "Days to due date" are left alone.
4. The number is in a plausible range (the years 1950 to 2100 by default).
5. At least half of the column's General numbers are in that range. A "Date Code" column holding 12, 15 and 7 is not a date column.

If your files use other header words ("Start", "Loss"), pass your own list: `date_keywords=["start", "loss"]` replaces the built-in one.

Not covered: times of day stored as fractions in General cells (`0.6041...`), and plain-number formats other than General (for example `0`).

---

## 6. Design in short

1. **Read every cell that really exists** in every sheet (visible, hidden, very hidden) with its merged ranges, hidden rows and columns, and the saved results of formulas. Only stored cells are read: sheets with formatting far past the data (a common real-world case) would otherwise use up all memory.
2. **Find blocks.** Filled cells that touch each other (side by side, above and below, or diagonally) belong together. A merged cell counts as filling its whole area. Each block is one table, one stray cell, or a list of scattered cells.
3. **Print each table with Excel's coordinates** in the margins: column letters across the top, row numbers down the left side. This keeps tables readable (the model can compare across a row) while any cell can still be cited.
4. **Add notes under the table**: merged ranges, then (if asked) comments, date-detection notes and formulas.
5. **Never stop on a problem.** An error in one table falls back to a cell-by-cell list for that table. An error in one sheet becomes a warning plus a note in the output.

Headers are never guessed in code when the table is written: the first row of a block is shown as it is, and the model decides what is a header. Only the opt-in date detection looks for headers, using the rules in section 5. Fully blank rows and columns inside a block are left out (the row numbers skip, so gaps are visible).

---

## 7. Limits

- `.xlsx` and `.xlsm` only. `.xls` and `.xlsb` go to the pandas method. Charts, images and text boxes are ignored.
- Two tables that touch with no blank row or column between them become **one** block. Tables separated by one blank row stay separate (`gap=1` joins them on purpose).
- The text `"3"` and the number 3 look the same.
- Display formats are not applied: `15%` is written as `0.15`.
- Dates before 1 March 1900 can be off by a day (an Excel quirk).
- The whole sheet is loaded into memory, so very large sheets are slow. The 2000 x 200 limit applies by default through `excelExtractText`.
- Replies in newer "threaded" comments are flattened into one line.

---

## 8. The edge-case workbook, the notebook and the tests

### The workbook: `samples/excel_edge_cases.xlsx`
Fourteen tabs, each a different kind of mess. Two are hidden on purpose. It is committed to the repo, and `make_samples.py` copies it along with the other samples, so `S["excel_edge_cases.xlsx"]` works like any other sample.

| # | Tab | What it contains |
|---|---|---|
| 1 | `Summary` | Plain table, header in row 1 |
| 2 | `Q3 Report` | Merged title, notes above, two-level merged header, formulas, notes below |
| 3 | `Multi Tables` | Four tables (two side by side, one with no header, one small key-value table) and a stray note |
| 4 | `No Header Data` | Numbers only, gaps, a date, true/false, text |
| 5 | `Hidden Rates` | **Hidden** sheet |
| 6 | `Config Secret` | **Very hidden** sheet |
| 7 | `Hidden Rows Cols` | Hidden rows 3 and 5, hidden column C |
| 8 | `Empty` | Nothing |
| 9 | `Big Log` | 450 rows (split into 3 parts) |
| 10 | `Messy` | Scattered cells, `\|` characters, line breaks, `#DIV/0!`, a 1,500-character cell, a blank-looking cell, diagonal neighbours, one big merge |
| 11 | `Merged Cells` | Horizontal, vertical and combined merges, plus a merge with no value |
| 12 | `Comments` | Comments on a header, a data cell, an empty cell in a table, a stray cell, an empty cell far from any table, a threaded comment |
| 13 | `Dates` | Date, month-only date, date and time, time, elapsed hours, date from a formula, zero date, date as text, plain number |
| 14 | `General Dates` | Dates as plain numbers: under date-like and non-date headers, a merged group header, a `TBD` between header and numbers, a two-row header, text in a date column, a real date column, numbers with no header |

To rebuild it (needs LibreOffice, which opens the file once so formulas have saved values):

```bash
python make_excel_edge_cases.py
python src/excel_to_markdown.py samples/excel_edge_cases.xlsx -o docs/excel_edge_cases_output.md --comments --detect-general-dates
```

### The notebook: `notebooks/02_excel_extraction_lab.ipynb`
Runs both methods on this workbook, tab by tab, so you can see what each one does with the same messy input.

### The tests

```bash
pytest                                        # everything
pytest tests/test_excel_to_markdown.py        # the converter itself
pytest tests/test_excel_openpyxl_method.py    # the method switch and options
pytest tests/test_docs_and_notebooks.py       # this document and the notebooks still match the code
```

- **The converter:** layouts, formulas, hidden content, chunking, the messy tab, merged cells (all three kinds), comments (including that they are off by default), Excel-formatted dates and times, General-format date detection (the header rules and the other rules from section 5, custom keywords, year range, 1904 files), reading from bytes, per-sheet text, row and column limits, command-line flags, and robustness (far-away formatting, one bad table, one bad sheet, unreadable files).
- **The method switch:** pandas is still the default and unchanged; unknown methods and misplaced options are errors; the same output shape; the improvements over pandas on the lab's own sample workbooks; comments, formulas, the legend and date detection through the payload; the `.xls` fallback (and turning it off); misnamed `.xlsx`; encrypted files failing the same way with both methods; the 2000 x 200 limits; `extractText` unchanged.
- **Documentation and notebooks:** every example in this document is real output; every payload field, converter option and command-line flag is documented; every tab of the workbook is listed; the saved output file is current; the notebooks are valid and contain no error output.
