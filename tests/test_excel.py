"""Excel behaviour of the CURRENT parser - including its known limitations.

The limitation tests document today's behaviour so that a new region-aware parser can be
compared against them (generated headers, lost formula text, flattened side-by-side tables, ...).
"""
import io
import zipfile

import openpyxl
import pandas as pd
import pytest


def _rows(table):
    return [[c.strip() for c in ln.strip("|").split("|")] for ln in table.splitlines()[2:]]


def test_complex_single_sheet_loss_run(h, samples):
    (table,) = h.run("extractText", samples["loss_run.xlsx"])
    assert "Loss Run - Harbor Mutual Insurance (fictional) - amounts in USD" in table   # title cell at B2
    assert "2026-04-30 00:00:00" in table                          # real dates become timestamps
    assert "0.62" in table and "0.6234" not in table               # loss ratio rounded to 2 dp
    assert "12500.75" in table
    assert "Reserve figures as of the valuation date" in table     # footnote
    # cached formula results are read (xlsxwriter stored them)
    total = [r for r in _rows(table) if "Total" in r][0]
    assert "80936.5" in total and "92950.5" in total


def test_side_by_side_tables_are_flattened_into_one_grid(h, samples):
    (table,) = h.run("extractText", samples["loss_run.xlsx"])
    row = [ln for ln in table.splitlines() if "CL-10421" in ln and "Property" in ln]
    assert row, "coverage table shares rows with the claims table"
    assert "500000" in row[0] and "2500" in row[0]


def test_stacked_second_table_and_merged_cell(h, samples):
    (table,) = h.run("extractText", samples["loss_run.xlsx"])
    assert "Payment history - CL-10421" in table
    rows = _rows(table)
    pay = [r for r in rows if "Restoration vendor" in r][0]
    assert "CL-10421" not in pay             # merged claim cell only has a value in its first row
    assert "1800.25" in pay


def test_textbox_chart_and_comment_are_not_extracted(h, samples):
    # the sheet really holds all three ...
    wb = openpyxl.load_workbook(samples["loss_run.xlsx"])
    assert any(c.comment for row in wb.active.iter_rows() for c in row)
    with zipfile.ZipFile(samples["loss_run.xlsx"]) as z:
        names = z.namelist()
        assert any(n.startswith("xl/charts/chart") for n in names)
        assert "Reviewer note: CL-10447" in z.read("xl/drawings/drawing1.xml").decode()
    # ... but only cell values reach the output
    (table,) = h.run("extractText", samples["loss_run.xlsx"])
    assert "Reviewer note" not in table
    assert "salvage recovery" not in table


def test_extract_text_requires_exactly_one_sheet(h, samples):
    with pytest.raises(ValueError, match="has 4 sheets"):
        h.run("extractText", samples["claims_workbook.xlsx"])


def test_excel_extract_text_returns_one_entry_per_sheet(h, samples):
    sheets = h.run("excelExtractText", samples["claims_workbook.xlsx"])
    assert [s["sheet_name"] for s in sheets] == ["Claims", "Payments", "Reserves", "Internal Notes"]
    claims, payments, reserves, _ = (s["text"] for s in sheets)
    assert "CL-10435" in claims and "36750.4" in claims
    assert "Independent adjuster fee" in payments and "700.5" in payments
    assert "Reserve development by month" in reserves and "92950.5" in reserves


def test_hidden_sheet_is_still_extracted(h, samples):
    wb = openpyxl.load_workbook(samples["claims_workbook.xlsx"])
    assert wb["Internal Notes"].sheet_state == "hidden"
    sheets = {s["sheet_name"]: s["text"] for s in h.run("excelExtractText", samples["claims_workbook.xlsx"])}
    assert "Reserve rationale (internal - not for distribution)" in sheets["Internal Notes"]


def test_whitespace_is_collapsed(h, samples):
    for s in h.run("excelExtractText", samples["claims_workbook.xlsx"]):
        body = "\n".join(s["text"].splitlines()[2:])
        assert "  " not in body.replace("| ", "").replace(" |", "")


def test_excel_extract_text_rejects_non_excel(h, samples):
    with pytest.raises(Exception, match="Invalid File Extension"):
        h.run("excelExtractText", samples["fnol_form.docx"])


def test_legacy_xls(h, samples):
    (table,) = h.run("extractText", samples["legacy_register.xls"])
    assert "Legacy claims register (1997-style export)" in table
    assert "Claim ID" in table and "12500.75" in table and "2026-01-14 00:00:00" in table


def test_misnamed_xlsx_is_read_by_signature(h, samples):
    # content is OOXML, name says .xls -> resolve_excel_engine picks openpyxl from the ZIP signature
    sheets = h.run("excelExtractText", samples["loss_run_misnamed.xls"])
    assert sheets[0]["sheet_name"] == "Loss Run"
    assert "CL-10463" in sheets[0]["text"]


def test_engine_resolution(h, samples):
    resolve = h.text_extract.resolve_excel_engine
    assert resolve(".xlsx", samples["legacy_register.xls"].read_bytes()) == "xlrd"       # OLE2 beats extension
    assert resolve(".xls", samples["loss_run.xlsx"].read_bytes()) == "openpyxl"          # ZIP beats extension
    assert resolve(".xlsb", b"unknown bytes") == "pyxlsb"                                # fall back to extension
    assert resolve(".xls", b"unknown bytes") == "xlrd"
    assert resolve("", b"unknown bytes") == "openpyxl"


# ---------------------------------------------------------------- known limitations (documented)
def test_LIMITATION_headers_are_generated_column_numbers(h, samples):
    (table,) = h.run("extractText", samples["claims_register.csv"])
    header_row = table.splitlines()[0]
    assert header_row.replace(" ", "").startswith("|0|1|2|3|")   # business headers become a data row
    (table,) = h.run("extractText", samples["loss_run.xlsx"])
    assert table.splitlines()[0].replace(" ", "").startswith("|0|1|2|3|")


def test_LIMITATION_formula_without_cached_value_is_blank(h, tmp_path):
    # files written by openpyxl (or never opened in Excel) hold formulas but no cached results
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Paid", 100])
    ws.append(["Paid", 250])
    ws.append(["Total", "=SUM(B1:B2)"])
    path = tmp_path / "no_cache.xlsx"
    wb.save(path)
    (table,) = h.run("extractText", path)
    total_row = [ln for ln in table.splitlines() if ln.startswith("| Total")][0]
    assert "SUM" not in total_row                                # formula text is not preserved
    assert total_row.replace("|", "").split() == ["Total"]       # and no value either


def test_LIMITATION_row_and_column_caps(h):
    wide = pd.DataFrame([[1] * (h.text_extract.MAX_COLUMNS + 50)] * 3)
    buf = io.BytesIO()
    wide.to_excel(buf, index=False, header=False)
    resp = h.text_extract.DocumentContent(buf.getvalue())
    (table,) = h.text_extract.ExcelExtractor().extract(".xlsx", resp)
    first_row = table.splitlines()[0]
    assert first_row.count("|") - 1 == h.text_extract.MAX_COLUMNS   # silently truncated to 200 columns
