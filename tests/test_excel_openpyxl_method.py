"""The openpyxl option of excelExtractText: how it is switched on, what it returns, and how it compares
with the original pandas method on the lab's sample workbooks.

(The converter itself - merged cells, comments, dates, hidden content, ... - is tested in
test_excel_to_markdown.py. These tests cover the wiring: payload options, fallback, limits.)
"""
import io

import openpyxl
import pandas as pd
import pytest


def _text(sheets, name=None):
    return (sheets[0] if name is None else next(s for s in sheets if s["sheet_name"] == name))["text"]


def _line(text, starts_with):
    return next(ln for ln in text.splitlines() if ln.startswith(starts_with))


# ---------------------------------------------------------------- choosing the method
def test_pandas_is_still_the_default(h, samples):
    default = h.run("excelExtractText", samples["claims_workbook.xlsx"])
    explicit = h.run("excelExtractText", samples["claims_workbook.xlsx"], method="pandas")
    assert default == explicit
    assert default[0]["text"].startswith("| 0 | 1 | 2 |")                  # generated column numbers: the old behaviour


def test_unknown_method_is_rejected(h, samples):
    with pytest.raises(ValueError, match="Unknown Excel method 'nope'"):
        h.run("excelExtractText", samples["claims_workbook.xlsx"], method="nope")


def test_openpyxl_options_are_rejected_with_the_pandas_method(h, samples):
    # silently ignoring a requested option would be worse than failing
    with pytest.raises(ValueError, match="only work with method='openpyxl'"):
        h.run("excelExtractText", samples["claims_workbook.xlsx"], include_comments=True)
    with pytest.raises(ValueError, match="only work with method='openpyxl'"):
        h.run("excelExtractText", samples["claims_workbook.xlsx"], method="pandas", detect_general_dates=True)


def test_same_output_shape_as_the_pandas_method(h, samples):
    old = h.run("excelExtractText", samples["claims_workbook.xlsx"], method="pandas")
    new = h.run("excelExtractText", samples["claims_workbook.xlsx"], method="openpyxl")
    assert [s["sheet_name"] for s in new] == [s["sheet_name"] for s in old]
    assert all(set(s) == {"sheet_name", "text"} and isinstance(s["text"], str) for s in new)


def test_non_excel_files_are_rejected_by_both_methods(h, samples):
    for method in ("pandas", "openpyxl"):
        with pytest.raises(Exception, match="Invalid File Extension"):
            h.run("excelExtractText", samples["fnol_form.docx"], method=method)


# ---------------------------------------------------------------- what the openpyxl method fixes
def test_real_headers_instead_of_generated_column_numbers(h, samples):
    text = _text(h.run("excelExtractText", samples["loss_run.xlsx"], method="openpyxl"))
    assert "| 5 | Claim ID | Insured | Date of loss | Cause | Status | Paid | Reserve | Loss ratio |" in text
    assert "| 0 | 1 | 2 |" not in text
    assert "Table 3 — 'Loss Run'!B5:I14" in text                            # cell coordinates for citations


def test_side_by_side_tables_are_kept_apart(h, samples):
    old = _text(h.run("excelExtractText", samples["loss_run.xlsx"], method="pandas"))
    new = _text(h.run("excelExtractText", samples["loss_run.xlsx"], method="openpyxl"))
    # pandas: the coverage table (limit 500000) is glued onto the same row as claim CL-10421
    glued = [ln for ln in old.splitlines() if "CL-10421" in ln and "Property" in ln]
    assert glued and "500000" in glued[0]
    # openpyxl: separate blocks
    assert "Table 2 — 'Loss Run'!K4:M8" in new
    assert "500000" not in _line(new, "| 6 | CL-10421")
    assert "| 6 | Property | 500000 | 2500 |" in new


def test_merged_cell_is_reported_not_lost(h, samples):
    new = _text(h.run("excelExtractText", samples["loss_run.xlsx"], method="openpyxl"))
    assert 'B19:B21 ("CL-10421")' in new                                    # tells the reader the claim covers rows 19-21
    assert _line(new, "| 20 |").startswith("| 20 |  | Expense")             # value only in the first cell, as in Excel


def test_hidden_sheet_is_read_and_labelled(h, samples):
    sheets = h.run("excelExtractText", samples["claims_workbook.xlsx"], method="openpyxl")
    hidden = _text(sheets, "Internal Notes")
    assert "## Sheet 'Internal Notes' — HIDDEN SHEET" in hidden
    assert "Reserve rationale (internal - not for distribution)" in hidden


def test_hidden_sheets_can_be_left_out(h, samples):
    sheets = h.run("excelExtractText", samples["claims_workbook.xlsx"], method="openpyxl", include_hidden_sheets=False)
    assert [s["sheet_name"] for s in sheets] == ["Claims", "Payments", "Reserves"]


def test_cached_values_come_out_with_full_precision(h, samples):
    new = _text(h.run("excelExtractText", samples["loss_run.xlsx"], method="openpyxl"))
    assert "0.6234" in new                       # the pandas method rounds this to 0.62
    assert "| 14 | Total |  |  |  |  | 80936.5 | 92950.5 |  |" in new


# ---------------------------------------------------------------- opt-in extras
def test_comments_are_opt_in(h, samples):
    off = _text(h.run("excelExtractText", samples["loss_run.xlsx"], method="openpyxl"))
    on = _text(h.run("excelExtractText", samples["loss_run.xlsx"], method="openpyxl", include_comments=True))
    assert "salvage recovery" not in off
    assert "Comment on G7: Includes salvage recovery of $1,200.00" in on   # placeholder author "None" is not shown


def test_formulas_are_opt_in(h, samples):
    off = _text(h.run("excelExtractText", samples["loss_run.xlsx"], method="openpyxl"))
    on = _text(h.run("excelExtractText", samples["loss_run.xlsx"], method="openpyxl", include_formulas=True))
    assert "=SUM" not in off
    assert "Formula in G14: `=SUM(G6:G13)`" in on


def test_legend_can_be_switched_off(h, samples):
    with_legend = _text(h.run("excelExtractText", samples["loss_run.xlsx"], method="openpyxl"))
    without = _text(h.run("excelExtractText", samples["loss_run.xlsx"], method="openpyxl", include_legend=False))
    assert with_legend.startswith("> **How to read this file.**")
    assert without.startswith("## Sheet 'Loss Run'")


def test_general_date_detection_through_the_payload(h, tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Claims"
    ws.append(["Claim ID", "Date of Loss", "Amount Due", "Adjuster Visit"])
    ws.append([46300, 46300, 46300, 46301])
    ws.append([46301, 46330, 46301, 46331])
    path = tmp_path / "general_dates.xlsx"
    wb.save(path)

    off = _text(h.run("excelExtractText", path, method="openpyxl"))
    assert "(serial" not in off and "| 2 | 46300 | 46300 | 46300 | 46301 |" in off

    on = _text(h.run("excelExtractText", path, method="openpyxl", detect_general_dates=True))
    assert "| 2 | 46300 | 2026-10-05 (serial 46300) | 46300 | 46301 |" in on          # only "Date of Loss" is date-like
    assert "Adjuster Visit" in on and "2026-10-06 (serial 46301)" not in on           # "Visit" is not a date word

    named = _text(h.run("excelExtractText", path, method="openpyxl", detect_general_dates=True, date_keywords=["visit"]))
    assert "| 2 | 46300 | 46300 | 46300 | 2026-10-06 (serial 46301) |" in named        # keywords replace the built-in list


# ---------------------------------------------------------------- file types
def test_legacy_xls_falls_back_to_pandas(h, samples):
    new = h.run("excelExtractText", samples["legacy_register.xls"], method="openpyxl")
    old = h.run("excelExtractText", samples["legacy_register.xls"], method="pandas")
    assert new == old


def test_fallback_can_be_turned_off(h, samples):
    with pytest.raises(ValueError, match="cannot read this file type"):
        h.run("excelExtractText", samples["legacy_register.xls"], method="openpyxl", fallback_to_pandas=False)


def test_misnamed_xlsx_is_read_by_signature(h, samples):
    # OOXML content with a .xls name: the file signature decides, so openpyxl can read it
    sheets = h.run("excelExtractText", samples["loss_run_misnamed.xls"], method="openpyxl")
    assert sheets[0]["sheet_name"] == "Loss Run"
    assert "| 5 | Claim ID | Insured |" in sheets[0]["text"]


def test_encrypted_workbooks_fail_the_same_way_with_both_methods(h, samples):
    for method in ("pandas", "openpyxl"):
        with pytest.raises(Exception):
            h.run("excelExtractText", samples["loss_run_encrypted.xlsx"], method=method)


# ---------------------------------------------------------------- limits and the extractor class itself
def test_row_limit_matches_the_pandas_cap(h, tmp_path):
    n = h.text_extract.MAX_ROWS + 100
    buf = io.BytesIO()
    pd.DataFrame([[f"r{i}c{j}" for j in range(3)] for i in range(n)]).to_excel(buf, index=False, header=False)
    path = tmp_path / "tall.xlsx"
    path.write_bytes(buf.getvalue())
    text = _text(h.run("excelExtractText", path, method="openpyxl"))
    assert f"r{h.text_extract.MAX_ROWS - 1}c0" in text and f"r{h.text_extract.MAX_ROWS}c0" not in text
    assert f"Only the first {h.text_extract.MAX_ROWS} rows" in text            # and the reader is told


def test_column_limit_matches_the_pandas_cap(h, tmp_path):
    n = h.text_extract.MAX_COLUMNS + 50
    buf = io.BytesIO()
    pd.DataFrame([[f"r{i}c{j}" for j in range(n)] for i in range(3)]).to_excel(buf, index=False, header=False)
    path = tmp_path / "wide.xlsx"
    path.write_bytes(buf.getvalue())
    text = _text(h.run("excelExtractText", path, method="openpyxl"))
    assert f"r0c{h.text_extract.MAX_COLUMNS - 1}" in text and f"r0c{h.text_extract.MAX_COLUMNS}" not in text
    assert f"Only the first {h.text_extract.MAX_COLUMNS} columns" in text


def test_extract_returns_one_string_per_sheet_without_the_one_sheet_limit(h, samples):
    doc = h.text_extract.DocumentContent(samples["claims_workbook.xlsx"].read_bytes())
    with pytest.raises(ValueError, match="has 4 sheets"):                   # pandas extract(): exactly one sheet only
        h.text_extract.ExcelExtractor().extract(".xlsx", doc)
    texts = h.text_extract.OpenpyxlExcelExtractor().extract(".xlsx", doc)   # openpyxl extract(): all sheets
    assert len(texts) == 4 and all(isinstance(t, str) for t in texts)


def test_extract_text_for_xlsx_is_unchanged(h, samples):
    # extractText keeps using the pandas extractor: the new method is only reachable through excelExtractText
    assert isinstance(h.text_extract.EXTENSION_TO_EXTRACTOR[".xlsx"], h.text_extract.ExcelExtractor)
    (table,) = h.run("extractText", samples["loss_run.xlsx"])
    assert table.splitlines()[0].replace(" ", "").startswith("|0|1|2|3|")
