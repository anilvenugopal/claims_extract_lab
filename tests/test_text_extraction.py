"""extractText across formats + markdown_utils + the extension registry."""
import pytest


def test_supported_extensions_match_registry(h):
    exts = h.call("getSupportedFileExtensions")
    assert exts == sorted(h.text_extract.EXTENSION_TO_EXTRACTOR)
    for e in (".pdf", ".docx", ".docm", ".xlsx", ".xls", ".xlsm", ".xlsb", ".pptx", ".csv", ".rtf", ".md"):
        assert e in exts
    # archives / images / email are NOT extractable through extractText
    for e in (".zip", ".gz", ".png", ".eml", ".msg"):
        assert e not in exts


@pytest.mark.parametrize("name", ["claim_notes.txt", "claim_report.html", "claim_feed.xml"])
def test_plain_text_formats_pass_through(h, samples, name):
    # html / xml are NOT parsed: markup is returned verbatim
    (text,) = h.run("extractText", samples[name])
    assert text == samples[name].read_text(encoding="utf-8")


def test_html_markup_and_comments_are_kept(h, samples):
    (text,) = h.run("extractText", samples["claim_report.html"])
    assert "<table" in text and "<!-- internal: reserve reviewed" in text


def test_markdown_strips_base64_and_masks_comments(h, samples):
    original = samples["adjuster_notes.md"].read_text(encoding="utf-8")
    (text,) = h.run("extractText", samples["adjuster_notes.md"])
    assert "base64" not in text                                      # embedded signature image removed
    assert "![scene](https://example.com/photos/CL-10447.png)" in text   # ordinary image kept
    assert "internal" not in text.lower()                            # HTML comment masked ...
    assert "| 2026-03-27 | Indemnity | 14,000.00 |" in text          # tables untouched
    start = original.index("<!--")
    end = original.index("-->") + 3
    # ... with spaces, so everything before the comment keeps its exact offsets
    assert text[:start] == original[:start]
    assert text[start:end].replace("\n", "").strip() == ""
    assert text[start:end].count("\n") == original[start:end].count("\n")   # line boundaries kept
    assert "Signature of adjuster: \n" in text


def test_docx_merged_cells_lists_and_tables(h, samples):
    pages = h.run("extractText", samples["fnol_form.docx"])
    assert len(pages) == 2                                           # page break -> second page
    p = pages[0]
    assert "Policy number | PL-77310" in p                           # table cells joined with ' | '
    assert "Coverage | Limit | Deductible | Applies" in p
    assert "Property | 500,000 | 2,500 | Yes" in p
    assert "Business interruption | 250,000 | 10,000 | No" in p
    assert "Photographs (12)" in p                                   # bullet list text survives
    assert "Coverage summary - amounts in USD" in p                  # merged header cell appears once
    assert p.count("Coverage summary - amounts in USD") == 1
    assert pages[1].startswith("4. Payment history")                 # content after the page break
    assert "Signed: M. Okafor, Senior Adjuster" in pages[1]


def test_docm_uses_patched_loader(h, samples):
    assert h.run("extractText", samples["fnol_form.docm"]) == h.run("extractText", samples["fnol_form.docx"])


def test_pptx_text_tables_and_notes(h, samples):
    slides = h.run("extractText", samples["claims_review.pptx"])
    assert len(slides) == 3
    assert slides[0].startswith("Q1 2026 Claims Review")
    assert "Status | Count | Paid (USD)" in slides[1] and "Open | 4 | 43,260.65" in slides[1]
    assert "Re-reserve CL-10447 after supplemental invoice" in slides[2]
    assert "Key risk: reserve adequacy" in slides[2]                 # grouped shapes are read too


def test_rtf(h, samples):
    (text,) = h.run("extractText", samples["settlement_letter.rtf"])
    assert "Claim CL-10421" in text and "$12,500.75" in text
    assert "\\par" not in text and "\\b" not in text                # control words stripped
    assert "Dear Alder Logistics, LLC," in text


def test_csv_is_rendered_as_markdown_table(h, samples):
    (table,) = h.run("extractText", samples["claims_register.csv"])
    assert table.startswith("| 0 ")                                 # generated column numbers, not business headers
    assert "Claim ID" in table and "CL-10463" in table
    assert "12,500.75" in table                                     # quoted thousands separator survives


def test_unsupported_extension_raises(h, samples):
    with pytest.raises(Exception, match="Invalid File Extension: .png"):
        h.run("extractText", samples["damage_photo.png"])


def test_zip_is_recognised_but_not_extractable(h, samples):
    assert h.app.get_extension_from_file_signature(samples["claim_package.zip"].read_bytes()) == ".zip"
    with pytest.raises(Exception, match="Invalid File Extension: .zip"):
        h.run("extractText", samples["claim_package.zip"])


# ---- markdown_utils ---------------------------------------------------------------------
def test_remove_base64_helper(h):
    rm = h.markdown_utils.removeMarkdownEmbeddedBase64
    assert rm(None) == "" and rm("") == ""
    plain = "text ![a](https://x/y.png) more"
    assert rm(plain) == plain                                # fast path: untouched
    assert rm("a ![x](data:image/png;base64,AAAA) b") == "a  b"
    assert rm("![nested [brackets]](data:image/jpeg;base64,QUJD)") == ""


def test_remove_base64_no_catastrophic_backtracking(h):
    import time

    nasty = "![" + "[" * 2000 + "](" + "x" * 5000
    start = time.time()
    h.markdown_utils.removeMarkdownEmbeddedBase64(nasty)
    assert time.time() - start < 2
