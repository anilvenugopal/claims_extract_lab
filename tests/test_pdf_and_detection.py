"""Formatted PDF output, password detection, and file-signature / extension resolution."""
import pytest


# --------------------------------------------------------------------- formatted PDF
def test_formatted_pdf_headings_bold_italic_indent(h, samples):
    pages = h.run("extractFormattedPDFText", samples["loss_summary.pdf"])
    assert len(pages) == 3
    p1, p2, p3 = pages
    assert p1.splitlines()[0] == "# Q1 2026 Loss Summary"
    assert "*Harbor Mutual Insurance (fictional) - valuation date 30 April 2026*" in p1   # italic subtitle
    assert "\n## 1. Executive summary" in p1
    assert "**$80,936.50**" in p1 and "**$92,950.50**" in p1                                # bold spans
    assert p2.startswith("## 2. Large losses")
    assert "### 2.1 CL-10435 - Fire (Brightwater Dental)" in p2                           # heading level 3
    assert "\n    - Paid to date: $36,750.40" in p2                                      # indented bullets
    assert "## 3. Recommendations" in p3 and "## 4. Sign-off" in p3


def test_formatted_pdf_tables_become_one_cell_per_line(h, samples):
    # LIMITATION: table cells are emitted one per line with growing indentation, not as a grid
    p1 = h.run("extractFormattedPDFText", samples["loss_summary.pdf"])[0]
    assert "**Claim ID**\n" in p1 and "\nCL-10421\n" in p1
    assert "CL-10421 " not in p1.replace("CL-10421\n", "")   # row is never kept on one line


def test_formatted_pdf_config_overrides_are_applied(h, samples):
    default = h.run("extractFormattedPDFText", samples["loss_summary.pdf"])[0]
    stricter = h.run("extractFormattedPDFText", samples["loss_summary.pdf"], h1_ratio=3.0)[0]
    assert default.startswith("# Q1 2026 Loss Summary")
    assert stricter.startswith("## Q1 2026 Loss Summary")   # title no longer reaches the H1 band


def test_formatted_pdf_rejects_non_pdf(h, samples):
    with pytest.raises(Exception, match="Must pass in pdf"):
        h.run("extractFormattedPDFText", samples["fnol_form.docx"])


def test_pdf_with_preamble_bytes_is_still_a_pdf(h, samples):
    assert h.run("extractFormattedPDFText", samples["loss_summary_padded.pdf"], ext=None)[0].startswith("# Q1 2026 Loss Summary")


def test_plain_embedded_pdf_text_keeps_table_rows_and_footer(h, samples):
    pages = h.run("extractEmbeddedText", samples["loss_summary.pdf"])
    assert len(pages) == 3 and "Executive summary" in pages[0]
    assert "CL-10421" in pages[0] and "12,500.75" in pages[0]
    assert "Confidential - Harbor Mutual (fictional) - page 2 of 3" in pages[1]   # footer is not filtered
    assert h.run("extractEmbeddedText", samples["fnol_form.docx"]) == []         # not a PDF -> []


def test_extract_text_on_pdf_matches_embedded_text_page_count(h, samples):
    assert len(h.run("extractText", samples["loss_summary.pdf"])) == 3


def test_scanned_pdf_has_no_text_layer(h, samples):
    # image-only page: no OCR in this code base, so the result is blank
    assert [t.strip() for t in h.run("extractEmbeddedText", samples["scanned_receipt.pdf"])] == [""]
    assert [t.strip() for t in h.run("extractText", samples["scanned_receipt.pdf"])] == [""]


# --------------------------------------------------------------------- protection
@pytest.mark.parametrize("ext", ["auto", None], ids=["ext-given", "ext-sniffed"])
@pytest.mark.parametrize(
    "name, expected",
    [
        ("loss_summary.pdf", "NOT_PASSWORD_PROTECTED"),
        ("loss_summary_encrypted.pdf", "PASSWORD_PROTECTED"),
        ("fnol_form.docx", "NOT_PASSWORD_PROTECTED"),
        ("fnol_form_encrypted.docx", "PASSWORD_PROTECTED"),
        ("loss_run.xlsx", "NOT_PASSWORD_PROTECTED"),
        ("loss_run_encrypted.xlsx", "PASSWORD_PROTECTED"),
        ("legacy_register.xls", "NOT_PASSWORD_PROTECTED"),
        ("fnol_no_extension", "NOT_PASSWORD_PROTECTED"),
        ("damage_photo.png", "NOT_APPLICABLE"),
        ("claim_notes.txt", "NOT_APPLICABLE"),
    ],
)
def test_inspect_document_protection(h, samples, name, expected, ext):
    assert h.run("inspectDocumentProtection", samples[name], ext=ext) == expected


def test_protection_of_garbage_office_file_is_unknown(h, tmp_path):
    bad = tmp_path / "broken.docx"
    bad.write_bytes(b"this is not an office file")
    assert h.run("inspectDocumentProtection", bad) == "UNKNOWN"


def test_encrypted_files_cannot_be_extracted(h, samples):
    # protection is only *detected*; nothing decrypts, so extraction fails
    with pytest.raises(Exception):                     # extractText: the Office parser fails outright
        h.run("extractText", samples["fnol_form_encrypted.docx"])
    assert h.run("extractEmbeddedText", samples["loss_summary_encrypted.pdf"]) == []   # swallowed -> empty list


# --------------------------------------------------------------------- signatures
@pytest.mark.parametrize(
    "name, ext",
    [
        ("loss_summary.pdf", ".pdf"),
        ("loss_summary_padded.pdf", ".pdf"),
        ("scanned_receipt.pdf", ".pdf"),
        ("damage_photo.png", ".png"),
        ("fnol_form.docx", ".docx"),
        ("fnol_form.docm", ".docx"),        # macro-enabled Word is also a ZIP with word/ entries
        ("fnol_no_extension", ".docx"),
        ("loss_run.xlsx", ".xlsx"),
        ("loss_run_misnamed.xls", ".xlsx"),
        ("claims_review.pptx", ".pptx"),
        ("claim_package.zip", ".zip"),
    ],
)
def test_signature_sniffing(h, samples, name, ext):
    assert h.app.get_extension_from_file_signature(samples[name].read_bytes()) == ext


@pytest.mark.parametrize(
    "blob, ext",
    [
        (b"\xff\xd8\xff\xe0....", ".jpg"),
        (b"GIF89a....", ".gif"),
        (b"\x1f\x8b\x08....", ".gz"),
        (b"Rar!\x1a\x07\x00", ".rar"),
        (b"7z\xbc\xaf\x27\x1c....", ".7z"),
        (b"just some text", ""),
    ],
)
def test_signature_sniffing_misc(h, blob, ext):
    assert h.app.get_extension_from_file_signature(blob) == ext


def test_plain_zip_is_identified_as_zip_but_has_no_extractor(h, tmp_path):
    import zipfile

    z = tmp_path / "bundle.zip"
    with zipfile.ZipFile(z, "w") as zf:
        zf.writestr("readme.txt", "hello")
    assert h.app.get_extension_from_file_signature(z.read_bytes()) == ".zip"
    with pytest.raises(Exception, match="Invalid File Extension: .zip"):
        h.run("extractText", z)


def test_caller_extension_wins_over_signature(h, samples):
    blob = samples["fnol_form.docx"].read_bytes()
    assert h.app.resolve_extension("PDF", blob) == ".pdf"       # normalised, trusted as given
    assert h.app.resolve_extension(None, blob) == ".docx"       # sniffed when absent
    assert h.app.normalize_extension(" .XLSX ") == ".xlsx"


def test_LIMITATION_every_ole2_file_is_reported_as_doc(h, samples):
    # The OLE2 signature is not disambiguated: legacy .xls and encrypted Office files all come back as '.doc'.
    # (Only matters when no extension is supplied; a supplied extension is trusted.)
    for name in ("legacy_register.xls", "fnol_form_encrypted.docx", "loss_run_encrypted.xlsx"):
        assert h.app.get_extension_from_file_signature(samples[name].read_bytes()) == ".doc"
