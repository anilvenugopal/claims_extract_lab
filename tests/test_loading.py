"""Loading a document from disk: extension resolution and error handling."""
import pytest


def test_missing_file_raises(h, tmp_path):
    with pytest.raises(FileNotFoundError, match="No such file"):
        h.call("extractText", {"path": str(tmp_path / "nope.txt"), "file_extension": ".txt"})


def test_extension_sniffed_when_not_supplied(h, samples):
    # FormattedPDFPayload has no extension field, so the signature is sniffed
    pages = h.call("extractFormattedPDFText", {"path": str(samples["loss_summary.pdf"])})
    assert pages[0].startswith("# Q1 2026 Loss Summary")


def test_supplied_extension_wins(h, tmp_path):
    f = tmp_path / "notes.bin"
    f.write_bytes(b"hello")
    assert h.call("extractText", {"path": str(f), "file_extension": "txt"}) == ["hello"]
