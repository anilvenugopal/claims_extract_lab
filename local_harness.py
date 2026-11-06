"""
Local test harness for the Claims Assist extraction code.

Exercises the files in ``src/`` (app.py, text_extract.py, ...) directly:

* :func:`call` dispatches a function by name with a JSON-style payload dict.
* :func:`run` is a convenience wrapper that takes a file path (or a name under ``samples/``).

Everything - extension resolution, extractors, protection checks, PDF formatting - is the real code.
"""
from __future__ import annotations

import sys
import warnings
from pathlib import Path
from typing import Any, Callable, Dict, Optional

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
SAMPLES = ROOT / "samples"

warnings.filterwarnings("ignore", message=".*fitz.*deprecated.*")
warnings.filterwarnings("ignore", category=FutureWarning, module="fitz")


if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

import app  # noqa: E402
import text_extract  # noqa: E402
import formatted_text_extraction  # noqa: E402
import markdown_utils  # noqa: E402


# --------------------------------------------------------------------------- dispatch
FUNCTIONS: Dict[str, Callable] = {
    fn.__name__: fn
    for fn in (
        app.extractText,
        app.extractFormattedPDFText,
        app.excelExtractText,
        app.extractEmbeddedText,
        app.getSupportedFileExtensions,
        app.inspectDocumentProtection,
    )
}

PAYLOAD_TYPES: Dict[str, Optional[type]] = {
    "extractText": app.Payload,
    "extractFormattedPDFText": app.FormattedPDFPayload,
    "excelExtractText": app.Payload,
    "extractEmbeddedText": app.Payload,
    "inspectDocumentProtection": app.Payload,
    "getSupportedFileExtensions": None,
}


def call(query_type: str, payload: Optional[dict] = None) -> Any:
    """Invoke an extraction function by name with a JSON-style payload dict."""
    fn = FUNCTIONS[query_type]
    payload_cls = PAYLOAD_TYPES[query_type]
    arg = payload_cls(**(payload or {})) if payload_cls else (payload or {})
    return fn(arg)


def run(query_type: str, path: str | Path, ext: Optional[str] = "auto", **extra: Any) -> Any:
    """Convenience wrapper: ``run('extractText', 'samples/fnol_form.docx')``.

    ext="auto"  -> use the file's own suffix as ``file_extension``
    ext=None    -> pass no extension, forcing the file-signature sniffing path
    ext=".xls"  -> pass an explicit (possibly wrong!) extension
    Extra keyword arguments are payload fields (e.g. ``h1_ratio=1.5`` for extractFormattedPDFText).
    """
    path = Path(path)
    if not path.is_file() and (SAMPLES / path).is_file():
        path = SAMPLES / path
    payload: Dict[str, Any] = {"path": str(path), **extra}
    if PAYLOAD_TYPES[query_type] is app.Payload:
        payload["file_extension"] = Path(path).suffix if ext == "auto" else ext
    return call(query_type, payload)


def available_functions() -> list[str]:
    return sorted(FUNCTIONS)
