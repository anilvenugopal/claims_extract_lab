# Document text extraction functions.
# - Each function takes a payload dataclass whose `path` points at the file to read.
# - Each function returns something serializable by `json.dumps`.

from dataclasses import dataclass
from pathlib import Path
from typing import Optional
import logging
import zipfile
import fitz
import msoffcrypto
import io
from text_extract import DocumentContent, ExcelExtractor, ExcelSheetResponse, PdfExtractor, EXTENSION_TO_EXTRACTOR, FormattedPDFTextExtractor
from formatted_text_extraction import ExtractionConfig

# Settings
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
logger.setLevel(level="WARN")


@dataclass
class Payload:
    path: str
    file_extension: Optional[str] = None


## see formatted_text_extraction.py for param definitions
@dataclass
class FormattedPDFPayload:
    path: str
    h1_ratio: Optional[float] = None
    h2_ratio: Optional[float] = None
    h3_ratio: Optional[float] = None
    indent_step: Optional[float] = None
    indent_dead_zone: Optional[float] = None
    gap_threshold: Optional[float] = None
    fallback_char_width: Optional[float] = None
    fallback_body_size: Optional[float] = None


# Reconciles incoming file extension (e.g. 'pdf') with derived extensions (e.g. '.pdf')
def normalize_extension(file_extension: Optional[str]) -> str:
    ext = (file_extension or "").strip().lower().lstrip(".")
    return f".{ext}" if ext else ""


# Prefers the extension the caller passed in and only falls back to sniffing the file signature if unavailable
def resolve_extension(file_extension: Optional[str], content: bytes) -> str:
    ext = normalize_extension(file_extension)
    if ext:
        return ext

    return get_extension_from_file_signature(content)


def extractFormattedPDFText(payload: FormattedPDFPayload) -> list[str]:
    document, file_extension = load_doc(payload.path)

    if file_extension != ".pdf":
        raise Exception(f"Invalid File Extension={file_extension}. Must pass in pdf.")

    # only pass through fields that were explicitly set — ExtractionConfig owns the defaults
    overrides = {
        k: v for k, v in {
            "h1_ratio": payload.h1_ratio,
            "h2_ratio": payload.h2_ratio,
            "h3_ratio": payload.h3_ratio,
            "indent_step": payload.indent_step,
            "indent_dead_zone": payload.indent_dead_zone,
            "gap_threshold": payload.gap_threshold,
            "fallback_char_width": payload.fallback_char_width,
            "fallback_body_size": payload.fallback_body_size,
        }.items() if v is not None
    }
    config = ExtractionConfig(**overrides)

    return FormattedPDFTextExtractor().extract(file_extension, document, config)

# pdf header not guaranteed at offset 0 — spec permits preceding bytes, and MuPDF's normal
# open path locates the %PDF- version marker by scanning only the first ~1KB of the stream
# (same tolerance Acrobat documents). match that window so anything fitz can open gets
# classified as pdf.
PDF_HEADER_SEARCH_WINDOW = 1024


def get_extension_from_file_signature(content: bytes) -> str:
    FILE_SIGNATURES = [
        # Documents
        (b"%PDF", ".pdf"),

        # Images
        (b"\x89PNG\r\n\x1a\n", ".png"),
        (b"\xff\xd8\xff", ".jpg"),
        (b"GIF87a", ".gif"),
        (b"GIF89a", ".gif"),
        (b"BM", ".bmp"),
        (b"II\x2a\x00", ".tiff"),  # Little-endian TIFF
        (b"MM\x00\x2a", ".tiff"),  # Big-endian TIFF
        (b"RIFF", ".webp"),

        # Archives
        (b"\x1f\x8b", ".gz"),
        (b"Rar!\x1a\x07", ".rar"),
        (b"7z\xbc\xaf\x27\x1c", ".7z"),

        # OLE2 compound (.doc, .xls, .ppt) - not disambiguated further; reported as .doc
        (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", ".doc"),

        # ZIP-based
        # (includes .docx, .xlsx, and .pptx)
        (b"PK\x03\x04", ".zip"),
    ]

    for signature, extension in FILE_SIGNATURES:
        if content[:len(signature)] == signature:
            # Disambiguate ZIP-based Office formats
            if extension == ".zip":
                try:
                    with zipfile.ZipFile(io.BytesIO(content)) as zf:
                        names = zf.namelist()
                        if any(name.startswith("word/") for name in names):
                            return ".docx"
                        elif any(name.startswith("xl/") for name in names):
                            return ".xlsx"
                        elif any(name.startswith("ppt/") for name in names):
                            return ".pptx"
                except zipfile.BadZipFile:
                    pass

            return extension

    # fallback for pdfs carrying bytes before the header (mail preamble, BOM, scanner padding).
    # runs after strict matching so zip/ole containers that embed a pdf still classify as
    # zip/ole, not pdf.
    if content[:PDF_HEADER_SEARCH_WINDOW].find(b"%PDF") != -1:
        return ".pdf"

    return ""


def load_doc(path: str, file_extension: Optional[str] = None) -> tuple[DocumentContent, str]:
    """Read a local file; returns its content and the resolved extension."""
    file_path = Path(path)
    if not file_path.is_file():
        raise FileNotFoundError(f"No such file: {path}")

    document = DocumentContent(file_path.read_bytes())
    file_extension = resolve_extension(file_extension, document.content)
    logger.info('Extension: ' + file_extension)

    return document, file_extension


OFFICE_EXTENSIONS = {
    ".doc",
    ".docx",
    ".docm",
    ".xls",
    ".xlsx",
    ".xlsm",
    ".xlsb",
    ".ppt",
    ".pptx",
}


def inspectDocumentProtection(
    payload: Payload,
) -> str:
    """
    Returns:
        PASSWORD_PROTECTED
        NOT_PASSWORD_PROTECTED
        UNKNOWN
        NOT_APPLICABLE
    """

    document, file_extension = load_doc(
        payload.path,
        payload.file_extension,
    )

    content = document.content
    file_extension = file_extension.lower()

    # PDF detection
    if file_extension == ".pdf":
        pdf_document = None

        try:
            pdf_document = fitz.open(
                stream=content,
                filetype="pdf",
            )

            if pdf_document.needs_pass:
                return "PASSWORD_PROTECTED"

            return "NOT_PASSWORD_PROTECTED"

        except Exception:
            return "UNKNOWN"

        finally:
            if pdf_document is not None:
                pdf_document.close()

    # Microsoft Office detection
    if file_extension in OFFICE_EXTENSIONS:
        try:
            office_document = msoffcrypto.OfficeFile(
                io.BytesIO(content)
            )

            if office_document.is_encrypted():
                return "PASSWORD_PROTECTED"

            return "NOT_PASSWORD_PROTECTED"

        except Exception:
            return "UNKNOWN"

    return "NOT_APPLICABLE"


def extractText(payload: Payload) -> list[str]:
    document, file_extension = load_doc(payload.path, payload.file_extension)

    if file_extension not in EXTENSION_TO_EXTRACTOR:
        error_msg = 'Invalid File Extension: ' + (file_extension or '<EMPTY FILE EXTENSION>')
        logger.error(error_msg)
        raise Exception(error_msg)

    extractor = EXTENSION_TO_EXTRACTOR[file_extension]
    return extractor.extract(file_extension, document)


def extractEmbeddedText(payload: Payload) -> list[str]:
    """
    Given a file path, read the PDF, load it with pymupdf (fitz), pass the doc to foo(), and return the list of pages as a string per page.
    """
    document, file_extension = load_doc(payload.path, payload.file_extension)

    return extractEmbeddedPdfText_helper(file_extension, document, payload)


def extractEmbeddedPdfText_helper(file_extension, document, payload):
    if file_extension != '.pdf':
        logger.error('File is not a PDF: ' + file_extension)
        return []
    try:
        pages = PdfExtractor().extract(file_extension, document)
        logger.info(f"Parsed {len(pages)} pages from PDF ({file_extension}) path={payload.path}")
        return pages
    except Exception as _:
        logger.error(f"Could not load PDF: ({file_extension}) path={payload.path}", exc_info=True)
        return []


def getSupportedFileExtensions(payload) -> list[str]:
    """
    Returns the list of all file extensions that can be extracted text from.
    """
    return sorted(EXTENSION_TO_EXTRACTOR.keys())


def excelExtractText(payload: Payload) -> list[ExcelSheetResponse]:
    document, file_extension = load_doc(payload.path, payload.file_extension)

    if file_extension not in ['.xls', '.xlsm', '.xlsb', '.xlsx']:
        error_msg = f"Invalid File Extension={file_extension}"
        logger.error(error_msg)
        raise Exception(error_msg)

    return ExcelExtractor().extractBySheet(file_extension, document)
