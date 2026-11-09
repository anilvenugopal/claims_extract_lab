from abc import ABC, abstractmethod
import io
from docx import Document
from docm import DOCMDocument
import pandas as pd
import re
from typing import TypedDict
import zipfile
import logging
import chardet
from striprtf.striprtf import rtf_to_text
import fitz
from formatted_text_extraction import _analyze_document, _render_page, ExtractionConfig
from markdown_utils import removeMarkdownEmbeddedBase64
from excel_to_markdown import parse_workbook


HTML_COMMENT_PATTERN = re.compile(r"<!--[\s\S]*?-->")

logger = logging.getLogger(__name__)


def maskMarkdownHtmlCommentsPreservingOffsets(text: str) -> str:
    """Mask HTML comments without changing text offsets or line boundaries."""
    if text is None:
        return ""

    def mask_comment(match: re.Match) -> str:
        return re.sub(r"[^\r\n]", " ", match.group(0))

    return HTML_COMMENT_PATTERN.sub(mask_comment, text)


class DocumentContent:
    """A loaded document: raw bytes, plus the text view the plain-text extractor uses."""

    def __init__(self, content: bytes):
        self.content = content

    @property
    def text(self) -> str:
        return self.content.decode("utf-8", errors="replace")


class ExcelSheetResponse(TypedDict):
    sheet_name: str
    text: str


# Interface
class TextExtractor(ABC):
    @abstractmethod
    def extract(self, file_extension: str, response: DocumentContent) -> list[str]:
        pass


class PlainTextExtractor(TextExtractor):
    def extract(self, file_extension: str, response: DocumentContent) -> list[str]:
        text = removeMarkdownEmbeddedBase64(response.text)

        if file_extension.lower().endswith('.md'):
            text = maskMarkdownHtmlCommentsPreservingOffsets(text)

        return [text]


class DocxExtractor(TextExtractor):
    DOC_TO_USE = staticmethod(Document)

    def extract(self, file_extension: str, response: DocumentContent) -> list[str]:  # noqa: C901
        byte_stream = io.BytesIO(response.content)
        document = self.DOC_TO_USE(byte_stream)
        pages = []
        current_page_text = []
        last_cell_in_table = None
        for element in document.element.body.iter():
            # move to last cell in table
            if last_cell_in_table is not None:
                if element == last_cell_in_table:
                    last_cell_in_table = None
                continue

            # Check for page breaks
            if element.tag.endswith('}br') and element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == 'page':
                if current_page_text:
                    pages.append(' '.join(current_page_text))
                    current_page_text = []

            # Handle tables
            elif element.tag.endswith('}tbl'):
                # print('a')
                # print(element.text)
                table_texts = []
                # Process rows in table
                for row in element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr'):
                    row_text = []
                    # Process cells in row
                    for cell in row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'):
                        # Get all text elements within the cell
                        cell_text = []
                        for t in cell.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                            if t.text:
                                cell_text.append(t.text)
                                last_cell_in_table = t
                        row_text.append(' '.join(cell_text))
                    table_texts.append(" | ".join(row_text))

                # dont add text from the same table # of cells in table times
                if len(current_page_text) == 0 or current_page_text[-1] != table_texts:
                    # print(table_texts)
                    current_page_text.extend(table_texts)

            # Handle text content
            elif element.tag.endswith('}t') and element.text:
                current_page_text.append(element.text)

        # Add the last page
        if current_page_text:
            pages.append(' '.join(current_page_text))
        return pages


class DocmExtractor(DocxExtractor):
    DOC_TO_USE = staticmethod(DOCMDocument)


MAX_ROWS = 2000
MAX_COLUMNS = 200
CHAR_LIMIT = 5_000_000
OLE2_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_SIGNATURE = b"PK\x03\x04"


def resolve_excel_engine(file_extension: str, content: bytes) -> str:
    """
    Attachments are frequently mis-named (e.g. an .xlsx exported with a '.xls' name).
    Pick the pandas read_excel engine from the file's signature.
    Falls back to the extension only if the signature is unrecognised.
    """
    if content.startswith(OLE2_SIGNATURE):
        return "xlrd"

    if content.startswith(ZIP_SIGNATURE):
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                names = set(zf.namelist())
            if "xl/workbook.bin" in names:
                return "pyxlsb"
            if "xl/workbook.xml" in names:
                return "openpyxl"
        except zipfile.BadZipFile:
            pass
        return "openpyxl"

    ext = (file_extension or "").lower()
    if ext.endswith(".xls"):
        return "xlrd"
    if ext.endswith(".xlsb"):
        return "pyxlsb"
    return "openpyxl"


class ExcelExtractor(TextExtractor):
    def read_workbook(self, file_extension: str, response: DocumentContent) -> dict[str, pd.DataFrame]:
        content = response.content
        engine = resolve_excel_engine(file_extension, content)
        # READ EXCEL WITHOUT HEADERS TO AVOID PANDAS DEFAULT LABELING
        return pd.read_excel(io.BytesIO(content), sheet_name=None, header=None, engine=engine, nrows=5000)

    def extract(self, file_extension: str, response: DocumentContent) -> list[str]:
        ws_dict = self.read_workbook(file_extension, response)

        if len(ws_dict.items()) != 1:
            raise ValueError('Excel document has ' + str(len(ws_dict.items())) + ' sheets')

        tables = []
        for excel_sheet_index, (sheet_name, df) in enumerate(ws_dict.items()):
            # Limit table size to not congest the pipeline with huge Excels
            df = df.iloc[:MAX_ROWS, :MAX_COLUMNS]

            # Round numeric values to 2 decimal places
            df = df.map(lambda x: round(x, 2) if isinstance(x, (int, float)) else x)

            # Strip whitespace and handle special characters
            df = df.map(lambda x: str(x).strip().replace('|', r'\|') if pd.notnull(x) else "")

            # Get table
            md_table = df.to_markdown(index=False)
            tables.append(md_table)

        return tables

    def extractBySheet(self, file_extension: str, response: DocumentContent) -> list[ExcelSheetResponse]:
        ws_dict = self.read_workbook(file_extension, response)

        tables = []
        for excel_sheet_index, (sheet_name, df) in enumerate(ws_dict.items()):
            # Limit table size to not congest the pipeline with huge Excels
            df = df.iloc[:MAX_ROWS, :MAX_COLUMNS]

            # Round numeric values to 2 decimal places
            df = df.map(lambda x: round(x, 2) if isinstance(x, (int, float)) else x)

            # Strip whitespace and handle special characters
            df = df.map(lambda x: str(x).strip().replace('|', r'\|') if pd.notnull(x) else "")

            # Get table
            md_table = df.to_markdown(index=False)

            # Remove consecutive spaces while preserving newlines and table formatting
            md_table = re.sub(r' {2,}', ' ', md_table)  # Replace 2 or more spaces with a single space
            md_table = re.sub(r' +\n', '\n', md_table)  # Remove spaces at the end of lines
            md_table = re.sub(r'\n +', '\n', md_table)  # Remove spaces at the beginning of lines

            # Truncate to 5 million characters
            md_table = md_table[:CHAR_LIMIT]

            tables.append(ExcelSheetResponse({"sheet_name": sheet_name, "text": md_table}))

        return tables


class OpenpyxlExcelExtractor(TextExtractor):
    """Excel -> coordinate-aware markdown (see excel_to_markdown.py), as an alternative to ExcelExtractor.

    Compared with the pandas ExcelExtractor it keeps the real header rows, finds several tables on one
    sheet, shows merged cells, hidden sheets/rows/columns, and (optionally) comments, formulas and
    General-format dates, and gives every table Excel's row numbers and column letters for citations.

    Same interface and output shape as ExcelExtractor, so it can be swapped in:
      * extractBySheet() -> [{"sheet_name": ..., "text": ...}, ...]   (one entry per sheet)
      * extract()        -> [text, ...]                               (one string per sheet; no one-sheet limit)

    Only .xlsx / .xlsm can be read this way. Legacy .xls and .xlsb files are handed to the pandas
    ExcelExtractor automatically (set fallback_to_pandas=False to raise instead).

    Extra keyword arguments are the options of excel_to_markdown.Options, for example
    include_comments=True, detect_general_dates=True, include_formulas=True, include_hidden_sheets=False.
    Like the pandas version, only the first MAX_ROWS rows and MAX_COLUMNS columns of a sheet are read.
    """

    def __init__(self, include_legend: bool = True, fallback_to_pandas: bool = True, **options):
        self.include_legend = include_legend
        self.fallback_to_pandas = fallback_to_pandas
        self.options = {"row_limit": MAX_ROWS, "column_limit": MAX_COLUMNS, **options}

    def _parse(self, file_extension: str, response):
        """Returns the ParseResult, or None if this file type has to go to the pandas extractor."""
        engine = resolve_excel_engine(file_extension, response.content)
        if engine != "openpyxl":
            if not self.fallback_to_pandas:
                raise ValueError(f"The openpyxl method cannot read this file type (engine would be '{engine}')")
            logger.info("openpyxl method cannot read %s files; using the pandas extractor", engine)
            return None
        result = parse_workbook(response.content, legend=False, **self.options)
        for warning in result.warnings:
            logger.warning(warning)
        return result

    def _to_sheets(self, result) -> list[ExcelSheetResponse]:
        sheets = []
        for sheet in result.sheets:
            if sheet.get("skipped"):
                continue
            text = sheet["markdown"]
            if self.include_legend:
                text = result.legend + "\n\n" + text
            sheets.append(ExcelSheetResponse({"sheet_name": sheet["name"], "text": text[:CHAR_LIMIT]}))
        return sheets

    def extractBySheet(self, file_extension: str, response) -> list[ExcelSheetResponse]:
        result = self._parse(file_extension, response)
        if result is None:
            return ExcelExtractor().extractBySheet(file_extension, response)
        return self._to_sheets(result)

    def extract(self, file_extension: str, response) -> list[str]:
        result = self._parse(file_extension, response)
        if result is None:
            return ExcelExtractor().extract(file_extension, response)
        return [sheet["text"] for sheet in self._to_sheets(result)]


class CsvExtractor(TextExtractor):
    def extract(self, file_extension: str, response: DocumentContent) -> list[str]:
        byte_stream = io.BytesIO(response.content)
        detection = chardet.detect(response.content)
        encoding = detection.get('encoding', 'utf-8')

        df = pd.read_csv(byte_stream, header=None, nrows=MAX_ROWS, encoding=encoding)
        # Limit table size
        df = df.iloc[:MAX_ROWS, :MAX_COLUMNS]

        # Round numeric values, clean whitespace, handle special characters
        df = df.map(lambda x: round(x, 2) if isinstance(x, (int, float)) else x)
        df = df.map(lambda x: str(x).strip().replace('|', r'\|') if pd.notnull(x) else "")

        # Convert to markdown table and truncate
        md_table = df.to_markdown(index=False)
        md_table = md_table[:CHAR_LIMIT]

        return [md_table]


class RtfExtractor(TextExtractor):
    def extract(self, file_extension: str, response: DocumentContent) -> list[str]:
        rtf_text = response.text
        raw_text = rtf_to_text(rtf_text)
        return [raw_text]


class PptxExtractor(TextExtractor):
    def _extract_shapes_text(self, shapes) -> list[str]:
        """Recursively extract text from shapes, including grouped shapes and tables."""
        texts = []
        for shape in shapes:
            if hasattr(shape, "shapes"):
                texts.extend(self._extract_shapes_text(shape.shapes))
            if shape.has_text_frame:
                texts.append(shape.text_frame.text)
            if shape.has_table:
                for row in shape.table.rows:
                    row_text = " | ".join(cell.text for cell in row.cells)
                    texts.append(row_text)
        return texts

    def extract(self, file_extension: str, response: DocumentContent) -> list[str]:
        from pptx import Presentation
        byte_stream = io.BytesIO(response.content)
        prs = Presentation(byte_stream)
        pages = []
        for slide in prs.slides:
            slide_texts = self._extract_shapes_text(slide.shapes)
            slide_text = "\n".join(t for t in slide_texts if t.strip())
            if slide_text.strip():
                pages.append(slide_text)
        return pages


class PdfExtractor(TextExtractor):
    def extract(self, file_extension: str, response: DocumentContent) -> list[str]:
        pdf_bytes = response.content
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        return [page.get_text("text", sort=True).strip() for page in pdf_doc]


class FormattedPDFTextExtractor(TextExtractor):
    def extract(self, file_extension: str, response: DocumentContent,
                config: ExtractionConfig | None = None) -> list[str]:
        if config is None:
            config = ExtractionConfig()

        pdf_bytes = response.content
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        # previously returned
        # return [page.get_text("text", sort=True).strip() for page in pdf_doc]
        pages_data = [page.get_text("dict", sort=True) for page in pdf_doc]
        body_size, left_margin = _analyze_document(pages_data, config)

        return [_render_page(pd, body_size, left_margin, config) for pd in pages_data]


# Canonical mapping from file extension to extractor instance.
EXTENSION_TO_EXTRACTOR: dict[str, TextExtractor] = {
    ".txt": PlainTextExtractor(),
    ".html": PlainTextExtractor(),
    ".htm": PlainTextExtractor(),
    ".md": PlainTextExtractor(),
    ".xml": PlainTextExtractor(),
    ".docx": DocxExtractor(),
    ".docm": DocmExtractor(),
    ".xls": ExcelExtractor(),
    ".xlsm": ExcelExtractor(),
    ".xlsb": ExcelExtractor(),
    ".xlsx": ExcelExtractor(),
    ".csv": CsvExtractor(),
    ".rtf": RtfExtractor(),
    ".pptx": PptxExtractor(),
    ".pdf": PdfExtractor(),
}
