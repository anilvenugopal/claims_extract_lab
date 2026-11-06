from __future__ import annotations
import docx
from typing import IO, TYPE_CHECKING, cast

if TYPE_CHECKING:
    from docx.document import Document as DocumentObject
    from docx.parts.document import DocumentPart

from docx.opc.constants import CONTENT_TYPE as CT
from docx.package import Package


# patched version of docx.Doucment to allow docm files
def DOCMDocument(docx_in: str | IO[bytes] | None = None) -> DocumentObject:
    WML_DOCUMENT_MACRO_ENABLED_MAIN = "application/vnd.ms-word.document.macroEnabled.main+xml"
    docx.opc.part.PartFactory.part_type_for[WML_DOCUMENT_MACRO_ENABLED_MAIN] = docx.parts.document.DocumentPart

    """Return a |Document| object loaded from `docx_in`, where `docx_in` can be either a path
    to a ``.docx`` file (a string) or a file-like object.

    If `docx_in` is missing or ``None``, the built-in default document "template" is
    loaded.
    """
    docx_in = docx._default_docx_path() if docx_in is None else docx_in
    document_part = cast("DocumentPart", Package.open(docx_in).main_document_part)
    if (document_part.content_type != CT.WML_DOCUMENT_MAIN) and (document_part.content_type != WML_DOCUMENT_MACRO_ENABLED_MAIN):
        tmpl = "file '%s' is not a Word file, content type is '%s'"
        raise ValueError(tmpl % (docx_in, document_part.content_type))
    return document_part.document
