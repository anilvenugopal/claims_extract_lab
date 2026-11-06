import re


BASE64_MARKDOWN_IMAGE_PATTERN = re.compile(
    r"!\[(?:[^\[\]]|\[[^\[\]]*\])*\]"
    r"\(\s*data:image/[^;\s)]+;base64,[^)]*\)",
    re.IGNORECASE,
)


def removeMarkdownEmbeddedBase64(text: str | None) -> str:
    """Remove embedded base64 Markdown images without pathological backtracking."""
    if not text:
        return ""

    # Avoid evaluating the pattern for ordinary Markdown images. The previous
    # nested-quantifier expression could backtrack exponentially after seeing
    # ``![`` when the target was a regular URL rather than a data URI.
    if "data:image/" not in text.lower():
        return text

    return BASE64_MARKDOWN_IMAGE_PATTERN.sub("", text)
