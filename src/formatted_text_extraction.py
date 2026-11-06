from pathlib import Path
from collections import Counter
from dataclasses import dataclass
import fitz


@dataclass(frozen=True)
class ExtractionConfig:
    """Tuning parameters for formatted PDF text extraction. All fields are optional with sensible defaults."""
    h1_ratio: float = 1.8           # font-size ratio (vs body text) at or above which a line becomes an H1 heading
    h2_ratio: float = 1.4           # same for H2 — band is [h2_ratio, h1_ratio)
    h3_ratio: float = 1.15          # same for H3 — band is [h3_ratio, h2_ratio); below this is body text
    indent_step: float = 24.0       # points of horizontal offset per indentation level (~1/3 inch)
    indent_dead_zone: float = 0.4   # fraction of indent_step below which offset is ignored (avoids false indents from minor x-jitter)
    gap_threshold: float = 1.5      # minimum gap in points between spans to trigger space insertion
    fallback_char_width: float = 6.0   # avg character width fallback (pts) when a line has no measurable text (~12pt font)
    fallback_body_size: float = 12.0   # body font size fallback when document has no text at all


def _is_bold(flags: int) -> bool:
    return bool(flags & 16)


def _is_italic(flags: int) -> bool:
    return bool(flags & 2)


def _analyze_document(pages_data: list[dict], config: ExtractionConfig) -> tuple[float, float]:
    """
    Determine baseline body font size (most chars) and left margin
    (most common first-span x0) across the whole document.
    """
    size_chars: Counter[float] = Counter()
    x0_counts: Counter[int] = Counter()

    for page_dict in pages_data:
        for block in page_dict["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                first_seen = False
                for span in line["spans"]:
                    text = span["text"].strip()
                    if text:
                        size_chars[round(span["size"])] += len(text) # round size to nearest integer so resilient to tiny variance (<0.5) in font sizes
                        if not first_seen:
                            x0_counts[round(span["bbox"][0])] += 1 # increment left boundary counter
                            first_seen = True

    body_size = size_chars.most_common(1)[0][0] if size_chars else config.fallback_body_size
    left_margin = float(x0_counts.most_common(1)[0][0]) if x0_counts else 0.0
    return body_size, left_margin


def _heading_level(size: float, body_size: float, config: ExtractionConfig) -> str | None:
    if body_size <= 0:
        return None
    ratio = size / body_size
    if ratio >= config.h1_ratio:
        return "#"
    if ratio >= config.h2_ratio:
        return "##"
    if ratio >= config.h3_ratio:
        return "###"
    return None


def _indent_prefix(x0: float, margin: float, config: ExtractionConfig) -> str:
    offset = x0 - margin
    if offset < config.indent_step * config.indent_dead_zone:
        return ""
    return "    " * max(1, int(offset / config.indent_step))


def _wrap_markers(text: str, bold: bool, italic: bool) -> str:
    """
    Apply markers with whitespace OUTSIDE so the markdown is valid.
    '**bold** next'  NOT  '**bold **next'
    """
    if not bold and not italic:
        return text

    stripped = text.strip()
    if not stripped:
        return text  # pure whitespace — don't wrap

    leading = text[: len(text) - len(text.lstrip())]
    trailing = text[len(text.rstrip()) :]

    if bold and italic:
        return f"{leading}***{stripped}***{trailing}"
    if bold:
        return f"{leading}**{stripped}**{trailing}"
    return f"{leading}*{stripped}*{trailing}"


def _merge_line_spans(spans: list[dict], config: ExtractionConfig) -> tuple[str, float, float]:
    """
    Merge a line's spans into one string with inline formatting.

    - Groups consecutive same-formatted spans before wrapping, so **a b** instead of **a****b**.
    - Inserts proportional spacing when the x-gap between spans suggests a columnar or tabular layout, preserving spatial structure.
    - Returns (formatted_text, weighted_avg_font_size, first_x0).
    """
    if not spans:
        return ("", 0.0, 0.0)

    # --- compute average character width from this line's spans ---
    total_width = sum(s["bbox"][2] - s["bbox"][0] for s in spans if s["text"])
    total_chars = sum(len(s["text"]) for s in spans if s["text"])
    avg_char_width = total_width / total_chars if total_chars else config.fallback_char_width

    # --- group consecutive same-formatted spans ---
    groups: list[dict] = []

    for span in spans:
        raw = span["text"]
        if not raw:
            continue
        bold = _is_bold(span["flags"])
        italic = _is_italic(span["flags"])
        x0 = span["bbox"][0]
        x1 = span["bbox"][2]

        gap_spaces = 0
        if groups:
            gap = x0 - groups[-1]["x1"]
            if gap > config.gap_threshold and not groups[-1]["text"].endswith((" ", "\t")) and not raw.startswith((" ", "\t")):
                gap_spaces = max(1, round(gap / avg_char_width)) # clamp to one space

        if groups and groups[-1]["bold"] == bold and groups[-1]["italic"] == italic:
            # Same formatting → extend current group
            if gap_spaces:
                groups[-1]["text"] += " " * gap_spaces
            groups[-1]["text"] += raw
            groups[-1]["x1"] = x1
        else:
            # New formatting → close previous, start new group
            if gap_spaces:
                # Attach gap-spaces to the PREVIOUS group's trailing text so they stay outside that group's closing marker
                groups[-1]["text"] += " " * gap_spaces
            groups.append({"bold": bold, "italic": italic, "text": raw, "x1": x1})

    if not groups:
        return ("", 0.0, 0.0)

    # --- apply markers per group ---
    parts = [_wrap_markers(g["text"], g["bold"], g["italic"]) for g in groups]
    merged = "".join(parts)

    # --- weighted average font size for heading detection ---
    total_size = sum(s["size"] * len(s["text"]) for s in spans if s["text"])
    avg_size = total_size / total_chars if total_chars else 0.0

    first_x0 = next((s["bbox"][0] for s in spans if s["text"]), 0.0)
    return (merged, avg_size, first_x0)


def _render_page(page_dict: dict, body_size: float, left_margin: float, config: ExtractionConfig) -> str:
    block_strs: list[str] = []

    for block in page_dict["blocks"]:
        if block.get("type") != 0:
            ## NOTE this drops all non-text blocks
            continue

        line_strs: list[str] = []
        for line in block["lines"]:
            merged, avg_size, first_x0 = _merge_line_spans(line["spans"], config)
            clean = merged.strip()
            if not clean:
                continue

            heading = _heading_level(avg_size, body_size, config)
            indent = _indent_prefix(first_x0, left_margin, config)

            if heading:
                # Strip bold markers from headings — the # already
                # conveys structural weight; bold inside is redundant
                clean = clean.replace("***", "*").replace("**", "")
                line_strs.append(f"{heading} {clean}")
            elif indent:
                line_strs.append(f"{indent}{clean}")
            else:
                line_strs.append(clean)

        if line_strs:
            block_strs.append("\n".join(line_strs))

    return "\n\n".join(block_strs)


def extract_pdf_for_llm(filename: str, config: ExtractionConfig | None = None) -> list[str]:
    """
    One string per PyMuPDF Page object.
    Formatting encoded as inline markdown for LLM consumption:
      - Heading hierarchy (# / ## / ###) from relative font size
      - **bold** and *italic* inline, with valid marker placement
      - Indentation from x-position (4-space tab stops)
      - Continuous reading flow within lines (no span fragmentation)
    """
    if config is None:
        config = ExtractionConfig()

    pdf_path = Path(__file__).parent / filename
    pdf_doc = fitz.open(str(pdf_path))

    pages_data = [page.get_text("dict", sort=True) for page in pdf_doc]
    body_size, left_margin = _analyze_document(pages_data, config)

    return [_render_page(pd, body_size, left_margin, config) for pd in pages_data]
