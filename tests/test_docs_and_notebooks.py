"""Checks that the documentation and the notebooks match the code, so they cannot quietly go stale.

If one of these fails, either the code changed and the document did not follow, or the other way
round. The failure message says what to refresh.
"""
import dataclasses
import re
import sys
from pathlib import Path

import nbformat
import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
import excel_to_markdown as m  # noqa: E402

DOC_PATH = ROOT / "docs" / "excel_to_markdown.md"
SAVED_PATH = ROOT / "docs" / "excel_edge_cases_output.md"
BOOK = ROOT / "samples" / "excel_edge_cases.xlsx"
DOC = DOC_PATH.read_text(encoding="utf-8")
SAVED = SAVED_PATH.read_text(encoding="utf-8")

REFRESH = ("python src/excel_to_markdown.py samples/excel_edge_cases.xlsx "
           "-o docs/excel_edge_cases_output.md --comments --detect-general-dates")


def _fenced_blocks(text):
    """[(language, [lines])] for every ``` fenced block (language is '' for a plain fence)."""
    blocks, inside, lang, buf = [], False, "", []
    for line in text.splitlines():
        if line.startswith("```"):
            if inside:
                blocks.append((lang, buf))
                inside, buf = False, []
            else:
                inside, lang = True, line[3:].strip()
        elif inside:
            buf.append(line)
    return blocks


def _first_column(text, header_starts_with):
    """First-column cells of the markdown table whose header row starts with the given text."""
    lines = text.splitlines()
    start = next(i for i, ln in enumerate(lines) if ln.startswith(header_starts_with))
    cells = []
    for ln in lines[start + 2:]:
        if not ln.startswith("|"):
            break
        cells.append(ln.split("|")[1].strip())
    return cells


# ---------------------------------------------------------------- the saved output and the examples
def test_saved_output_is_current():
    result = m.parse_workbook(str(BOOK), include_comments=True, detect_general_dates=True)
    assert result.markdown == SAVED, f"docs/excel_edge_cases_output.md is out of date. Refresh it with:\n  {REFRESH}"


def test_every_example_in_the_docs_is_real_output():
    # plain ``` fences in the docs are output excerpts; they must appear, line by line and in order, in the saved output
    examples = [lines for lang, lines in _fenced_blocks(DOC) if lang == ""]
    assert len(examples) >= 10
    for lines in examples:
        position = 0
        for line in (ln for ln in lines if ln.strip()):
            found = SAVED.find(line, position)
            assert found != -1, f"docs example line not found (in order) in the saved output: {line[:100]!r}"
            position = found + len(line)


# ---------------------------------------------------------------- what is documented = what exists
def test_docs_list_every_tab_of_the_edge_case_workbook():
    documented = re.findall(r"^\| \d+ \| `([^`]+)` \|", DOC, re.M)
    actual = openpyxl.load_workbook(BOOK).sheetnames
    assert documented == actual


def test_docs_list_every_payload_field(h):
    documented = {t.strip("`") for t in _first_column(DOC, "| Payload field |")}
    inherited = {f.name for f in dataclasses.fields(h.app.Payload)}
    actual = {f.name for f in dataclasses.fields(h.app.ExcelPayload)} - inherited
    assert documented == actual


def test_docs_list_every_converter_option_and_command_line_flag():
    tokens = [t for cell in _first_column(DOC, "| Option (Python / command line) |")
              for t in re.findall(r"`([^`]+)`", cell)]
    documented_names = {t for t in tokens if not t.startswith("--")}
    documented_flags = {t.split()[0] for t in tokens if t.startswith("--")}

    actual_names = {f.name for f in dataclasses.fields(m.Options)}
    actual_flags = {f for f in m.build_parser()._option_string_actions if f.startswith("--")} - {"--help", "--output"}
    assert documented_names == actual_names
    assert documented_flags == actual_flags


def test_docs_describe_the_header_words_the_code_uses():
    # the lists in section 5 are the real ones
    def listed(after, before):
        part = DOC.split(after)[1].split(before)[0]
        return set(re.findall(r"`([a-z]+)`", part))
    assert listed("contains a word from the built-in list:", "Words are split") == set(m.DEFAULT_DATE_KEYWORDS)
    assert listed("suggests an amount, count or id:", "So \"Amount Due\"") == set(m.NOT_DATE_WORDS)


# ---------------------------------------------------------------- links and paths
def test_local_links_and_listed_files_exist():
    for md in [ROOT / "README.md", *sorted((ROOT / "docs").glob("*.md"))]:
        for target in re.findall(r"\]\(([^)#\s]+)(?:#[^)]*)?\)", md.read_text(encoding="utf-8")):
            if target.startswith(("http://", "https://", "mailto:")):
                continue
            assert (md.parent / target).exists(), f"{md.name} links to a missing file: {target}"

    tree = next(lines for lang, lines in _fenced_blocks(DOC) if lang == "text")
    names = re.findall(r"[\w.-]+\.(?:py|md|xlsx|ipynb)", "\n".join(tree))
    assert len(names) >= 12
    for name in names:
        assert list(ROOT.rglob(name)), f"the file tree in docs/excel_to_markdown.md lists a file that does not exist: {name}"


def test_no_references_to_files_that_were_moved():
    stale = ("tests/fixtures", "test_book", "build_test_workbook")
    for path in [ROOT / "README.md", *(ROOT / "docs").glob("*.md"), *(ROOT / "notebooks").glob("*.ipynb"),
                 *(ROOT / "src").glob("*.py"), *ROOT.glob("*.py"),
                 *(p for p in (ROOT / "tests").glob("*.py") if p.name != Path(__file__).name)]:
        text = path.read_text(encoding="utf-8")
        for needle in stale:
            assert needle not in text, f"{path.relative_to(ROOT)} still mentions {needle}"


# ---------------------------------------------------------------- notebooks
def test_notebooks_are_valid_and_contain_no_error_output():
    notebooks = sorted((ROOT / "notebooks").glob("*.ipynb"))
    assert {n.name for n in notebooks} >= {"01_claims_extraction_lab.ipynb", "02_excel_extraction_lab.ipynb"}
    for path in notebooks:
        nb = nbformat.read(path, as_version=4)
        nbformat.validate(nb)
        errors = [(i, o.ename) for i, c in enumerate(nb.cells) if c.cell_type == "code"
                  for o in c.get("outputs", []) if o.output_type == "error"]
        assert not errors, f"{path.name} has cells that ended in an error: {errors}"
