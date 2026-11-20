"""Golden fixtures of the pre-feature converter output.

The fixtures under ``tests/golden/`` were generated with the converter as it
was BEFORE the sibling-link feature. Property 12 (``--no-sibling-links``
reproduces the pre-feature output) compares against them.

Regenerating is only valid on pre-feature code::

    python tests/golden_fixtures.py --write

Never regenerate after the converter has been changed for the feature.
"""

from __future__ import annotations

import sys
import tempfile
import zipfile
from pathlib import Path

GOLDEN_DIR = Path(__file__).parent / "golden"

# name -> Markdown source. Covers headings, duplicate and long headings,
# internal anchors, [[#Heading]], external links, sibling links, wiki links.
DOCUMENTS: dict[str, str] = {
    "headings": "# Title\n\n## Section One\n\nText.\n\n### Sub *Section*\n\n#### Deep\n",
    "duplicate_headings": (
        "# Intro\n\n## Setup\n\nA\n\n## Setup\n\nB\n\n## Setup\n\nC\n\n"
        "See [first](#setup) and [again](#setup-1).\n"
    ),
    "long_headings": (
        "# " + "A very long heading title that goes on and on " * 5 + "\n\n"
        "## " + "Another extremely long heading with many words " * 4 + "\n\n"
        "## " + "Another extremely long heading with many words " * 4 + "\n\n"
        "Link to [long](#"
        + "a-very-long-heading-title-that-goes-on-and-on-" * 5
        + ").\n"
    ),
    "internal_anchors": (
        "# Top\n\n## Details\n\nGo to [top](#top) or [details](#details) "
        "or [missing](#nope).\n"
    ),
    "wiki_heading_links": (
        "# Notes\n\n## Plan\n\nSee [[#Plan]], [[#Plan|the plan]], "
        "[[#Missing]] and [[#Notes|Home]].\n"
    ),
    "external_links": (
        "# Links\n\n[site](https://example.com/a?b=1#c) and "
        "<https://example.org> and [mail](mailto:a@b.com) and "
        "[ftp](ftp://host/file.md).\n"
    ),
    "sibling_links": (
        "# Siblings\n\n[a](other.md) [b](./sub/page.md#sec) "
        "[c](../up/Doc.markdown) [d](<with space.md>) [e](docs/My%20File.md) "
        "[f](/abs/path.md) [g](page.md#Heading Two).\n"
    ),
    "wiki_links": (
        "# Wiki\n\n[[Page]] [[folder/Page]] [[Page|alias]] "
        "[[Page#Heading]] [[Page#Heading|alias]] [[folder/Page.md]].\n"
    ),
    "mixed": (
        "# Mixed\n\n## Part\n\n## Part\n\n"
        "[ext](https://example.com) [sib](other.md) [anc](#part) "
        "[[Wiki]] [[#Part]] [[#Part|x]] [rel](notes.txt) [img](pic.png).\n\n"
        "- item [li](sibling.md)\n- item [[Wiki Item]]\n\n"
        "> quote [q](q.md)\n\n```\n[code](code.md) [[code]]\n```\n"
    ),
}


def convert_to_document_xml(markdown: str) -> str:
    """Convert Markdown with the current converter and return word/document.xml."""
    from md_to_docx import build_docx, parse_markdown

    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        build_docx(parse_markdown(markdown), str(out))
        with zipfile.ZipFile(out) as z:
            return z.read("word/document.xml").decode("utf-8")


def golden_path(name: str) -> Path:
    return GOLDEN_DIR / f"{name}.xml"


def write_golden() -> list[Path]:
    """Generate fixtures (and the Markdown sources) from the current converter."""
    GOLDEN_DIR.mkdir(exist_ok=True)
    written = []
    for name, md in DOCUMENTS.items():
        (GOLDEN_DIR / f"{name}.md").write_text(md, encoding="utf-8", newline="\n")
        p = golden_path(name)
        p.write_text(convert_to_document_xml(md), encoding="utf-8", newline="\n")
        written.append(p)
    return written


def load_golden() -> dict[str, tuple[str, str]]:
    """Load fixtures for Property 12: ``{name: (markdown, expected_document_xml)}``."""
    result = {}
    for name in DOCUMENTS:
        md = (GOLDEN_DIR / f"{name}.md").read_text(encoding="utf-8")
        xml = golden_path(name).read_text(encoding="utf-8")
        result[name] = (md, xml)
    return result


if __name__ == "__main__":
    if "--write" not in sys.argv:
        sys.exit("pass --write to regenerate (only valid on pre-feature code)")
    for p in write_golden():
        print(p)
