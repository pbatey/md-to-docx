"""Unit tests for sibling Markdown link rendering (task 6.10).

Covers the worked examples of the design's Data Models table, label formatting,
long-heading bookmark naming, malformed/passthrough inputs and heading bookmark
naming with links enabled or disabled.
"""

import re

import pytest
from docx import Document
from docx.oxml.ns import qn

from md_to_docx import build_docx, parse_markdown
from md_to_docx.md_to_docx import (
    STYLE_HYPERLINK,
    LinkContext,
    add_formatted_text,
    bookmark_for_fragment,
    bookmark_name,
    build_heading_anchors,
    normalize_base_url,
)

BOOKMARK_RE = re.compile(r"^_[A-Za-z0-9_]{1,39}$")


def _ctx(base=None, rel_dir=(), enabled=True, warnings=None):
    sink = warnings if warnings is not None else []
    return LinkContext(
        enabled=enabled,
        base_url=normalize_base_url(base) if base else None,
        source_rel_dir=tuple(rel_dir),
        warn=sink.append,
    )


def _para(md, ctx, anchors=None):
    p = Document().add_paragraph()
    add_formatted_text(p, md, anchors=anchors or {}, link_ctx=ctx)
    return p


def _links(p):
    return p._p.findall(qn("w:hyperlink"))


def _rel(p, link):
    """Return (target_ref, is_external) of a hyperlink's relationship."""
    rel = p.part.rels[link.get(qn("r:id"))]
    return rel.target_ref, rel.is_external


def _anchor(link):
    return link.get(qn("w:anchor"))


def _text(p):
    return "".join(t.text or "" for t in p._p.iter(qn("w:t")))


def _check(md, target, anchor, base=None, rel_dir=()):
    p = _para(md, _ctx(base, rel_dir))
    links = _links(p)
    assert len(links) == 1
    link = links[0]
    ref, external = _rel(p, link)
    assert external
    assert "#" not in ref
    assert ref == target
    assert _anchor(link) == anchor
    for r in link.findall(qn("w:r")):
        styles = [s.get(qn("w:val")) for s in r.iter(qn("w:rStyle"))]
        assert STYLE_HYPERLINK in styles
    return p


# --------------------------------------------------------------------------- #
# Worked examples
# --------------------------------------------------------------------------- #
def test_example_relative_path():
    _check("[S](b/c.md)", "b/c.docx", None)


def test_example_dot_slash_uppercase_ext_fragment_text():
    _check(
        "[S](./My%20Notes.MD#Install Steps)", "My%20Notes.docx", "_toc_install_steps"
    )


def test_example_angle_brackets_with_space():
    _check("[S](<My Notes.md>)", "My%20Notes.docx", None)


def test_example_backslash_separator():
    _check(r"[S](sub\x.md)", "sub/x.docx", None)


def test_example_base_url_single_file():
    _check("[S](b/c.md)", "https://x/docs/b/c.docx", None, base="https://x/docs/")


def test_example_base_url_recursive_subdir():
    _check(
        "[S](../y.md#Intro)",
        "https://x/docs/a/y.docx",
        "_toc_intro",
        base="https://x/docs/",
        rel_dir=("a", "b"),
    )


def test_example_base_url_climb_warns_and_clamps():
    warnings = []
    p = _para("[S](../y.md)", _ctx("https://x/docs/", warnings=warnings))
    (link,) = _links(p)
    assert _rel(p, link)[0] == "https://x/docs/y.docx"
    assert _anchor(link) is None
    assert len(warnings) == 1


@pytest.mark.parametrize(
    "md", ["[S](https://e.com/a.md)", "[S](/abs/a.md)", "[S](a.txt)"]
)
@pytest.mark.parametrize("base", [None, "https://x/docs/"])
def test_example_non_sibling_unchanged(md, base):
    with_ctx = _para(md, _ctx(base))
    without = _para(md, None)
    assert with_ctx._p.xml.count("w:hyperlink") == without._p.xml.count("w:hyperlink")
    assert _text(with_ctx) == _text(without)
    for p in (with_ctx, without):
        for link in _links(p):
            assert not _rel(p, link)[0].endswith(".docx")


def test_example_wiki_note_heading():
    p = _check("[[Note#Heading]]", "Note.docx", "_toc_heading")
    assert _text(p) == "Note#Heading"


def test_example_wiki_non_markdown_extension_with_bold_label():
    p = _check("[[a.txt|**A**]]", "a.txt.docx", None)
    (link,) = _links(p)
    runs = link.findall(qn("w:r"))
    assert len(runs) == 1
    assert runs[0].find(qn("w:rPr")).find(qn("w:b")) is not None
    assert _text(p) == "A"


def test_example_long_heading_bookmark():
    heading = "word " * 12  # 60 chars
    p = _para(f"[S](x.md#{heading.strip()})", _ctx())
    (link,) = _links(p)
    assert _rel(p, link)[0] == "x.docx"
    name = _anchor(link)
    assert len(name) == 40
    assert re.search(r"_[0-9a-f]{8}$", name)


def test_example_disabled_is_plain_text():
    p = _para("[S](b.md)", _ctx(enabled=False))
    assert _links(p) == []
    assert _text(p) == "S"


# --------------------------------------------------------------------------- #
# Label formatting
# --------------------------------------------------------------------------- #
def test_label_with_bold_plain_code_runs_in_one_hyperlink():
    p = _para("[**bold** plain `code`](x.md)", _ctx())
    (link,) = _links(p)
    runs = link.findall(qn("w:r"))
    assert len(runs) == 3
    texts = ["".join(t.text for t in r.findall(qn("w:t"))) for r in runs]
    assert texts == ["bold", " plain ", "code"]
    assert runs[0].find(qn("w:rPr")).find(qn("w:b")) is not None
    assert runs[1].find(qn("w:rPr")).find(qn("w:b")) is None
    code_rpr = runs[2].find(qn("w:rPr"))
    assert code_rpr.find(qn("w:b")) is None
    assert (
        code_rpr.find(qn("w:rFonts")) is not None
        or len(code_rpr.findall(qn("w:rStyle"))) > 1
    )
    assert not p._p.xpath(".//w:hyperlink//w:hyperlink")


def test_wiki_label_italic_run():
    p = _para("[[a.md|*Label*]]", _ctx())
    (link,) = _links(p)
    runs = link.findall(qn("w:r"))
    assert len(runs) == 1
    assert runs[0].find(qn("w:rPr")).find(qn("w:i")) is not None
    assert _text(p) == "Label"
    assert _rel(p, link)[0] == "a.docx"


# --------------------------------------------------------------------------- #
# Long headings
# --------------------------------------------------------------------------- #
LONG = "This is a really long heading that goes past forty characters"


def test_long_heading_bookmark_is_40_chars_and_repeatable():
    name = bookmark_for_fragment(LONG)
    assert len(name) == 40
    assert BOOKMARK_RE.match(name)
    assert name == bookmark_for_fragment(LONG)


def test_long_headings_with_common_prefix_are_distinct():
    a = bookmark_for_fragment(LONG + " alpha")
    b = bookmark_for_fragment(LONG + " beta")
    assert len(a) == len(b) == 40
    assert a != b


def test_link_by_full_text_or_slug_same_name():
    from md_to_docx.md_to_docx import slugify_heading

    slug = slugify_heading(LONG)
    assert bookmark_for_fragment(LONG) == bookmark_for_fragment(slug)
    assert bookmark_for_fragment(LONG) == bookmark_name(slug)


def test_long_heading_bookmark_matches_heading_and_links(tmp_path):
    md = f"# {LONG}\n\nSee [x](#{LONG}) and [y](other.md#{LONG}).\n"
    out = tmp_path / "o.docx"
    build_docx(parse_markdown(md), str(out), link_ctx=_ctx(), announce=False)
    doc = Document(str(out))
    starts = [b.get(qn("w:name")) for b in doc.element.body.iter(qn("w:bookmarkStart"))]
    anchors = [h.get(qn("w:anchor")) for h in doc.element.body.iter(qn("w:hyperlink"))]
    expected = bookmark_for_fragment(LONG)
    assert expected in starts
    assert anchors.count(expected) >= 2


# --------------------------------------------------------------------------- #
# Malformed and passthrough
# --------------------------------------------------------------------------- #
def test_wiki_same_document_heading_resolves_internal():
    anchors = build_heading_anchors(parse_markdown("# Heading\n"))
    p = _para("[[#Heading]]", _ctx(), anchors=anchors)
    (link,) = _links(p)
    assert link.get(qn("r:id")) is None
    assert _anchor(link) == anchors["heading"]
    assert _text(p) == "Heading"


def test_wiki_missing_heading_creates_no_external_link():
    p = _para("[[#Missing]]", _ctx(), anchors={})
    for link in _links(p):
        assert link.get(qn("r:id")) is None
    assert "Missing" in _text(p)


@pytest.mark.parametrize("md", ["[[]]", "[[Note|]]", "[[Note"])
def test_malformed_wiki_has_no_hyperlink_and_text_kept(md):
    p = _para(md, _ctx())
    assert _links(p) == []
    assert _text(p) == md


@pytest.mark.parametrize("md", ["[S](<My Notes.md)", "[S](<>)"])
def test_unclosed_or_empty_angle_target_is_plain(md):
    p = _para(md, _ctx())
    assert _links(p) == []
    assert _text(p) == "S"


def test_angle_target_with_closing_bracket_is_link():
    _check("[S](<a b.md>)", "a%20b.docx", None)


# --------------------------------------------------------------------------- #
# Heading bookmarks
# --------------------------------------------------------------------------- #
def _bookmarks(tmp_path, md, ctx):
    out = tmp_path / "o.docx"
    build_docx(parse_markdown(md), str(out), link_ctx=ctx, announce=False)
    doc = Document(str(out))
    return [
        b.get(qn("w:name"))
        for b in doc.element.body.iter(qn("w:bookmarkStart"))
        if (b.get(qn("w:name")) or "").startswith("_toc")
    ]


def test_heading_bookmarks_enabled_have_no_index_and_are_unique(tmp_path):
    md = "# Intro\n\n## Intro\n\n## Intro\n\n### a-b\n\n### a_b\n"
    names = _bookmarks(tmp_path, md, _ctx())
    assert len(names) == 5
    assert len(set(names)) == 5
    assert names[0] == "_toc_intro"
    for n in names:
        assert BOOKMARK_RE.match(n)
        assert not re.match(r"^_toc_\d+_", n)


def test_heading_bookmarks_legacy_keep_index_naming(tmp_path):
    names = _bookmarks(tmp_path, "# Intro\n\n## Details\n", _ctx(enabled=False))
    assert names == ["_toc_0_intro", "_toc_1_details"]
    blocks = parse_markdown("# Intro\n\n## Details\n")
    build_heading_anchors(blocks, legacy=True)
    assert [b["_bookmark"] for b in blocks if "_bookmark" in b] == names
