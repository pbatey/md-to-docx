"""Property test: link labels keep their inline formatting."""

from docx import Document
from docx.oxml.ns import qn
from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import (
    STYLE_CODE_CHAR,
    STYLE_HYPERLINK,
    LinkContext,
    _render_inline,
    add_formatted_text,
)

_text = st.text(alphabet="abcXYZ019 ", min_size=1, max_size=6).filter(
    lambda s: s.strip() != ""
)


def _wrap(children):
    return st.one_of(
        children.map(lambda s: f"**{s}**"),
        children.map(lambda s: f"*{s}*"),
        children.map(lambda s: f"~~{s}~~"),
        children.map(lambda s: f"***{s}***"),
        _text.map(lambda s: f"`{s}`"),  # code spans are literal
    )


_inline = st.recursive(
    _text,
    lambda children: st.one_of(
        _wrap(children),
        st.lists(children, min_size=2, max_size=3).map("".join),
    ),
    max_leaves=6,
)


def _run_info(r):
    """Return (text, flags, rStyle ids) for a w:r element."""
    rPr = r.find(qn("w:rPr"))
    flags = set()
    styles = []
    if rPr is not None:
        for tag, name in (("w:b", "bold"), ("w:i", "italic"), ("w:strike", "strike")):
            el = rPr.find(qn(tag))
            if el is not None and el.get(qn("w:val")) not in ("0", "false"):
                flags.add(name)
        for rs in rPr.findall(qn("w:rStyle")):
            styles.append(rs.get(qn("w:val")))
        if STYLE_CODE_CHAR in styles or rPr.find(qn("w:rFonts")) is not None:
            flags.add("code")
    text = "".join(t.text or "" for t in r.findall(qn("w:t")))
    return text, frozenset(flags), styles


def _merge(items):
    out = []
    for text, flags in items:
        if text == "":
            continue
        if out and out[-1][1] == flags:
            out[-1] = (out[-1][0] + text, flags)
        else:
            out.append((text, flags))
    return out


def _outside(label):
    p = Document().add_paragraph()
    _render_inline(p, label)
    runs = p._p.findall(qn("w:r"))
    assert p._p.find(qn("w:hyperlink")) is None
    return _merge([_run_info(r)[:2] for r in runs])


def _inside(markdown):
    p = Document().add_paragraph()
    ctx = LinkContext(base_url=None, source_rel_dir=(), warn=lambda m: None)
    add_formatted_text(p, markdown, anchors={}, link_ctx=ctx)
    links = p._p.findall(qn("w:hyperlink"))
    assert len(links) == 1
    assert not p._p.xpath(".//w:hyperlink//w:hyperlink")
    infos = [_run_info(r) for r in links[0].findall(qn("w:r"))]
    for _, _, styles in infos:
        assert STYLE_HYPERLINK in styles
    return _merge([i[:2] for i in infos])


# Feature: sibling-markdown-links, Property 10: Link labels keep their inline formatting
# **Validates: Requirements 1.1, 1.4, 1.5, 1.10, 4.2, 13.15**
@settings(max_examples=150, deadline=None)
@given(label=_inline)
def test_sibling_link_label_keeps_formatting(label):
    assert _inside(f"[{label}](other.md)") == _outside(label)


# Feature: sibling-markdown-links, Property 10: Link labels keep their inline formatting
# **Validates: Requirements 1.1, 1.4, 1.5, 1.10, 4.2, 13.15**
@settings(max_examples=150, deadline=None)
@given(label=_inline)
def test_wiki_link_label_keeps_formatting(label):
    # wiki_parts trims surrounding whitespace from the label
    assert _inside(f"[[Note|{label}]]") == _outside(label.strip())
