"""Property test for wiki link rewriting."""

from docx import Document
from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import (
    LinkContext,
    add_formatted_text,
    bookmark_for_fragment,
    encode_path,
    rewrite_wiki_link,
    to_docx_target,
)

_word = st.text(alphabet="abcdefXYZ019", min_size=1, max_size=6)
_note_base = st.lists(_word, min_size=1, max_size=3).map("_".join)
_dirs = st.lists(_word, max_size=3).map(lambda d: "".join(s + "/" for s in d))
_ext = st.sampled_from(["", ".md", ".markdown", ".MD", ".txt", ".pdf", ".v2"])
_heading = st.lists(_word, min_size=1, max_size=4).map(" ".join)
_label = st.lists(_word, min_size=1, max_size=3).map(" ".join)


def _ctx():
    return LinkContext(base_url=None, warn=lambda m: None)


def _expected_path(note):
    if note.lower().endswith((".md", ".markdown")):
        return encode_path(to_docx_target(note))
    return encode_path(note + ".docx")


# Feature: sibling-markdown-links, Property 14: Wiki link rewrite
# **Validates: Requirements 4.1, 4.2, 4.3, 4.5, 4.6, 4.7**
@settings(max_examples=200)
@given(
    dirs=_dirs,
    base=_note_base,
    ext=_ext,
    heading=st.one_of(st.just(""), _heading),
    label=st.one_of(st.just(""), _label),
)
def test_wiki_link_rewrite(dirs, base, ext, heading, label):
    note = dirs + base + ext
    inner = note + ("#" + heading if heading else "") + ("|" + label if label else "")

    result = rewrite_wiki_link(inner, _ctx())
    assert result is not None
    rw, shown = result

    assert rw.target == _expected_path(note)
    assert "#" not in rw.target
    assert rw.anchor == (bookmark_for_fragment(heading) if heading else None)

    if label:
        assert shown == label
    else:
        # Visible text is the note's last path segment, plus "#Heading".
        expected = note.split("/")[-1] + ("#" + heading if heading else "")
        assert shown == expected


@settings(max_examples=100)
@given(note=_note_base, prefix=_label, suffix=_label)
def test_malformed_wiki_links_stay_plain_text(note, prefix, suffix):
    ctx = _ctx()
    # Rewrite layer: empty note without heading, empty label.
    assert rewrite_wiki_link("", ctx) is None
    assert rewrite_wiki_link(note + "|", ctx) is None

    # Rendering layer: original text kept, no hyperlink written.
    for wiki in ("[[]]", f"[[{note}|]]", f"[[{note}"):
        doc = Document()
        para = doc.add_paragraph()
        add_formatted_text(para, f"{prefix} {wiki} {suffix}", link_ctx=ctx)
        assert para._p.xpath(".//w:hyperlink") == []
        assert wiki in para.text
