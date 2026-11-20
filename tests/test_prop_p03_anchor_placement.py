"""Property test for anchor placement in sibling Markdown / wiki links."""

from docx import Document
from docx.oxml.ns import qn
from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import (
    LinkContext,
    add_formatted_text,
    bookmark_for_fragment,
)

_name = st.text(alphabet="abcdefXYZ019", min_size=1, max_size=8)
_path = st.builds(
    lambda segs, stem, ext: "".join(s + "/" for s in segs) + stem + "." + ext,
    st.lists(st.one_of(_name, st.just("..")), max_size=3),
    _name,
    st.sampled_from(["md", "markdown", "MD"]),
)

# Fragments: empty, spaces, '%' sequences (valid and stray), non-ASCII.
_frag_alphabet = st.sampled_from(
    list("abcXYZ019 -_%") + ["%20", "%C3%A9", "é", "ü", "日本", "%zz"]
)
_fragment = (
    st.lists(_frag_alphabet, max_size=6).map("".join).filter(lambda f: f == f.strip())
)
_base_url = st.one_of(st.none(), st.just("https://example.com/docs"))


def _render(text, base_url):
    ctx = LinkContext(enabled=True, base_url=base_url, warn=lambda _m: None)
    doc = Document()
    paragraph = doc.add_paragraph()
    add_formatted_text(paragraph, text, anchors={}, link_ctx=ctx)
    links = paragraph._p.findall(qn("w:hyperlink"))
    rels = [r for r in paragraph.part.rels.values() if r.reltype.endswith("/hyperlink")]
    return links, rels


def _check(links, rels, fragment):
    assert len(links) == 1
    assert rels, "expected a hyperlink relationship"
    for rel in rels:
        assert "#" not in rel.target_ref
    anchor = links[0].get(qn("w:anchor"))
    if fragment:
        assert anchor == bookmark_for_fragment(fragment)
    else:
        assert anchor is None


# Feature: sibling-markdown-links, Property 3: Anchor placement
# **Validates: Requirements 1.3, 2.1, 2.3, 2.4, 2.5, 2.6, 4.3, 8.6, 13.16**
@settings(max_examples=150, deadline=None)
@given(path=_path, fragment=_fragment, base_url=_base_url)
def test_markdown_link_anchor_placement(path, fragment, base_url):
    raw = path + ("#" + fragment if fragment else "")
    links, rels = _render(f"[label]({raw})", base_url)
    _check(links, rels, fragment)


# Feature: sibling-markdown-links, Property 3: Anchor placement
# **Validates: Requirements 1.3, 2.1, 2.3, 2.4, 2.5, 2.6, 4.3, 8.6, 13.16**
@settings(max_examples=150, deadline=None)
@given(
    note=_name,
    fragment=_fragment,
    alias=st.one_of(st.none(), _name),
    base_url=_base_url,
)
def test_wiki_link_anchor_placement(note, fragment, alias, base_url):
    inner = note + ("#" + fragment if fragment else "")
    if alias is not None:
        inner += "|" + alias
    links, rels = _render(f"[[{inner}]]", base_url)
    _check(links, rels, fragment)
