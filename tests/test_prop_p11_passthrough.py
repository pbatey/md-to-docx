"""Property test: non-sibling links render exactly as without a link context."""

from docx import Document
from hypothesis import assume, given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import (
    LinkContext,
    add_formatted_text,
    parse_link_target,
)

_name = st.text(alphabet="abcdefXYZ0123-_", min_size=1, max_size=8)
_label = st.text(alphabet="abc XYZ", min_size=1, max_size=10).filter(
    lambda s: s.strip()
)

_external = st.builds(
    lambda scheme, host, path: f"{scheme}://{host}.example/{path}",
    st.sampled_from(["http", "https"]),
    _name,
    st.one_of(
        _name,
        st.builds(lambda n, e: n + "." + e, _name, st.sampled_from(["md", "html"])),
    ),
)
_mailto = _name.map(lambda n: f"mailto:{n}@example.com")
_anchor = _name.map(lambda n: "#" + n)
_non_md = st.builds(
    lambda d, n, e: "".join(x + "/" for x in d) + n + "." + e,
    st.lists(_name, max_size=3),
    _name,
    st.sampled_from(["png", "pdf", "txt", "html", "docx", "mdx", "py"]),
)
_absolute = st.builds(
    lambda d, n, e: "/" + "".join(x + "/" for x in d) + n + "." + e,
    st.lists(_name, max_size=3),
    _name,
    st.sampled_from(["md", "markdown", "pdf"]),
)
_target = st.one_of(_external, _mailto, _anchor, _non_md, _absolute)

_ctx = st.one_of(
    st.just(LinkContext(base_url=None, warn=lambda m: None)),
    st.just(LinkContext(base_url="https://example.com/docs", warn=lambda m: None)),
    st.just(
        LinkContext(
            base_url="https://example.com/docs",
            source_rel_dir=("a", "b"),
            warn=lambda m: None,
        )
    ),
)


def _render(text, ctx):
    doc = Document()
    para = doc.add_paragraph()
    add_formatted_text(para, text, anchors=None, link_ctx=ctx)
    rels = sorted(
        (r.reltype, r.target_ref, r.is_external) for r in doc.part.rels.values()
    )
    return para._p.xml, rels


# Feature: sibling-markdown-links, Property 11: Non-sibling links render as before
# **Validates: Requirements 3.6, 8.14**
@settings(max_examples=100, deadline=None)
@given(target=_target, label=_label, ctx=_ctx)
def test_non_sibling_links_render_as_before(target, label, ctx):
    assume(parse_link_target(target) is None)
    text = f"see [{label}]({target}) end"
    assert _render(text, ctx) == _render(text, None)
