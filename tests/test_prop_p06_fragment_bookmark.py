"""Property test: a link fragment equals the target heading's bookmark."""

from hypothesis import assume, given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import (
    bookmark_for_fragment,
    build_heading_anchors,
    slugify_heading,
)

_word = st.text(alphabet="abcXYZ019é日-", min_size=1, max_size=12)
_heading = st.lists(_word, min_size=1, max_size=8).map(" ".join)
# Long headings exercise the 40-character truncate-and-hash rule.
_long_heading = st.lists(_word, min_size=6, max_size=14).map(" ".join)
_any_heading = st.one_of(_heading, _long_heading)


def _blocks(texts):
    return [{"type": "h2", "text": t} for t in texts]


def _base_slug(text):
    return slugify_heading(text) or "section"


# Feature: sibling-markdown-links, Property 6: Fragment equals the heading bookmark
# **Validates: Requirements 2.2, 4.3, 12.9**
@settings(max_examples=200)
@given(texts=st.lists(_any_heading, min_size=1, max_size=6))
def test_fragment_equals_bookmark_for_distinct_headings(texts):
    slugs = [_base_slug(t) for t in texts]
    assume(len(set(slugs)) == len(slugs))
    blocks = _blocks(texts)
    build_heading_anchors(blocks)
    for block, text, slug in zip(blocks, texts, slugs):
        bookmark = block["_bookmark"]
        assert bookmark_for_fragment(slug) == bookmark
        assert bookmark_for_fragment(text) == bookmark


# Feature: sibling-markdown-links, Property 6: Fragment equals the heading bookmark
# **Validates: Requirements 2.2, 4.3, 12.9**
@settings(max_examples=100)
@given(text=_any_heading, copies=st.integers(min_value=2, max_value=4))
def test_numbered_duplicate_slug_equals_bookmark(text, copies):
    blocks = _blocks([text] * copies)
    build_heading_anchors(blocks)
    base = _base_slug(text)
    assert bookmark_for_fragment(text) == blocks[0]["_bookmark"]
    for n, block in enumerate(blocks):
        slug = base if n == 0 else f"{base}-{n}"
        assert bookmark_for_fragment(slug) == block["_bookmark"]
