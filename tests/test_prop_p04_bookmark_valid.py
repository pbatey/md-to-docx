# Feature: sibling-markdown-links, Property 4: Bookmark names are valid Word bookmark names
"""**Validates: Requirements 2.7, 13.17**"""

import re

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import (
    bookmark_for_fragment,
    bookmark_name,
    build_heading_anchors,
)

VALID = re.compile(r"^_[A-Za-z0-9_]{1,39}$")

text_st = st.text(max_size=120)
# Bias toward long, shared-prefix and repeated headings.
heading_st = st.one_of(
    text_st,
    st.text(alphabet="abc -", min_size=30, max_size=80),
    st.sampled_from(["Install", "install", "Intro", "a-b", "a_b", "日本語", "!!!"]),
)


@settings(max_examples=200)
@given(text_st)
def test_bookmark_name_valid(slug):
    assert VALID.match(bookmark_name(slug))


@settings(max_examples=200)
@given(text_st)
def test_bookmark_for_fragment_valid(fragment):
    assert VALID.match(bookmark_for_fragment(fragment))


@settings(max_examples=200)
@given(st.lists(st.tuples(st.integers(1, 6), heading_st), max_size=15))
def test_every_heading_anchor_valid(headings):
    blocks = [{"type": f"h{lvl}", "text": t} for lvl, t in headings]
    anchors = build_heading_anchors(blocks)
    for name in anchors.values():
        assert VALID.match(name)
    for block in blocks:
        assert VALID.match(block["_bookmark"])
