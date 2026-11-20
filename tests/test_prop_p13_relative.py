"""Property test for relative links without a base URL."""

import re

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import (
    LinkContext,
    encode_path,
    rewrite_link,
    to_docx_target,
)

_name = st.text(
    alphabet=st.characters(
        min_codepoint=48, max_codepoint=122, whitelist_categories=("L", "N")
    ),
    min_size=1,
    max_size=8,
)
_segment = st.one_of(_name, st.just(".."), st.just("."))
_ext = st.sampled_from(["md", "markdown", "MD", "Markdown"])
_fragment = st.text(alphabet="abcdefXYZ0-_", max_size=8)


@st.composite
def md_paths(draw):
    segs = draw(st.lists(_segment, max_size=5))
    stem = draw(_name)
    prefix = draw(st.sampled_from(["", "./"]))
    return prefix + "".join(s + "/" for s in segs) + stem + "." + draw(_ext)


# Feature: sibling-markdown-links, Property 13: Relative links without a base URL
# **Validates: Requirements 1.6, 8.3, 9.3**
@settings(max_examples=200)
@given(
    path=md_paths(),
    fragment=_fragment,
    rel_dir_a=st.lists(_name, max_size=4).map(tuple),
    rel_dir_b=st.lists(_name, max_size=4).map(tuple),
)
def test_relative_target_without_base_url(path, fragment, rel_dir_a, rel_dir_b):
    raw = path + ("#" + fragment if fragment else "")
    expected = encode_path(to_docx_target(path))

    results = []
    for rel_dir in (rel_dir_a, rel_dir_b):
        ctx = LinkContext(base_url=None, source_rel_dir=rel_dir, warn=lambda m: None)
        out = rewrite_link(raw, ctx)
        assert out is not None
        results.append(out.target)

    for target in results:
        assert not re.match(r"^[A-Za-z][A-Za-z0-9+.\-]*:", target)
        assert not target.startswith("/")
        assert "#" not in target
        assert target == expected
    # Same string regardless of where the output file lives.
    assert results[0] == results[1]
