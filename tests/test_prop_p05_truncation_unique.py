"""Property test for bookmark truncation determinism and uniqueness."""

import re

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import bookmark_name, build_heading_anchors

_PREFIX = "a very long shared heading prefix that exceeds forty characters "

_words = st.text(alphabet="abcXY 09-_", max_size=15)
_headings = st.one_of(
    _words,
    st.builds(lambda s: _PREFIX + s, _words),
    st.sampled_from(["Setup", "setup", "a-b", "a_b", "a b", "Same", "Same"]),
)
_slugs = st.one_of(
    st.builds(lambda s: _PREFIX.replace(" ", "-") + s, _words),
    st.text(alphabet="abc-_09", max_size=80),
)

_HASHED_RE = re.compile(r"^.{31}_[0-9a-f]{8}$")


def _blocks(texts):
    return [{"type": "h2", "text": t} for t in texts]


# Feature: sibling-markdown-links, Property 5: Truncation is deterministic and names are unique
# **Validates: Requirements 2.8, 13.18**
@settings(max_examples=200)
@given(st.lists(_headings, min_size=1, max_size=12))
def test_names_unique_and_deterministic(texts):
    first = _blocks(texts)
    second = _blocks(texts)
    build_heading_anchors(first)
    build_heading_anchors(second)
    names = [b["_bookmark"] for b in first]
    assert names == [b["_bookmark"] for b in second]
    assert len(set(names)) == len(names)
    assert all(len(n) <= 40 for n in names)


# Feature: sibling-markdown-links, Property 5: Truncation is deterministic and names are unique
# **Validates: Requirements 2.8, 13.18**
@settings(max_examples=200)
@given(_slugs, _slugs)
def test_long_names_are_exactly_40_with_hash_suffix(a, b):
    na, nb = bookmark_name(a), bookmark_name(b)
    assert na == bookmark_name(a)
    assert len(na) <= 40 and len(nb) <= 40
    for slug, name in ((a, na), (b, nb)):
        body = re.sub(r"[^A-Za-z0-9]+", "_", slug).strip("_")
        if not body:
            continue
        full = "_toc_" + body
        if len(full) > 40:
            assert len(name) == 40
            assert _HASHED_RE.match(name), name
        else:
            assert name == full
    fa = re.sub(r"[^A-Za-z0-9]+", "_", a).strip("_")
    fb = re.sub(r"[^A-Za-z0-9]+", "_", b).strip("_")
    if fa != fb and len(na) == 40 and len(nb) == 40 and len(fa) > 35 and len(fb) > 35:
        assert na != nb
