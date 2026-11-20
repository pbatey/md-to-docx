"""Property 1: Sibling classification.

Feature: sibling-markdown-links, Property 1: Sibling classification

**Validates: Requirements 1.9, 3.1, 3.3, 3.4, 3.5**
"""

import re

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import SiblingTarget, parse_link_target

_ALPHABET = st.characters(
    blacklist_characters="%<>\\\x00",
    blacklist_categories=("Cs",),
)
_text = st.text(_ALPHABET, max_size=20)
_ext = st.sampled_from(
    ["", ".md", ".markdown", ".MD", ".Md", ".MarkDown", ".txt", ".mdx", ".md.bak"]
)
_prefix = st.sampled_from(
    ["", "#", "/", "http://", "https://x/", "mailto:", "a+b.c-d:", "./", "../", "sub/"]
)
_frag = st.one_of(st.just(""), _text.map(lambda t: "#" + t.replace("#", "")))
_wrap = st.booleans()
_pad = st.sampled_from(["", " ", "\t", "  "])


def _expected(raw: str) -> bool:
    s = raw.strip()
    if s.startswith("<"):
        end = s.find(">")
        if end == -1:
            return False
        s = s[1:end]
    if not s or s.startswith(("#", "/")):
        return False
    if re.match(r"^[A-Za-z][A-Za-z0-9+.\-]*:", s):
        return False
    path = s.split("#", 1)[0]
    return path.lower().endswith((".md", ".markdown"))


@settings(max_examples=300)
@given(_pad, _wrap, _prefix, _text, _ext, _frag, _pad)
def test_classification_iff(lpad, wrap, prefix, body, ext, frag, rpad):
    # Feature: sibling-markdown-links, Property 1: Sibling classification
    core = prefix + body + ext + frag
    raw = lpad + (f"<{core}>" if wrap else core) + rpad
    result = parse_link_target(raw)
    assert (result is not None) == _expected(raw)
    if result is not None:
        assert isinstance(result, SiblingTarget)


@settings(max_examples=200)
@given(st.text(max_size=40))
def test_never_raises_on_arbitrary_text(raw):
    # Feature: sibling-markdown-links, Property 1: Sibling classification
    result = parse_link_target(raw)
    assert result is None or isinstance(result, SiblingTarget)


def test_examples():
    # Feature: sibling-markdown-links, Property 1: Sibling classification
    assert parse_link_target("") is None
    assert parse_link_target("   ") is None
    assert parse_link_target("#a.md") is None
    assert parse_link_target("/a.md") is None
    assert parse_link_target("https://x/a.md") is None
    assert parse_link_target("a.txt") is None
    assert parse_link_target("a.MD#h").path == "a.MD"
    assert parse_link_target("<my note.markdown>").path == "my note.markdown"
