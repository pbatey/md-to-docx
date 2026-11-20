"""Property test for URL-safe rewritten targets."""

import re

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import encode_path, percent_decode, to_docx_target

_SAFE_RE = re.compile(r"^(?:[A-Za-z0-9\-._~/]|%[0-9A-F]{2})*$")

_path_chars = st.one_of(
    st.sampled_from(list(" \\/%.-_~#?&é日本\u00a0")),
    st.sampled_from(["%20", "%C3%A9", "%zz", "%C3", "%FF", "%5C", "%", "%2"]),
    st.characters(blacklist_categories=("Cs",)),
)
_stems = st.lists(_path_chars, max_size=30).map("".join)
_paths = st.builds(
    lambda stem, ext, dot: ("./" if dot else "") + stem + ext,
    _stems,
    st.sampled_from(["", ".md", ".markdown", ".MD", ".txt"]),
    st.booleans(),
)


# Feature: sibling-markdown-links, Property 7: Rewritten target contains only URL-safe characters
# **Validates: Requirements 5.1, 5.3, 5.4**
@settings(max_examples=200)
@given(_paths)
def test_rewritten_target_is_url_safe(path):
    out = encode_path(to_docx_target(percent_decode(path)))
    assert "\\" not in out
    assert " " not in out
    assert _SAFE_RE.match(out), out
