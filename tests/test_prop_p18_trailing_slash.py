"""Property test: a trailing slash on the base URL does not change the result."""

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import normalize_base_url, resolve_absolute

_seg = st.text(
    alphabet=st.characters(
        min_codepoint=33, max_codepoint=0x2FF, blacklist_characters="/\\?#%"
    ),
    min_size=1,
    max_size=8,
)
_path_seg = st.one_of(_seg, st.just(".."), st.just("."), st.just(""))

_bases = st.builds(
    lambda scheme, host, segs: normalize_base_url(
        f"{scheme}://{host}" + "".join("/" + s for s in segs)
    ),
    st.sampled_from(["http", "https"]),
    st.from_regex(r"[a-z][a-z0-9]{0,7}(\.[a-z]{2,4}){0,2}", fullmatch=True),
    st.lists(st.text(alphabet="abcXYZ09-_ ", min_size=1, max_size=6), max_size=3),
)
_source_dirs = st.lists(_seg.filter(lambda s: s not in (".", "..")), max_size=4).map(
    tuple
)
_docx_paths = st.lists(_path_seg, min_size=1, max_size=6).map(
    lambda s: "/".join(s) + ".docx"
)


# Feature: sibling-markdown-links, Property 18: Trailing slash does not matter
# **Validates: Requirements 8.8, 13.12**
@settings(max_examples=200)
@given(base=_bases, source_dir=_source_dirs, docx_path=_docx_paths)
def test_trailing_slash_does_not_matter(base, source_dir, docx_path):
    without = resolve_absolute(base, source_dir, docx_path)
    with_slash = resolve_absolute(base + "/", source_dir, docx_path)
    assert without == with_slash
    assert "//" not in without[0].split("://", 1)[1]
