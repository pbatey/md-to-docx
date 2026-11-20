"""Property test: climbing above the base root discards the extra ``..`` segments."""

from urllib.parse import urlsplit

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import encode_path, normalize_base_url, resolve_absolute

_seg = st.text(
    alphabet=st.characters(
        min_codepoint=33, max_codepoint=0x2FF, blacklist_characters="/\\?#%"
    ),
    min_size=1,
    max_size=8,
).filter(lambda s: s not in (".", ".."))

_bases = st.builds(
    lambda scheme, host, segs: normalize_base_url(
        f"{scheme}://{host}" + "".join("/" + s for s in segs)
    ),
    st.sampled_from(["http", "https"]),
    st.from_regex(r"[a-z][a-z0-9]{0,7}(\.[a-z]{2,4}){0,2}", fullmatch=True),
    st.lists(st.text(alphabet="abcXYZ09-_ ", min_size=1, max_size=6), max_size=3),
)

_cases = st.tuples(
    st.lists(_seg, max_size=4),  # source_dir
    st.integers(min_value=1, max_value=4),  # extra ".." beyond source_dir depth
    st.lists(_seg, min_size=1, max_size=4),  # segments after climbing
).map(lambda t: (tuple(t[0]), t[1], t[2]))


# Feature: sibling-markdown-links, Property 19: Climbing above the base root
# **Validates: Requirements 8.9, 12.6**
@settings(max_examples=200)
@given(base=_bases, case=_cases)
def test_climbing_above_base_is_discarded(base, case):
    source_dir, extra, tail = case
    docx_path = "/".join(
        [".."] * (len(source_dir) + extra) + tail[:-1] + [tail[-1] + ".docx"]
    )

    url, climbed = resolve_absolute(base, source_dir, docx_path)

    assert climbed is True
    expected_tail = "/".join(encode_path(s) for s in tail[:-1] + [tail[-1] + ".docx"])
    assert url == base + "/" + expected_tail

    parts = urlsplit(url)
    assert parts.scheme in ("http", "https") and parts.netloc
    assert ".." not in url[len(base) + 1 :].split("/")
    assert "//" not in url.split("://", 1)[1]
