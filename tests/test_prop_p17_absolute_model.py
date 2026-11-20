"""Property 17: Absolute link matches the path-resolution model.

Feature: sibling-markdown-links, Property 17: Absolute link matches the path-resolution model
**Validates: Requirements 8.4, 8.5, 8.6, 12.3, 12.4, 13.13**
"""
import posixpath

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import encode_path, normalize_base_url, resolve_absolute

_name = st.text(
    alphabet=st.characters(
        blacklist_characters="/\\\x00",
        blacklist_categories=("Cs", "Cc"),
    ),
    min_size=1,
    max_size=8,
).filter(lambda s: s not in (".", ".."))

_host = st.from_regex(r"[a-z][a-z0-9]{0,8}(\.[a-z]{2,4})?", fullmatch=True)
_base_path = st.lists(st.from_regex(r"[A-Za-z0-9_-]{1,6}", fullmatch=True), max_size=3)


@st.composite
def _base(draw):
    scheme = draw(st.sampled_from(["http", "https"]))
    host = draw(_host)
    segs = draw(_base_path)
    raw = f"{scheme}://{host}" + "".join("/" + s for s in segs)
    return normalize_base_url(raw + draw(st.sampled_from(["", "/"])))


@st.composite
def _case(draw):
    source_dir = tuple(draw(st.lists(_name, max_size=4)))
    depth = len(source_dir)
    segs = []
    for _ in range(draw(st.integers(0, 6))):
        kind = draw(st.sampled_from(["name", "dot", "up", "empty"]))
        if kind == "up" and depth > 0:
            segs.append("..")
            depth -= 1
        elif kind == "dot":
            segs.append(".")
        elif kind == "empty" and segs:  # a leading "" would make an absolute path
            segs.append("")
        else:
            segs.append(draw(_name))
            depth += 1
    segs.append(draw(_name))  # always end on a real file name
    return source_dir, "/".join(segs)


@settings(max_examples=200, deadline=None)
@given(base=_base(), case=_case())
def test_absolute_link_matches_path_model(base, case):
    source_dir, docx_path = case
    url, climbed = resolve_absolute(base, source_dir, docx_path)

    joined = posixpath.join("/".join(source_dir), docx_path)
    expected = base + "/" + encode_path(posixpath.normpath(joined))

    assert url == expected
    assert not climbed
    tail = url[len(base) + 1 :]
    assert url.startswith(base + "/") and not tail.startswith("/")
    assert "#" not in url
    assert not {".", ".."} & set(tail.split("/"))
