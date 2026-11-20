"""Property test for the sibling-markdown-links path rewrite."""

import re

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import parse_link_target, to_docx_target

_name = st.text(
    alphabet=st.characters(min_codepoint=48, max_codepoint=122, whitelist_categories=("L", "N")),
    min_size=1,
    max_size=8,
)
_segment = st.one_of(_name, st.just(".."), st.just("."))
_ext = st.sampled_from(["md", "markdown"]).flatmap(
    lambda e: st.lists(st.booleans(), min_size=len(e), max_size=len(e)).map(
        lambda flags: "".join(c.upper() if f else c for c, f in zip(e, flags))
    )
)
_fragment = st.text(alphabet="abcdefXYZ0-_", max_size=8)


@st.composite
def md_paths(draw):
    segs = draw(st.lists(_segment, max_size=5))
    stem = draw(_name)
    ext = draw(_ext)
    prefix = draw(st.sampled_from(["", "./"]))
    return prefix + "".join(s + "/" for s in segs) + stem + "." + ext


# Feature: sibling-markdown-links, Property 2: Path rewrite preserves everything except the extension
# **Validates: Requirements 1.2, 1.7, 1.8, 9.1**
@settings(max_examples=200)
@given(path=md_paths())
def test_path_rewrite_preserves_all_but_extension(path):
    out = to_docx_target(path)
    assert out.endswith(".docx")
    assert not out.endswith(".DOCX") or out.endswith(".docx")

    stripped = path[2:] if path.startswith("./") else path
    base_in = re.sub(r"\.(?:md|markdown)$", "", stripped, flags=re.IGNORECASE)
    assert out[: -len(".docx")] == base_in

    # Segment sequence (including "..", ".") is unchanged apart from the
    # one leading "./" and the last segment's extension.
    in_segs = stripped.split("/")
    out_segs = out.split("/")
    assert in_segs[:-1] == out_segs[:-1]
    assert out_segs[-1] == base_in.split("/")[-1] + ".docx"


# Feature: sibling-markdown-links, Property 2: Path rewrite preserves everything except the extension
# **Validates: Requirements 1.2, 1.7, 1.8, 9.1**
@settings(max_examples=200)
@given(path=md_paths(), fragment=_fragment)
def test_parse_then_rewrite_keeps_path_and_fragment(path, fragment):
    raw = path + ("#" + fragment if fragment else "")
    target = parse_link_target(raw)
    assert target is not None
    assert target.path == path
    assert target.fragment == fragment
    assert to_docx_target(target.path).endswith(".docx")
