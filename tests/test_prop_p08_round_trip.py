"""Property test for docx path round trip."""

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import encode_path, percent_decode, to_docx_target

# Feature: sibling-markdown-links, Property 8: Docx path round trip

_text = st.text(alphabet=st.characters(blacklist_categories=("Cs",)), max_size=40)


@st.composite
def md_paths(draw):
    """Paths ending in lowercase .md that do not start with './'."""
    stem = draw(_text)
    path = stem + ".md"
    if path.startswith("./"):
        path = "x" + path
    return path


@settings(max_examples=200)
@given(md_paths())
def test_docx_target_round_trip(path):
    """**Validates: Requirements 13.9**"""
    docx = to_docx_target(path)
    assert docx.endswith(".docx")
    assert docx[: -len(".docx")] + ".md" == path


@settings(max_examples=200)
@given(md_paths())
def test_encode_decode_round_trip_of_docx_path(path):
    """**Validates: Requirements 13.9, 13.11**"""
    docx = to_docx_target(path)
    assert percent_decode(encode_path(docx)) == docx
    restored = percent_decode(encode_path(docx))
    assert restored[: -len(".docx")] + ".md" == path
