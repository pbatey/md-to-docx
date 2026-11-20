from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import to_docx_target

# Bias generation toward strings that exercise the extension swap.
_pieces = st.sampled_from(
    [".md", ".markdown", ".MD", ".docx", "./", "../", "/", "#", " ", "a", "b", "."]
)
_paths = st.one_of(
    st.text(),
    st.lists(_pieces, max_size=8).map("".join),
)


# Feature: sibling-markdown-links, Property 9: Docx path idempotence
# **Validates: Requirements 13.10, 13.11**
@settings(max_examples=200)
@given(p=_paths)
def test_to_docx_target_is_idempotent(p):
    once = to_docx_target(p)
    assert to_docx_target(once) == once
