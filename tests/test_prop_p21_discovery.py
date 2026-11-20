# Feature: sibling-markdown-links, Property 21: Tree discovery
"""Property test for discover_tree.

**Validates: Requirements 10.8, 10.9, 10.11**
"""

import tempfile
from pathlib import Path

from hypothesis import given, settings, strategies as st

from md_to_docx.md_to_docx import discover_tree

SKIPPED = [".git", ".venv", "node_modules", "venv", "__pycache__"]
STEM = st.text(alphabet="abcXYZ01_-", min_size=1, max_size=4)
DIR = st.one_of(STEM, st.sampled_from(SKIPPED))
EXT = st.sampled_from(
    [".md", ".MD", ".Md", ".markdown", ".MarkDown", ".txt", ".docx", ".mdx", ""]
)
HIDDEN = st.booleans()


@st.composite
def entries(draw):
    dirs = draw(st.lists(DIR, max_size=3))
    stem = draw(STEM)
    ext = draw(EXT)
    hidden = draw(HIDDEN)
    return dirs, ("." if hidden else "") + stem + ext


def _is_md(name):
    return name.lower().endswith((".md", ".markdown"))


@settings(max_examples=100, deadline=None)
@given(st.lists(entries(), max_size=15))
def test_discover_tree_returns_exactly_visible_markdown(items):
    expected = set()
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        for dirs, fname in items:
            d = root.joinpath(*dirs)
            d.mkdir(parents=True, exist_ok=True)
            p = d / fname
            if p.is_dir():
                continue
            p.write_text("x")
            skipped = any(x in SKIPPED for x in dirs)
            if _is_md(fname) and not fname.startswith(".") and not skipped:
                # Case-insensitive filesystems may merge names; use real name.
                expected.add(p.relative_to(root).as_posix())
        result = discover_tree(root)
        actual_lower = {r.lower() for r in result}
        assert actual_lower == {e.lower() for e in expected}
        assert result == sorted(result)
        assert len(result) == len(set(result))
        assert all("\\" not in r and not r.startswith("/") for r in result)
