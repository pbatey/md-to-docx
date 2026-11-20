"""Property test for relative links resolving to the linked output (model tree)."""

import posixpath
from pathlib import Path
from urllib.parse import unquote

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import LinkContext, plan_outputs, rewrite_link

_seg = st.text(alphabet="abcdefghij0123456789", min_size=1, max_size=4)
_rel = st.builds(
    lambda dirs, name: "/".join([*dirs, name + ".md"]),
    st.lists(_seg, max_size=3),
    _seg,
)
_tree = st.lists(_rel, min_size=1, max_size=8, unique=True)


# Feature: sibling-markdown-links, Property 23: Relative links resolve to the linked output
# **Validates: Requirements 12.7**
@settings(max_examples=200, deadline=None)
@given(rels=_tree, data=st.data())
def test_relative_link_resolves_to_linked_output(rels, data):
    outdir = Path("/out")
    mapping, collisions = plan_outputs(rels, Path("/in"), outdir)
    assert not collisions
    s = data.draw(st.sampled_from(rels))
    t = data.draw(st.sampled_from(rels))

    link = posixpath.relpath(t, posixpath.dirname(s) or ".")
    rewritten = rewrite_link(link, LinkContext(source_path=s))
    assert rewritten is not None

    s_out = mapping[s].as_posix()
    resolved = posixpath.normpath(
        posixpath.join(posixpath.dirname(s_out), unquote(rewritten.target))
    )
    assert resolved == mapping[t].as_posix()
