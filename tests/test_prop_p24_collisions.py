"""Property test: files sharing a directory and stem are reported as collisions."""

from pathlib import Path

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import output_relpath, plan_outputs

_name = st.text(alphabet="abcxyz019_-", min_size=1, max_size=5)
_dirs = st.lists(_name, max_size=2).map("/".join)  # "" is the root
_exts = st.sets(st.sampled_from([".md", ".markdown"]), min_size=1)

# {(dir, stem): set of extensions}
_trees = st.dictionaries(st.tuples(_dirs, _name), _exts, min_size=1, max_size=8)


def _rel(d, stem, ext):
    return f"{d}/{stem}{ext}" if d else f"{stem}{ext}"


# Feature: sibling-markdown-links, Property 24: Stem collisions are detected
# **Validates: Requirements 11.11**
@settings(max_examples=200, deadline=None)
@given(tree=_trees, use_outdir=st.booleans())
def test_stem_collisions_detected(tree, use_outdir):
    rels = sorted(_rel(d, s, e) for (d, s), exts in tree.items() for e in exts)
    outdir = Path("/out") if use_outdir else None
    mapping, collisions = plan_outputs(rels, Path("/in"), outdir)

    expected = {
        (d, s): tuple(sorted(_rel(d, s, e) for e in exts))
        for (d, s), exts in tree.items()
        if len(exts) > 1
    }
    assert {(c.dir, c.stem): c.files for c in collisions} == expected

    colliding = {f for files in expected.values() for f in files}
    assert colliding.isdisjoint(mapping)
    assert set(mapping) == set(rels) - colliding

    base = outdir if outdir is not None else Path("/in")
    for rel, out in mapping.items():
        assert out == base / output_relpath(rel)
