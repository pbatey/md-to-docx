"""Property test for the mirrored output set.

Feature: sibling-markdown-links, Property 22: Mirrored output set

**Validates: Requirements 9.5, 10.2, 11.2, 11.4, 13.7**
"""

from pathlib import Path

from hypothesis import given, settings, strategies as st

from md_to_docx.md_to_docx import output_relpath, plan_outputs

_seg = st.text(alphabet="abcdefgh_-0123", min_size=1, max_size=5)
_ext = st.sampled_from([".md", ".markdown"])


@st.composite
def _rels(draw):
    dirs = draw(st.lists(st.lists(_seg, max_size=3), min_size=1, max_size=4))
    stems = draw(st.lists(_seg, min_size=1, max_size=8))
    seen, rels = set(), []
    for stem in stems:
        d = "/".join(draw(st.sampled_from(dirs)))
        if (d, stem) in seen:
            continue
        seen.add((d, stem))
        rels.append((d + "/" if d else "") + stem + draw(_ext))
    return rels


@settings(max_examples=100, deadline=None)
@given(rels=_rels(), use_outdir=st.booleans())
def test_p22_mirrored_output_set(rels, use_outdir):
    root = Path("/in/root")
    outdir = Path("/out/dir") if use_outdir else None
    base = outdir if use_outdir else root

    mapping, collisions = plan_outputs(rels, root, outdir)

    assert collisions == []
    assert set(mapping) == set(rels)
    expected = {Path(output_relpath(r)).as_posix() for r in rels}
    produced = {p.relative_to(base).as_posix() for p in mapping.values()}
    assert produced == expected
    assert len(produced) == len(rels)  # one-to-one
    for rel, out in mapping.items():
        r = Path(rel)
        assert out == base / r.parent / (r.stem + ".docx")
        assert out.suffix == ".docx"
