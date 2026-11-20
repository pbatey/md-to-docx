"""Property test: one warning per occurrence of a link to a missing target."""

import tempfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import LinkContext, add_formatted_text

# Lowercase only so names stay distinct on case-insensitive file systems.
_stem = st.text(alphabet="abcdefgh0123", min_size=1, max_size=6)
_fragment = st.one_of(st.just(""), st.sampled_from(["#intro", "#Some%20Heading"]))

# (existing stems, missing stems, shuffled occurrence list of (is_missing, stem, fragment))
_cases = st.tuples(
    st.sets(_stem, max_size=3),
    st.sets(_stem, max_size=3),
).flatmap(
    lambda sets: st.tuples(
        st.just(sorted("e" + s for s in sets[0])),
        st.just(sorted("m" + s for s in sets[1])),
    )
).flatmap(
    lambda names: st.tuples(
        st.just(names[0]),
        st.just(names[1]),
        st.lists(
            st.tuples(
                st.sampled_from(
                    [(False, n) for n in names[0]] + [(True, n) for n in names[1]]
                    or [(True, "mx")]
                ),
                _fragment,
            ),
            max_size=8,
        ),
    )
)


# Feature: sibling-markdown-links, Property 20: One warning per missing target occurrence
# **Validates: Requirements 6.1, 6.2, 6.3, 6.5, 12.1**
@settings(max_examples=150, deadline=None)
@given(case=_cases)
def test_one_warning_per_missing_occurrence(case):
    existing, missing, occurrences = case
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        for stem in existing:
            (root / f"{stem}.md").write_text("# T\n", encoding="utf-8")

        warnings = []
        source_path = str(root / "source.md")
        ctx = LinkContext(
            enabled=True,
            source_path=source_path,
            source_dir=root,
            warn=warnings.append,
        )

        parts, expected_targets = [], []
        for (is_missing, stem), frag in occurrences:
            raw = f"{stem}.md{frag}"
            parts.append(f"[lbl]({raw})")
            if is_missing and stem in missing:
                expected_targets.append(raw)
            elif is_missing:
                # placeholder name used only when no names were generated
                expected_targets.append(raw)
        text = " and ".join(parts)

        doc = Document()
        paragraph = doc.add_paragraph()
        add_formatted_text(paragraph, text, anchors={}, link_ctx=ctx)

        assert len(warnings) == len(expected_targets)
        assert sorted(
            w.rsplit("link target not found: ", 1)[1] for w in warnings
        ) == sorted(expected_targets)
        for w in warnings:
            assert source_path in w
            assert "link target not found: " in w

        # A hyperlink is written for every Sibling_Link, missing or not.
        assert len(paragraph._p.findall(qn("w:hyperlink"))) == len(occurrences)
