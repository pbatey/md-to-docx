import re
import tempfile
import zipfile
from pathlib import Path

from golden_fixtures import load_golden
from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx import build_docx, parse_markdown
from md_to_docx.md_to_docx import LinkContext

GOLDEN = load_golden()


def _norm(xml: str) -> str:
    """Renumber bookmark w:id values by order of appearance.

    The ids come from a process-global counter, so they depend on how many
    documents were converted earlier in the process (the golden files were
    generated in a single run). Everything else is compared verbatim.
    """
    seen: dict[str, str] = {}

    def sub(m):
        return f'{m.group(1)}w:id="{seen.setdefault(m.group(2), str(len(seen) + 1))}"'

    return re.sub(r'(<w:bookmark(?:Start|End) )w:id="(\d+)"', sub, xml)


def _convert(markdown: str, link_ctx) -> str:
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        build_docx(
            parse_markdown(markdown), str(out), link_ctx=link_ctx, announce=False
        )
        with zipfile.ZipFile(out) as z:
            return _norm(z.read("word/document.xml").decode("utf-8"))


_base_urls = st.one_of(
    st.none(),
    st.sampled_from(
        ["https://example.com", "https://example.com/docs", "http://host:8080/a/b"]
    ),
)


def _disabled_ctx(base_url):
    return LinkContext(enabled=False, base_url=base_url, source_path="src.md")


_fragments = st.sampled_from(
    [
        "# Title\n\n",
        "## Setup\n\n",
        "## Setup\n\n",
        "[a](other.md) ",
        "[b](./sub/page.md#sec) ",
        "[c](#setup) ",
        "[d](https://example.com/x.md) ",
        "[[Page]] ",
        "[[Page#Heading|alias]] ",
        "[[#Setup]] ",
        "plain text ",
        "\n\n",
        "- item [li](sibling.md)\n",
    ]
)
_documents = st.lists(_fragments, min_size=1, max_size=12).map("".join)


# Feature: sibling-markdown-links, Property 12: --no-sibling-links reproduces the pre-feature output
# **Validates: Requirements 7.2, 8.13, 2.9, 13.19**
@settings(max_examples=100, deadline=None)
@given(
    name=st.sampled_from(sorted(GOLDEN)),
    base_url=_base_urls,
    use_none=st.booleans(),
)
def test_opt_out_matches_golden(name, base_url, use_none):
    md, expected = GOLDEN[name]
    expected = _norm(expected)
    ctx = None if use_none else _disabled_ctx(base_url)
    assert _convert(md, ctx) == expected


# Feature: sibling-markdown-links, Property 12: --no-sibling-links reproduces the pre-feature output
# **Validates: Requirements 7.2, 8.13, 2.9, 13.19**
@settings(max_examples=100, deadline=None)
@given(md=_documents, base_url=_base_urls)
def test_disabled_context_equals_no_context(md, base_url):
    assert _convert(md, _disabled_ctx(base_url)) == _convert(md, None)


def test_all_golden_with_and_without_base_url():
    for name, (md, expected) in GOLDEN.items():
        expected = _norm(expected)
        for base in (None, "https://example.com/docs"):
            assert _convert(md, _disabled_ctx(base)) == expected, name
        assert _convert(md, None) == expected, name
