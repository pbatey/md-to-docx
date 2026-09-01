"""Unit tests for md_to_docx.

The pure parsing/slug helpers are tested directly. A couple of end-to-end tests
run the full conversion and read the resulting .docx back with python-docx to
confirm the document is well-formed and carries the expected content.
"""

from pathlib import Path

import pytest
from docx import Document

from md_to_docx import build_docx, find_template, parse_markdown
from md_to_docx.md_to_docx import build_heading_anchors, slugify_heading


# --------------------------------------------------------------------------- #
# slugify_heading
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize(
    "text, expected",
    [
        ("Hello World", "hello-world"),
        ("Getting Started!", "getting-started"),
        ("**Bold** and `code`", "bold-and-code"),
        ("Multiple   Spaces", "multiple-spaces"),
        ("Trailing --- dashes", "trailing-dashes"),
        ("[link](http://example.com)", "link"),
    ],
)
def test_slugify_heading(text, expected):
    assert slugify_heading(text) == expected


# --------------------------------------------------------------------------- #
# parse_markdown
# --------------------------------------------------------------------------- #
def test_parse_headings():
    blocks = parse_markdown("# H1\n## H2\n### H3\n#### H4")
    assert [b["type"] for b in blocks] == ["h1", "h2", "h3", "h4"]
    assert [b["text"] for b in blocks] == ["H1", "H2", "H3", "H4"]


def test_parse_skips_frontmatter():
    md = "---\ntitle: Test\nauthor: Me\n---\n# Real Heading"
    blocks = parse_markdown(md)
    assert blocks == [{"type": "h1", "text": "Real Heading"}]


def test_parse_bullet_list():
    blocks = parse_markdown("- one\n- two\n- three")
    assert len(blocks) == 1
    assert blocks[0]["type"] == "list"
    # Items are dicts; a plain bullet has checked == None.
    assert [it["text"] for it in blocks[0]["items"]] == ["one", "two", "three"]
    assert all(it["checked"] is None for it in blocks[0]["items"])


def test_parse_numbered_list_with_subitems():
    md = "1. first\n   - sub a\n   - sub b\n2. second"
    blocks = parse_markdown(md)
    assert len(blocks) == 1
    assert blocks[0]["type"] == "numbered_list"
    items = blocks[0]["items"]
    assert items[0]["text"] == "first"
    assert items[0]["sub_items"] == ["sub a", "sub b"]
    assert items[1]["text"] == "second"


def test_parse_blockquote_preserves_lines():
    md = "> line one\n> line two"
    blocks = parse_markdown(md)
    assert blocks[0]["type"] == "blockquote"
    assert blocks[0]["lines"] == ["line one", "line two"]


def test_parse_code_block():
    md = "```\nprint('hi')\nx = 1\n```"
    blocks = parse_markdown(md)
    assert blocks[0]["type"] == "code_block"
    assert blocks[0]["lines"] == ["print('hi')", "x = 1"]


def test_parse_table():
    md = "| A | B |\n|---|---|\n| 1 | 2 |"
    blocks = parse_markdown(md)
    assert blocks[0]["type"] == "table"
    # separator row is kept in raw lines; it's filtered during rendering
    assert blocks[0]["lines"][0].startswith("| A")


def test_parse_hr():
    blocks = parse_markdown("text\n\n---\n\nmore")
    assert any(b["type"] == "hr" for b in blocks)


# --------------------------------------------------------------------------- #
# build_heading_anchors
# --------------------------------------------------------------------------- #
def test_build_heading_anchors_unique_slugs():
    blocks = parse_markdown("# Intro\n## Intro\n## Details")
    anchors = build_heading_anchors(blocks)
    # Both "Intro" headings register; the second gets a numeric suffix.
    assert "intro" in anchors
    assert "intro-1" in anchors
    assert "details" in anchors
    # Each heading block gets a bookmark tag.
    assert all("_bookmark" in b for b in blocks)


# --------------------------------------------------------------------------- #
# find_template
# --------------------------------------------------------------------------- #
def test_find_template_env_override(tmp_path, monkeypatch):
    fake = tmp_path / "custom.docx"
    fake.write_bytes(b"not really a docx")
    monkeypatch.setenv("MD_TO_DOCX_TEMPLATE", str(fake))
    assert find_template() == fake


def test_find_template_default_exists():
    # The repo/package ships a template; it should resolve without an override.
    template = find_template()
    assert template is not None
    assert template.exists()
    assert template.name == "md-template.docx"


def test_find_template_sibling_file_first(tmp_path, monkeypatch):
    """A md-template.docx sitting directly beside the script wins over the
    bundled templates/ copy (candidate 2 in the resolution order)."""
    import md_to_docx.md_to_docx as mod

    # Point the module's __file__ at a temp dir and drop a template beside it.
    fake_script = tmp_path / "md_to_docx.py"
    fake_script.write_text("# fake")
    sibling = tmp_path / "md-template.docx"
    sibling.write_bytes(b"stub")
    monkeypatch.setattr(mod, "__file__", str(fake_script))
    monkeypatch.delenv("MD_TO_DOCX_TEMPLATE", raising=False)
    # Run from a CWD with no templates/ so only script-relative lookups apply.
    monkeypatch.chdir(tmp_path)

    assert find_template() == sibling


# --------------------------------------------------------------------------- #
# End-to-end conversion
# --------------------------------------------------------------------------- #
def test_build_docx_roundtrip(tmp_path):
    md = (
        "# Title\n\n"
        "Some **bold** and *italic* text.\n\n"
        "## Section\n\n"
        "- item one\n- item two\n\n"
        "> a quote line\n\n"
        "| Col A | Col B |\n|-------|-------|\n| 1 | 2 |\n"
    )
    blocks = parse_markdown(md)
    out = tmp_path / "out.docx"
    build_docx(blocks, str(out), title="T", author="A", date="today")

    assert out.exists()
    doc = Document(str(out))
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "Title" in text
    assert "Section" in text
    assert "item one" in text
    # Table content survives the round-trip.
    assert len(doc.tables) == 1
    cell_texts = {c.text for row in doc.tables[0].rows for c in row.cells}
    assert {"Col A", "Col B", "1", "2"}.issubset(cell_texts)


def test_build_docx_internal_link_resolves(tmp_path):
    md = "# Overview\n\nJump to [Details](#details).\n\n## Details\n\nHere."
    blocks = parse_markdown(md)
    out = tmp_path / "linked.docx"
    build_docx(blocks, str(out), title="T", author="A", date="today")

    doc = Document(str(out))
    # An internal hyperlink (w:hyperlink with w:anchor) should exist in the body.
    xml = doc.element.body.xml
    assert "w:anchor" in xml


# --------------------------------------------------------------------------- #
# Helpers for inline-rendering tests
# --------------------------------------------------------------------------- #
def _render_para(md, tmp_path, base_dir=None):
    """Convert a markdown string and return the resulting Document."""
    blocks = parse_markdown(md)
    out = tmp_path / "r.docx"
    build_docx(blocks, str(out), title="T", author="A", date="d", base_dir=base_dir)
    return Document(str(out))


def _all_runs(doc):
    for p in doc.paragraphs:
        for r in p.runs:
            yield r


def _write_tiny_png(path, w=2, h=2):
    """Write a small but valid RGB PNG that python-docx can parse."""
    import struct
    import zlib

    def _chunk(typ, data):
        body = typ + data
        return struct.pack(">I", len(data)) + body + struct.pack(
            ">I", zlib.crc32(body) & 0xFFFFFFFF)

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)  # 8-bit RGB
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * w for _ in range(h))
    png = (sig + _chunk(b"IHDR", ihdr)
           + _chunk(b"IDAT", zlib.compress(raw))
           + _chunk(b"IEND", b""))
    Path(path).write_bytes(png)


# --------------------------------------------------------------------------- #
# R7: Backslash escapes
# --------------------------------------------------------------------------- #
def test_escape_prevents_emphasis(tmp_path):
    doc = _render_para(r"\*not italic\*", tmp_path)
    text = "".join(r.text for r in _all_runs(doc))
    assert text == "*not italic*"
    assert not any(r.italic for r in _all_runs(doc))


def test_escape_backtick_and_tilde(tmp_path):
    doc = _render_para(r"a \`b\` c \~\~d\~\~", tmp_path)
    text = "".join(r.text for r in _all_runs(doc))
    assert "`b`" in text
    assert "~~d~~" in text
    assert not any(r.font.strike for r in _all_runs(doc))


# --------------------------------------------------------------------------- #
# R1: Headings 5 and 6
# --------------------------------------------------------------------------- #
def test_parse_headings_5_6():
    blocks = parse_markdown("##### H5\n###### H6")
    assert [b["type"] for b in blocks] == ["h5", "h6"]
    assert [b["text"] for b in blocks] == ["H5", "H6"]


def test_headings_5_6_register_anchors():
    blocks = parse_markdown("##### Deep Section")
    anchors = build_heading_anchors(blocks)
    assert "deep-section" in anchors
    assert blocks[0].get("_bookmark")


def test_headings_5_6_render(tmp_path):
    doc = _render_para("##### Five\n\n###### Six", tmp_path)
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "Five" in text
    assert "Six" in text


# --------------------------------------------------------------------------- #
# R3: Strikethrough
# --------------------------------------------------------------------------- #
def test_strikethrough(tmp_path):
    doc = _render_para("~~gone~~", tmp_path)
    struck = [r for r in _all_runs(doc) if r.text == "gone"]
    assert struck and struck[0].font.strike


def test_strikethrough_stacks_with_bold(tmp_path):
    doc = _render_para("**~~x~~**", tmp_path)
    runs = [r for r in _all_runs(doc) if r.text == "x"]
    assert runs and runs[0].font.strike and runs[0].bold


def test_single_tilde_is_literal(tmp_path):
    doc = _render_para("a ~ b", tmp_path)
    text = "".join(r.text for r in _all_runs(doc))
    assert "~" in text
    assert not any(r.font.strike for r in _all_runs(doc))


# --------------------------------------------------------------------------- #
# R4: Code fence language info string
# --------------------------------------------------------------------------- #
def test_code_fence_language_captured():
    blocks = parse_markdown("```python\nprint(1)\n```")
    assert blocks[0]["type"] == "code_block"
    assert blocks[0]["language"] == "python"
    assert blocks[0]["lines"] == ["print(1)"]


def test_code_fence_no_language():
    blocks = parse_markdown("```\nplain\n```")
    assert blocks[0]["language"] == ""


def test_code_block_has_padded_shaded_box(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("```\ncode line\n```", tmp_path)
    # Find the paragraph that carries the code text.
    code_p = None
    for p in doc.paragraphs:
        if "code line" in p.text:
            code_p = p
            break
    assert code_p is not None
    pPr = code_p._p.find(qn("w:pPr"))
    assert pPr is not None
    # Shading fill present.
    shd = pPr.find(qn("w:shd"))
    assert shd is not None and shd.get(qn("w:fill")) == "F2F2F2"
    # Four-sided border, same color as the fill, with a nonzero space (padding).
    pBdr = pPr.find(qn("w:pBdr"))
    assert pBdr is not None
    for side in ("top", "left", "bottom", "right"):
        el = pBdr.find(qn(f"w:{side}"))
        assert el is not None, f"missing {side} border"
        assert el.get(qn("w:color")) == "F2F2F2"
        assert int(el.get(qn("w:space"))) > 0

    # And it should no longer carry the old left indent.
    assert code_p.paragraph_format.left_indent is None


# --------------------------------------------------------------------------- #
# R5: Task list checkboxes
# --------------------------------------------------------------------------- #
def test_parse_task_items():
    blocks = parse_markdown("- [ ] todo\n- [x] done\n- normal")
    items = blocks[0]["items"]
    assert items[0] == {"text": "todo", "checked": False}
    assert items[1] == {"text": "done", "checked": True}
    assert items[2] == {"text": "normal", "checked": None}


def test_task_items_render_glyphs(tmp_path):
    doc = _render_para("- [ ] todo\n- [x] done", tmp_path)
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "\u2610" in text  # unchecked box
    assert "\u2611" in text  # checked box
    assert "todo" in text and "done" in text


# --------------------------------------------------------------------------- #
# R2: Images
# --------------------------------------------------------------------------- #
def test_image_embeds_local_file(tmp_path):
    img = tmp_path / "pic.png"
    _write_tiny_png(img)
    doc = _render_para("![a picture](pic.png)", tmp_path, base_dir=tmp_path)
    assert len(doc.inline_shapes) == 1


def test_image_missing_falls_back_to_alt(tmp_path):
    doc = _render_para("![missing alt](nope.png)", tmp_path, base_dir=tmp_path)
    text = "".join(r.text for r in _all_runs(doc))
    assert "missing alt" in text
    assert len(doc.inline_shapes) == 0


def test_image_remote_falls_back_to_alt(tmp_path):
    doc = _render_para("![remote](https://example.com/x.png)", tmp_path, base_dir=tmp_path)
    text = "".join(r.text for r in _all_runs(doc))
    assert "remote" in text
    assert len(doc.inline_shapes) == 0


def _render_para_remote(md, tmp_path, allow_remote_images):
    """Like _render_para but toggles remote image fetching."""
    from md_to_docx import build_docx as _build
    blocks = parse_markdown(md)
    out = tmp_path / "r.docx"
    _build(blocks, str(out), title="T", author="A", date="d",
           base_dir=tmp_path, allow_remote_images=allow_remote_images)
    return Document(str(out))


def test_remote_image_embedded_when_allowed(tmp_path, monkeypatch):
    import io
    import md_to_docx.md_to_docx as mod

    # Serve a valid PNG from a fake fetcher — no real network.
    png = tmp_path / "buf.png"
    _write_tiny_png(png)
    payload = png.read_bytes()
    monkeypatch.setattr(mod, "_fetch_remote_image",
                        lambda url: io.BytesIO(payload))

    doc = _render_para_remote("![pic](https://example.com/a.png)", tmp_path,
                              allow_remote_images=True)
    assert len(doc.inline_shapes) == 1


def test_remote_image_not_fetched_when_disabled(tmp_path, monkeypatch):
    import md_to_docx.md_to_docx as mod

    called = {"n": 0}

    def _spy(url):
        called["n"] += 1
        raise AssertionError("should not fetch when disabled")

    monkeypatch.setattr(mod, "_fetch_remote_image", _spy)

    doc = _render_para_remote("![pic alt](https://example.com/a.png)", tmp_path,
                              allow_remote_images=False)
    assert called["n"] == 0
    assert len(doc.inline_shapes) == 0
    assert "pic alt" in "".join(r.text for r in _all_runs(doc))


def test_remote_image_fetch_failure_falls_back(tmp_path, monkeypatch):
    import md_to_docx.md_to_docx as mod

    def _boom(url):
        raise OSError("network down")

    monkeypatch.setattr(mod, "_fetch_remote_image", _boom)

    doc = _render_para_remote("![pic alt](https://example.com/a.png)", tmp_path,
                              allow_remote_images=True)
    assert len(doc.inline_shapes) == 0
    assert "pic alt" in "".join(r.text for r in _all_runs(doc))


# --------------------------------------------------------------------------- #
# R6: Autolinks and bare URLs
# --------------------------------------------------------------------------- #
def test_bare_url_becomes_hyperlink(tmp_path):
    doc = _render_para("see https://example.com for more", tmp_path)
    xml = doc.element.body.xml
    assert "hyperlink" in xml
    # The trailing word should remain text, not part of the URL.
    assert "for more" in "".join(r.text for r in _all_runs(doc)) or "for more" in xml


def test_angle_autolink(tmp_path):
    doc = _render_para("<https://example.com>", tmp_path)
    xml = doc.element.body.xml
    assert "hyperlink" in xml
    assert "<https" not in "".join(r.text for r in _all_runs(doc))


def test_url_in_code_span_not_linkified(tmp_path):
    doc = _render_para("`https://example.com`", tmp_path)
    # The code run should carry the URL literally as text.
    code_runs = [r for r in _all_runs(doc) if "example.com" in r.text]
    assert code_runs
    assert code_runs[0].font.name == "Consolas"


def test_trailing_punctuation_trimmed_from_url(tmp_path):
    doc = _render_para("go to https://example.com.", tmp_path)
    text = "".join(r.text for r in _all_runs(doc))
    # The period should survive as literal text.
    assert text.rstrip().endswith(".")


# --------------------------------------------------------------------------- #
# Hyperlink character style (links look clickable even without a template style)
# --------------------------------------------------------------------------- #
def _has_hyperlink_style(doc):
    from docx.oxml.ns import qn
    for s in doc.styles.element.findall(qn("w:style")):
        if (s.get(qn("w:styleId")) == "Hyperlink"
                and s.get(qn("w:type")) == "character"):
            return s
    return None


def test_hyperlink_style_injected(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("see https://example.com now", tmp_path)
    style = _has_hyperlink_style(doc)
    assert style is not None
    # It should carry a color and an underline so links render as links.
    rpr = style.find(qn("w:rPr"))
    assert rpr is not None
    assert rpr.find(qn("w:color")) is not None
    assert rpr.find(qn("w:u")) is not None


def test_internal_link_references_hyperlink_style(tmp_path):
    from docx.oxml.ns import qn
    md = "# Overview\n\nJump to [Details](#details).\n\n## Details\n\nHere."
    doc = _render_para(md, tmp_path)
    # The internal hyperlink run should reference the Hyperlink style.
    styled = False
    for h in doc.element.body.iter(qn("w:hyperlink")):
        for rstyle in h.iter(qn("w:rStyle")):
            if rstyle.get(qn("w:val")) == "Hyperlink":
                styled = True
    assert styled
    assert _has_hyperlink_style(doc) is not None
