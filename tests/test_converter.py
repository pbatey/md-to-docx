"""Unit tests for md_to_docx.

The pure parsing/slug helpers are tested directly. A couple of end-to-end tests
run the full conversion and read the resulting .docx back with python-docx to
confirm the document is well-formed and carries the expected content.
"""

from pathlib import Path

import pytest
from docx import Document

from md_to_docx import build_docx, parse_markdown
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
# Hyperlinks are styled with direct color + underline (no named style)
# --------------------------------------------------------------------------- #
def test_hyperlink_direct_color_and_underline(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("see https://example.com now", tmp_path)
    found = False
    for h in doc.element.body.iter(qn("w:hyperlink")):
        rPr = h.find(qn("w:r") + "/" + qn("w:rPr"))
        if rPr is None:
            continue
        color = rPr.find(qn("w:color"))
        u = rPr.find(qn("w:u"))
        if color is not None and color.get(qn("w:val")) == "0563C1" and u is not None:
            found = True
    assert found


def test_internal_link_directly_styled(tmp_path):
    from docx.oxml.ns import qn
    md = "# Overview\n\nJump to [Details](#details).\n\n## Details\n\nHere."
    doc = _render_para(md, tmp_path)
    # An internal-anchor hyperlink exists and its run carries a direct color.
    styled = False
    for h in doc.element.body.iter(qn("w:hyperlink")):
        if h.get(qn("w:anchor")) is None:
            continue
        rPr = h.find(qn("w:r") + "/" + qn("w:rPr"))
        if rPr is not None and rPr.find(qn("w:color")) is not None:
            styled = True
    assert styled


def test_link_color_override(tmp_path):
    from docx.oxml.ns import qn
    from md_to_docx import build_docx as _b, load_style
    cfg_file = tmp_path / "s.yaml"
    cfg_file.write_text("links:\n  color: FF0000\n", encoding="utf-8")
    blocks = parse_markdown("visit https://example.com")
    out = tmp_path / "o.docx"
    _b(blocks, str(out), title="T", author="A", date="d",
       style=load_style(str(cfg_file)))
    doc = Document(str(out))
    reds = [c for c in doc.element.body.iter(qn("w:color"))
            if c.get(qn("w:val")) == "FF0000"]
    assert reds


# --------------------------------------------------------------------------- #
# Style config: defaults, loading, merging (YAML)
# --------------------------------------------------------------------------- #
def test_load_style_defaults_when_no_path():
    from md_to_docx import DEFAULT_STYLE, load_style
    cfg = load_style()
    assert cfg == DEFAULT_STYLE
    # It's a copy, not the same object (so callers can't mutate defaults).
    assert cfg is not DEFAULT_STYLE


def test_load_style_deep_merge(tmp_path):
    from md_to_docx import load_style
    yaml_file = tmp_path / "style.yaml"
    yaml_file.write_text(
        "body:\n  size_pt: 13\nheadings:\n  1:\n    color: FF0000\n",
        encoding="utf-8")
    cfg = load_style(str(yaml_file))
    # Overridden values applied.
    assert cfg["body"]["size_pt"] == 13
    assert cfg["headings"][1]["color"] == "FF0000"
    # Sibling defaults preserved (deep merge, not wholesale replace).
    assert cfg["body"]["font"] == "Aptos"
    assert cfg["headings"][1]["bold"] is True
    assert cfg["headings"][2]["color"] == "4F81BD"


def test_load_style_invalid_yaml_raises(tmp_path):
    from md_to_docx import StyleError, load_style
    bad = tmp_path / "bad.yaml"
    bad.write_text("body: : : not valid\n", encoding="utf-8")
    with pytest.raises(StyleError):
        load_style(str(bad))


def test_load_style_missing_file_raises():
    from md_to_docx import StyleError, load_style
    with pytest.raises(StyleError):
        load_style("does-not-exist.yaml")


def test_load_style_unknown_keys_ignored(tmp_path):
    from md_to_docx import load_style
    yaml_file = tmp_path / "style.yaml"
    yaml_file.write_text("totally_unknown_key: 42\nbody:\n  size_pt: 12\n",
                         encoding="utf-8")
    cfg = load_style(str(yaml_file))
    # Unknown key is merged in but harmless; known override still works.
    assert cfg["body"]["size_pt"] == 12


def test_dump_config_roundtrips_to_defaults():
    import yaml as _yaml
    from md_to_docx import DEFAULT_STYLE, dump_default_style
    text = dump_default_style()
    parsed = _yaml.safe_load(text)
    # Heading keys come back as ints via normalization path; compare loosely.
    assert parsed["body"] == DEFAULT_STYLE["body"]
    assert parsed["links"] == DEFAULT_STYLE["links"]


def test_styleconfig_coercion(tmp_path):
    from md_to_docx import StyleConfig, load_style
    sc = StyleConfig(load_style())
    # number + color coercion
    assert sc.num("body", "size_pt") == 11.0
    assert sc.color("links", "color") is not None
    # malformed color falls back (warns), does not raise
    bad = tmp_path / "s.yaml"
    bad.write_text("links:\n  color: nothex\n", encoding="utf-8")
    sc2 = StyleConfig(load_style(str(bad)))
    # falls back to default link color, still a valid RGBColor
    assert sc2.color("links", "color", default="0563C1") is not None


# --------------------------------------------------------------------------- #
# Direct-formatting: headings, body, page (no template)
# --------------------------------------------------------------------------- #
def test_heading_direct_formatting_and_outline(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("# Big Title", tmp_path)
    p = next(p for p in doc.paragraphs if "Big Title" in p.text)
    # Direct run color/size from defaults (H1 = 365F91, 20pt).
    run = p.runs[0]
    assert run.font.color.rgb is not None
    assert run.bold is True
    # Outline level set to 0 for h1.
    outline = p._p.find(qn("w:pPr") + "/" + qn("w:outlineLvl"))
    assert outline is not None and outline.get(qn("w:val")) == "0"


def test_heading_color_override(tmp_path):
    from md_to_docx import build_docx as _b, load_style
    cfg = tmp_path / "s.yaml"
    cfg.write_text("headings:\n  1:\n    color: FF0000\n", encoding="utf-8")
    blocks = parse_markdown("# Title")
    out = tmp_path / "o.docx"
    _b(blocks, str(out), title="T", author="A", date="d",
       style=load_style(str(cfg)))
    doc = Document(str(out))
    run = next(p for p in doc.paragraphs if "Title" in p.text).runs[0]
    assert str(run.font.color.rgb) == "FF0000"


def test_body_font_applied(tmp_path):
    doc = _render_para("plain paragraph text", tmp_path)
    p = next(p for p in doc.paragraphs if "plain paragraph" in p.text)
    assert p.runs[0].font.name == "Aptos"


def test_page_margins_applied(tmp_path):
    from docx.shared import Inches
    from md_to_docx import build_docx as _b, load_style
    cfg = tmp_path / "s.yaml"
    cfg.write_text("page:\n  margins_in:\n    left: 2.0\n", encoding="utf-8")
    blocks = parse_markdown("# t")
    out = tmp_path / "o.docx"
    _b(blocks, str(out), title="T", author="A", date="d",
       style=load_style(str(cfg)))
    doc = Document(str(out))
    assert doc.sections[0].left_margin == Inches(2.0)


# --------------------------------------------------------------------------- #
# Config-driven code block, blockquote, inline code
# --------------------------------------------------------------------------- #
def _render_with_style(md, tmp_path, yaml_text):
    from md_to_docx import build_docx as _b, load_style
    cfg = tmp_path / "s.yaml"
    cfg.write_text(yaml_text, encoding="utf-8")
    blocks = parse_markdown(md)
    out = tmp_path / "o.docx"
    _b(blocks, str(out), title="T", author="A", date="d",
       base_dir=tmp_path, style=load_style(str(cfg)))
    return Document(str(out))


def test_code_block_fill_override(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_with_style("```\nx=1\n```", tmp_path,
                             "code_block:\n  fill: ABCDEF\n")
    code_p = next(p for p in doc.paragraphs if "x=1" in p.text)
    shd = code_p._p.find(qn("w:pPr") + "/" + qn("w:shd"))
    assert shd is not None and shd.get(qn("w:fill")) == "ABCDEF"


def test_blockquote_bar_color_override(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_with_style("> quoted", tmp_path,
                             "blockquote:\n  bar_color: 112233\n")
    q = next(p for p in doc.paragraphs if "quoted" in p.text)
    left = q._p.find(qn("w:pPr") + "/" + qn("w:pBdr") + "/" + qn("w:left"))
    assert left is not None and left.get(qn("w:color")) == "112233"


def test_inline_code_font_override(tmp_path):
    doc = _render_with_style("use `foo` here", tmp_path,
                             "inline_code:\n  font: Courier New\n")
    code_runs = [r for p in doc.paragraphs for r in p.runs if r.text == "foo"]
    assert code_runs and code_runs[0].font.name == "Courier New"


# --------------------------------------------------------------------------- #
# Lists/numbering without a template
# --------------------------------------------------------------------------- #
def test_bullet_list_has_numbering(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("- one\n- two", tmp_path)
    numprs = list(doc.element.body.iter(qn("w:numPr")))
    assert len(numprs) == 2


def test_two_ordered_lists_each_restart(tmp_path):
    from docx.oxml.ns import qn
    md = "1. a\n2. b\n\ntext\n\n1. c\n2. d"
    doc = _render_para(md, tmp_path)
    # Two separate numId values (one per block) so each restarts at 1.
    num_ids = set()
    for numPr in doc.element.body.iter(qn("w:numPr")):
        nid = numPr.find(qn("w:numId"))
        if nid is not None:
            num_ids.add(nid.get(qn("w:val")))
    assert len(num_ids) >= 2
    # Each created num has a startOverride of 1.
    numbering = doc.part.numbering_part._element
    overrides = [o.get(qn("w:val"))
                 for o in numbering.iter(qn("w:startOverride"))]
    assert overrides and all(v == "1" for v in overrides)


def test_numbered_list_subitems_nest(tmp_path):
    from docx.oxml.ns import qn
    md = "1. first\n   - sub a\n   - sub b\n2. second"
    doc = _render_para(md, tmp_path)
    # Sub-bullets render at ilvl 1.
    ilvls = [e.get(qn("w:val")) for e in doc.element.body.iter(qn("w:ilvl"))]
    assert "1" in ilvls


def test_task_glyphs_preserved(tmp_path):
    doc = _render_para("- [ ] todo\n- [x] done", tmp_path)
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "\u2610" in text and "\u2611" in text


# --------------------------------------------------------------------------- #
# Table direct borders + header
# --------------------------------------------------------------------------- #
def test_table_direct_borders_and_header(tmp_path):
    from docx.oxml.ns import qn
    md = "| A | B |\n|---|---|\n| 1 | 2 |"
    doc = _render_with_style(
        md, tmp_path,
        "table:\n  border:\n    color: 123456\n  header:\n    bold: true\n    fill: EEEEEE\n")
    t = doc.tables[0]
    tblBorders = t._tbl.find(qn("w:tblPr") + "/" + qn("w:tblBorders"))
    assert tblBorders is not None
    top = tblBorders.find(qn("w:top"))
    assert top is not None and top.get(qn("w:color")) == "123456"
    # Header row bold + shaded.
    hdr_cell = t.rows[0].cells[0]
    assert any(r.bold for para in hdr_cell.paragraphs for r in para.runs)
    shd = hdr_cell._tc.find(qn("w:tcPr") + "/" + qn("w:shd"))
    assert shd is not None and shd.get(qn("w:fill")) == "EEEEEE"


def test_table_horizontal_rules_only_by_default(tmp_path):
    from docx.oxml.ns import qn
    md = "| A | B |\n|---|---|\n| 1 | 2 |"
    doc = _render_para(md, tmp_path)
    t = doc.tables[0]
    tblBorders = t._tbl.find(qn("w:tblPr") + "/" + qn("w:tblBorders"))
    assert tblBorders is not None
    # Horizontal edges drawn, vertical edges turned off (val="nil").
    assert tblBorders.find(qn("w:top")).get(qn("w:val")) == "single"
    assert tblBorders.find(qn("w:bottom")).get(qn("w:val")) == "single"
    assert tblBorders.find(qn("w:insideH")).get(qn("w:val")) == "single"
    assert tblBorders.find(qn("w:insideV")).get(qn("w:val")) == "nil"
    assert tblBorders.find(qn("w:left")).get(qn("w:val")) == "nil"
    assert tblBorders.find(qn("w:right")).get(qn("w:val")) == "nil"
    # No header fill by default; a header underline (per-cell bottom border) is set.
    hdr = t.rows[0].cells[0]._tc
    assert hdr.find(qn("w:tcPr") + "/" + qn("w:shd")) is None
    ul = hdr.find(qn("w:tcPr") + "/" + qn("w:tcBorders") + "/" + qn("w:bottom"))
    assert ul is not None and ul.get(qn("w:val")) == "single"


# --------------------------------------------------------------------------- #
# CLI: --dump-config
# --------------------------------------------------------------------------- #
def test_dump_config_cli(tmp_path, monkeypatch, capsys):
    import yaml as _yaml
    from md_to_docx import DEFAULT_STYLE, main
    monkeypatch.setattr("sys.argv", ["md-to-docx", "--dump-config"])
    with pytest.raises(SystemExit) as exc:
        main()
    assert exc.value.code == 0
    out = capsys.readouterr().out
    parsed = _yaml.safe_load(out)
    assert parsed["body"] == DEFAULT_STYLE["body"]