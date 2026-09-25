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


def test_bullet_continuation_joined_with_space():
    # A soft-wrapped continuation line joins with a space, not a newline, so
    # emphasis spanning the wrap still pairs up.
    blocks = parse_markdown("- the **Doc\n  Repo** value")
    assert blocks[0]["items"][0]["text"] == "the **Doc Repo** value"


def test_numbered_continuation_joined_with_space():
    blocks = parse_markdown("1. the **Doc\n   Repo** value")
    assert blocks[0]["items"][0]["text"] == "the **Doc Repo** value"


def test_bold_spanning_bullet_wrap_renders(tmp_path):
    doc = _render_para("- stored in the **Doc\n  Repo** now", tmp_path)
    p = next(p for p in doc.paragraphs if "Doc Repo" in p.text)
    assert any(r.bold and "Doc Repo" in r.text for r in p.runs)


def test_bold_spanning_numbered_wrap_renders(tmp_path):
    doc = _render_para("1. uploaded to the **Doc\n   Repo**, then done", tmp_path)
    p = next(p for p in doc.paragraphs if "Doc Repo" in p.text)
    assert any(r.bold and "Doc Repo" in r.text for r in p.runs)


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


def _hr_bottom_border(doc):
    """Return the bottom border of the first HR paragraph (empty Normal para
    with a pBdr), or None."""
    from docx.oxml.ns import qn
    for p in doc.paragraphs:
        pPr = p._p.find(qn("w:pPr"))
        if pPr is None or p.text.strip() or p.style.name != "Normal":
            continue
        b = pPr.find(qn("w:pBdr") + "/" + qn("w:bottom"))
        if b is not None:
            return b
    return None


def test_hr_draws_line_by_default(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("text\n\n---\n\nmore", tmp_path)
    b = _hr_bottom_border(doc)
    assert b is not None and b.get(qn("w:val")) == "single"
    assert b.get(qn("w:color")) == "BFBFBF"


def test_hr_rule_disabled(tmp_path):
    doc = _render_with_style("a\n\n---\n\nb", tmp_path, "hr:\n  rule: false\n")
    assert _hr_bottom_border(doc) is None


def test_hr_rule_color_and_width_override(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_with_style("a\n\n---\n\nb", tmp_path,
                             "hr:\n  color: FF0000\n  width_pt: 2\n")
    b = _hr_bottom_border(doc)
    assert b is not None
    assert b.get(qn("w:color")) == "FF0000"
    assert b.get(qn("w:sz")) == "16"  # 2pt * 8


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
        yield from p.runs


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
    # The shaded padded box now lives on the Code Block STYLE, not the paragraph.
    pPr = doc.styles["Code Block"].element.find(qn("w:pPr"))
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
    # The code paragraph references the style and carries no direct box.
    code_p = next(p for p in doc.paragraphs if "code line" in p.text)
    assert code_p.style.name == "Code Block"
    p_pPr = code_p._p.find(qn("w:pPr"))
    assert p_pPr is None or p_pPr.find(qn("w:pBdr")) is None


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
    # The code run should carry the URL literally as text and use the inline-code style.
    code_runs = [r for r in _all_runs(doc) if "example.com" in r.text]
    assert code_runs
    assert code_runs[0].style.name == "Inline Code"


def test_trailing_punctuation_trimmed_from_url(tmp_path):
    doc = _render_para("go to https://example.com.", tmp_path)
    text = "".join(r.text for r in _all_runs(doc))
    # The period should survive as literal text.
    assert text.rstrip().endswith(".")


# --------------------------------------------------------------------------- #
# Hyperlinks reference the Hyperlink character style (named style, not direct)
# --------------------------------------------------------------------------- #
def test_hyperlink_uses_char_style(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("see https://example.com now", tmp_path)
    found = False
    for h in doc.element.body.iter(qn("w:hyperlink")):
        rPr = h.find(qn("w:r") + "/" + qn("w:rPr"))
        if rPr is None:
            continue
        rStyle = rPr.find(qn("w:rStyle"))
        if rStyle is not None and rStyle.get(qn("w:val")) == "Hyperlink":
            found = True
    assert found


def test_internal_link_uses_char_style(tmp_path):
    from docx.oxml.ns import qn
    md = "# Overview\n\nJump to [Details](#details).\n\n## Details\n\nHere."
    doc = _render_para(md, tmp_path)
    # An internal-anchor hyperlink exists and its run uses the Hyperlink style.
    styled = False
    for h in doc.element.body.iter(qn("w:hyperlink")):
        if h.get(qn("w:anchor")) is None:
            continue
        rPr = h.find(qn("w:r") + "/" + qn("w:rPr"))
        if rPr is not None:
            rStyle = rPr.find(qn("w:rStyle"))
            if rStyle is not None and rStyle.get(qn("w:val")) == "Hyperlink":
                styled = True
    assert styled


def test_link_color_on_hyperlink_style(tmp_path):
    from md_to_docx import build_docx as _b
    from md_to_docx import load_style
    cfg_file = tmp_path / "s.yaml"
    cfg_file.write_text("links:\n  color: FF0000\n", encoding="utf-8")
    blocks = parse_markdown("visit https://example.com")
    out = tmp_path / "o.docx"
    _b(blocks, str(out), title="T", author="A", date="d",
       style=load_style(str(cfg_file)))
    doc = Document(str(out))
    # The Hyperlink character style should carry the overridden color.
    hl_style = doc.styles["Hyperlink"]
    assert str(hl_style.font.color.rgb) == "FF0000"


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
    # body.font is None (inherit the global default).
    assert cfg["body"]["font"] is None
    assert cfg["font"] == "Aptos"
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
# Named-styles: headings, body, page
# --------------------------------------------------------------------------- #
def test_heading_style_and_outline(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("# Big Title", tmp_path)
    p = next(p for p in doc.paragraphs if "Big Title" in p.text)
    # Paragraph uses the named "Heading 1" style.
    assert p.style.name == "Heading 1"
    # The style carries the color/bold/outline — the run itself has no direct
    # color (it inherits from the style).
    style = doc.styles["Heading 1"]
    assert style.font.color.rgb is not None
    assert style.font.bold is True
    # Outline level set to 0 for h1 on the style element.
    outline = style.element.find(qn("w:pPr") + "/" + qn("w:outlineLvl"))
    assert outline is not None and outline.get(qn("w:val")) == "0"


def test_heading_color_override(tmp_path):
    from md_to_docx import build_docx as _b
    from md_to_docx import load_style
    cfg = tmp_path / "s.yaml"
    cfg.write_text("headings:\n  1:\n    color: FF0000\n", encoding="utf-8")
    blocks = parse_markdown("# Title")
    out = tmp_path / "o.docx"
    _b(blocks, str(out), title="T", author="A", date="d",
       style=load_style(str(cfg)))
    doc = Document(str(out))
    # Color lives on the Heading 1 style, not the run.
    style = doc.styles["Heading 1"]
    assert str(style.font.color.rgb) == "FF0000"


def test_body_font_inherited_from_doc_defaults(tmp_path):
    """Body text has no direct font — it inherits from docDefaults (Aptos)."""
    from docx.oxml.ns import qn
    doc = _render_para("plain paragraph text", tmp_path)
    p = next(p for p in doc.paragraphs if "plain paragraph" in p.text)
    # The run has no direct font (inherits from Normal → docDefaults).
    assert p.runs[0].font.name is None
    # The docDefaults rFonts has Aptos.
    styles_el = doc.styles.element
    rFonts = styles_el.find(
        qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/" + qn("w:rPr") + "/" + qn("w:rFonts"))
    assert rFonts is not None
    assert rFonts.get(qn("w:ascii")) == "Aptos"


def test_page_margins_applied(tmp_path):
    from docx.shared import Inches

    from md_to_docx import build_docx as _b
    from md_to_docx import load_style
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
    from md_to_docx import build_docx as _b
    from md_to_docx import load_style
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
    # The fill lives on the Code Block style (shd + border color).
    pPr = doc.styles["Code Block"].element.find(qn("w:pPr"))
    shd = pPr.find(qn("w:shd"))
    assert shd is not None and shd.get(qn("w:fill")) == "ABCDEF"
    top = pPr.find(qn("w:pBdr") + "/" + qn("w:top"))
    assert top is not None and top.get(qn("w:color")) == "ABCDEF"


def test_blockquote_bar_color_override(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_with_style("> quoted", tmp_path,
                             "blockquote:\n  bar_color: 112233\n")
    # The left bar now lives on the Blockquote style's pPr (editable in Word),
    # not stamped on each paragraph.
    style = doc.styles["Blockquote"]
    left = style.element.find(
        qn("w:pPr") + "/" + qn("w:pBdr") + "/" + qn("w:left"))
    assert left is not None and left.get(qn("w:color")) == "112233"
    # And the paragraph itself carries no direct border.
    q = next(p for p in doc.paragraphs if "quoted" in p.text)
    pPr = q._p.find(qn("w:pPr"))
    assert pPr is None or pPr.find(qn("w:pBdr")) is None


def test_inline_code_font_override(tmp_path):
    doc = _render_with_style("use `foo` here", tmp_path,
                             "inline_code:\n  font: Courier New\n")
    # The inline-code character style should carry the overridden font.
    style = doc.styles["Inline Code"]
    assert style.font.name == "Courier New"
    # The run uses the inline-code style.
    code_runs = [r for p in doc.paragraphs for r in p.runs if r.text == "foo"]
    assert code_runs and code_runs[0].style.name == "Inline Code"


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
    # The between-row rule (drawn by default) carries the overridden color.
    inside_h = tblBorders.find(qn("w:insideH"))
    assert inside_h is not None and inside_h.get(qn("w:color")) == "123456"
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
    # Only the between-row rule is drawn by default; top/bottom/vertical off.
    assert tblBorders.find(qn("w:insideH")).get(qn("w:val")) == "single"
    assert tblBorders.find(qn("w:top")).get(qn("w:val")) == "nil"
    assert tblBorders.find(qn("w:bottom")).get(qn("w:val")) == "nil"
    assert tblBorders.find(qn("w:insideV")).get(qn("w:val")) == "nil"
    assert tblBorders.find(qn("w:left")).get(qn("w:val")) == "nil"
    assert tblBorders.find(qn("w:right")).get(qn("w:val")) == "nil"
    # No header fill by default; a header underline (per-cell bottom border) is set.
    hdr = t.rows[0].cells[0]._tc
    assert hdr.find(qn("w:tcPr") + "/" + qn("w:shd")) is None
    ul = hdr.find(qn("w:tcPr") + "/" + qn("w:tcBorders") + "/" + qn("w:bottom"))
    assert ul is not None and ul.get(qn("w:val")) == "single"


def test_table_column_alignment(tmp_path):
    from docx.enum.text import WD_ALIGN_PARAGRAPH as A
    md = ("| L | C | R |\n"
          "|:--|:-:|--:|\n"
          "| a | b | c |\n")
    doc = _render_para(md, tmp_path)
    t = doc.tables[0]
    # Header and body share the column alignment.
    for r in range(len(t.rows)):
        assert t.rows[r].cells[0].paragraphs[0].alignment == A.LEFT
        assert t.rows[r].cells[1].paragraphs[0].alignment == A.CENTER
        assert t.rows[r].cells[2].paragraphs[0].alignment == A.RIGHT


def test_table_unspecified_alignment_is_default(tmp_path):
    md = "| A | B |\n|---|---|\n| 1 | 2 |"
    doc = _render_para(md, tmp_path)
    # No colons -> alignment left as default (None).
    for cell in doc.tables[0].rows[1].cells:
        assert cell.paragraphs[0].alignment is None


def test_parse_table_alignments_helper():
    from docx.enum.text import WD_ALIGN_PARAGRAPH as A

    from md_to_docx.md_to_docx import _parse_table_alignments
    aligns = _parse_table_alignments("|:---|:--:|---:|---|")
    assert aligns == [A.LEFT, A.CENTER, A.RIGHT, None]


def test_table_escaped_pipe_in_cell(tmp_path):
    doc = _render_para("| A | B |\n|---|---|\n| a \\| b | c |", tmp_path)
    t = doc.tables[0]
    # The escaped pipe stays inside one cell as a literal '|'.
    assert t.rows[1].cells[0].text == "a | b"
    assert t.rows[1].cells[1].text == "c"


def test_table_br_becomes_line_break_in_cell(tmp_path):
    doc = _render_para("| A |\n|---|\n| line1<br>line2 |", tmp_path)
    assert doc.tables[0].rows[1].cells[0].text == "line1\nline2"


def test_split_table_row_helper():
    from md_to_docx.md_to_docx import _split_table_row
    assert _split_table_row("| a \\| b | c |") == ["a | b", "c"]
    assert _split_table_row("| x<br>y |") == ["x\ny"]


def test_line_starting_with_dashes_does_not_hang(tmp_path):
    # A line like "---|---" starts with "---" but isn't an <hr>; it must not
    # send the parser into an infinite loop.
    blocks = parse_markdown("A | B\n---|---\n1 | 2")
    assert [b["type"] for b in blocks] == ["paragraph", "paragraph", "paragraph"]


def test_extra_dashes_line_parses(tmp_path):
    blocks = parse_markdown("----\ntext")
    # "----" is not a valid hr (only exactly "---" is) so it's a paragraph.
    assert blocks[0]["type"] == "paragraph"


def test_table_no_top_bottom_edges_by_default(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("| A | B |\n|---|---|\n| 1 | 2 |", tmp_path)
    tblBorders = doc.tables[0]._tbl.find(
        qn("w:tblPr") + "/" + qn("w:tblBorders"))
    assert tblBorders.find(qn("w:top")).get(qn("w:val")) == "nil"
    assert tblBorders.find(qn("w:bottom")).get(qn("w:val")) == "nil"


def test_table_top_bottom_can_be_re_enabled(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_with_style(
        "| A | B |\n|---|---|\n| 1 | 2 |", tmp_path,
        "table:\n  border:\n    edges: [top, bottom, insideH]\n")
    tblBorders = doc.tables[0]._tbl.find(
        qn("w:tblPr") + "/" + qn("w:tblBorders"))
    assert tblBorders.find(qn("w:top")).get(qn("w:val")) == "single"
    assert tblBorders.find(qn("w:bottom")).get(qn("w:val")) == "single"


# --------------------------------------------------------------------------- #
# Code-block and table left indent
# The inset is applied on the RIGHT only: code blocks and tables stay flush-left
# with body text (and the heading rules) and just stop short of the right margin.
# --------------------------------------------------------------------------- #
def test_code_block_title_style_exists_and_based_on_code_block(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("```python\nx = 1\n```", tmp_path)
    title = doc.styles["Code Block Title"]
    # Based on Code Block (inherits the box/indent), overrides run look.
    assert title.base_style is not None and title.base_style.name == "Code Block"
    assert title.font.italic is True
    # "next" style returns to Code Block.
    nxt = title.element.find(qn("w:next"))
    assert nxt is not None


def test_code_block_language_uses_title_style(tmp_path):
    doc = _render_para("```python\nx = 1\n```", tmp_path)
    langs = [p for p in doc.paragraphs
             if p.text == "python" and p.style.name == "Code Block Title"]
    assert langs
    # The code itself is a separate Code Block paragraph.
    code = [p for p in doc.paragraphs
            if "x = 1" in p.text and p.style.name == "Code Block"]
    assert code


def test_code_block_no_language_has_no_title(tmp_path):
    doc = _render_para("```\nplain code\n```", tmp_path)
    titles = [p for p in doc.paragraphs if p.style.name == "Code Block Title"]
    assert not titles


def test_code_block_box_left_aligns_with_body(tmp_path):
    """The shaded box's left edge lines up with body text: the paragraph's left
    indent equals the border padding (which offsets the border outward)."""
    from docx.oxml.ns import qn
    doc = _render_para("```\ncode\n```", tmp_path)
    ind = doc.styles["Code Block"].element.find(
        qn("w:pPr") + "/" + qn("w:ind"))
    assert ind is not None
    # padding_pt default 6 -> 120 twips left indent.
    assert int(ind.get(qn("w:left"))) == 120


def test_code_block_right_inset_by_default(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("```\ncode\n```", tmp_path)
    ind = doc.styles["Code Block"].element.find(
        qn("w:pPr") + "/" + qn("w:ind"))
    # right = indent_in (0.25in = 360) + padding (6pt = 120) = 480 twips.
    assert int(ind.get(qn("w:right"))) == 480


def test_code_block_indent_override(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_with_style("```\ncode\n```", tmp_path,
                             "code_block:\n  indent_in: 0.5\n")
    ind = doc.styles["Code Block"].element.find(
        qn("w:pPr") + "/" + qn("w:ind"))
    # right = 0.5in (720) + padding (120) = 840 twips.
    assert int(ind.get(qn("w:right"))) == 840


def test_table_indent_matches_cell_left_margin(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("| A | B |\n|---|---|\n| 1 | 2 |", tmp_path)
    tblInd = doc.tables[0]._tbl.find(qn("w:tblPr") + "/" + qn("w:tblInd"))
    # tblInd equals the cell left margin (5pt -> 100 twips) so the table border
    # is positioned to line the cell text up near the body margin.
    assert tblInd is not None
    assert int(tblInd.get(qn("w:w"))) == 100


def test_table_indent_narrows_width_to_avoid_overflow(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("| A | B |\n|---|---|\n| 1 | 2 |", tmp_path)
    tblW = doc.tables[0]._tbl.find(qn("w:tblPr") + "/" + qn("w:tblW"))
    # Full-width table with an inset must be < 100% (5000 pct) so it fits.
    assert tblW.get(qn("w:type")) == "pct"
    assert int(tblW.get(qn("w:w"))) < 5000


def test_table_indent_override_to_zero(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_with_style("| A | B |\n|---|---|\n| 1 | 2 |", tmp_path,
                             "table:\n  indent_in: 0\n")
    tblPr = doc.tables[0]._tbl.find(qn("w:tblPr"))
    # Right inset removed => full width. The left alignment (negative tblInd)
    # stays regardless of the right inset.
    assert int(tblPr.find(qn("w:tblW")).get(qn("w:w"))) == 5000


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


# --------------------------------------------------------------------------- #
# Named-styles feature: document defaults, paragraph/character styles
# --------------------------------------------------------------------------- #
def test_doc_defaults_font_is_aptos(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("hello", tmp_path)
    styles_el = doc.styles.element
    rFonts = styles_el.find(
        qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/" + qn("w:rPr") + "/" + qn("w:rFonts"))
    assert rFonts is not None
    assert rFonts.get(qn("w:ascii")) == "Aptos"
    assert rFonts.get(qn("w:hAnsi")) == "Aptos"


def test_doc_defaults_font_override(tmp_path):
    """Overriding the top-level `font` changes docDefaults."""
    from docx.oxml.ns import qn
    doc = _render_with_style("hello", tmp_path, "font: 'Comic Sans MS'\n")
    styles_el = doc.styles.element
    rFonts = styles_el.find(
        qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/" + qn("w:rPr") + "/" + qn("w:rFonts"))
    assert rFonts is not None
    assert rFonts.get(qn("w:ascii")) == "Comic Sans MS"


def test_heading_font_default(tmp_path):
    """Headings use the top-level heading_font (Aptos Display) by default."""
    doc = _render_para("# Title\n\n## Sub", tmp_path)
    assert doc.styles["Heading 1"].font.name == "Aptos Display"
    assert doc.styles["Heading 2"].font.name == "Aptos Display"


def _heading_bottom_rule(doc, style_name):
    """Return the bottom-border element of a heading style, or None."""
    from docx.oxml.ns import qn
    pPr = doc.styles[style_name].element.find(qn("w:pPr"))
    if pPr is None:
        return None
    return pPr.find(qn("w:pBdr") + "/" + qn("w:bottom"))


def test_h1_h2_have_bottom_rule_by_default(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_para("# H1\n\n## H2", tmp_path)
    h1_rule = _heading_bottom_rule(doc, "Heading 1")
    h2_rule = _heading_bottom_rule(doc, "Heading 2")
    assert h1_rule is not None and h1_rule.get(qn("w:val")) == "single"
    assert h1_rule.get(qn("w:color")) == "365F91"
    assert h2_rule is not None and h2_rule.get(qn("w:color")) == "4F81BD"


def test_h3_h6_have_no_bottom_rule_by_default(tmp_path):
    doc = _render_para("### H3\n\n#### H4\n\n##### H5\n\n###### H6", tmp_path)
    for name in ("Heading 3", "Heading 4", "Heading 5", "Heading 6"):
        assert _heading_bottom_rule(doc, name) is None


def test_heading_rule_override_color_and_width(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_with_style(
        "# Title", tmp_path,
        "headings:\n  1:\n    rule:\n      color: FF0000\n      width_pt: 3\n")
    rule = _heading_bottom_rule(doc, "Heading 1")
    assert rule is not None
    assert rule.get(qn("w:color")) == "FF0000"
    assert rule.get(qn("w:sz")) == "24"  # 3pt * 8


def test_heading_rule_disabled_with_null(tmp_path):
    doc = _render_with_style("# Title", tmp_path,
                             "headings:\n  1:\n    rule: null\n")
    assert _heading_bottom_rule(doc, "Heading 1") is None


def test_heading_rule_added_to_lower_level(tmp_path):
    from docx.oxml.ns import qn
    doc = _render_with_style(
        "### Sub", tmp_path,
        "headings:\n  3:\n    rule:\n      color: '00AA00'\n      width_pt: 1\n")
    rule = _heading_bottom_rule(doc, "Heading 3")
    assert rule is not None and rule.get(qn("w:color")) == "00AA00"


def test_heading_style_has_no_theme_font(tmp_path):
    """The built-in Heading styles reference theme fonts (majorHAnsi) which
    override an explicit font; setting the font must strip those theme attrs so
    the configured face (Aptos Display) actually renders in Word."""
    from docx.oxml.ns import qn
    doc = _render_para("# Title", tmp_path)
    rFonts = doc.styles["Heading 1"].element.find(
        qn("w:rPr") + "/" + qn("w:rFonts"))
    assert rFonts is not None
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        assert rFonts.get(qn(f"w:{attr}")) is None, f"{attr} should be stripped"
    assert rFonts.get(qn("w:ascii")) == "Aptos Display"


def test_heading_font_override(tmp_path):
    """Overriding heading_font changes all headings that don't set their own."""
    doc = _render_with_style("# H1\n\n## H2", tmp_path,
                             "heading_font: Georgia\n")
    assert doc.styles["Heading 1"].font.name == "Georgia"
    assert doc.styles["Heading 2"].font.name == "Georgia"


def test_per_level_heading_font_wins_over_heading_font(tmp_path):
    """A per-level font overrides heading_font for just that level."""
    doc = _render_with_style("# H1\n\n## H2", tmp_path,
                             "heading_font: Georgia\nheadings:\n  1:\n    font: Impact\n")
    assert doc.styles["Heading 1"].font.name == "Impact"
    assert doc.styles["Heading 2"].font.name == "Georgia"


def test_heading_paragraph_uses_style_name(tmp_path):
    doc = _render_para("### Third", tmp_path)
    p = next(p for p in doc.paragraphs if "Third" in p.text)
    assert p.style.name == "Heading 3"


def test_normal_style_inherits_font(tmp_path):
    """Normal style has no font set (None) — it inherits from docDefaults."""
    doc = _render_para("text", tmp_path)
    normal = doc.styles["Normal"]
    assert normal.font.name is None


def test_unused_template_styles_hidden_from_gallery(tmp_path):
    """Template quick styles we never apply (Title, Quote, List Paragraph, ...)
    are pruned from the gallery, leaving only our styles + headings/Normal."""
    doc = _render_para("# H\n\ntext\n\n> quote", tmp_path)
    visible = {s.name for s in doc.styles if _safe_quick(s)}
    expected = {
        "Normal", "Heading 1", "Heading 2", "Heading 3", "Heading 4",
        "Heading 5", "Heading 6", "Blockquote", "Code Block",
        "Code Block Title", "Inline Code",
    }
    assert visible == expected
    # A few template styles that must NOT be in the gallery anymore.
    for hidden in ("Title", "Subtitle", "Quote", "Intense Quote", "List Paragraph"):
        assert doc.styles[hidden].quick_style is False


def _safe_quick(style):
    try:
        return bool(style.quick_style)
    except Exception:
        return False


def test_custom_styles_are_visible_in_word(tmp_path):
    """Our custom styles are flagged qFormat (quick_style) + uiPriority so Word
    shows them in the Styles gallery / pane instead of hiding them."""
    from docx.oxml.ns import qn
    doc = _render_para("> quote", tmp_path)
    for name in ("Blockquote", "Code Block", "Inline Code"):
        style = doc.styles[name]
        assert style.quick_style is True, f"{name} should be a quick style"
        assert style.element.find(qn("w:qFormat")) is not None
        assert style.priority is not None


def test_blockquote_uses_md_quote_style(tmp_path):
    doc = _render_para("> quoted text", tmp_path)
    p = next(p for p in doc.paragraphs if "quoted text" in p.text)
    assert p.style.name == "Blockquote"


def test_md_quote_style_carries_left_bar(tmp_path):
    """The Blockquote style's pPr carries the left bar, in schema order (before
    spacing/ind), so Word accepts it and the whole quote is editable via the
    style."""
    from docx.oxml.ns import qn
    doc = _render_para("> quoted", tmp_path)
    pPr = doc.styles["Blockquote"].element.find(qn("w:pPr"))
    assert pPr is not None
    left = pPr.find(qn("w:pBdr") + "/" + qn("w:left"))
    assert left is not None
    assert left.get(qn("w:val")) == "single"
    # pBdr must appear before spacing and ind in the child ordering.
    children = [c.tag for c in pPr]
    i_bdr = children.index(qn("w:pBdr"))
    for tag in ("w:spacing", "w:ind"):
        if qn(tag) in children:
            assert i_bdr < children.index(qn(tag))


def test_code_block_uses_md_code_style(tmp_path):
    doc = _render_para("```\ncode line\n```", tmp_path)
    p = next(p for p in doc.paragraphs if "code line" in p.text)
    assert p.style.name == "Code Block"


def test_code_block_style_has_consolas(tmp_path):
    doc = _render_para("```\ncode\n```", tmp_path)
    style = doc.styles["Code Block"]
    assert style.font.name == "Consolas"


def test_note_callout_renders_as_blockquote(tmp_path):
    """A "> **Note:** ..." line is just a normal blockquote (no special note)."""
    doc = _render_para("> **Note:** heads up", tmp_path)
    p = next(p for p in doc.paragraphs if "heads up" in p.text)
    assert p.style.name == "Blockquote"


def test_hyperlink_style_exists_with_color(tmp_path):
    doc = _render_para("visit https://example.com", tmp_path)
    hl = doc.styles["Hyperlink"]
    assert str(hl.font.color.rgb) == "0563C1"
    assert hl.font.underline is True


def test_inline_code_uses_code_char_style(tmp_path):
    doc = _render_para("use `foo` here", tmp_path)
    code_runs = [r for p in doc.paragraphs for r in p.runs if r.text == "foo"]
    assert code_runs
    assert code_runs[0].style.name == "Inline Code"


def test_code_char_style_has_consolas(tmp_path):
    doc = _render_para("use `foo` here", tmp_path)
    style = doc.styles["Inline Code"]
    assert style.font.name == "Consolas"


def test_emphasis_on_styled_paragraph(tmp_path):
    """Bold/italic still render as direct run properties on styled paragraphs."""
    doc = _render_para("**bold** and *italic*", tmp_path)
    bold_runs = [r for r in _all_runs(doc) if r.text == "bold"]
    italic_runs = [r for r in _all_runs(doc) if r.text == "italic"]
    assert bold_runs and bold_runs[0].bold
    assert italic_runs and italic_runs[0].italic
