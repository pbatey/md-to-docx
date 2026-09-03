#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["python-docx", "pyyaml"]
# ///
"""
Convert onboarding-emails-detail.md to DOCX with proper line break preservation.

Usage:
    uv run md_to_docx.py <input.md> <output.docx>

Uses python-docx for fine-grained control over formatting,
especially preserving line breaks within email body blockquotes.
"""

import io
import os
import sys
import re
import urllib.request
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from lxml import etree
from copy import deepcopy
import yaml


# --------------------------------------------------------------------------- #
# Styling configuration
#
# All styling is applied as *direct formatting* (fonts, sizes, colors, indents,
# spacing, borders, shading) rather than via named Word styles or a template.
# The defaults below reproduce the look the project shipped with its old .docx
# template. A YAML file may override any subset of these keys (see load_style).
#
# Units are explicit in key names: ``*_pt`` = points, ``*_in`` = inches. Colors
# are 6-digit hex strings without a leading ``#``.
# --------------------------------------------------------------------------- #
DEFAULT_STYLE = {
    "page": {
        # Letter by default. Either a named size or explicit width/height inches.
        "size": "letter",              # "letter" | "a4"
        "margins_in": {"top": 1.0, "bottom": 1.0, "left": 1.25, "right": 1.25},
    },
    "body": {
        "font": "Aptos",
        "size_pt": 11,
        "color": "000000",
        "space_before_pt": 0,
        "space_after_pt": 8,
        "line_spacing": 1.15,
    },
    # Per-level heading formatting (levels 1-6). Colors mirror the old template.
    "headings": {
        1: {"font": "Aptos Display", "size_pt": 20, "color": "365F91", "bold": True,
            "italic": False, "space_before_pt": 18, "space_after_pt": 4},
        2: {"font": "Aptos Display", "size_pt": 16, "color": "4F81BD", "bold": True,
            "italic": False, "space_before_pt": 12, "space_after_pt": 4},
        3: {"font": "Aptos Display", "size_pt": 13, "color": "4F81BD", "bold": True,
            "italic": False, "space_before_pt": 10, "space_after_pt": 2},
        4: {"font": "Aptos Display", "size_pt": 12, "color": "4F81BD", "bold": True,
            "italic": False, "space_before_pt": 10, "space_after_pt": 2},
        5: {"font": "Aptos Display", "size_pt": 11, "color": "243F60", "bold": True,
            "italic": False, "space_before_pt": 8, "space_after_pt": 2},
        6: {"font": "Aptos Display", "size_pt": 11, "color": "243F60", "bold": False,
            "italic": True, "space_before_pt": 8, "space_after_pt": 2},
    },
    "inline_code": {
        "font": "Consolas",
        "size_pt": 10,
        "color": None,                 # optional run color
        "fill": None,                  # optional run shading (hex) or None
    },
    "code_block": {
        "font": "Consolas",
        "size_pt": 9,
        "fill": "F2F2F2",
        "padding_pt": 6,
        "space_before_pt": 8,
        "space_after_pt": 8,
        "caption": {"size_pt": 8, "color": "808080", "italic": True},
    },
    "blockquote": {
        "bar_color": "365F91",         # same blue as Heading 1
        "bar_width_pt": 2.25,
        "bar_gap_pt": 12,
        "indent_in": 0.25,
        "space_after_pt": 4,
    },
    "lists": {
        "bullet": {
            # Per-level bullet glyph + the font that renders it. These mirror
            # Word's own defaults: level 0 is Symbol's filled bullet (U+F0B7),
            # level 1 is a lowercase "o" in Courier New (a hollow bullet). Each
            # glyph must be paired with a font that actually contains it — the
            # Unicode bullet U+2022 is NOT present in the Symbol font.
            # Both levels use the Symbol font's filled bullet (U+F0B7). The
            # Symbol font has no open/hollow bullet glyph, so the sub-bullet
            # reuses the same mark rather than a broken glyph.
            "glyphs": ["\uF0B7", "\uF0B7"],     # level 0, level 1
            "glyph_fonts": ["Symbol", "Symbol"],
            # Level-0 indent matches the ordered list (0.5") so bullet and
            # numbered lists align; level 1 nests one step deeper.
            "indent_in": [0.5, 0.75],
            "space_after_pt": 2,
        },
        "ordered": {
            "indent_in": 0.5,
            "space_after_pt": 2,
            "restart_each_block": True,
        },
    },
    "table": {
        # Horizontal-rules-only look: a line above and below the header and
        # under each body row, no vertical lines or side borders.
        # "edges" lists which borders to draw; omit an edge to leave it off.
        # Valid edges: top, bottom, left, right, insideH, insideV.
        "border": {
            "style": "single",
            "width_pt": 0.75,
            "color": "808080",
            "edges": ["top", "bottom", "insideH"],
        },
        # A heavier rule directly under the header row.
        "header": {
            "bold": True,
            "fill": None,                       # no shading behind the header
            "underline_width_pt": 1.0,
            "underline_color": "404040",
        },
        # Balanced vertical padding with a little left inset.
        "cell_margins_pt": {"top": 4, "bottom": 4, "left": 5, "right": 10},
        "width": "full",               # "full" | "auto"
    },
    "note": {
        "indent_in": 0.3,
        "italic": True,
        "size_pt": 10,
    },
    "hr": {
        "space_after_pt": 6,
        "rule": False,                 # True -> draw a bottom border line
    },
    "links": {
        "color": "0563C1",
        "underline": True,
    },
}


# Named page sizes in (width_in, height_in).
_PAGE_SIZES = {
    "letter": (8.5, 11.0),
    "a4": (8.27, 11.69),
}


class StyleError(Exception):
    """Raised when a style config cannot be parsed."""


def _deep_merge(base: dict, override: dict) -> dict:
    """Recursively merge ``override`` into a copy of ``base``.

    Nested dicts merge key-by-key; scalars and lists replace. Keys present only
    in ``base`` are kept (so a partial YAML overrides just what it names).
    """
    result = deepcopy(base)
    for key, value in override.items():
        if (key in result and isinstance(result[key], dict)
                and isinstance(value, dict)):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = deepcopy(value)
    return result


def _normalize_heading_keys(cfg: dict) -> dict:
    """Coerce YAML heading keys (which arrive as strings) to ints 1-6."""
    headings = cfg.get("headings")
    if isinstance(headings, dict):
        normalized = {}
        for k, v in headings.items():
            try:
                normalized[int(k)] = v
            except (ValueError, TypeError):
                normalized[k] = v
        cfg["headings"] = normalized
    return cfg


def load_style(path=None) -> dict:
    """Return the effective style config: defaults, optionally overridden by YAML.

    ``path`` points at a YAML file whose keys override the corresponding
    defaults (deep merge). Unknown keys are ignored. Invalid YAML raises
    :class:`StyleError`.
    """
    if path is None:
        return deepcopy(DEFAULT_STYLE)

    p = Path(path)
    if not p.is_file():
        raise StyleError(f"style config not found: {path}")
    try:
        loaded = yaml.safe_load(p.read_text(encoding="utf-8"))
    except yaml.YAMLError as exc:
        raise StyleError(f"invalid YAML in {path}: {exc}") from exc

    if loaded is None:
        return deepcopy(DEFAULT_STYLE)
    if not isinstance(loaded, dict):
        raise StyleError(f"style config must be a mapping, got {type(loaded).__name__}")

    loaded = _normalize_heading_keys(loaded)
    return _deep_merge(DEFAULT_STYLE, loaded)


def dump_default_style() -> str:
    """Return the default style config serialized as YAML (for --dump-config)."""
    return yaml.safe_dump(DEFAULT_STYLE, sort_keys=False, allow_unicode=True)


_HEX_COLOR_RE = re.compile(r"^[0-9A-Fa-f]{6}$")


def _as_hex(value):
    """Normalize a color value to a 6-digit hex string, or None.

    YAML parses an all-digit color like ``123456`` as an int, so we stringify
    and zero-pad to 6 digits before validating.
    """
    if value is None:
        return None
    if isinstance(value, int):
        value = f"{value:06d}"
    value = str(value).strip()
    return value.upper() if _HEX_COLOR_RE.match(value) else None


def _coerce_color(value, default, where):
    """Return an ``RGBColor`` from a hex string, or ``None`` when unset.

    Malformed values warn to stderr and fall back to ``default`` (which may be
    ``None``), so one bad color doesn't abort a conversion.
    """
    if value is None:
        value = default
    if value is None:
        return None
    hex_value = _as_hex(value)
    if hex_value:
        return RGBColor.from_string(hex_value)
    print(f"WARNING: invalid color {value!r} for {where}; using default",
          file=sys.stderr)
    default_hex = _as_hex(default)
    return RGBColor.from_string(default_hex) if default_hex else None


def _coerce_number(value, default, where):
    """Return a float from a numeric value, warning + falling back on error."""
    if value is None:
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        print(f"WARNING: invalid number {value!r} for {where}; using default",
              file=sys.stderr)
        return default


class StyleConfig:
    """Typed, validated accessor over the merged style dict.

    Wraps the plain config dict (defaults deep-merged with any YAML) and exposes
    coercion helpers so render code can pull points/inches/colors without
    repeating validation. Malformed individual fields fall back to the default
    for that field (with a warning) rather than raising.
    """

    def __init__(self, cfg: dict = None):
        self._cfg = cfg if cfg is not None else deepcopy(DEFAULT_STYLE)

    @property
    def raw(self) -> dict:
        return self._cfg

    def section(self, name: str) -> dict:
        return self._cfg.get(name, {}) or {}

    # -- typed getters ------------------------------------------------------- #
    def num(self, section: str, key: str, default=0):
        return _coerce_number(self.section(section).get(key), default,
                              f"{section}.{key}")

    def color(self, section: str, key: str, default=None):
        return _coerce_color(self.section(section).get(key), default,
                             f"{section}.{key}")

    def flag(self, section: str, key: str, default=False):
        val = self.section(section).get(key, default)
        return bool(val)

    def text(self, section: str, key: str, default=None):
        val = self.section(section).get(key, default)
        return val if val is not None else default

    def heading(self, level: int) -> dict:
        headings = self._cfg.get("headings", {})
        return headings.get(level) or DEFAULT_STYLE["headings"].get(level, {})


# The style in effect for the current build_docx call. Inline helpers
# (_style_run/_emit_run) read this so they can apply body and inline-code
# formatting without threading a config argument through every call site. It is
# set at the start of build_docx and reset when done.
_ACTIVE_STYLE = StyleConfig(DEFAULT_STYLE)


def slugify_heading(text: str) -> str:
    """Convert heading text to a GitHub-style anchor slug.

    Mirrors the common markdown TOC convention:
    - lowercase
    - strip inline markdown (**bold**, *italic*, `code`, links)
    - spaces -> hyphens
    - drop characters that aren't word chars, spaces, or hyphens
    """
    # Strip wiki links and inline markdown so the slug matches the rendered text
    text = re.sub(r"\[\[([^|\]]+)\|([^\]]+)\]\]", r"\2", text)
    text = re.sub(r"\[\[([^\]]+)\]\]", lambda m: m.group(1).split("/")[-1], text)
    text = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", text)  # [text](url) -> text
    text = re.sub(r"[*_`]", "", text)  # drop emphasis/code markers
    text = text.strip().lower()
    text = re.sub(r"[^\w\s-]", "", text)  # drop punctuation
    text = re.sub(r"\s+", "-", text)      # spaces -> hyphens
    text = re.sub(r"-+", "-", text)       # collapse repeated hyphens
    return text.strip("-")


def build_heading_anchors(blocks: list) -> dict:
    """Build a map of anchor keys -> unique bookmark name for every heading.

    Registers each heading under both its GitHub-style slug (for
    ``[text](#slug)`` links) and its lowercased raw text (for Obsidian
    ``[[#Heading Text]]`` links). Duplicate slugs get a numeric suffix,
    matching GitHub's behavior.
    """
    anchors = {}
    slug_counts = {}
    for idx, block in enumerate(blocks):
        if block.get("type") in ("h1", "h2", "h3", "h4", "h5", "h6"):
            base_slug = slugify_heading(block["text"])
            if not base_slug:
                base_slug = "section"
            count = slug_counts.get(base_slug, 0)
            slug = base_slug if count == 0 else f"{base_slug}-{count}"
            slug_counts[base_slug] = count + 1

            bookmark = f"_toc_{idx}_{slug}"
            block["_bookmark"] = bookmark
            # Key by slug (GitHub anchors) and by raw lowercased text (Obsidian)
            anchors.setdefault(slug, bookmark)
            raw_key = re.sub(r"\s+", " ", block["text"].strip().lower())
            anchors.setdefault(raw_key, bookmark)
            # Also key by the raw text's own slug variations already covered by slug
    return anchors


def parse_markdown(text: str) -> list:
    """Parse markdown into a list of blocks with type and content."""
    blocks = []
    lines = text.split("\n")
    i = 0

    # Skip YAML frontmatter if present
    if lines and lines[0].strip() == "---":
        i = 1
        while i < len(lines) and lines[i].strip() != "---":
            i += 1
        i += 1  # skip the closing ---

    while i < len(lines):
        line = lines[i]

        # Heading (ATX, levels 1-6). Level derives from the number of leading
        # '#' so ##### / ###### work and ordering bugs can't creep in.
        heading_match = re.match(r"^(#{1,6})\s+(.*)$", line)
        if heading_match:
            level = len(heading_match.group(1))
            blocks.append({"type": f"h{level}", "text": heading_match.group(2).strip()})

        # Horizontal rule
        elif line.strip() == "---":
            blocks.append({"type": "hr"})

        # Fenced code block (``` ... ```), optionally with a language info string.
        elif line.startswith("```"):
            info = line[3:].strip()
            language = info.split()[0] if info else ""
            code_lines = []
            i += 1  # skip the opening ```
            while i < len(lines) and not lines[i].startswith("```"):
                code_lines.append(lines[i])
                i += 1
            if i < len(lines):
                i += 1  # skip the closing ```
            blocks.append({"type": "code_block", "lines": code_lines, "language": language})
            continue

        # Blockquote (email body) — collect all consecutive > lines
        elif line.startswith("> ") or line.strip() == ">":
            quote_lines = []
            while i < len(lines) and (lines[i].startswith("> ") or lines[i].strip() == ">"):
                content = lines[i][2:] if lines[i].startswith("> ") else ""
                quote_lines.append(content)
                i += 1
            blocks.append({"type": "blockquote", "lines": quote_lines})
            continue  # skip the i += 1 at the end

        # Table
        elif line.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].startswith("|"):
                table_lines.append(lines[i])
                i += 1
            blocks.append({"type": "table", "lines": table_lines})
            continue

        # Bold metadata lines (From:, To:, Subject:, etc.) — group consecutive ones
        elif line.startswith("**From:**") or line.startswith("**To:**") or \
             line.startswith("**Subject:**") or line.startswith("**CC:**") or \
             line.startswith("**Body:**") or line.startswith("**Body"):
            meta_lines = []
            while i < len(lines) and (lines[i].startswith("**From:") or lines[i].startswith("**To:") or \
                  lines[i].startswith("**Subject:") or lines[i].startswith("**CC:") or \
                  lines[i].startswith("**Body")):
                meta_lines.append(lines[i].rstrip())
                i += 1
            blocks.append({"type": "meta", "lines": meta_lines})
            continue

        # Note/callout (starts with >  but after blockquote check)
        elif line.startswith("> **Note"):
            blocks.append({"type": "note", "text": line[2:]})

        # Bullet list items (including GFM task items: - [ ] / - [x])
        elif line.startswith("- "):
            list_items = []
            while i < len(lines):
                if lines[i].startswith("- "):
                    # Start of a bullet item — collect it plus continuation lines
                    item_text = lines[i][2:].strip()
                    # Detect a task-list checkbox prefix.
                    task_match = re.match(r"\[([ xX])\]\s+(.*)", item_text)
                    if task_match:
                        checked = task_match.group(1).lower() == "x"
                        item_text = task_match.group(2)
                    else:
                        checked = None  # not a task item
                    i += 1
                    # Collect continuation lines (indented, not a new bullet or block element)
                    while i < len(lines):
                        next_line = lines[i]
                        # Indented continuation (2+ spaces, not a new bullet)
                        if re.match(r"^\s{2,}\S", next_line) and not next_line.strip().startswith("- "):
                            item_text += "\n" + next_line.strip()
                            i += 1
                        else:
                            break
                    list_items.append({"text": item_text, "checked": checked})
                else:
                    break
            blocks.append({"type": "list", "items": list_items})
            continue

        # Numbered list items (1. , 2. , etc.)
        elif re.match(r"^\d+\.\s", line):
            list_items = []
            while i < len(lines):
                if re.match(r"^\d+\.\s", lines[i]):
                    # Start of a numbered item — collect it plus any continuation/sub-items
                    item_text = re.sub(r"^\d+\.\s", "", lines[i]).strip()
                    sub_items = []
                    i += 1
                    # Collect continuation lines (indented, not a new numbered item or block)
                    while i < len(lines):
                        next_line = lines[i]
                        # Indented sub-bullet (e.g., "   - something")
                        if re.match(r"^\s{2,}- ", next_line):
                            sub_items.append(next_line.strip()[2:])  # strip "- " prefix
                            i += 1
                        # Indented continuation text (e.g., "   OR no match found")
                        elif re.match(r"^\s{2,}\S", next_line) and not re.match(r"^\d+\.\s", next_line):
                            item_text += "\n" + next_line.strip()
                            i += 1
                        else:
                            break
                    list_items.append({"text": item_text, "sub_items": sub_items})
                # Allow blank lines between numbered items (loose lists)
                elif lines[i].strip() == "":
                    # Peek ahead: if the next non-blank line is a numbered item, skip the blank
                    peek = i + 1
                    while peek < len(lines) and lines[peek].strip() == "":
                        peek += 1
                    if peek < len(lines) and re.match(r"^\d+\.\s", lines[peek]):
                        i = peek  # skip blank lines, continue collecting
                    else:
                        break
                else:
                    break
            blocks.append({"type": "numbered_list", "items": list_items})
            continue

        # Regular paragraph
        elif line.strip():
            # Collect consecutive non-empty lines as a paragraph
            # Preserve trailing double-space as line breaks
            para_lines = []
            while i < len(lines) and lines[i].strip() and not lines[i].startswith("#") \
                  and not lines[i].startswith("|") and not lines[i].startswith("> ") \
                  and not lines[i].startswith("---") and not lines[i].startswith("- ") \
                  and not re.match(r"^\d+\.\s", lines[i]) \
                  and not lines[i].startswith("**From:") \
                  and not lines[i].startswith("**To:") and not lines[i].startswith("**Subject:") \
                  and not lines[i].startswith("**Body"):
                para_lines.append(lines[i])
                i += 1
            # Join lines: if a line ends with 2+ spaces (markdown line break), use \n; otherwise space
            joined = ""
            prev_had_break = False
            for idx, pl in enumerate(para_lines):
                if idx > 0 and not prev_had_break:
                    joined += " "
                if pl.endswith("  "):
                    joined += pl.rstrip() + "\n"
                    prev_had_break = True
                else:
                    joined += pl
                    prev_had_break = False
            blocks.append({"type": "paragraph", "text": joined})
            continue

        i += 1

    return blocks


def _resolve_internal_anchor(target: str, anchors: dict):
    """Resolve a link target to a heading bookmark, or None if not internal/known.

    Handles:
    - ``#section-name``        (GitHub anchor)
    - ``#Section Name``        (raw heading text)
    - ``#Heading`` from ``[[#Heading]]`` (Obsidian, passed in without the #)
    """
    if not anchors:
        return None
    key = target.strip()
    if key.startswith("#"):
        key = key[1:]
    # Try GitHub-style slug match first, then raw lowercased text match
    slug = key.lower()
    if slug in anchors:
        return anchors[slug]
    raw = re.sub(r"\s+", " ", key.strip().lower())
    if raw in anchors:
        return anchors[raw]
    # Last resort: slugify the target and try again (handles spaced anchors)
    slugged = slugify_heading(key)
    if slugged in anchors:
        return anchors[slugged]
    return None


def _build_hyperlink_run(text: str):
    """Build a ``w:r`` for a hyperlink, styled with direct color + underline.

    Reads link color/underline from the active style so links render as links
    without depending on a named "Hyperlink" character style.
    """
    sc = _ACTIVE_STYLE
    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    color_hex = _as_hex(sc.text("links", "color", "0563C1"))
    if color_hex:
        color = OxmlElement('w:color')
        color.set(qn('w:val'), color_hex)
        rPr.append(color)
    if sc.flag("links", "underline", True):
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)
    run.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    run.append(t)
    return run


def _add_internal_hyperlink(paragraph, text: str, bookmark: str):
    """Add a run that hyperlinks to an internal bookmark (w:anchor)."""
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('w:anchor'), bookmark)
    hyperlink.append(_build_hyperlink_run(text))
    paragraph._p.append(hyperlink)


def _add_external_hyperlink(paragraph, text: str, url: str):
    """Add a run that hyperlinks to an external URL (creates a relationship)."""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    hyperlink.append(_build_hyperlink_run(text))
    paragraph._p.append(hyperlink)


def _apply_page_setup(doc, sc: "StyleConfig"):
    """Set page size and margins on the document's first section from config."""
    section = doc.sections[0]
    page = sc.section("page")
    size = str(page.get("size", "letter")).lower()
    if isinstance(page.get("size"), dict):
        w = _coerce_number(page["size"].get("width_in"), 8.5, "page.size.width_in")
        h = _coerce_number(page["size"].get("height_in"), 11.0, "page.size.height_in")
    else:
        w, h = _PAGE_SIZES.get(size, _PAGE_SIZES["letter"])
    section.page_width = Inches(w)
    section.page_height = Inches(h)

    margins = page.get("margins_in", {}) or {}
    section.top_margin = Inches(_coerce_number(margins.get("top"), 1.0, "page.margins_in.top"))
    section.bottom_margin = Inches(_coerce_number(margins.get("bottom"), 1.0, "page.margins_in.bottom"))
    section.left_margin = Inches(_coerce_number(margins.get("left"), 1.25, "page.margins_in.left"))
    section.right_margin = Inches(_coerce_number(margins.get("right"), 1.25, "page.margins_in.right"))


def _apply_paragraph_format(paragraph, *, space_before_pt=None, space_after_pt=None,
                            line_spacing=None, left_indent_in=None):
    """Apply direct paragraph-format properties, skipping any left as ``None``."""
    pf = paragraph.paragraph_format
    if space_before_pt is not None:
        pf.space_before = Pt(space_before_pt)
    if space_after_pt is not None:
        pf.space_after = Pt(space_after_pt)
    if line_spacing is not None:
        pf.line_spacing = line_spacing
    if left_indent_in is not None:
        pf.left_indent = Inches(left_indent_in)


def _apply_run_format(run, *, font=None, size_pt=None, color=None, bold=None,
                      italic=None, strike=None, fill=None):
    """Apply direct run/character formatting, skipping any left as ``None``.

    ``color`` and ``fill`` accept an ``RGBColor`` (or hex string); ``fill`` adds
    run-level shading via ``w:shd`` on the run properties.
    """
    if font is not None:
        run.font.name = font
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if strike is not None:
        run.font.strike = strike
    if color is not None:
        run.font.color.rgb = color if isinstance(color, RGBColor) else RGBColor.from_string(str(color).upper())
    if fill is not None:
        hexfill = _as_hex(fill)
        if hexfill:
            rPr = run._r.get_or_add_rPr()
            existing = rPr.find(qn('w:shd'))
            if existing is not None:
                rPr.remove(existing)
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), hexfill)
            rPr.append(shd)


def _content_width_emu(paragraph):
    """Return the usable content width (page width minus margins) in EMU.

    Falls back to 6.5 inches when section geometry is unavailable.
    """
    try:
        section = paragraph.part.document.sections[0]
        width = section.page_width - section.left_margin - section.right_margin
        if width and width > 0:
            return width
    except Exception:
        pass
    return Inches(6.5)


# Remote-image fetch limits. Kept conservative so a slow or oversized response
# can't stall or blow up a conversion. python-docx only accepts raster formats
# (PNG/JPEG/GIF/BMP/TIFF); anything else falls back to alt text.
REMOTE_IMAGE_TIMEOUT = 10          # seconds
REMOTE_IMAGE_MAX_BYTES = 10 * 1024 * 1024  # 10 MiB


def _fetch_remote_image(url: str):
    """Download a remote image and return it as a ``BytesIO`` stream.

    Enforces a timeout and a maximum size. Raises on any network error, an
    oversized response, or a non-image content type; callers are expected to
    catch and fall back to alt text.
    """
    req = urllib.request.Request(url, headers={"User-Agent": "md-to-docx"})
    with urllib.request.urlopen(req, timeout=REMOTE_IMAGE_TIMEOUT) as resp:
        content_type = (resp.headers.get("Content-Type") or "").lower()
        if content_type and not content_type.startswith("image/"):
            raise ValueError(f"not an image content type: {content_type!r}")
        # Read one byte past the cap so we can detect an over-limit response.
        data = resp.read(REMOTE_IMAGE_MAX_BYTES + 1)
    if len(data) > REMOTE_IMAGE_MAX_BYTES:
        raise ValueError("remote image exceeds size limit")
    return io.BytesIO(data)


def _add_image(paragraph, src: str, alt: str, base_dir=None, allow_remote=False):
    """Embed an image into ``paragraph``; fall back to alt text otherwise.

    Local files are resolved against ``base_dir``. Remote ``http(s)`` sources are
    fetched only when ``allow_remote`` is true (opt-in), subject to a timeout and
    size cap. Missing/unreadable files, disallowed or failed remote fetches, and
    unsupported formats render the alt text (or a visible placeholder) instead of
    raising, so a bad image never aborts the conversion. Images wider than the
    content area are scaled down to fit while preserving aspect ratio.
    """
    def _fallback():
        paragraph.add_run(alt if alt else "[image]")

    def _embed(image_source):
        run = paragraph.add_run()
        # Add at natural size first so we can measure, then cap the width.
        picture = run.add_picture(image_source)
        max_width = _content_width_emu(paragraph)
        if picture.width and picture.width > max_width:
            ratio = max_width / picture.width
            picture.width = int(picture.width * ratio)
            picture.height = int(picture.height * ratio)

    # Remote source.
    if re.match(r"^https?://", src, re.IGNORECASE):
        if not allow_remote:
            _fallback()
            return
        try:
            _embed(_fetch_remote_image(src))
        except Exception:
            # Network error, timeout, oversize, unsupported/corrupt format, etc.
            _fallback()
        return

    # Local source.
    path = Path(src)
    if not path.is_absolute() and base_dir is not None:
        path = Path(base_dir) / path

    if not path.is_file():
        _fallback()
        return

    try:
        _embed(str(path))
    except Exception:
        # Unreadable/corrupt image, unsupported format, etc.
        _fallback()


# CommonMark backslash escapes: a backslash before ASCII punctuation makes that
# punctuation literal (it cannot start emphasis, a link, code, etc.). We hide
# escaped punctuation inside a \x00ESC\x00<codepoint>\x00 sentinel so none of the
# inline regexes below can match it, then restore the literal character when the
# text is finally emitted as a run (see _unescape_sentinels / _emit_run).
_ASCII_PUNCT = r"""!"#$%&'()*+,\-./:;<=>?@[\\\]^_`{|}~"""
_ESCAPE_RE = re.compile(r"\\([" + _ASCII_PUNCT + r"])")


def _escape_punct(text: str) -> str:
    """Replace ``\\<punct>`` with an ESC sentinel so it survives inline parsing.

    A backslash before a non-punctuation character is left untouched (CommonMark
    treats it as a literal backslash).
    """
    return _ESCAPE_RE.sub(lambda m: f"\x00ESC\x00{ord(m.group(1))}\x00", text)


def _unescape_sentinels(text: str) -> str:
    """Restore ESC sentinels produced by :func:`_escape_punct` to literal chars."""
    if "\x00ESC\x00" not in text:
        return text
    return re.sub(r"\x00ESC\x00(\d+)\x00", lambda m: chr(int(m.group(1))), text)


def add_formatted_text(paragraph, text: str, anchors: dict = None, base_dir=None,
                       allow_remote_images=False):
    """Add text to a paragraph, handling **bold**, *italic*, `code`, links, and \\n line breaks.

    Link handling:
    - ``[text](#anchor)`` and ``[[#Heading]]`` / ``[[#Heading|alias]]`` resolve to
      internal hyperlinks pointing at heading bookmarks (used for a clickable TOC).
    - ``[text](http...)`` becomes an external hyperlink.
    - Other ``[[wiki links]]`` degrade to plain text (as before).

    ``base_dir`` is the directory used to resolve relative image paths (see
    :func:`_add_image`); it defaults to the current working directory.
    ``allow_remote_images`` enables fetching ``http(s)`` image sources.
    """
    # Hide backslash-escaped punctuation before any tokenization so escaped
    # markers (e.g. ``\*not italic\*``) cannot start a construct.
    text = _escape_punct(text)
    # Internal Obsidian heading links: [[#Heading]] or [[#Heading|alias]]
    def _obsidian_heading_link(m):
        target = m.group(1)
        alias = m.group(2) if m.group(2) else target
        # Placeholder token; resolved during segment rendering below
        return f"\x00LINK\x00{target}\x00{alias}\x00"

    text = re.sub(r"\[\[#([^|\]]+)\|([^\]]+)\]\]", lambda m: f"\x00LINK\x00#{m.group(1)}\x00{m.group(2)}\x00", text)
    text = re.sub(r"\[\[#([^\]]+)\]\]", lambda m: f"\x00LINK\x00#{m.group(1)}\x00{m.group(1)}\x00", text)

    # Images ![alt](src) -> IMG token. Done BEFORE links so the leading '!' is
    # consumed and the src/alt aren't mistaken for a plain link.
    text = re.sub(r"!\[([^\]]*)\]\(([^)]+)\)",
                  lambda m: f"\x00IMG\x00{m.group(2)}\x00{m.group(1)}\x00", text)

    # Standard markdown links [text](target) -> placeholder token
    text = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", lambda m: f"\x00LINK\x00{m.group(2)}\x00{m.group(1)}\x00", text)

    # Angle autolinks <https://...> / <mailto:...> -> LINK token (label == URL).
    text = re.sub(r"<((?:https?://|mailto:)[^>\s]+)>",
                  lambda m: f"\x00LINK\x00{m.group(1)}\x00{m.group(1)}\x00", text)

    # Remaining (non-heading) wiki links: [[path/page]] -> page, [[page|alias]] -> alias
    text = re.sub(r"\[\[([^|\]]+)\|([^\]]+)\]\]", r"\2", text)  # [[target|alias]] -> alias
    text = re.sub(r"\[\[([^\]]+)\]\]", lambda m: m.group(1).split("/")[-1], text)  # [[path/page]] -> page

    # Split on newlines first to handle line breaks
    segments = text.split("\n")
    for seg_idx, segment in enumerate(segments):
        if seg_idx > 0:
            # Add a line break between segments
            paragraph.add_run().add_break()

        # Split out link/image placeholders so they aren't chopped by the emphasis regex
        link_parts = re.split(
            r"((?:\x00LINK\x00|\x00IMG\x00)[^\x00]*\x00[^\x00]*\x00)", segment)
        for lp in link_parts:
            img_match = re.match(r"\x00IMG\x00([^\x00]*)\x00([^\x00]*)\x00", lp)
            if img_match:
                src = _unescape_sentinels(img_match.group(1)).strip()
                alt = _unescape_sentinels(img_match.group(2))
                _add_image(paragraph, src, alt, base_dir,
                           allow_remote=allow_remote_images)
                continue

            link_match = re.match(r"\x00LINK\x00([^\x00]*)\x00([^\x00]*)\x00", lp)
            if link_match:
                target = _unescape_sentinels(link_match.group(1))
                label = _unescape_sentinels(link_match.group(2))
                bookmark = _resolve_internal_anchor(target, anchors)
                if bookmark:
                    _add_internal_hyperlink(paragraph, label, bookmark)
                elif re.match(r"^(https?://|mailto:)", target.strip()):
                    _add_external_hyperlink(paragraph, label, target.strip())
                else:
                    # Unresolved internal anchor or relative link: render label as plain text
                    paragraph.add_run(label)
                continue

            # Render emphasis/code recursively so nested markers combine, e.g.
            # **`code`** is both bold and monospace, or ***bold italic***.
            _render_inline(paragraph, lp)


# Inline-emphasis tokens, ordered so the longest/most-specific markers win.
# Each entry: (compiled regex, formatting flag to add, whether inner text is
# itself markdown that should be re-processed). Code spans are literal — their
# contents are not re-parsed for emphasis.
# Underscore emphasis follows CommonMark's intra-word rule: an underscore
# delimiter is only a valid opener/closer at a word boundary. This prevents
# identifiers like ``date_updated`` or ``snake_case`` from being italicized.
# The delimiter must be preceded by start-of-string or a non-word char, and
# followed by start-of-string or a non-word char on the closing side. Asterisk
# emphasis has no such restriction (it may appear mid-word).
_INLINE_PATTERNS = [
    (re.compile(r"\*\*\*(.+?)\*\*\*", re.DOTALL), ("bold", "italic"), True),
    (re.compile(r"(?<![\w_])___(?!_)(.+?)(?<!_)___(?![\w_])", re.DOTALL), ("bold", "italic"), True),
    (re.compile(r"\*\*(.+?)\*\*", re.DOTALL), ("bold",), True),
    (re.compile(r"(?<![\w_])__(?!_)(.+?)(?<!_)__(?![\w_])", re.DOTALL), ("bold",), True),
    (re.compile(r"`([^`]+)`", re.DOTALL), ("code",), False),
    # Strikethrough (GFM). Placed after code so a ``~~`` inside a code span stays
    # literal, mirroring how emphasis is ignored inside code.
    (re.compile(r"~~(.+?)~~", re.DOTALL), ("strike",), True),
    (re.compile(r"\*(.+?)\*", re.DOTALL), ("italic",), True),
    (re.compile(r"(?<![\w_])_(?!_)(.+?)(?<!_)_(?![\w_])", re.DOTALL), ("italic",), True),
]


# Bare URL matcher for autolinking running text. Trailing punctuation is
# trimmed separately so a URL at the end of a sentence doesn't swallow the
# period/paren/etc.
_BARE_URL_RE = re.compile(r"https?://[^\s<>()]+")
_URL_TRAILING_PUNCT = ".,;:!?)]}'\""


def _style_run(run, flags):
    """Apply the accumulated formatting flags to a run, using the active style.

    Body font/size/color come from ``body.*``; inline ``code`` spans use
    ``inline_code.*``. Bold/italic/strike are markdown-driven flags.
    """
    sc = _ACTIVE_STYLE
    # Base body formatting for every inline run.
    _apply_run_format(run,
                      font=sc.text("body", "font"),
                      size_pt=sc.num("body", "size_pt", 11),
                      color=sc.color("body", "color", default="000000"))
    if "bold" in flags:
        run.bold = True
    if "italic" in flags:
        run.italic = True
    if "strike" in flags:
        run.font.strike = True
    if "code" in flags:
        _apply_run_format(run,
                          font=sc.text("inline_code", "font", "Consolas"),
                          size_pt=sc.num("inline_code", "size_pt", 10),
                          color=sc.color("inline_code", "color"),
                          fill=sc.text("inline_code", "fill"))


def _emit_run(paragraph, text, flags):
    """Add a run with the accumulated formatting flags applied.

    Non-code text is scanned for bare URLs, which become external hyperlinks.
    URLs inside a code span are left literal.
    """
    if text == "":
        return
    # Restore any backslash-escaped punctuation hidden during inline parsing.
    text = _unescape_sentinels(text)
    if text == "":
        return

    # Inside a code span, don't linkify — emit a single styled run verbatim.
    if "code" in flags:
        _style_run(paragraph.add_run(text), flags)
        return

    # Split the text around bare URLs, emitting external hyperlinks for them and
    # normally-styled runs for the rest.
    pos = 0
    for m in _BARE_URL_RE.finditer(text):
        url = m.group(0)
        # Trim trailing sentence punctuation off the URL, leaving it as text.
        trimmed = url.rstrip(_URL_TRAILING_PUNCT)
        trailing = url[len(trimmed):]
        before = text[pos:m.start()]
        if before:
            _style_run(paragraph.add_run(before), flags)
        _add_external_hyperlink(paragraph, trimmed, trimmed)
        if trailing:
            _style_run(paragraph.add_run(trailing), flags)
        pos = m.end()
    remainder = text[pos:]
    if remainder:
        _style_run(paragraph.add_run(remainder), flags)


def _render_inline(paragraph, text, flags=frozenset()):
    """Recursively render inline markdown emphasis/code, stacking formatting.

    Finds the earliest emphasis or code token, emits the text before it with the
    current flags, then recurses into the token's inner content with the added
    flag so combinations like **`code`** (bold + monospace) or ***bold italic***
    render correctly. Code spans are treated as literal and not re-parsed.
    """
    # Find the earliest-matching pattern in the text.
    best = None  # (start, match, add_flags, recurse_inner)
    for pattern, add_flags, recurse_inner in _INLINE_PATTERNS:
        m = pattern.search(text)
        if m is not None and (best is None or m.start() < best[0]):
            best = (m.start(), m, add_flags, recurse_inner)

    if best is None:
        _emit_run(paragraph, text, flags)
        return

    start, m, add_flags, recurse_inner = best
    # Text before the token keeps the current flags.
    _emit_run(paragraph, text[:start], flags)

    new_flags = flags.union(add_flags)
    inner = m.group(1)
    if recurse_inner:
        _render_inline(paragraph, inner, new_flags)
    else:
        _emit_run(paragraph, inner, new_flags)

    # Continue with whatever follows the token.
    _render_inline(paragraph, text[m.end():], flags)


def _make_element(tag, **attrs):
    """Create an OxmlElement with w: namespace attributes."""
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn(f'w:{k}'), v)
    return el


# Blockquote left-bar styling. Matches the "Blockquote Example" in the template
# (see create_template.py): a gray vertical bar on the left plus a hanging indent.
BLOCKQUOTE_BAR_COLOR = '999999'   # medium gray vertical bar
BLOCKQUOTE_BAR_SIZE = '18'        # border thickness in eighths of a point (~2.25pt)
BLOCKQUOTE_BAR_SPACE = '12'       # space between bar and text, in points


def _apply_blockquote_bar(paragraph, color=BLOCKQUOTE_BAR_COLOR,
                          width_pt=2.25, gap_pt=12):
    """Add a left vertical bar (paragraph border) to a blockquote paragraph.

    ``width_pt`` is the bar thickness in points (converted to the eighths-of-a-
    point ``sz`` unit); ``gap_pt`` is the space between bar and text in points.
    Applied as a direct border so it needs no named style.
    """
    pPr = paragraph._p.get_or_add_pPr()
    # Remove any existing borders so repeated calls stay idempotent.
    existing = pPr.find(qn('w:pBdr'))
    if existing is not None:
        pPr.remove(existing)
    pBdr = OxmlElement('w:pBdr')
    sz = str(max(1, int(round(width_pt * 8))))   # points -> eighths of a point
    left = _make_element('w:left', val='single', sz=sz,
                         space=str(int(round(gap_pt))),
                         color=_as_hex(color) or "999999")
    pBdr.append(left)
    pPr.append(pBdr)


# Code-block styling: a light gray fill plus a same-color border on all four
# sides. The border's w:space attribute (in points) is what gives the code text
# breathing room inside the shaded box — Word has no separate padding property
# for shading, so a matching-color border is the standard way to fake it.
CODE_BLOCK_FILL = 'F2F2F2'        # light gray background
CODE_BLOCK_BORDER_SPACE = '6'     # padding between border and text, in points


def _apply_code_block_box(paragraph, fill=CODE_BLOCK_FILL,
                          space=CODE_BLOCK_BORDER_SPACE):
    """Give a code paragraph a shaded background with interior padding.

    Adds paragraph shading and a four-sided border whose color matches the fill,
    so the border is invisible but its ``w:space`` acts as left/right padding.
    Vertical breathing room comes from the paragraph's space before/after.
    """
    pPr = paragraph._p.get_or_add_pPr()
    fill = _as_hex(fill) or "F2F2F2"

    # Shading (the gray fill).
    existing_shd = pPr.find(qn('w:shd'))
    if existing_shd is not None:
        pPr.remove(existing_shd)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    pPr.append(shd)

    # Four-sided border, same color as the fill, used purely for its spacing.
    existing_bdr = pPr.find(qn('w:pBdr'))
    if existing_bdr is not None:
        pPr.remove(existing_bdr)
    pBdr = OxmlElement('w:pBdr')
    for side in ('top', 'left', 'bottom', 'right'):
        pBdr.append(_make_element(f'w:{side}', val='single', sz='4',
                                  space=space, color=fill))
    pPr.append(pBdr)


def _apply_custom_table_style(doc, table, sc: "StyleConfig"):
    """Apply direct table formatting (borders, margins, width) from config.

    No named table style is referenced; borders are drawn directly via
    ``w:tblBorders`` so the look is fully controlled by ``table.*``.
    """
    border = sc.section("table").get("border", {}) or {}
    b_style = border.get("style", "single")
    b_color = _as_hex(border.get("color")) or "808080"
    b_size = str(max(2, int(round(_coerce_number(border.get("width_pt"), 0.75, "table.border.width_pt") * 8))))
    # Which edges get a visible line; any others are explicitly turned off.
    enabled_edges = border.get("edges")
    if enabled_edges is None:
        enabled_edges = ["top", "bottom", "insideH"]
    enabled_edges = set(enabled_edges)

    margins = sc.section("table").get("cell_margins_pt", {}) or {}
    def _pt_to_dxa(v, d):
        return str(int(round(_coerce_number(v, d, "table.cell_margins_pt") * 20)))

    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    # Draw only the enabled edges; disabled edges are set to 'nil' so no line
    # shows (this is how you get a horizontal-rules-only table).
    existing_bdr = tblPr.find(qn('w:tblBorders'))
    if existing_bdr is not None:
        tblPr.remove(existing_bdr)
    tblBorders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        if edge in enabled_edges:
            tblBorders.append(_make_element(f'w:{edge}', val=b_style, sz=b_size,
                                            space='0', color=b_color))
        else:
            tblBorders.append(_make_element(f'w:{edge}', val='nil'))
    tblPr.append(tblBorders)

    # tblLook: emphasize first row/column.
    existing_look = tblPr.find(qn('w:tblLook'))
    if existing_look is not None:
        tblPr.remove(existing_look)
    tblPr.append(_make_element('w:tblLook', val='04A0', firstRow='1', lastRow='0',
                               firstColumn='1', lastColumn='0', noHBand='0', noVBand='1'))

    # Cell margins from config (points -> dxa).
    existing_mar = tblPr.find(qn('w:tblCellMar'))
    if existing_mar is not None:
        tblPr.remove(existing_mar)
    tblCellMar = OxmlElement('w:tblCellMar')
    tblCellMar.append(_make_element('w:top', w=_pt_to_dxa(margins.get("top"), 3.6), type='dxa'))
    tblCellMar.append(_make_element('w:left', w=_pt_to_dxa(margins.get("left"), 5.4), type='dxa'))
    tblCellMar.append(_make_element('w:bottom', w=_pt_to_dxa(margins.get("bottom"), 3.6), type='dxa'))
    tblCellMar.append(_make_element('w:right', w=_pt_to_dxa(margins.get("right"), 5.4), type='dxa'))
    tblPr.append(tblCellMar)

    # Width: full (100%) or auto.
    existing_w = tblPr.find(qn('w:tblW'))
    if existing_w is not None:
        tblPr.remove(existing_w)
    if sc.text("table", "width", "full") == "full":
        tblPr.append(_make_element('w:tblW', w='5000', type='pct'))
    else:
        tblPr.append(_make_element('w:tblW', w='0', type='auto'))

    existing_layout = tblPr.find(qn('w:tblLayout'))
    if existing_layout is not None:
        tblPr.remove(existing_layout)
    tblPr.append(_make_element('w:tblLayout', type='autofit'))


def _apply_table_header(table, sc: "StyleConfig"):
    """Style the header row: bold text, optional fill, and an underline rule.

    The header underline is drawn as a per-cell bottom border so it can be
    heavier/darker than the row rules, matching the reference style.
    """
    header = sc.section("table").get("header", {}) or {}
    if not table.rows:
        return
    make_bold = bool(header.get("bold", True))
    fill = _as_hex(header.get("fill"))
    ul_color = _as_hex(header.get("underline_color")) or "404040"
    ul_size = header.get("underline_width_pt")
    ul_size = str(max(2, int(round(_coerce_number(ul_size, 1.0, "table.header.underline_width_pt") * 8)))) \
        if ul_size is not None else None

    for cell in table.rows[0].cells:
        tcPr = cell._tc.get_or_add_tcPr()
        if fill:
            existing = tcPr.find(qn('w:shd'))
            if existing is not None:
                tcPr.remove(existing)
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), fill)
            tcPr.append(shd)
        # Per-cell bottom border = the header underline rule.
        if ul_size is not None:
            existing_bdr = tcPr.find(qn('w:tcBorders'))
            if existing_bdr is not None:
                tcPr.remove(existing_bdr)
            tcBorders = OxmlElement('w:tcBorders')
            tcBorders.append(_make_element('w:bottom', val='single', sz=ul_size,
                                           space='0', color=ul_color))
            tcPr.append(tcBorders)
        if make_bold:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.bold = True


def _fix_narrow_column_widths(table):
    """After cells are populated, set proportional column widths based on content."""
    tbl = table._tbl
    num_cols = len(table.columns)
    if num_cols < 2:
        return

    # Calculate max content length per column
    col_max_len = [0] * num_cols
    for row in table.rows:
        for c_idx, cell in enumerate(row.cells):
            if c_idx < num_cols:
                text_len = len(cell.text)
                col_max_len[c_idx] = max(col_max_len[c_idx], text_len)

    # Total page width in twips (6.5 inches = 9360 twips)
    PAGE_WIDTH_TWIPS = 9360

    # Calculate proportional widths based on content length
    # Use sqrt to dampen the ratio — prevents huge disparities
    import math
    col_weights = [math.sqrt(max(length, 1)) for length in col_max_len]
    total_weight = sum(col_weights)

    col_widths_twips = [int(PAGE_WIDTH_TWIPS * w / total_weight) for w in col_weights]

    # Set gridCol widths in tblGrid
    tblGrid = tbl.find(qn('w:tblGrid'))
    if tblGrid is not None:
        gridCols = tblGrid.findall(qn('w:gridCol'))
        for c_idx, gridCol in enumerate(gridCols):
            if c_idx < num_cols:
                gridCol.set(qn('w:w'), str(col_widths_twips[c_idx]))

    # Set cell widths to match
    for row in table.rows:
        for c_idx, cell in enumerate(row.cells):
            if c_idx < num_cols:
                tcPr = cell._tc.find(qn('w:tcPr'))
                if tcPr is None:
                    tcPr = OxmlElement('w:tcPr')
                    cell._tc.insert(0, tcPr)
                existing_tcW = tcPr.find(qn('w:tcW'))
                if existing_tcW is not None:
                    tcPr.remove(existing_tcW)
                tcW = _make_element('w:tcW', w=str(col_widths_twips[c_idx]), type='dxa')
                tcPr.append(tcW)


def _new_num_id(doc, abstract_num_id):
    """Create a new num entry referencing the given abstractNumId. Returns the new numId.
    
    Includes a level override to restart numbering at 1, which ensures
    each numbered list block starts fresh.
    """
    numbering_elm = doc.part.numbering_part._element
    nums = numbering_elm.findall(qn('w:num'))
    max_num_id = max((int(n.get(qn('w:numId'))) for n in nums), default=0)
    new_num_id = max_num_id + 1

    new_num = OxmlElement('w:num')
    new_num.set(qn('w:numId'), str(new_num_id))
    new_num.append(_make_element('w:abstractNumId', val=str(abstract_num_id)))

    # Add level override to force restart at 1
    lvl_override = OxmlElement('w:lvlOverride')
    lvl_override.set(qn('w:ilvl'), '0')
    start_override = OxmlElement('w:startOverride')
    start_override.set(qn('w:val'), '1')
    lvl_override.append(start_override)
    new_num.append(lvl_override)

    numbering_elm.append(new_num)
    return new_num_id


def _set_list_numbering(paragraph, num_id: int, ilvl: int = 0):
    """Attach a ``w:numPr`` (numId + level) to a paragraph for list numbering."""
    pPr = paragraph._p.get_or_add_pPr()
    numPr = OxmlElement('w:numPr')
    numPr.append(_make_element('w:ilvl', val=str(ilvl)))
    numPr.append(_make_element('w:numId', val=str(num_id)))
    pPr.append(numPr)


def _get_numbering_element(doc):
    """Return the document's numbering part XML element, creating it if absent.

    A blank ``Document()`` may not ship a numbering part until a list is used;
    ``numbering_part`` on python-docx creates a default one on access.
    """
    return doc.part.numbering_part._element


def _make_list_level(ilvl: int, num_fmt: str, lvl_text: str, left_twips: int,
                     hanging_twips: int = 360, font: str = None):
    """Build a ``w:lvl`` element for an abstractNum (bullet or decimal)."""
    lvl = OxmlElement('w:lvl')
    lvl.set(qn('w:ilvl'), str(ilvl))
    lvl.append(_make_element('w:start', val='1'))
    lvl.append(_make_element('w:numFmt', val=num_fmt))
    lvl.append(_make_element('w:lvlText', val=lvl_text))
    lvl.append(_make_element('w:lvlJc', val='left'))
    pPr = OxmlElement('w:pPr')
    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), str(left_twips))
    ind.set(qn('w:hanging'), str(hanging_twips))
    pPr.append(ind)
    lvl.append(pPr)
    if font:
        rPr = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), font)
        rFonts.set(qn('w:hAnsi'), font)
        rFonts.set(qn('w:cs'), font)
        rFonts.set(qn('w:hint'), 'default')
        rPr.append(rFonts)
        lvl.append(rPr)
    return lvl


# abstractNumIds we create in code. Large offset to avoid colliding with any
# default abstracts python-docx may ship.
_ABSTRACT_BULLET = 9001
_ABSTRACT_ORDERED = 9002


def _ensure_numbering(doc, sc: "StyleConfig"):
    """Create our own bullet and decimal abstractNum definitions once per doc.

    Returns ``(bullet_abstract_id, ordered_abstract_id)``. Indents and glyphs
    come from ``lists.*``. Idempotent: only injects if not already present.
    """
    numbering_elm = _get_numbering_element(doc)
    existing_ids = {int(a.get(qn('w:abstractNumId')))
                    for a in numbering_elm.findall(qn('w:abstractNum'))}

    bullet = sc.section("lists").get("bullet", {}) or {}
    # Per-level glyphs and their fonts. Each glyph must live in its paired font
    # (e.g. U+F0B7 exists in Symbol, but U+2022 does not).
    glyphs = bullet.get("glyphs") or ["\uF0B7", "o"]
    # Backward compat: an older single "bullet_font" applies to every level.
    if bullet.get("glyph_fonts"):
        glyph_fonts = bullet["glyph_fonts"]
    else:
        one_font = bullet.get("bullet_font", "Symbol")
        glyph_fonts = [one_font, one_font]
    b_indents = bullet.get("indent_in") or [0.25, 0.5]
    ordered = sc.section("lists").get("ordered", {}) or {}
    o_indent = _coerce_number(ordered.get("indent_in"), 0.5, "lists.ordered.indent_in")

    def _in_to_twips(v):
        return int(round(float(v) * 1440))

    def _glyph(i, fallback):
        return glyphs[i] if i < len(glyphs) else fallback

    def _gfont(i):
        return glyph_fonts[i] if i < len(glyph_fonts) else (glyph_fonts[-1] if glyph_fonts else "Symbol")

    def _bindent(i, fallback):
        return b_indents[i] if i < len(b_indents) else fallback

    if _ABSTRACT_BULLET not in existing_ids:
        abn = OxmlElement('w:abstractNum')
        abn.set(qn('w:abstractNumId'), str(_ABSTRACT_BULLET))
        abn.append(_make_list_level(0, 'bullet', _glyph(0, "\uF0B7"),
                                    _in_to_twips(_bindent(0, 0.25)), font=_gfont(0)))
        abn.append(_make_list_level(1, 'bullet', _glyph(1, "o"),
                                    _in_to_twips(_bindent(1, 0.5)), font=_gfont(1)))
        # abstractNum must precede num elements; insert at the front.
        numbering_elm.insert(0, abn)

    if _ABSTRACT_ORDERED not in existing_ids:
        abn = OxmlElement('w:abstractNum')
        abn.set(qn('w:abstractNumId'), str(_ABSTRACT_ORDERED))
        abn.append(_make_list_level(0, 'decimal', '%1.', _in_to_twips(o_indent)))
        numbering_elm.insert(0, abn)

    return _ABSTRACT_BULLET, _ABSTRACT_ORDERED


_bookmark_id_counter = [0]


def _add_bookmark(paragraph, name: str):
    """Wrap a paragraph's content start with a Word bookmark so internal
    hyperlinks (w:anchor) can jump to it.

    The bookmarkStart is inserted *after* the paragraph's <w:pPr> (paragraph
    properties), not at index 0. OOXML requires <w:pPr> to be the first child
    of <w:p>; inserting the bookmark ahead of it produces a malformed paragraph
    that desktop Word tolerates but strict renderers (SharePoint / Word Online)
    reject — they drop the <w:pPr>, which silently demotes styled headings to
    Normal. Placing the bookmark after pPr keeps the paragraph well-formed."""
    _bookmark_id_counter[0] += 1
    bm_id = str(_bookmark_id_counter[0])
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), bm_id)
    start.set(qn('w:name'), name)
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), bm_id)
    p = paragraph._p
    # Insert bookmarkStart directly after <w:pPr> if present, else at the start.
    pPr = p.find(qn('w:pPr'))
    if pPr is not None:
        pPr.addnext(start)
    else:
        p.insert(0, start)
    p.append(end)


def _set_outline_level(paragraph, level: int):
    """Set the paragraph's Word outline level (0-based) for nav pane / bookmarks.

    We no longer use named "Heading N" styles, so we mark the outline level
    directly. This keeps the document navigable and preserves PDF bookmarks.
    """
    pPr = paragraph._p.get_or_add_pPr()
    existing = pPr.find(qn('w:outlineLvl'))
    if existing is not None:
        pPr.remove(existing)
    pPr.append(_make_element('w:outlineLvl', val=str(level - 1)))


def _add_heading(doc, text: str, level: int, sc: "StyleConfig"):
    """Add a heading as a directly-formatted paragraph (no named style).

    Applies the per-level font/size/color/bold/italic and spacing from the style
    config, sets the outline level, and returns the paragraph. Inline markdown in
    the heading text is not re-parsed (headings are plain text today).
    """
    h = sc.heading(level)
    p = doc.add_paragraph()
    _apply_paragraph_format(
        p,
        space_before_pt=_coerce_number(h.get("space_before_pt"), 0, f"headings.{level}.space_before_pt"),
        space_after_pt=_coerce_number(h.get("space_after_pt"), 0, f"headings.{level}.space_after_pt"),
    )
    run = p.add_run(text)
    _apply_run_format(
        run,
        font=h.get("font"),
        size_pt=_coerce_number(h.get("size_pt"), 11, f"headings.{level}.size_pt"),
        color=_coerce_color(h.get("color"), "000000", f"headings.{level}.color"),
        bold=bool(h.get("bold", False)),
        italic=bool(h.get("italic", False)),
    )
    _set_outline_level(p, level)
    return p


def build_docx(blocks: list, output_path: str, title: str, author: str, date: str,
               base_dir=None, allow_remote_images=False, style=None):
    """Build the DOCX from parsed blocks, applying direct formatting from ``style``.

    ``base_dir`` resolves relative image paths; it defaults to the current
    working directory when not supplied. ``allow_remote_images`` enables
    fetching ``http(s)`` image sources (off by default). ``style`` is a merged
    style dict (see :func:`load_style`); when ``None`` the built-in defaults are
    used.
    """
    if base_dir is None:
        base_dir = Path.cwd()
    if style is None:
        style = load_style()
    sc = StyleConfig(style)

    # Publish the active style so inline helpers can read body/inline-code config.
    global _ACTIVE_STYLE
    _ACTIVE_STYLE = sc

    # Start from a blank document and style everything as direct formatting.
    doc = Document()
    _apply_page_setup(doc, sc)

    # Build the heading anchor map (also tags each heading block with a
    # "_bookmark" name) so TOC links can resolve to real internal hyperlinks.
    anchors = build_heading_anchors(blocks)

    # Create our own list numbering definitions (bullet + decimal) so lists work
    # without a template's numbering part. Both bullet levels share one abstract.
    bullet_abstract, numbered_abstract = _ensure_numbering(doc, sc)
    bullet2_abstract = bullet_abstract

    for block in blocks:
        if block["type"] in ("h1", "h2", "h3", "h4", "h5", "h6"):
            level = int(block["type"][1])
            h = _add_heading(doc, block["text"], level, sc)
            if block.get("_bookmark"):
                _add_bookmark(h, block["_bookmark"])

        elif block["type"] == "hr":
            p = doc.add_paragraph()
            _apply_paragraph_format(p, space_after_pt=sc.num("hr", "space_after_pt", 6))
            # Optionally draw an actual horizontal rule as a bottom border.
            if sc.flag("hr", "rule", False):
                pPr = p._p.get_or_add_pPr()
                pBdr = OxmlElement('w:pBdr')
                pBdr.append(_make_element('w:bottom', val='single', sz='6',
                                          space='1', color='BFBFBF'))
                pPr.append(pBdr)

        elif block["type"] == "meta":
            p = doc.add_paragraph()
            _apply_paragraph_format(p, space_after_pt=2)
            first = True
            for meta_line in block["lines"]:
                if not first:
                    p.add_run("\n")
                add_formatted_text(p, meta_line, anchors, base_dir=base_dir,
                                   allow_remote_images=allow_remote_images)
                first = False

        elif block["type"] == "blockquote":
            # Blockquote — a directly-formatted paragraph with a left bar. Split
            # into paragraphs on empty lines; one docx paragraph per chunk.
            bq_indent = sc.num("blockquote", "indent_in", 0.25)
            bq_after = sc.num("blockquote", "space_after_pt", 4)
            bq_color = sc.text("blockquote", "bar_color", "999999")
            bq_width = sc.num("blockquote", "bar_width_pt", 2.25)
            bq_gap = sc.num("blockquote", "bar_gap_pt", 12)

            def _emit_quote(para_lines):
                p = doc.add_paragraph()
                _apply_paragraph_format(p, space_after_pt=bq_after,
                                        left_indent_in=bq_indent)
                _apply_blockquote_bar(p, color=bq_color, width_pt=bq_width,
                                      gap_pt=bq_gap)
                first = True
                for pl in para_lines:
                    if not first:
                        p.add_run("\n")
                    add_formatted_text(p, pl, anchors, base_dir=base_dir,
                                   allow_remote_images=allow_remote_images)
                    first = False

            current_para_lines = []
            for line in block["lines"]:
                if line.strip() == "":
                    # Empty line = paragraph break — emit current paragraph
                    if current_para_lines:
                        _emit_quote(current_para_lines)
                        current_para_lines = []
                else:
                    current_para_lines.append(line)
            # Emit any remaining lines
            if current_para_lines:
                _emit_quote(current_para_lines)

        elif block["type"] == "code_block":
            # Render as monospace text in a shaded, padded box (all from config).
            cb_font = sc.text("code_block", "font", "Consolas")
            cb_size = sc.num("code_block", "size_pt", 9)
            cb_fill = sc.text("code_block", "fill", "F2F2F2")
            cb_pad = sc.num("code_block", "padding_pt", 6)
            caption_cfg = sc.section("code_block").get("caption", {}) or {}

            p = doc.add_paragraph()
            _apply_paragraph_format(
                p,
                space_before_pt=sc.num("code_block", "space_before_pt", 8),
                space_after_pt=sc.num("code_block", "space_after_pt", 8),
            )
            # Shaded background plus a same-color border that pads the text.
            _apply_code_block_box(p, fill=cb_fill, space=str(int(round(cb_pad))))
            # Optional language caption; only when a language was captured.
            language = block.get("language")
            first_line = True
            if language:
                caption = p.add_run(language)
                _apply_run_format(
                    caption,
                    font=cb_font,
                    size_pt=_coerce_number(caption_cfg.get("size_pt"), 8, "code_block.caption.size_pt"),
                    color=_coerce_color(caption_cfg.get("color"), "808080", "code_block.caption.color"),
                    italic=bool(caption_cfg.get("italic", True)),
                )
                first_line = False
            # Add each line with line breaks between them
            for line_idx, code_line in enumerate(block["lines"]):
                if line_idx > 0 or not first_line:
                    p.add_run().add_break()
                run = p.add_run(code_line)
                _apply_run_format(run, font=cb_font, size_pt=cb_size)

        elif block["type"] == "table":
            # Parse table into rows
            rows = []
            for tl in block["lines"]:
                if re.match(r"^\|[-| :]+\|$", tl.strip()):
                    continue  # skip separator row
                cells = [c.strip() for c in tl.split("|")[1:-1]]
                if cells:
                    rows.append(cells)

            if rows:
                table = doc.add_table(rows=len(rows), cols=len(rows[0]))
                _apply_custom_table_style(doc, table, sc)
                for r_idx, row in enumerate(rows):
                    for c_idx, cell in enumerate(row):
                        if c_idx < len(table.columns):
                            tc = table.cell(r_idx, c_idx)
                            tc.text = ""
                            para = tc.paragraphs[0]
                            # Tight paragraph spacing inside cells so the row
                            # height is controlled by cell margins, not the
                            # body paragraph's space_before/after.
                            _apply_paragraph_format(para, space_before_pt=0,
                                                    space_after_pt=0)
                            add_formatted_text(para, cell, anchors, base_dir=base_dir,
                                   allow_remote_images=allow_remote_images)
                # Emphasize the header row, then size columns to content.
                _apply_table_header(table, sc)
                _fix_narrow_column_widths(table)
                # Add a small spacer paragraph after the table
                spacer = doc.add_paragraph()
                spacer.paragraph_format.space_before = Pt(8)
                spacer.paragraph_format.space_after = Pt(0)
                spacer_run = spacer.add_run()
                spacer_run.font.size = Pt(2)

        elif block["type"] == "note":
            p = doc.add_paragraph()
            _apply_paragraph_format(p, left_indent_in=sc.num("note", "indent_in", 0.3))
            run = p.add_run(block["text"])
            _apply_run_format(run,
                              size_pt=sc.num("note", "size_pt", 10),
                              italic=sc.flag("note", "italic", True))

        elif block["type"] == "paragraph":
            p = doc.add_paragraph()
            _apply_paragraph_format(
                p,
                space_before_pt=sc.num("body", "space_before_pt", 0),
                space_after_pt=sc.num("body", "space_after_pt", 8),
                line_spacing=sc.num("body", "line_spacing", 1.15),
            )
            add_formatted_text(p, block["text"], anchors, base_dir=base_dir,
                                   allow_remote_images=allow_remote_images)

        elif block["type"] == "list":
            # Create a new numId for this bullet list block
            list_num_id = _new_num_id(doc, bullet_abstract)
            bullet_after = _coerce_number(
                sc.section("lists").get("bullet", {}).get("space_after_pt"),
                2, "lists.bullet.space_after_pt")
            for item in block["items"]:
                # Items are dicts: {"text": str, "checked": Optional[bool]}.
                # checked is None for a normal bullet, True/False for a task item.
                item_text = item["text"] if isinstance(item, dict) else item
                checked = item.get("checked") if isinstance(item, dict) else None
                p = doc.add_paragraph()
                _apply_paragraph_format(p, space_after_pt=bullet_after)
                _set_list_numbering(p, list_num_id, ilvl=0)
                # Task items get a checkbox glyph prefix (☑ checked / ☐ unchecked).
                # Give it the body font so Word doesn't substitute another font
                # (e.g. Cambria) for the Unicode checkbox symbol.
                if checked is not None:
                    cb_run = p.add_run("\u2611 " if checked else "\u2610 ")
                    _apply_run_format(cb_run, font=sc.text("body", "font"))
                add_formatted_text(p, item_text, anchors, base_dir=base_dir,
                                   allow_remote_images=allow_remote_images)

        elif block["type"] == "numbered_list":
            # Create a new numId for this numbered list block (ensures restart)
            list_num_id = _new_num_id(doc, numbered_abstract)
            ordered_after = _coerce_number(
                sc.section("lists").get("ordered", {}).get("space_after_pt"),
                2, "lists.ordered.space_after_pt")
            bullet_after = _coerce_number(
                sc.section("lists").get("bullet", {}).get("space_after_pt"),
                2, "lists.bullet.space_after_pt")
            for item in block["items"]:
                item_text = item["text"] if isinstance(item, dict) else item
                sub_items = item.get("sub_items", []) if isinstance(item, dict) else []
                p = doc.add_paragraph()
                _apply_paragraph_format(p, space_after_pt=ordered_after)
                _set_list_numbering(p, list_num_id, ilvl=0)
                add_formatted_text(p, item_text, anchors, base_dir=base_dir,
                                   allow_remote_images=allow_remote_images)
                # Render sub-items as indented bullets (level 1 of the bullet list).
                if sub_items:
                    sub_num_id = _new_num_id(doc, bullet2_abstract)
                    for sub in sub_items:
                        sp = doc.add_paragraph()
                        _apply_paragraph_format(sp, space_after_pt=bullet_after)
                        _set_list_numbering(sp, sub_num_id, ilvl=1)
                        add_formatted_text(sp, sub, anchors, base_dir=base_dir,
                                   allow_remote_images=allow_remote_images)

    doc.save(output_path)
    print(f"Saved: {output_path}")


def check_output_writable(output_path: str):
    """Fail fast if the output .docx is open in Word (or otherwise locked).

    When Word has a document open it creates a hidden owner file named
    ``~$<filename>`` alongside it and holds a write lock on the file itself.
    Saving over it would raise a PermissionError partway through the run, so we
    detect both signals up front and print an actionable message.

    Returns None on success; prints an error and exits(1) on a detected lock.
    """
    out = Path(output_path)

    # 1. Word owner/lock file: ~$<filename> in the same directory.
    lock_file = out.parent / f"~${out.name}"
    if lock_file.exists():
        print(f"ERROR: '{out.name}' appears to be open in Word "
              f"(found lock file '{lock_file.name}').")
        print("       Close the document in Word and run the conversion again.")
        sys.exit(1)

    # 2. If the file exists, confirm we can actually open it for writing.
    #    On Windows an open Word doc is locked and this raises PermissionError.
    if out.exists():
        try:
            with open(out, "a"):
                pass
        except PermissionError:
            print(f"ERROR: cannot write to '{out.name}' — the file is locked, "
                  f"likely open in Word or another program.")
            print("       Close the document and run the conversion again.")
            sys.exit(1)


def main():
    # Separate flags from positional args so ordering is flexible.
    allow_remote_images = False
    style_path = os.environ.get("MD_TO_DOCX_STYLE")
    positional = []
    args = sys.argv[1:]
    i = 0
    while i < len(args):
        arg = args[i]
        if arg in ("--fetch-remote-images", "--fetch-remote"):
            allow_remote_images = True
        elif arg == "--dump-config":
            print(dump_default_style(), end="")
            sys.exit(0)
        elif arg == "--style":
            if i + 1 >= len(args):
                print("ERROR: --style requires a path argument")
                sys.exit(1)
            style_path = args[i + 1]
            i += 1
        elif arg.startswith("--style="):
            style_path = arg.split("=", 1)[1]
        else:
            positional.append(arg)
        i += 1

    if len(positional) < 2:
        print("Usage: uv run md_to_docx.py [--fetch-remote-images] "
              "[--style config.yaml] [--dump-config] <input.md> <output.docx>")
        sys.exit(1)

    input_path = positional[0]
    output_path = positional[1]

    if not Path(input_path).exists():
        print(f"ERROR: {input_path} not found")
        sys.exit(1)

    try:
        style = load_style(style_path)
    except StyleError as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)

    # Preflight: bail early with a clear message if the output is locked/open.
    check_output_writable(output_path)

    text = Path(input_path).read_text(encoding="utf-8")
    blocks = parse_markdown(text)

    build_docx(
        blocks,
        output_path,
        title="Onboarding Journey — Email Detail",
        author="Phil Batey",
        date="July 2, 2026",
        # Resolve relative image paths against the markdown file's directory.
        base_dir=Path(input_path).resolve().parent,
        allow_remote_images=allow_remote_images,
        style=style,
    )


if __name__ == "__main__":
    main()
