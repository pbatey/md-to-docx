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

import argparse
import io
import os
import re
import sys
import urllib.request
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path

import yaml
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

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
    # Global document-default font. Written to docDefaults/rPrDefault so every
    # style and run that does not set its own font inherits it. ``null`` here is
    # not meaningful (there must be a document default); it falls back to Aptos.
    "font": "Aptos",
    # Default heading typeface. Every heading level whose own ``font`` is
    # null/absent inherits this, so the heading face is set in one place. Set a
    # per-level ``font`` to override just that level.
    "heading_font": "Aptos Display",
    "page": {
        # Letter by default. Either a named size or explicit width/height inches.
        "size": "letter",              # "letter" | "a4"
        "margins_in": {"top": 1.0, "bottom": 1.0, "left": 1.25, "right": 1.25},
    },
    "body": {
        # null => inherit the global default font (see top-level ``font``).
        "font": None,
        "size_pt": 11,
        "color": "000000",
        "space_before_pt": 0,
        "space_after_pt": 8,
        "line_spacing": 1.15,
    },
    # Per-level heading formatting (levels 1-6). Colors mirror the old template.
    # ``font: null`` (the default) means "inherit the top-level heading_font";
    # set a level's ``font`` to override just that level.
    #
    # ``rule`` draws a horizontal line under the heading. It's a mapping of
    # ``width_pt`` (line thickness), ``color`` (hex), and ``space_pt`` (gap
    # between the text and the line). ``rule: null`` (the default for H3-H6)
    # means no line. H1 and H2 get a rule by default.
    "headings": {
        1: {"font": None, "size_pt": 20, "color": "365F91", "bold": True,
            "italic": False, "space_before_pt": 18, "space_after_pt": 4,
            "rule": {"width_pt": 1.0, "color": "365F91", "space_pt": 2}},
        2: {"font": None, "size_pt": 16, "color": "4F81BD", "bold": True,
            "italic": False, "space_before_pt": 12, "space_after_pt": 4,
            "rule": {"width_pt": 0.75, "color": "4F81BD", "space_pt": 2}},
        3: {"font": None, "size_pt": 13, "color": "4F81BD", "bold": True,
            "italic": False, "space_before_pt": 10, "space_after_pt": 2,
            "rule": None},
        4: {"font": None, "size_pt": 12, "color": "4F81BD", "bold": True,
            "italic": False, "space_before_pt": 10, "space_after_pt": 2,
            "rule": None},
        5: {"font": None, "size_pt": 11, "color": "243F60", "bold": True,
            "italic": False, "space_before_pt": 8, "space_after_pt": 2,
            "rule": None},
        6: {"font": None, "size_pt": 11, "color": "243F60", "bold": False,
            "italic": True, "space_before_pt": 8, "space_after_pt": 2,
            "rule": None},
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
        "indent_in": 0.25,             # left inset so the box doesn't run full-width
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
        # Horizontal-rules-only look: a rule under each body row plus the header
        # underline (see "header" below). No outer top/bottom edges, no vertical
        # lines or side borders. "edges" lists which borders to draw; omit an
        # edge to leave it off. Valid edges: top, bottom, left, right, insideH,
        # insideV. Add "top"/"bottom" back to box the table.
        "border": {
            "style": "single",
            "width_pt": 0.75,
            "color": "808080",
            "edges": ["insideH"],
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
        "indent_in": 0.25,             # left inset so the table doesn't run full-width
    },
    "hr": {
        # A markdown ``---`` renders as a horizontal line when ``rule`` is true.
        # ``width_pt``/``color`` control the line; ``space_after_pt`` is the gap
        # below it. Set ``rule: false`` to render just blank space instead.
        "space_after_pt": 6,
        "rule": True,
        "width_pt": 0.75,
        "color": "BFBFBF",
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

    def top(self, key: str, default=None):
        """Return a top-level scalar config value (e.g. ``font``), or default."""
        val = self._cfg.get(key, default)
        return val if val is not None else default

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
        elif line.startswith(("**From:**", "**To:**", "**Subject:**",
                              "**CC:**", "**Body:**", "**Body")):
            meta_lines = []
            while i < len(lines) and (lines[i].startswith("**From:") or lines[i].startswith("**To:") or \
                  lines[i].startswith("**Subject:") or lines[i].startswith("**CC:") or \
                  lines[i].startswith("**Body")):
                meta_lines.append(lines[i].rstrip())
                i += 1
            blocks.append({"type": "meta", "lines": meta_lines})
            continue

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
                            # A soft-wrapped line joins with a space (so inline
                            # emphasis spanning the wrap still pairs up); only a
                            # trailing double-space forces a hard line break.
                            if item_text.endswith("  "):
                                item_text = item_text.rstrip() + "\n" + next_line.strip()
                            else:
                                item_text = item_text.rstrip("\n") + " " + next_line.strip()
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
                            # Soft wrap joins with a space so inline emphasis
                            # spanning the wrap still pairs; a trailing
                            # double-space forces a hard line break.
                            if item_text.endswith("  "):
                                item_text = item_text.rstrip() + "\n" + next_line.strip()
                            else:
                                item_text = item_text.rstrip("\n") + " " + next_line.strip()
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
            # Guard against no progress: if the paragraph collector consumed no
            # lines (e.g. the current line starts with "---" but isn't exactly a
            # horizontal rule, so it matches no block and the inner while stops
            # immediately), emit it as a one-line paragraph and advance so we
            # don't spin forever on the same line.
            if not para_lines:
                blocks.append({"type": "paragraph", "text": lines[i].strip()})
                i += 1
                continue
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
    key = key.removeprefix("#")
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
    """Build a ``w:r`` for a hyperlink, referencing the ``Hyperlink`` char style.

    The generated ``Hyperlink`` character style carries the configured color and
    underline, so the link is styled consistently and editable in Word.
    """
    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), STYLE_HYPERLINK)
    rPr.append(rStyle)
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


def _set_default_font(doc, font: str):
    """Set the document-default font on ``docDefaults/rPrDefault``.

    This is the lowest rung of Word's font-resolution ladder: any run, character
    style, or paragraph style that does not name its own font inherits this. We
    write ``w:rFonts`` (ascii/hAnsi/cs) so Latin, high-ANSI, and complex-script
    text all use it. A non-string/empty ``font`` is ignored (Word then falls back
    to its own default).
    """
    if not font or not isinstance(font, str):
        return
    styles_el = doc.styles.element  # <w:styles>
    docDefaults = styles_el.find(qn('w:docDefaults'))
    if docDefaults is None:
        docDefaults = OxmlElement('w:docDefaults')
        # docDefaults must be the first child of <w:styles>.
        styles_el.insert(0, docDefaults)
    rPrDefault = docDefaults.find(qn('w:rPrDefault'))
    if rPrDefault is None:
        rPrDefault = OxmlElement('w:rPrDefault')
        docDefaults.append(rPrDefault)
    rPr = rPrDefault.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        rPrDefault.append(rPr)
    existing = rPr.find(qn('w:rFonts'))
    if existing is not None:
        rPr.remove(existing)
    rFonts = OxmlElement('w:rFonts')
    for attr in ('ascii', 'hAnsi', 'cs'):
        rFonts.set(qn(f'w:{attr}'), font)
    # rFonts must lead the run properties.
    rPr.insert(0, rFonts)


def _get_or_add_style(doc, name: str, style_type):
    """Return the named style, creating it if the document doesn't have it.

    ``Normal`` (and the built-in ``Heading N``) already exist in a blank doc, so
    we fetch and update them in place; custom ids are added fresh.
    """
    try:
        return doc.styles[name]
    except KeyError:
        return doc.styles.add_style(name, style_type)


def _style_outline_level(style, level: int):
    """Set ``w:outlineLvl`` (0-based) on a paragraph style's ``pPr``.

    Putting the outline level on the style (not each paragraph) is what makes
    Word's navigation pane and PDF bookmarks treat the style as a heading.
    """
    pPr = style.element.get_or_add_pPr()
    existing = pPr.find(qn('w:outlineLvl'))
    if existing is not None:
        pPr.remove(existing)
    pPr.append(_make_element('w:outlineLvl', val=str(level)))


def _style_bottom_rule(style, *, width_pt=1.0, color="000000", space_pt=2):
    """Add a bottom border (horizontal rule) to a paragraph style's ``pPr``.

    Used for the under-heading rules on H1/H2. ``width_pt`` is the thickness
    (converted to eighths of a point for ``w:sz``), ``color`` the hex line color,
    and ``space_pt`` the gap between the text and the line. Replaces any existing
    ``w:pBdr`` so the call is idempotent, and inserts it in schema order (before
    ``w:spacing``/``w:ind``).
    """
    pPr = style.element.get_or_add_pPr()
    existing = pPr.find(qn('w:pBdr'))
    if existing is not None:
        pPr.remove(existing)
    pBdr = OxmlElement('w:pBdr')
    sz = str(max(1, int(round(width_pt * 8))))   # points -> eighths of a point
    bottom = _make_element('w:bottom', val='single', sz=sz,
                           space=str(int(round(space_pt))),
                           color=_as_hex(color) or "000000")
    pBdr.append(bottom)
    # w:pBdr must precede w:shd/w:spacing/w:ind in pPr (schema order).
    following = None
    for tag in ('w:shd', 'w:spacing', 'w:ind'):
        following = pPr.find(qn(tag))
        if following is not None:
            break
    if following is not None:
        following.addprevious(pBdr)
    else:
        pPr.append(pBdr)


# Theme-font attributes on a style's rFonts. python-docx's default template
# ships built-in styles (Heading N, Normal) whose rFonts reference these theme
# fonts (e.g. asciiTheme="majorHAnsi"). A theme reference WINS over an explicit
# w:ascii/w:hAnsi on the same element, so setting ``style.font.name`` alone is
# silently ignored by Word — it keeps using the theme font (Calibri). We strip
# these theme attributes whenever we set an explicit font on a style.
_RFONTS_THEME_ATTRS = ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme")


def _set_style_font(style, font: str):
    """Set an explicit font on a style and remove any theme-font references.

    ``style.font.name = font`` writes ``w:ascii``/``w:hAnsi``, but leaves any
    ``*Theme`` attributes in place, which override the explicit face. Removing
    the theme attributes makes the explicit font actually take effect in Word.
    """
    style.font.name = font
    rPr = style.element.find(qn('w:rPr'))
    if rPr is None:
        return
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        return
    for attr in _RFONTS_THEME_ATTRS:
        theme_qn = qn(f'w:{attr}')
        if rFonts.get(theme_qn) is not None:
            del rFonts.attrib[theme_qn]


# Custom style ids for our block/char styles. Heading N reuse Word's built-in
# names so Word treats them as real headings (nav pane, TOC fields).
STYLE_QUOTE = "MdQuote"
STYLE_CODE = "MdCode"
STYLE_CODE_TITLE = "MdCodeTitle"
STYLE_CODE_CHAR = "MdCodeChar"
STYLE_HYPERLINK = "Hyperlink"

# Human-friendly display names shown in Word's style gallery. The styleId (the
# constants above) stays stable so paragraph references don't break; only the
# label changes. Avoid colliding with Word's built-in names (e.g. "Quote",
# "Code") so our styles show up as distinct gallery entries.
STYLE_DISPLAY_NAMES = {
    STYLE_QUOTE: "Blockquote",
    STYLE_CODE: "Code Block",
    STYLE_CODE_TITLE: "Code Block Title",
    STYLE_CODE_CHAR: "Inline Code",
}


def _hide_style_from_gallery(style):
    """Hide a style from Word's gallery / recommended view (keep the definition).

    Clears ``w:qFormat`` (so it drops out of the quick-styles gallery and the
    "recommended" filter) and sets ``w:semiHidden`` + ``w:unhideWhenUsed`` (so
    the Styles pane's default view also omits it until it's actually used). The
    style itself is left intact so any content already using it still renders.
    """
    pr = style.element
    qf = pr.find(qn('w:qFormat'))
    if qf is not None:
        pr.remove(qf)
    # semiHidden + unhideWhenUsed: Word hides these from the default pane view.
    for tag in ('w:semiHidden', 'w:unhideWhenUsed'):
        if pr.find(qn(tag)) is None:
            pr.append(_make_element(tag))


# styleIds we keep visible in Word's Styles gallery. Everything else that the
# base template flags as a quick style gets hidden so the gallery isn't cluttered
# with styles md_to_docx never applies. Hyperlink is intentionally omitted (Word
# applies it automatically and doesn't surface it in the gallery anyway).
_GALLERY_KEEP_IDS = frozenset({
    "Normal",
    "Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6",
    STYLE_QUOTE, STYLE_CODE, STYLE_CODE_TITLE, STYLE_CODE_CHAR,
})


def _prune_gallery_styles(doc):
    """Remove gallery visibility from every style we don't apply.

    Sweeps all styles; any that the template flagged as a quick style but that
    isn't in :data:`_GALLERY_KEEP_IDS` is hidden via
    :func:`_hide_style_from_gallery`. Leaves our styles (and the essential
    built-ins) as the only entries in Word's gallery.
    """
    for style in doc.styles:
        try:
            style_id = style.style_id
            is_quick = style.quick_style
        except Exception:
            continue
        if is_quick and style_id not in _GALLERY_KEEP_IDS:
            _hide_style_from_gallery(style)


def _mark_style_visible(style, priority=None):
    """Make a style show up in Word's Styles gallery / pane.

    Word hides styles that aren't flagged as "recommended". Setting
    ``w:qFormat`` (via ``quick_style``) surfaces the style in the Home-tab
    gallery and the Styles pane's default view; ``w:uiPriority`` (via
    ``priority``) controls its sort order. Also clears ``semiHidden`` and
    ``unhideWhenUsed`` if present so nothing suppresses it.
    """
    style.quick_style = True
    if priority is not None:
        style.priority = priority
    # Clear any hide flags Word might otherwise honor.
    pr = style.element
    for tag in ('w:semiHidden', 'w:unhideWhenUsed'):
        el = pr.find(qn(tag))
        if el is not None:
            pr.remove(el)
    return style


def _apply_code_box_to_pPr(pPr, fill, space_pt):
    """Add the code-block shaded box (four-sided border + shading) to a ``pPr``.

    The border color matches the fill so it's invisible; its ``w:space`` acts as
    interior padding. Inserts ``w:pBdr`` then ``w:shd`` in schema order (both
    precede ``w:spacing``/``w:ind``). Idempotent: removes any existing pBdr/shd.
    Shared by the ``Code Block`` style generator and the direct-formatting helper.
    """
    hexfill = _as_hex(fill) or "F2F2F2"
    for tag in ('w:pBdr', 'w:shd'):
        existing = pPr.find(qn(tag))
        if existing is not None:
            pPr.remove(existing)

    pBdr = OxmlElement('w:pBdr')
    for side in ('top', 'left', 'bottom', 'right'):
        pBdr.append(_make_element(f'w:{side}', val='single', sz='4',
                                  space=str(int(round(space_pt))), color=hexfill))
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexfill)

    # pBdr and shd must precede spacing/ind in pPr (schema order).
    following = None
    for tag in ('w:spacing', 'w:ind'):
        following = pPr.find(qn(tag))
        if following is not None:
            break
    if following is not None:
        following.addprevious(pBdr)
        following.addprevious(shd)
    else:
        pPr.append(pBdr)
        pPr.append(shd)


def _define_paragraph_style(doc, name, *, display_name=None, font=None,
                            size_pt=None, color=None,
                            bold=None, italic=None, space_before_pt=None,
                            space_after_pt=None, line_spacing=None,
                            left_indent_in=None, right_indent_in=None,
                            outline_level=None, base="Normal"):
    """Create/update a named paragraph style from config values.

    ``name`` is the styleId (stable, referenced by paragraphs); ``display_name``
    is the human-friendly label Word shows in its style gallery (defaults to
    ``name``). ``font=None`` leaves the style's font unset so it inherits (from
    the base style or the document default) — this is how "null font means
    inherit" works. Every other ``None`` argument is likewise skipped.
    """
    style = _get_or_add_style(doc, name, WD_STYLE_TYPE.PARAGRAPH)
    if display_name is not None:
        style.name = display_name
    if base is not None and name != "Normal":
        # ``base`` may be a styleId (e.g. "MdCode"); prefer its display name so
        # the lookup doesn't hit python-docx's deprecated by-id path.
        base_key = STYLE_DISPLAY_NAMES.get(base, base)
        try:
            style.base_style = doc.styles[base_key]
        except (KeyError, Exception):
            pass
    fmt = style.font
    if font is not None:
        _set_style_font(style, font)
    if size_pt is not None:
        fmt.size = Pt(size_pt)
    if color is not None:
        fmt.color.rgb = color if isinstance(color, RGBColor) \
            else RGBColor.from_string(str(color).upper())
    if bold is not None:
        fmt.bold = bold
    if italic is not None:
        fmt.italic = italic
    pf = style.paragraph_format
    if space_before_pt is not None:
        pf.space_before = Pt(space_before_pt)
    if space_after_pt is not None:
        pf.space_after = Pt(space_after_pt)
    if line_spacing is not None:
        pf.line_spacing = line_spacing
    if left_indent_in is not None:
        pf.left_indent = Inches(left_indent_in)
    if right_indent_in is not None:
        pf.right_indent = Inches(right_indent_in)
    if outline_level is not None:
        _style_outline_level(style, outline_level)
    return style


def _define_char_style(doc, name, *, display_name=None, font=None, size_pt=None,
                       color=None, underline=None, fill=None):
    """Create/update a named character style from config values.

    ``name`` is the styleId; ``display_name`` is the gallery label (defaults to
    ``name``). Like the paragraph variant, ``None`` fields are left unset so they
    inherit. ``fill`` adds run-shading (``w:shd``) onto the style's run
    properties.
    """
    style = _get_or_add_style(doc, name, WD_STYLE_TYPE.CHARACTER)
    if display_name is not None:
        style.name = display_name
    fmt = style.font
    if font is not None:
        _set_style_font(style, font)
    if size_pt is not None:
        fmt.size = Pt(size_pt)
    if color is not None:
        fmt.color.rgb = color if isinstance(color, RGBColor) \
            else RGBColor.from_string(str(color).upper())
    if underline is not None:
        fmt.underline = underline
    if fill is not None:
        hexfill = _as_hex(fill)
        if hexfill:
            rPr = style.element.get_or_add_rPr()
            existing = rPr.find(qn('w:shd'))
            if existing is not None:
                rPr.remove(existing)
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), hexfill)
            rPr.append(shd)
    return style


def _build_styles(doc, sc: "StyleConfig"):
    """Generate the document default font + all named styles from config.

    Runs once at the start of ``build_docx`` before any content is added, so
    every paragraph/run can reference a style by name.
    """
    # 1. Document default font: everything inherits this unless overridden.
    _set_default_font(doc, sc.top("font", "Aptos"))

    # 2. Normal (body). font=None => inherit the document default.
    _define_paragraph_style(
        doc, "Normal",
        font=sc.text("body", "font"),
        size_pt=sc.num("body", "size_pt", 11),
        color=sc.color("body", "color", default="000000"),
        space_before_pt=sc.num("body", "space_before_pt", 0),
        space_after_pt=sc.num("body", "space_after_pt", 8),
        line_spacing=sc.num("body", "line_spacing", 1.15),
        base=None,
    )

    # 3. Heading 1..6. Own font wins; otherwise inherit the top-level
    #    heading_font (so the heading face is set in one place).
    heading_font = sc.top("heading_font", "Aptos Display")
    for level in range(1, 7):
        h = sc.heading(level)
        own_font = h.get("font")
        font = own_font if own_font is not None else heading_font
        heading_style = _define_paragraph_style(
            doc, f"Heading {level}",
            font=font,
            size_pt=_coerce_number(h.get("size_pt"), 11, f"headings.{level}.size_pt"),
            color=_coerce_color(h.get("color"), "000000", f"headings.{level}.color"),
            bold=bool(h.get("bold", False)),
            italic=bool(h.get("italic", False)),
            space_before_pt=_coerce_number(h.get("space_before_pt"), 0,
                                           f"headings.{level}.space_before_pt"),
            space_after_pt=_coerce_number(h.get("space_after_pt"), 0,
                                          f"headings.{level}.space_after_pt"),
            outline_level=level - 1,
        )
        # Optional under-heading rule (bottom border on the style).
        rule = h.get("rule")
        if isinstance(rule, dict):
            _style_bottom_rule(
                heading_style,
                width_pt=_coerce_number(rule.get("width_pt"), 1.0,
                                        f"headings.{level}.rule.width_pt"),
                color=_as_hex(rule.get("color")) or "000000",
                space_pt=_coerce_number(rule.get("space_pt"), 2,
                                        f"headings.{level}.rule.space_pt"),
            )

    # 4. Blockquote paragraph style (indent + spacing + the left bar). Folding
    #    the bar into the style's pPr — rather than stamping it on each
    #    paragraph — makes the whole blockquote look editable in Word by editing
    #    the one style.
    quote_style = _define_paragraph_style(
        doc, STYLE_QUOTE,
        display_name=STYLE_DISPLAY_NAMES[STYLE_QUOTE],
        space_after_pt=sc.num("blockquote", "space_after_pt", 4),
        left_indent_in=sc.num("blockquote", "indent_in", 0.25),
    )
    quote_pPr = quote_style.element.get_or_add_pPr()
    existing_bdr = quote_pPr.find(qn('w:pBdr'))
    if existing_bdr is not None:
        quote_pPr.remove(existing_bdr)
    quote_pBdr = _build_blockquote_pBdr(
        color=sc.text("blockquote", "bar_color", "999999"),
        width_pt=sc.num("blockquote", "bar_width_pt", 2.25),
        gap_pt=sc.num("blockquote", "bar_gap_pt", 12),
    )
    # w:pBdr must precede w:shd/w:spacing/w:ind in pPr (schema order); insert it
    # before whichever of those exists, else append.
    following = None
    for tag in ('w:shd', 'w:spacing', 'w:ind'):
        following = quote_pPr.find(qn(tag))
        if following is not None:
            break
    if following is not None:
        following.addprevious(quote_pBdr)
    else:
        quote_pPr.append(quote_pBdr)
    # Surface it in Word's Styles gallery / pane (priority just after headings).
    _mark_style_visible(quote_style, priority=20)

    # 5. Code-block paragraph style. The shaded box (four-sided border + fill)
    #    lives IN the style's pPr so it's editable in Word.
    #    The box border carries a w:space (from padding_pt) that offsets the
    #    border OUTWARD from the text. To keep the box's left edge flush with
    #    body text, indent the paragraph left by that same padding; the right
    #    inset is padding + the requested right inset so the box stops short of
    #    the margin.
    cb_pad_pt = sc.num("code_block", "padding_pt", 6)
    cb_pad_in = cb_pad_pt / 72.0
    cb_right_inset = sc.num("code_block", "indent_in", 0.25)
    cb_fill = sc.text("code_block", "fill", "F2F2F2")
    code_style = _define_paragraph_style(
        doc, STYLE_CODE,
        display_name=STYLE_DISPLAY_NAMES[STYLE_CODE],
        font=sc.text("code_block", "font", "Consolas"),
        size_pt=sc.num("code_block", "size_pt", 9),
        space_before_pt=sc.num("code_block", "space_before_pt", 8),
        space_after_pt=sc.num("code_block", "space_after_pt", 8),
        left_indent_in=cb_pad_in,
        right_indent_in=cb_right_inset + cb_pad_in,
    )
    _apply_code_box_to_pPr(code_style.element.get_or_add_pPr(), cb_fill, cb_pad_pt)
    _mark_style_visible(code_style, priority=21)

    # 5b. Code-block title/caption style (the language label). Based on Code
    #     Block so it inherits the box/indent/shading; overrides just the run
    #     look (small italic gray). ``next`` returns to Code Block after it.
    caption_cfg = sc.section("code_block").get("caption", {}) or {}
    title_style = _define_paragraph_style(
        doc, STYLE_CODE_TITLE,
        display_name=STYLE_DISPLAY_NAMES[STYLE_CODE_TITLE],
        base=STYLE_CODE,
        size_pt=_coerce_number(caption_cfg.get("size_pt"), 8,
                               "code_block.caption.size_pt"),
        color=_coerce_color(caption_cfg.get("color"), "808080",
                            "code_block.caption.color"),
        italic=bool(caption_cfg.get("italic", True)),
    )
    # "next style" = Code Block, so Enter after a title continues code.
    # w:next must sit after w:basedOn (schema order), before pPr/rPr.
    tel = title_style.element
    existing_next = tel.find(qn('w:next'))
    if existing_next is not None:
        tel.remove(existing_next)
    tnext = OxmlElement('w:next')
    tnext.set(qn('w:val'), STYLE_CODE)
    based = tel.find(qn('w:basedOn'))
    if based is not None:
        based.addnext(tnext)
    else:
        name_el = tel.find(qn('w:name'))
        (name_el if name_el is not None else tel).addnext(tnext) \
            if name_el is not None else tel.insert(0, tnext)
    _mark_style_visible(title_style, priority=21)

    # 6. Inline-code character style (monospace font/size, optional color/fill).
    code_char_style = _define_char_style(
        doc, STYLE_CODE_CHAR,
        display_name=STYLE_DISPLAY_NAMES[STYLE_CODE_CHAR],
        font=sc.text("inline_code", "font", "Consolas"),
        size_pt=sc.num("inline_code", "size_pt", 10),
        color=sc.color("inline_code", "color"),
        fill=sc.text("inline_code", "fill"),
    )
    _mark_style_visible(code_char_style, priority=22)

    # 8. Hyperlink character style (color + underline from links.*).
    _define_char_style(
        doc, STYLE_HYPERLINK,
        color=sc.color("links", "color", default="0563C1"),
        underline=sc.flag("links", "underline", True),
    )

    # 9. Declutter Word's gallery: hide the many template quick styles we never
    #    apply (Title, Subtitle, Quote, Intense Quote, List Paragraph, etc.),
    #    leaving only our styles and the essential headings/Normal.
    _prune_gallery_styles(doc)


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
        paragraph.add_run(alt or "[image]")

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
        alias = m.group(2) or target
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

    Normal runs inherit their font/size/color from the paragraph style (or the
    document default) — we do NOT stamp those properties directly. Only emphasis
    (bold/italic/strike) and the inline-code character style are set here.
    """
    if "bold" in flags:
        run.bold = True
    if "italic" in flags:
        run.italic = True
    if "strike" in flags:
        run.font.strike = True
    if "code" in flags:
        # Apply the inline-code character style so inline code is monospace.
        try:
            run.style = run.part.document.styles[STYLE_DISPLAY_NAMES[STYLE_CODE_CHAR]]
        except (KeyError, Exception):
            # Fallback: stamp the code font directly if the style is missing.
            sc = _ACTIVE_STYLE
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


def _build_blockquote_pBdr(color=BLOCKQUOTE_BAR_COLOR, width_pt=2.25, gap_pt=12):
    """Build a ``w:pBdr`` with a single left bar for a blockquote.

    ``width_pt`` is the bar thickness in points (converted to the eighths-of-a-
    point ``sz`` unit); ``gap_pt`` is the space between bar and text in points.
    Shared by the ``MdQuote`` style generator and the direct-formatting helper.
    """
    pBdr = OxmlElement('w:pBdr')
    sz = str(max(1, int(round(width_pt * 8))))   # points -> eighths of a point
    left = _make_element('w:left', val='single', sz=sz,
                         space=str(int(round(gap_pt))),
                         color=_as_hex(color) or "999999")
    pBdr.append(left)
    return pBdr


def _apply_blockquote_bar(paragraph, color=BLOCKQUOTE_BAR_COLOR,
                          width_pt=2.25, gap_pt=12):
    """Add a left vertical bar (paragraph border) to a blockquote paragraph.

    Applied as a direct border. Kept for callers that style a paragraph without
    the ``MdQuote`` style; normal blockquotes now get the bar from the style.
    """
    pPr = paragraph._p.get_or_add_pPr()
    # Remove any existing borders so repeated calls stay idempotent.
    existing = pPr.find(qn('w:pBdr'))
    if existing is not None:
        pPr.remove(existing)
    pPr.append(_build_blockquote_pBdr(color=color, width_pt=width_pt, gap_pt=gap_pt))


# Code-block styling: a light gray fill plus a same-color border on all four
# sides. The border's w:space attribute (in points) is what gives the code text
# breathing room inside the shaded box — Word has no separate padding property
# for shading, so a matching-color border is the standard way to fake it.
CODE_BLOCK_FILL = 'F2F2F2'        # light gray background
CODE_BLOCK_BORDER_SPACE = '6'     # padding between border and text, in points


def _apply_code_block_box(paragraph, fill=CODE_BLOCK_FILL,
                          space=CODE_BLOCK_BORDER_SPACE):
    """Give a code paragraph a shaded background with interior padding (direct).

    Kept as a direct-formatting utility; the normal render path now carries the
    box on the ``Code Block`` style instead. Delegates to the shared builder.
    """
    _apply_code_box_to_pPr(paragraph._p.get_or_add_pPr(), fill,
                           _coerce_number(space, 6, "code_block.padding_pt"))


def _split_table_row(line: str):
    """Split a markdown table row into cell strings.

    Splits on unescaped ``|`` only, so ``\\|`` is treated as a literal pipe
    inside a cell (GFM). The escaped pipe is unescaped to ``|`` in the result,
    and ``<br>`` is converted to a newline so it renders as an in-cell line
    break. Leading/trailing empty cells (from the surrounding ``|``) are dropped.
    """
    # Split on a pipe not preceded by a backslash.
    parts = re.split(r"(?<!\\)\|", line)
    # Drop the empty first/last segments produced by the leading/trailing pipe.
    if parts and parts[0].strip() == "":
        parts = parts[1:]
    if parts and parts[-1].strip() == "":
        parts = parts[:-1]
    cells = []
    for c in parts:
        c = c.replace("\\|", "|")            # literal pipe
        c = re.sub(r"<br\s*/?>", "\n", c, flags=re.IGNORECASE)  # in-cell break
        cells.append(c.strip())
    return cells


def _parse_table_alignments(separator_line: str):
    """Parse a markdown table separator row into per-column alignments.

    Returns a list of ``WD_ALIGN_PARAGRAPH`` values (or ``None`` for the default
    left) — one per column. Alignment comes from the colons in each cell:
    ``:---`` left, ``:---:`` center, ``---:`` right, ``---`` unspecified (None).
    """
    aligns = []
    cells = separator_line.strip().split("|")[1:-1]
    for cell in cells:
        c = cell.strip()
        left = c.startswith(":")
        right = c.endswith(":")
        if left and right:
            aligns.append(WD_ALIGN_PARAGRAPH.CENTER)
        elif right:
            aligns.append(WD_ALIGN_PARAGRAPH.RIGHT)
        elif left:
            aligns.append(WD_ALIGN_PARAGRAPH.LEFT)
        else:
            aligns.append(None)  # unspecified -> default (left)
    return aligns


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
        enabled_edges = ["insideH"]
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

    # Right-only inset: the table is narrowed on the right so it doesn't run the
    # full content width. For the LEFT edge, Word insets cell text from the table
    # border by the cell's left margin; a small positive tblInd equal to that
    # margin lines the table border up so cell text sits near the body margin.
    indent_in = _coerce_number(sc.section("table").get("indent_in"), 0.0,
                               "table.indent_in")
    cell_left_dxa = int(_pt_to_dxa(margins.get("left"), 5.4))
    existing_ind = tblPr.find(qn('w:tblInd'))
    if existing_ind is not None:
        tblPr.remove(existing_ind)
    tblPr.append(_make_element('w:tblInd', w=str(cell_left_dxa), type='dxa'))

    # Width: full (100%) or auto.
    existing_w = tblPr.find(qn('w:tblW'))
    if existing_w is not None:
        tblPr.remove(existing_w)
    if sc.text("table", "width", "full") == "full":
        # A full-width table is narrowed by the inset (as a fraction of the
        # content width) so its right edge stops short of the right margin.
        pct = 5000
        if indent_in > 0:
            try:
                section = doc.sections[0]
                content_in = (section.page_width - section.left_margin
                              - section.right_margin) / 914400.0  # EMU -> in
                if content_in > 0:
                    pct = max(0, int(round(5000 * (1 - indent_in / content_in))))
            except Exception:
                pass
        tblPr.append(_make_element('w:tblW', w=str(pct), type='pct'))
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


def _fix_narrow_column_widths(table, available_twips=9360):
    """After cells are populated, set proportional column widths based on content.

    ``available_twips`` is the usable width the columns must sum to (content
    width minus any table indent). Defaults to 6.5in (9360 twips).
    """
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

    # Calculate proportional widths based on content length
    # Use sqrt to dampen the ratio — prevents huge disparities
    import math
    col_weights = [math.sqrt(max(length, 1)) for length in col_max_len]
    total_weight = sum(col_weights)

    col_widths_twips = [int(available_twips * w / total_weight) for w in col_weights]

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
    """Add a heading using the generated ``Heading N`` paragraph style.

    The style carries font/size/color/bold/italic/spacing and the outline level.
    No direct run formatting is stamped — this is the key difference from the old
    direct-formatting approach.
    """
    p = doc.add_paragraph(style=f"Heading {level}")
    p.add_run(text)
    return p


def build_docx(blocks: list, output_path: str, title=None, author=None, date=None,
               base_dir=None, allow_remote_images=False, style=None):
    """Build the DOCX from parsed blocks, applying named styles from ``style``.

    ``title``, ``author``, and ``date`` are optional document metadata written to
    the file's core properties (Word's Title/Author fields); each is omitted when
    left as ``None``. ``base_dir`` resolves relative image paths; it defaults to
    the current working directory when not supplied. ``allow_remote_images``
    enables fetching ``http(s)`` image sources (off by default). ``style`` is a
    merged style dict (see :func:`load_style`); when ``None`` the built-in
    defaults are used.
    """
    if base_dir is None:
        base_dir = Path.cwd()
    if style is None:
        style = load_style()
    sc = StyleConfig(style)

    # Publish the active style so inline helpers can read body/inline-code config.
    global _ACTIVE_STYLE
    _ACTIVE_STYLE = sc

    # Start from a blank document, generate named styles, then add content.
    doc = Document()
    _apply_page_setup(doc, sc)
    _build_styles(doc, sc)

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
            # Draw the horizontal line as a bottom border, width/color from config.
            if sc.flag("hr", "rule", True):
                sz = str(max(1, int(round(sc.num("hr", "width_pt", 0.75) * 8))))
                pPr = p._p.get_or_add_pPr()
                pBdr = OxmlElement('w:pBdr')
                pBdr.append(_make_element('w:bottom', val='single', sz=sz,
                                          space='1',
                                          color=_as_hex(sc.text("hr", "color", "BFBFBF")) or "BFBFBF"))
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
            # Blockquote — uses the MdQuote paragraph style, which now carries
            # both the indent/spacing AND the left bar. Split into paragraphs on
            # empty lines.
            def _emit_quote(para_lines):
                p = doc.add_paragraph(style=STYLE_DISPLAY_NAMES[STYLE_QUOTE])
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
            # The shaded box + spacing come from the Code Block paragraph style.
            # A language caption (when present) is a separate paragraph using the
            # Code Block Title style, which is based on Code Block so it shares
            # the same box (the two paragraphs read as one continuous box).
            language = block.get("language")
            if language:
                cap = doc.add_paragraph(style=STYLE_DISPLAY_NAMES[STYLE_CODE_TITLE])
                cap.add_run(language)

            p = doc.add_paragraph(style=STYLE_DISPLAY_NAMES[STYLE_CODE])
            # Add each line with line breaks between them. Font/size/box come
            # from the Code Block paragraph style.
            for line_idx, code_line in enumerate(block["lines"]):
                if line_idx > 0:
                    p.add_run().add_break()
                p.add_run(code_line)

        elif block["type"] == "table":
            # Parse table into rows, capturing per-column alignment from the
            # separator row (:--- left, :---: center, ---: right).
            rows = []
            col_align = []
            for tl in block["lines"]:
                if re.match(r"^\|[-| :]+\|$", tl.strip()):
                    col_align = _parse_table_alignments(tl)
                    continue  # skip the separator row itself
                cells = _split_table_row(tl)
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
                            # Per-column alignment from the separator row.
                            align = col_align[c_idx] if c_idx < len(col_align) else None
                            if align is not None:
                                para.alignment = align
                            add_formatted_text(para, cell, anchors, base_dir=base_dir,
                                   allow_remote_images=allow_remote_images)
                # Emphasize the header row, then size columns to content.
                _apply_table_header(table, sc)
                # Columns must sum to the content width minus the table indent.
                try:
                    section = doc.sections[0]
                    content_twips = int((section.page_width - section.left_margin
                                         - section.right_margin) / 914400.0 * 1440)
                except Exception:
                    content_twips = 9360
                table_indent_in = _coerce_number(
                    sc.section("table").get("indent_in"), 0.0, "table.indent_in")
                avail_twips = max(1440, content_twips - int(round(table_indent_in * 1440)))
                _fix_narrow_column_widths(table, available_twips=avail_twips)
                # Add a small spacer paragraph after the table
                spacer = doc.add_paragraph()
                spacer.paragraph_format.space_before = Pt(8)
                spacer.paragraph_format.space_after = Pt(0)
                spacer_run = spacer.add_run()
                spacer_run.font.size = Pt(2)

        elif block["type"] == "paragraph":
            p = doc.add_paragraph()  # uses Normal style (body font/spacing)
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
                # The font inherits from the document default via Normal style.
                if checked is not None:
                    p.add_run("\u2611 " if checked else "\u2610 ")
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

    # Write optional document metadata to the file's core properties.
    cp = doc.core_properties
    if title is not None:
        cp.title = title
    if author is not None:
        cp.author = author
    if date is not None:
        # core_properties.created requires a datetime; accept one directly, or
        # parse an ISO-8601 date/datetime string. Anything else is ignored so a
        # freeform label never breaks the save.
        created = None
        if isinstance(date, datetime):
            created = date
        elif isinstance(date, str):
            try:
                created = datetime.fromisoformat(date)
            except ValueError:
                created = None
        if created is not None:
            cp.created = created

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


def _resolve_version() -> str:
    """Best-effort package version for --version.

    Works when installed (reads distribution metadata) and degrades gracefully
    when run as a standalone PEP 723 script where no distribution is present.
    """
    from importlib.metadata import PackageNotFoundError, version

    try:
        return version("md-to-docx")
    except PackageNotFoundError:
        return "0+unknown"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="md-to-docx",
        description="Convert a Markdown file to a styled Word .docx document.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {_resolve_version()}",
    )
    parser.add_argument(
        "--fetch-remote-images",
        "--fetch-remote",
        dest="allow_remote_images",
        action="store_true",
        help="Download and embed remote http(s) images (off by default).",
    )
    parser.add_argument(
        "--style",
        dest="style_path",
        metavar="config.yaml",
        default=os.environ.get("MD_TO_DOCX_STYLE"),
        help="Path to a YAML style config (overrides MD_TO_DOCX_STYLE).",
    )
    parser.add_argument(
        "--dump-config",
        action="store_true",
        help="Print the default style config as YAML to stdout and exit.",
    )
    # Positional args are optional so --dump-config can run without them.
    parser.add_argument(
        "input",
        nargs="?",
        metavar="input.md",
        help="Path to the Markdown input file.",
    )
    parser.add_argument(
        "output",
        nargs="?",
        metavar="output.docx",
        help="Path to write the .docx output.",
    )
    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    if args.dump_config:
        # Write UTF-8 bytes directly so non-UTF-8 consoles (e.g. Windows cp1252)
        # don't choke on characters like the bullet glyph in the default style.
        sys.stdout.buffer.write(dump_default_style().encode("utf-8"))
        sys.exit(0)

    if not args.input or not args.output:
        parser.error("the following arguments are required: input.md, output.docx")

    input_path = args.input
    output_path = args.output
    allow_remote_images = args.allow_remote_images
    style_path = args.style_path

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

    # Derive document metadata from the content/environment rather than
    # hard-coding it: title from the first level-1 heading (else the filename),
    # author from the OS user, date as now.
    doc_title = next(
        (b["text"] for b in blocks if b.get("type") == "h1" and b.get("text")),
        Path(input_path).stem,
    )
    doc_author = os.environ.get("USER") or os.environ.get("USERNAME")

    build_docx(
        blocks,
        output_path,
        title=doc_title,
        author=doc_author,
        date=datetime.now(timezone.utc),
        # Resolve relative image paths against the markdown file's directory.
        base_dir=Path(input_path).resolve().parent,
        allow_remote_images=allow_remote_images,
        style=style,
    )


if __name__ == "__main__":
    main()
