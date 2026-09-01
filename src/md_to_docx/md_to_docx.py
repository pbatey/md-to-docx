#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["python-docx"]
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


def _add_internal_hyperlink(paragraph, text: str, bookmark: str):
    """Add a run that hyperlinks to an internal bookmark (w:anchor)."""
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('w:anchor'), bookmark)
    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), 'Hyperlink')
    rPr.append(rStyle)
    run.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    run.append(t)
    hyperlink.append(run)
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
    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), 'Hyperlink')
    rPr.append(rStyle)
    run.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


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
    """Apply the accumulated formatting flags to an existing run."""
    if "bold" in flags:
        run.bold = True
    if "italic" in flags:
        run.italic = True
    if "strike" in flags:
        run.font.strike = True
    if "code" in flags:
        run.font.name = "Consolas"
        run.font.size = Pt(10)


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


def _apply_blockquote_bar(paragraph):
    """Add a left vertical bar (paragraph border) to a blockquote paragraph.

    This renders the blockquote as a paragraph with a vertical bar on the left
    hand side rather than a plain indented paragraph. The border is applied
    directly so it works whether or not the template's Quote style is present.
    """
    pPr = paragraph._p.get_or_add_pPr()
    # Remove any existing borders so repeated calls stay idempotent.
    existing = pPr.find(qn('w:pBdr'))
    if existing is not None:
        pPr.remove(existing)
    pBdr = OxmlElement('w:pBdr')
    left = _make_element('w:left', val='single', sz=BLOCKQUOTE_BAR_SIZE,
                         space=BLOCKQUOTE_BAR_SPACE, color=BLOCKQUOTE_BAR_COLOR)
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


# The table style provided by templates/md-template.docx. The template ships a
# sample table that uses Word's "Plain Table 2" style (styleId "PlainTable2"),
# which is already defined in the template's styles.xml. We reference that style
# directly instead of synthesizing our own, so table styling stays in sync with
# whatever the template author configures.
TEMPLATE_TABLE_STYLE_ID = 'PlainTable2'

# Cell margins (top/bottom in dxa) matching the template's sample table.
TEMPLATE_CELL_MARGIN_TOP_BOTTOM = '72'


def _table_style_available(doc, style_id):
    """Return True if the given table styleId exists in the document's styles."""
    styles_element = doc.styles._element
    for existing in styles_element.findall(qn('w:style')):
        if existing.get(qn('w:styleId')) == style_id and existing.get(qn('w:type')) == 'table':
            return True
    return False


# Standard Word hyperlink color (the same blue Word applies by default).
HYPERLINK_COLOR = '0563C1'


def _ensure_hyperlink_style(doc):
    """Guarantee a ``Hyperlink`` character style exists so links look clickable.

    Both internal and external links reference the built-in ``Hyperlink``
    character style (``rStyle val="Hyperlink"``). Many templates — including the
    bundled one — don't define it, so Word has nothing to resolve the reference
    to and renders link text as plain black with no underline. When the style is
    missing we inject a minimal one (blue + single underline) so links render as
    links regardless of the template.
    """
    styles_element = doc.styles._element
    for existing in styles_element.findall(qn('w:style')):
        if (existing.get(qn('w:styleId')) == 'Hyperlink'
                and existing.get(qn('w:type')) == 'character'):
            return  # already defined by the template

    style = OxmlElement('w:style')
    style.set(qn('w:type'), 'character')
    style.set(qn('w:styleId'), 'Hyperlink')

    name = OxmlElement('w:name')
    name.set(qn('w:val'), 'Hyperlink')
    style.append(name)

    # Character-only style: don't offer it in the quick style gallery.
    style.append(OxmlElement('w:unhideWhenUsed'))
    style.append(OxmlElement('w:semiHidden'))

    rPr = OxmlElement('w:rPr')
    color = OxmlElement('w:color')
    color.set(qn('w:val'), HYPERLINK_COLOR)
    rPr.append(color)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    style.append(rPr)

    styles_element.append(style)


def _apply_custom_table_style(doc, table):
    """Apply the template's table style (PlainTable2) and match its table-level
    properties (tblLook, cell margins) to the sample table in md-template.docx.

    Falls back to the built-in 'Table Grid' style if the template style is
    missing (e.g. when running without the template).
    """
    if _table_style_available(doc, TEMPLATE_TABLE_STYLE_ID):
        style_id = TEMPLATE_TABLE_STYLE_ID
    elif _table_style_available(doc, 'TableGrid'):
        style_id = 'TableGrid'
    else:
        style_id = TEMPLATE_TABLE_STYLE_ID  # reference by id; Word resolves if present

    # Set the style on the table
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    # Remove existing style ref if any
    existing_style = tblPr.find(qn('w:tblStyle'))
    if existing_style is not None:
        tblPr.remove(existing_style)

    tblStyle = _make_element('w:tblStyle', val=style_id)
    tblPr.insert(0, tblStyle)

    # Set tblLook: firstRow=1, lastRow=0, firstColumn=1, lastColumn=0, noHBand=0, noVBand=1
    existing_look = tblPr.find(qn('w:tblLook'))
    if existing_look is not None:
        tblPr.remove(existing_look)
    tblLook = _make_element('w:tblLook',
                            val='04A0',
                            firstRow='1',
                            lastRow='0',
                            firstColumn='1',
                            lastColumn='0',
                            noHBand='0',
                            noVBand='1')
    tblPr.append(tblLook)

    # Cell margins matching the template's sample table (72 dxa top/bottom).
    existing_mar = tblPr.find(qn('w:tblCellMar'))
    if existing_mar is not None:
        tblPr.remove(existing_mar)
    tblCellMar = OxmlElement('w:tblCellMar')
    tblCellMar.append(_make_element('w:top', w=TEMPLATE_CELL_MARGIN_TOP_BOTTOM, type='dxa'))
    tblCellMar.append(_make_element('w:left', w='108', type='dxa'))
    tblCellMar.append(_make_element('w:bottom', w=TEMPLATE_CELL_MARGIN_TOP_BOTTOM, type='dxa'))
    tblCellMar.append(_make_element('w:right', w='108', type='dxa'))
    tblPr.append(tblCellMar)

    # Auto-fit columns to content
    # Set table width to 100% (5000 fifths of a percent)
    existing_w = tblPr.find(qn('w:tblW'))
    if existing_w is not None:
        tblPr.remove(existing_w)
    tblW = _make_element('w:tblW', w='5000', type='pct')
    tblPr.append(tblW)

    # Set layout to autofit by default (overridden to fixed for tables with narrow columns)
    existing_layout = tblPr.find(qn('w:tblLayout'))
    if existing_layout is not None:
        tblPr.remove(existing_layout)
    tblLayout = _make_element('w:tblLayout', type='autofit')
    tblPr.append(tblLayout)

    # Column widths are set by _fix_narrow_column_widths() after cells are populated.


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


def _find_abstract_num(doc, fmt='bullet', left='360'):
    """Find an abstractNumId in the template matching the given format and left indent."""
    numbering_elm = doc.part.numbering_part._element
    for abstract in numbering_elm.findall(qn('w:abstractNum')):
        lvl = abstract.find(qn('w:lvl'))
        if lvl is None:
            continue
        numFmt_el = lvl.find(qn('w:numFmt'))
        if numFmt_el is None or numFmt_el.get(qn('w:val')) != fmt:
            continue
        pPr = lvl.find(qn('w:pPr'))
        if pPr is None:
            continue
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            continue
        if ind.get(qn('w:left')) == left:
            return int(abstract.get(qn('w:abstractNumId')))
    return None


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


def find_template() -> "Path | None":
    """Locate templates/md-template.docx across standalone and installed use.

    Resolution order (first match wins):
    1. ``MD_TO_DOCX_TEMPLATE`` environment variable (explicit override).
    2. ``md-template.docx`` directly beside this script — copy the script and the
       template into one folder and run with uv, no ``templates/`` needed.
    3. ``templates/md-template.docx`` next to this script (package data, and the
       "copy the script + templates folder" workflow).
    4. ``templates/md-template.docx`` one directory above the script.
    5. ``templates/md-template.docx`` two directories above the script — the
       in-repo ``src/md_to_docx/`` layout where the template lives at the root.
    6. ``templates/md-template.docx`` under the current working directory.

    Returns the first existing path, or ``None`` when no template is found (the
    caller falls back to a blank document).
    """
    env_override = os.environ.get("MD_TO_DOCX_TEMPLATE")
    if env_override:
        p = Path(env_override).expanduser()
        if p.exists():
            return p

    script_dir = Path(__file__).resolve().parent
    candidates = [
        # Template dropped directly beside the script (copy the two files and go).
        script_dir / "md-template.docx",
        # Bundled package data / template in a templates/ folder next to the script.
        script_dir / "templates" / "md-template.docx",
        # Repo root when the script lives directly under it.
        script_dir.parent / "templates" / "md-template.docx",
        # Repo root for the src/md_to_docx/ layout (root/templates/...).
        script_dir.parent.parent / "templates" / "md-template.docx",
        # Wherever the user is running from.
        Path.cwd() / "templates" / "md-template.docx",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def _add_heading_safe(doc, text: str, level: int):
    """Add a heading, degrading gracefully when the style is missing.

    Templates commonly define Heading 1-4 but omit Heading 5/6. ``add_heading``
    raises ``KeyError`` when the "Heading N" style is absent, so we fall back to
    a bold paragraph (with a size that steps down by level) rather than failing
    the whole conversion.
    """
    try:
        return doc.add_heading(text, level=level)
    except KeyError:
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        # Step the size down slightly for deeper (fallback) levels.
        run.font.size = Pt(max(11, 16 - level))
        return p


def build_docx(blocks: list, output_path: str, title: str, author: str, date: str,
               base_dir=None, allow_remote_images=False):
    """Build the DOCX from parsed blocks using the template for styles.

    ``base_dir`` resolves relative image paths; it defaults to the current
    working directory when not supplied. ``allow_remote_images`` enables
    fetching ``http(s)`` image sources (off by default).
    """
    if base_dir is None:
        base_dir = Path.cwd()
    template_path = find_template()
    if template_path is not None and template_path.exists():
        doc = Document(str(template_path))
        # Remove all placeholder content paragraphs and tables from template
        body = doc.element.body
        for child in list(body):
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag in ('p', 'tbl'):
                body.remove(child)
    else:
        doc = Document()

    # Guarantee a Hyperlink character style exists so internal/external links
    # render blue + underlined even when the template omits the style.
    _ensure_hyperlink_style(doc)

    # Build the heading anchor map (also tags each heading block with a
    # "_bookmark" name) so TOC links can resolve to real internal hyperlinks.
    anchors = build_heading_anchors(blocks)

    # Find the abstractNumIds we need from the template's numbering defs
    bullet_abstract = _find_abstract_num(doc, fmt='bullet', left='360')
    bullet2_abstract = _find_abstract_num(doc, fmt='bullet', left='720')
    numbered_abstract = _find_abstract_num(doc, fmt='decimal', left='720')

    # Fallback to template style numIds if abstracts not found
    if bullet_abstract is None:
        bullet_abstract = 8  # default template
    if bullet2_abstract is None:
        bullet2_abstract = 6
    if numbered_abstract is None:
        numbered_abstract = 7

    for block in blocks:
        if block["type"] in ("h1", "h2", "h3", "h4", "h5", "h6"):
            level = int(block["type"][1])
            h = _add_heading_safe(doc, block["text"], level)
            if block.get("_bookmark"):
                _add_bookmark(h, block["_bookmark"])

        elif block["type"] == "hr":
            # Add a thin line as paragraph border
            p = doc.add_paragraph()
            p.space_after = Pt(6)

        elif block["type"] == "meta":
            p = doc.add_paragraph()
            p.space_after = Pt(2)
            first = True
            for meta_line in block["lines"]:
                if not first:
                    p.add_run("\n")
                add_formatted_text(p, meta_line, anchors, base_dir=base_dir,
                                   allow_remote_images=allow_remote_images)
                first = False

        elif block["type"] == "blockquote":
            # Blockquote (e.g. an email body) — render as a paragraph with a
            # vertical bar on the left. Uses the template's "Quote" style when
            # available and applies a left border bar directly (see
            # _apply_blockquote_bar) so the bar shows with or without the template.
            # Split into paragraphs on empty lines, add each as separate paragraph.
            def _emit_quote(para_lines):
                try:
                    p = doc.add_paragraph(style="Quote")
                except KeyError:
                    p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.25)
                p.paragraph_format.space_after = Pt(4)
                _apply_blockquote_bar(p)
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
            # Render as Consolas text in a shaded, padded box.
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(8)
            # Shaded background plus a same-color border that pads the text.
            _apply_code_block_box(p)
            # Optional language caption, styled distinctly from the code so the
            # info string isn't lost. Only rendered when a language was captured,
            # so no-language blocks match the previous output exactly.
            language = block.get("language")
            first_line = True
            if language:
                caption = p.add_run(language)
                caption.font.name = "Consolas"
                caption.font.size = Pt(8)
                caption.italic = True
                caption.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
                first_line = False
            # Add each line with line breaks between them
            for line_idx, code_line in enumerate(block["lines"]):
                if line_idx > 0 or not first_line:
                    p.add_run().add_break()
                run = p.add_run(code_line)
                run.font.name = "Consolas"
                run.font.size = Pt(9)

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
                _apply_custom_table_style(doc, table)
                for r_idx, row in enumerate(rows):
                    for c_idx, cell in enumerate(row):
                        if c_idx < len(table.columns):
                            tc = table.cell(r_idx, c_idx)
                            tc.text = ""
                            para = tc.paragraphs[0]
                            add_formatted_text(para, cell, anchors, base_dir=base_dir,
                                   allow_remote_images=allow_remote_images)
                # Now that cells have content, fix column widths
                _fix_narrow_column_widths(table)
                # Add a small spacer paragraph after the table
                spacer = doc.add_paragraph()
                spacer.paragraph_format.space_before = Pt(8)
                spacer.paragraph_format.space_after = Pt(0)
                spacer_run = spacer.add_run()
                spacer_run.font.size = Pt(2)

        elif block["type"] == "note":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.3)
            run = p.add_run(block["text"])
            run.italic = True
            run.font.size = Pt(10)

        elif block["type"] == "paragraph":
            p = doc.add_paragraph()
            add_formatted_text(p, block["text"], anchors, base_dir=base_dir,
                                   allow_remote_images=allow_remote_images)

        elif block["type"] == "list":
            # Create a new numId for this bullet list block
            list_num_id = _new_num_id(doc, bullet_abstract)
            for item in block["items"]:
                # Items are dicts: {"text": str, "checked": Optional[bool]}.
                # checked is None for a normal bullet, True/False for a task item.
                item_text = item["text"] if isinstance(item, dict) else item
                checked = item.get("checked") if isinstance(item, dict) else None
                p = doc.add_paragraph(style="List Bullet")
                # Direct numPr override to use our specific numbering instance
                pPr = p._p.get_or_add_pPr()
                numPr = OxmlElement('w:numPr')
                numPr.append(_make_element('w:ilvl', val='0'))
                numPr.append(_make_element('w:numId', val=str(list_num_id)))
                pPr.append(numPr)
                # Task items get a checkbox glyph prefix (☑ checked / ☐ unchecked).
                if checked is not None:
                    p.add_run("\u2611 " if checked else "\u2610 ")
                add_formatted_text(p, item_text, anchors, base_dir=base_dir,
                                   allow_remote_images=allow_remote_images)

        elif block["type"] == "numbered_list":
            # Create a new numId for this numbered list block (ensures restart)
            list_num_id = _new_num_id(doc, numbered_abstract)
            for item in block["items"]:
                item_text = item["text"] if isinstance(item, dict) else item
                sub_items = item.get("sub_items", []) if isinstance(item, dict) else []
                p = doc.add_paragraph(style="List Number")
                # Direct numPr override
                pPr = p._p.get_or_add_pPr()
                numPr = OxmlElement('w:numPr')
                numPr.append(_make_element('w:ilvl', val='0'))
                numPr.append(_make_element('w:numId', val=str(list_num_id)))
                pPr.append(numPr)
                add_formatted_text(p, item_text, anchors, base_dir=base_dir,
                                   allow_remote_images=allow_remote_images)
                # Render sub-items as indented bullets
                if sub_items:
                    sub_num_id = _new_num_id(doc, bullet2_abstract)
                    for sub in sub_items:
                        sp = doc.add_paragraph(style="List Bullet 2")
                        sp_pPr = sp._p.get_or_add_pPr()
                        sp_numPr = OxmlElement('w:numPr')
                        sp_numPr.append(_make_element('w:ilvl', val='0'))
                        sp_numPr.append(_make_element('w:numId', val=str(sub_num_id)))
                        sp_pPr.append(sp_numPr)
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
    positional = []
    for arg in sys.argv[1:]:
        if arg in ("--fetch-remote-images", "--fetch-remote"):
            allow_remote_images = True
        else:
            positional.append(arg)

    if len(positional) < 2:
        print("Usage: uv run md_to_docx.py [--fetch-remote-images] "
              "<input.md> <output.docx>")
        sys.exit(1)

    input_path = positional[0]
    output_path = positional[1]

    if not Path(input_path).exists():
        print(f"ERROR: {input_path} not found")
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
    )


if __name__ == "__main__":
    main()
