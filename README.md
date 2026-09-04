# md-to-docx

Convert Markdown to a styled Word `.docx` document. The output uses **named Word
styles** generated from a set of built-in defaults that you can override with a
YAML config (see [Styling](#styling)). No Word template is required. Because
styles are named, you can open the `.docx` in Word and restyle all headings (or
body text, code blocks, etc.) at once by editing the style — just like any
normal Word document.

## Supported Markdown

- Headings, levels 1-6 (`#` through `######`)
- Paragraphs with hard-wrap preservation (two trailing spaces = line break)
- Bold, italic, bold-italic, inline code, and strikethrough (`~~text~~`)
- Backslash escapes (`\*`, `` \` ``, `\~`, `\[`, ...) render punctuation literally
- Bullet lists, numbered lists (with one level of sub-bullets), and GFM task
  lists (`- [ ]` / `- [x]`, rendered with checkbox glyphs)
- Fenced code blocks, including the language info string (```` ```python ````)
- Tables (pipe-delimited with a leading `|` on each row), with per-column
  alignment from the separator row (`:---` left, `:---:` center, `---:` right),
  escaped pipes (`\|`) as literal cell text, and `<br>` for in-cell line breaks
- Blockquotes (with a left bar)
- Horizontal rules
- Images `![alt](path)` (local PNG, JPEG, GIF, BMP, and TIFF files are embedded
  and scaled to fit the page width). Missing files and unsupported formats such
  as SVG fall back to the image's alt text. Remote `http(s)` images are fetched
  only with the opt-in `--fetch-remote-images` flag (see below); otherwise they
  fall back to alt text too.
- Links: inline `[text](url)`, internal anchors `[text](#heading)`, Obsidian
  `[[#Heading]]` wiki-links, bare URLs, and `<https://...>` autolinks
- Clickable internal table-of-contents links via heading bookmarks

Not supported: nested lists beyond one level, reference-style links, footnotes,
setext headings, indented code blocks, raw HTML, nested blockquotes, and
pipe-less tables.

## Install

```sh
uv tool install .
# or
pip install .
```

This exposes an `md-to-docx` command:

```sh
md-to-docx input.md output.docx
```

### Remote images

Remote `http(s)` image sources are not fetched by default (a document should not
make network requests unless you ask it to). Pass `--fetch-remote-images` to
download and embed them:

```sh
md-to-docx --fetch-remote-images input.md output.docx
```

Fetches enforce a timeout and a size cap; any failure falls back to the image's
alt text.

## Styling

The converter generates **named Word styles** from the YAML config (or built-in
defaults). Body text uses the `Normal` style, headings use `Heading 1` through
`Heading 6`, and so on. A global document-default font ensures everything
inherits a single typeface unless a style overrides it.

You do not need a config file for normal use. To customize, pass a YAML file
with `--style`:

```sh
md-to-docx --style style.yaml input.md output.docx
```

You can also set the `MD_TO_DOCX_STYLE` environment variable instead of the
flag. A YAML file overrides only the keys it names; everything else keeps its
default (deep merge).

To see every available key with its default value, dump the effective config:

```sh
md-to-docx --dump-config > style.yaml
```

`style.example.yaml` in this repo is that dump, ready to edit.

### Units and value formats

- Sizes end in `_pt` (points); indents and margins end in `_in` (inches).
- Colors are 6-digit hex strings, e.g. `"1F3864"` (quote all-digit colors like
  `"123456"` in YAML so they aren't read as numbers).
- Booleans are `true`/`false`.

### Key groups

- `font` — the global document-default font (default `"Aptos"`). Everything
  inherits this unless a style sets its own font.
- `heading_font` — the default heading typeface (default `"Aptos Display"`). All
  heading levels inherit this; set a per-level `font` to override just that level.
- `page` — `size` (`letter` | `a4` | `{width_in, height_in}`) and `margins_in`.
- `body` — `font` (null = inherit global `font`), `size_pt`, `color`,
  `space_before_pt`, `space_after_pt`, `line_spacing`. Maps to the `Normal` style.
- `headings` — per level `1`-`6`: `font` (null = inherit `heading_font`),
  `size_pt`, `color`, `bold`, `italic`, `space_before_pt`, `space_after_pt`,
  and an optional `rule` (an under-heading horizontal line). `rule` is a mapping
  of `width_pt` (line thickness), `color` (hex), and `space_pt` (gap between the
  text and the line); set `rule: null` to remove it. By default H1 and H2 have a
  rule and H3-H6 do not. Maps to `Heading 1` through `Heading 6` styles.
- `inline_code` — `font`, `size_pt`, optional `color` and `fill`. Maps to the
  `MdCodeChar` character style.
- `code_block` — `font`, `size_pt`, `fill`, `padding_pt` (interior padding; the
  box's left edge stays aligned with body text automatically), `space_before_pt`,
  `space_after_pt`, `indent_in` (extra right-side inset so the box stops short of
  the right margin), and a `caption` (`size_pt`, `color`, `italic`). The shaded
  box (border + fill) lives on the "Code Block" paragraph style. A fenced block's
  language info string renders as a "Code Block Title" paragraph (based on Code
  Block, so it shares the same box) styled from `caption`.
- `blockquote` — `bar_color`, `bar_width_pt`, `bar_gap_pt`, `indent_in`,
  `space_after_pt`. Maps to the `MdQuote` paragraph style (the left bar is
  direct formatting on top of the style).
- `lists` — `bullet` (`glyphs`, `indent_in`, `space_after_pt`) and `ordered`
  (`indent_in`, `space_after_pt`, `restart_each_block`).
- `table` — `border` (`style`, `width_pt`, `color`, `edges`), `header` (`bold`,
  `fill`), `cell_margins_pt` (cell padding; `left` also drives horizontal
  alignment — the table indent is set to this so cell text lines up near the body
  margin), `width` (`full` | `auto`), `indent_in` (right-side inset; the table is
  narrowed on the right). `edges` lists which borders to draw
  (`top`, `bottom`, `left`, `right`, `insideH`, `insideV`); by default only
  `insideH` (between-row rules) is drawn, so there are no outer top/bottom lines.
  Add `top`/`bottom` to box the table.
- `hr` — a markdown `---` line. `rule` draws a horizontal line (on by default);
  `width_pt` and `color` control it, and `space_after_pt` is the gap below.
  Set `rule: false` to render blank space instead of a line.
- `links` — `color`, `underline`. Maps to the `Hyperlink` character style.

## Standalone (no install)

The single file `src/md_to_docx/md_to_docx.py` carries
[PEP 723](https://peps.python.org/pep-0723/) inline metadata, so you can copy it
anywhere and run it directly with [uv](https://docs.astral.sh/uv/); uv resolves
its dependencies (`python-docx`, `pyyaml`) automatically:

```sh
uv run md_to_docx.py input.md output.docx
uv run md_to_docx.py --style style.yaml input.md output.docx
```

No template or other files are needed — the default styling is baked into the
script.

## Development

```sh
uv run pytest
```
