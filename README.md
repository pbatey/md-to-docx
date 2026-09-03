# md-to-docx

Convert Markdown to a styled Word `.docx` document. Styling is applied as direct
formatting from a set of built-in defaults that you can override with a YAML
config (see [Styling](#styling)). No Word template is required.

## Supported Markdown

- Headings, levels 1-6 (`#` through `######`)
- Paragraphs with hard-wrap preservation (two trailing spaces = line break)
- Bold, italic, bold-italic, inline code, and strikethrough (`~~text~~`)
- Backslash escapes (`\*`, `` \` ``, `\~`, `\[`, ...) render punctuation literally
- Bullet lists, numbered lists (with one level of sub-bullets), and GFM task
  lists (`- [ ]` / `- [x]`, rendered with checkbox glyphs)
- Fenced code blocks, including the language info string (```` ```python ````)
- Tables (pipe-delimited with a leading `|` on each row)
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

Styling is applied as direct formatting from built-in defaults. You do not need
a config file for normal use. To customize, pass a YAML file with `--style`:

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

- `page` — `size` (`letter` | `a4` | `{width_in, height_in}`) and `margins_in`.
- `body` — `font`, `size_pt`, `color`, `space_before_pt`, `space_after_pt`,
  `line_spacing` for normal paragraphs and inline runs.
- `headings` — per level `1`-`6`: `font`, `size_pt`, `color`, `bold`, `italic`,
  `space_before_pt`, `space_after_pt`.
- `inline_code` — `font`, `size_pt`, optional `color` and `fill`.
- `code_block` — `font`, `size_pt`, `fill`, `padding_pt`, `space_before_pt`,
  `space_after_pt`, and a `caption` (`size_pt`, `color`, `italic`).
- `blockquote` — `bar_color`, `bar_width_pt`, `bar_gap_pt`, `indent_in`,
  `space_after_pt`.
- `lists` — `bullet` (`glyphs`, `indent_in`, `space_after_pt`) and `ordered`
  (`indent_in`, `space_after_pt`, `restart_each_block`).
- `table` — `border` (`style`, `width_pt`, `color`), `header` (`bold`, `fill`),
  `cell_margins_pt`, `width` (`full` | `auto`).
- `note` — `indent_in`, `italic`, `size_pt`.
- `hr` — `space_after_pt`, `rule` (draw a horizontal line when `true`).
- `links` — `color`, `underline`.

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
