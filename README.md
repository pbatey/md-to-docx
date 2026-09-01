# md-to-docx

Convert Markdown to a styled Word `.docx` document. Styling comes from a Word
template (`templates/md-template.docx`).

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

See `.kiro/specs/markdown-feature-coverage/requirements.md` for the explicit
out-of-scope list (nested lists beyond one level, reference-style links,
footnotes, setext headings, indented code blocks, raw HTML, nested
blockquotes, remote image fetching, and pipe-less tables).

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

## Standalone (no install)

The single file `src/md_to_docx/md_to_docx.py` carries
[PEP 723](https://peps.python.org/pep-0723/) inline metadata, so you can copy
it together with a `templates/` folder and run it directly with
[uv](https://docs.astral.sh/uv/):

```sh
uv run md_to_docx.py input.md output.docx
```

The script finds the template by looking (in order) for `md-template.docx`
directly beside itself, then `templates/md-template.docx` next to it and in a
few parent directories, then under the current working directory. So you can
copy `md_to_docx.py` and `md-template.docx` into a single folder and run it.
Set `MD_TO_DOCX_TEMPLATE` to point at a specific template:

```sh
MD_TO_DOCX_TEMPLATE=/path/to/md-template.docx uv run md_to_docx.py input.md output.docx
```

## Development

```sh
uv run pytest
```
