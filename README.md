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
- Bullet and numbered lists with arbitrary nesting (up to five levels),
  including mixed ordered/unordered nesting; nesting is normalized by
  indentation order, so irregular hand-authored indentation still nests
  sensibly. GFM task lists (`- [ ]` / `- [x]`) render with checkbox glyphs.
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

Not supported: reference-style links, footnotes, setext headings, indented code
blocks, raw HTML, nested blockquotes, and pipe-less tables.

## Install

Install it as a global command with [uv](https://docs.astral.sh/uv/) (installs
into an isolated environment and puts `md-to-docx` on your `PATH`):

```sh
uv tool install .
# from anywhere, straight from GitHub:
uv tool install git+https://github.com/pbatey/md-to-docx.git
# or with pip:
pip install .
```

If the command isn't found in a new shell, run `uv tool update-shell` once and
restart your terminal. To pick up local code changes after editing, reinstall
with `uv tool install --from . md-to-docx --reinstall`.

This exposes an `md-to-docx` command:

```sh
md-to-docx input.md output.docx
```

### Usage

```
md-to-docx [-h] [--version] [--fetch-remote-images]
           [--style config.yaml] [--dump-config]
           [--no-sibling-links] [--link-base-url URL]
           [--sharepoint | --sp]
           [-r] [-o DIR]
           input output.docx
```

`input` (a Markdown file, or a directory with `-r`) is required for a
conversion. `output.docx` is required unless you pass `-o`/`--outdir` or `-r`;
`--dump-config` is the one mode that runs without them.

| Option | Description |
| --- | --- |
| `-h`, `--help` | Show usage and exit. |
| `--version` | Print the version and exit. |
| `--fetch-remote-images` | Download and embed remote `http(s)` images (off by default). |
| `--style config.yaml` | Path to a YAML style config (overrides `MD_TO_DOCX_STYLE`). |
| `--dump-config` | Print the default style config as YAML to stdout and exit. |
| `--no-sibling-links` | Disable link rewriting and restore the output from before this feature. |
| `--link-base-url URL` | Write links to other documents as absolute URLs under this `http(s)` folder URL. |
| `--sharepoint`, `--sp` | Append `?web=1` to absolute links so SharePoint opens them in Word for the web. Requires `--link-base-url`. |
| `-r`, `--recursive` | Convert every Markdown file under the input directory. |
| `-o DIR`, `--outdir DIR` | Write output under `DIR`, mirroring the source layout. |

Run `md-to-docx --help` to see this list at any time.

### Links between documents

Links to other local Markdown files are rewritten to hyperlinks to the matching
`.docx` files, so a converted set of documents can be navigated by clicking.

- A link whose target is a relative path ending in `.md` or `.markdown` (any
  letter case) becomes a hyperlink to the same path ending in `.docx`:
  `[Setup](docs/setup.md)` links to `docs/setup.docx`. Directory segments such
  as `../` and `sub/` are kept; a leading `./` is dropped.
- A heading fragment is kept: `[Install](setup.md#install)` links to the
  `install` heading bookmark in `setup.docx`.
- Obsidian wiki-links work the same way: `[[Note]]`, `[[Note|Label]]`, and
  `[[Note#Heading]]` link to `Note.docx`.
- Percent-encoded sequences (`%20`) are decoded, `<My Notes.md>` angle-bracket
  targets are accepted, and backslashes become `/`. The resulting address is
  percent-encoded.
- Links are left unchanged when they have a URL scheme (`https:`, `mailto:`),
  start with `#`, point to a non-Markdown file, or are absolute filesystem paths.
- A warning is written to standard error for each link whose target `.md` file
  does not exist. The hyperlink is still written and the exit status is not
  affected.

Relative links resolve only if the `.docx` files keep the same relative layout
as the Markdown files. Converting a single file does not convert the files it
links to; convert the whole tree (see [Converting a directory](#converting-a-directory))
and keep the output layout intact.

Heading bookmarks are named with at most 40 letters, digits, and underscores
(long names are truncated with a short hash) so that links into other documents
stay valid.

Cross-document heading anchors may not scroll to the heading in Word for the
web; the document still opens.

#### `--no-sibling-links`

Link rewriting is on by default. When `--no-sibling-links` is omitted, sibling
and wiki links are rewritten as described above. Pass `--no-sibling-links` to
restore the output from before this feature: those links are rendered as plain
text, other links are unchanged, and heading bookmarks use the legacy
`_toc_{idx}_{slug}` names. The flag takes no value. If combined with
`--link-base-url`, `--no-sibling-links` wins and no absolute links are written.

#### `--link-base-url`

By default links are relative. Relative links work only when the documents are
opened from local or synced folders. Word for the web (for example a `.docx`
opened in a browser from SharePoint) cannot resolve relative paths. For that
case, give the URL of the folder that holds the converted files:

```sh
md-to-docx --link-base-url https://contoso.sharepoint.com/sites/team/Shared%20Documents/docs/ a.md out/a.docx
```

A link to `b/c.md` becomes `https://contoso.sharepoint.com/sites/team/Shared%20Documents/docs/b/c.docx`.
The trailing `/` on the base URL is optional, existing `%XX` sequences are not
encoded again, and `./` and `../` segments are resolved. A `../` that would climb
above the base URL root is dropped and produces a warning. The URL must be
`http` or `https` with a host, otherwise the command exits with an error. In
single-file mode the base URL is the folder of the output file; in recursive
mode it is the URL of the output root. Links that are not to Markdown files are
not affected.

#### `--sharepoint` / `--sp`

SharePoint downloads a `.docx` when you open a plain link to it. Add
`--sharepoint` (alias `--sp`) to append `?web=1` to every absolute link to
another document, so SharePoint opens the document in Word for the web instead.
The option takes no value and requires `--link-base-url`; without it the command
exits with an error. It has no effect with `--no-sibling-links`, and links that
are not to Markdown files are unchanged.

You can paste the address of the library folder straight from your browser,
including the `.../Forms/AllItems.aspx?id=...` form:

```sh
md-to-docx -r docs/ --outdir out/ --sharepoint \
  --link-base-url 'https://contoso.sharepoint.com/sites/team/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2Fteam%2FShared%20Documents%2Fdocs'
```

With an `id` parameter, the base URL becomes the host plus the decoded `id`
path (here `https://contoso.sharepoint.com/sites/team/Shared Documents/docs`,
percent-encoded once in the links). Without `id`, the document library root is
used. Without `--sharepoint`, a library page URL is used as given and a warning
suggests `--sharepoint`.

Heading fragments are still written as an anchor in the link, not in the URL.
Tested in SharePoint: a link from one document to a heading in another opens
the linked document in a new tab at the top of the page. Word for the web does
not move to the heading, with or without `#fragment` appended to the URL. Links
to a heading in the same document do jump to it.

### Converting a directory

`-r`/`--recursive` converts every Markdown file under an input directory:

```sh
md-to-docx -r docs/                  # writes each .docx next to its .md
md-to-docx -r docs/ --outdir out/    # writes out/<same relative path>.docx
```

- Files ending in `.md` or `.markdown` (any letter case) are converted, at any
  depth, in sorted path order. Each output has the same stem with `.docx`.
- Directories whose names start with `.` and directories named `node_modules`,
  `venv`, or `__pycache__` are skipped, as are files whose names start with `.`.
  Symbolic links to directories are not followed.
- `-o`/`--outdir DIR` writes outputs under `DIR` with the same relative layout
  as the source tree, creating directories as needed, so relative links keep
  working. `--outdir` also works with a single input file (`DIR/<stem>.docx`).
  It cannot be combined with a positional output argument, and `-r` cannot be
  combined with one either.
- A directory input requires `-r`. `--dump-config` ignores `-r` and `--outdir`.
- One progress line is printed per file, and a final summary gives the number of
  files converted and failed. If a file fails, the error is reported, the other
  files are still converted, and the command exits non-zero after trying every
  file. Two files in one directory with the same stem (`a.md` and `a.markdown`)
  are reported as an error and neither is converted.
- An empty tree prints a warning and exits 0.
- `--style`, `--fetch-remote-images`, `--no-sibling-links`, and
  `--link-base-url` apply to every file.

SharePoint workflow: convert the tree, upload `out/` to the library folder, and
open the documents from there.

```sh
md-to-docx -r docs/ --outdir out/ --link-base-url https://.../docs/
```

Here `https://.../docs/` is the URL of the SharePoint folder where you upload
the contents of `out/`. Use the base URL form only for browser viewing; for
local or synced folders, omit `--link-base-url` and keep the relative links.

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

The same options apply as the installed command (`--help`, `--version`,
`--fetch-remote-images`, `--style`, `--dump-config`, `--no-sibling-links`,
`--link-base-url`, `-r`/`--recursive`, `-o`/`--outdir`). No template or other files
are needed — the default styling is baked into the script.

## Development

```sh
uv run pytest
```

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for release notes and notable changes.
