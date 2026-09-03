# Design: YAML-Based Styling (Replace the DOCX Template)

## Overview

Replace the Word template with a code-defined default style config that can be
overridden by an optional YAML file. All styling is applied as **direct
formatting** on runs, paragraphs, and tables. This removes every dependency on
named Word styles and on a template's numbering part.

The change is concentrated in `src/md_to_docx/md_to_docx.py`. The render loop in
`build_docx` keeps its structure; each element's styling now reads from a
`StyleConfig` object instead of `doc.add_heading(...)`, `doc.add_paragraph(style=...)`,
`_find_abstract_num`, and the `PlainTable2` reference.

### What breaks when the template is removed (from code inventory)

Template-dependent today, so each needs a direct-formatting replacement:

1. Heading styles Heading 1-6 (via `doc.add_heading`) — fonts/sizes/colors/spacing.
2. Named paragraph styles: `Normal` (implicit body), `Quote`, `List Bullet`,
   `List Number`, `List Bullet 2`.
3. Numbering: `_find_abstract_num` lookups, magic fallback numIds `8/6/7`,
   `_new_num_id` requiring a numbering part.
4. Table style `PlainTable2` visual (borders, header fill).
5. Page geometry (size/margins), currently inherited from the template and used
   by `_content_width_emu` and the `9360`-twip assumption.

Already direct formatting (just promote to config knobs): blockquote bar
(`_apply_blockquote_bar`), code box (`_apply_code_block_box`), inline code
(`_style_run`), hyperlink color (`_ensure_hyperlink_style` injects it), note,
table sizing/margins/column-width algorithm, and all `Pt`/`Inches` spacing.

## Configuration Model

### Representation

A single nested `dict` is the config. Defaults are defined in code as a module
constant `DEFAULT_STYLE` (a plain dict, so it is trivially YAML-dumpable). A
`StyleConfig` wrapper provides typed, validated access with dotted lookups and
unit conversion, so the render code reads e.g. `style.heading(1).size_pt`.

Design choice: keep `DEFAULT_STYLE` as a plain dict (round-trippable to YAML per
R2.4) and layer a thin typed accessor over it, rather than a dataclass tree that
would duplicate the schema. Validation/coercion happens in the accessor.

### Schema (top-level keys)

```yaml
page:
  size: letter            # letter | a4 | {width_in, height_in}
  margins_in: {top: 1.0, bottom: 1.0, left: 1.0, right: 1.0}
body:
  font: "Calibri"
  size_pt: 11
  color: "000000"
  space_before_pt: 0
  space_after_pt: 8
  line_spacing: 1.15
headings:                 # per-level 1..6
  1: {font: "Calibri Light", size_pt: 20, color: "1F3864", bold: true,
      space_before_pt: 12, space_after_pt: 4}
  2: {...}
  # ... through 6
inline_code:
  font: "Consolas"
  size_pt: 10
  color: null             # optional
  fill: null              # optional shading
code_block:
  font: "Consolas"
  size_pt: 9
  fill: "F2F2F2"
  padding_pt: 6
  space_before_pt: 8
  space_after_pt: 8
  caption: {size_pt: 8, color: "808080", italic: true}
blockquote:
  bar_color: "999999"
  bar_width_pt: 2.25
  bar_gap_pt: 12
  indent_in: 0.25
  space_after_pt: 4
lists:
  bullet:
    glyphs: ["\u2022", "\u25E6"]   # level 0, level 1
    indent_in: [0.25, 0.5]
    space_after_pt: 2
  ordered:
    format: decimal
    indent_in: 0.5
    space_after_pt: 2
    restart_each_block: true
table:
  border: {style: single, width_pt: 0.5, color: "BFBFBF"}
  header: {bold: true, fill: "F2F2F2"}
  cell_margins_pt: {top: 3.6, bottom: 3.6, left: 5.4, right: 5.4}
  width: full             # full | auto
note:
  indent_in: 0.3
  italic: true
  size_pt: 10
hr:
  space_after_pt: 6
  rule: false             # true -> draw a bottom border line
links:
  color: "0563C1"
  underline: true
```

Units are explicit in key names (`_pt`, `_in`). Colors are 6-hex strings without
`#`. This satisfies R7.1 (documented units/formats).

### Loading and merging (R3)

- `load_style(path=None) -> dict`:
  - Start from a deep copy of `DEFAULT_STYLE`.
  - If `path` given: `yaml.safe_load` it. On `yaml.YAMLError`, print a clear
    error naming the file and exit non-zero (R3.3). `safe_load` avoids arbitrary
    object construction.
  - Deep-merge the loaded mapping over the defaults (R3.2): recurse into nested
    dicts; scalar/list values replace.
  - Unknown top-level or nested keys are ignored, optionally collected and
    warned once to stderr (R3.4).
- Validation/coercion is lazy in the accessor: a malformed color or non-numeric
  size raises a specific `StyleError` with the key path, OR falls back to the
  default for that key. Design decision: **fall back + warn** for individual
  malformed fields (keeps a run productive), **hard error** only for unparseable
  YAML. This matches R3.5's "defined per field" latitude.

## Direct-Formatting Helpers

Central helpers so every element applies config uniformly:

- `_apply_paragraph_format(paragraph, *, space_before_pt, space_after_pt,
  line_spacing=None, left_indent_in=None)` — sets `paragraph_format` fields.
- `_apply_run_format(run, *, font=None, size_pt=None, color=None, bold=None,
  italic=None, strike=None, fill=None)` — sets `run.font.*`; `color`/`fill` via
  `RGBColor.from_string` and a `w:shd` on the run's `rPr` when `fill` given.
- Existing `_style_run` is refactored to read inline-code settings from config
  (font/size/color/fill) instead of the hardcoded `Consolas`/`Pt(10)`.

### Headings (replace `add_heading`)

`_add_heading_safe` is replaced by `_add_heading(doc, text, level, style)`:
- `p = doc.add_paragraph()`; add a run with the level's font/size/color/bold.
- Apply space before/after from config.
- Set the paragraph's **outline level** (`w:pPr/w:outlineLvl val=level-1`) so the
  Word navigation pane and PDF bookmarks still see the hierarchy even though we
  no longer use the named Heading styles (preserves R4.3 outline behavior).
- Keep `_add_bookmark` exactly as-is (it inserts after `pPr`).

No `KeyError` fallback is needed because we never reference a named style.

### Body paragraphs / meta / note

- `paragraph` block: create paragraph, apply body spacing/line-spacing; inline
  runs pick up body font/size/color via `_apply_run_format` inside `_emit_run`
  and `_style_run` (thread the config through `add_formatted_text`).
- `note`: apply `note.indent_in`, italic, size from config.
- `meta`: apply spacing from config.

Threading: `add_formatted_text`, `_render_inline`, `_emit_run`, `_style_run`,
`_add_image` already receive extra kwargs (`base_dir`, `allow_remote_images`).
Add a `style` argument along the same path so inline runs get body/inline-code
formatting. `build_docx` holds the `StyleConfig` and passes it down.

### Code block

`_apply_code_block_box` gains parameters from config (fill, padding). Code line
runs and the caption read font/size/color from `code_block`/`code_block.caption`.

### Blockquote

`_apply_blockquote_bar` reads `blockquote.bar_color/bar_width_pt/bar_gap_pt`
(convert pt to the eighths-of-a-point `sz` and the point `space`). Indent and
spacing from config. No more `Quote` named style.

### Lists and numbering (R5 — the hardest part)

We must produce bullets and restarting ordered numbers from a blank document.
python-docx's default `Document()` does include a numbering part, but not the
specific abstracts/indents we need, and relying on default numIds is what we are
removing. Two viable approaches:

- **Approach A (chosen): create our own `abstractNum` definitions in code.**
  Add a helper `_ensure_numbering(doc)` that, once per document, injects
  `w:abstractNum` entries into the numbering part for: a bullet list (level 0
  glyph + level 1 glyph at configured indents) and a decimal ordered list.
  `_new_num_id` then references these known abstractIds (created in code, not
  looked up by `_find_abstract_num`), keeping the per-block `startOverride=1`
  restart. This preserves real Word list semantics (proper hanging indents,
  selectable list formatting) and keeps parity with today's output.
  - `_find_abstract_num` and the magic `8/6/7` fallbacks are deleted.
  - Glyphs/indents/number format come from `lists.*`.
- **Approach B (rejected): fake lists with a literal glyph run + manual indent**
  (no numbering part). Simpler, but ordered-list auto-numbering and clean
  hanging indents are lost, and multi-digit numbers misalign. Rejected for
  fidelity.

Task-list glyphs (checkbox) continue to be prepended in the `list` branch.

### Table (replace `PlainTable2`)

`_apply_custom_table_style` is rewritten to apply **direct** table properties
from `table.*`:
- Instead of `tblStyle val="PlainTable2"`, set direct table borders via
  `w:tblBorders` (top/bottom/left/right/insideH/insideV) using
  `table.border.{style,width_pt,color}`.
- Header row: bold the first row's runs and/or shade its cells (`w:shd`) per
  `table.header`.
- Keep the existing `tblLook`, `tblCellMar` (now from `table.cell_margins_pt`),
  `tblW`, `tblLayout`, and `_fix_narrow_column_widths` logic. Replace the
  hardcoded `9360` with the configured content width in twips.

### Page geometry

`_apply_page_setup(doc, style)` sets `section.page_width/height` and the four
margins from `page.*`. `_content_width_emu` continues to read the section (now
configured), so image scaling and the table width use the real content width.

## CLI and Packaging

- `main()` parses `--style <path>` (config file) and `--dump-config`. Existing
  positional args and `--fetch-remote-images` unchanged. Optional
  `MD_TO_DOCX_STYLE` env var mirrors `--style`.
- `--dump-config`: print `yaml.safe_dump(DEFAULT_STYLE)` and exit 0.
- `build_docx` gains a `style=None` parameter; when `None` it uses
  `load_style()` (defaults). `main()` calls `load_style(style_path)`.
- Add `pyyaml` to the PEP 723 block and `[project].dependencies`.
- Remove template packaging: drop `[tool.hatch.build.targets.wheel.force-include]`
  for the template and delete the `templates/` package data.
- Delete `find_template`, `MD_TO_DOCX_TEMPLATE`, `templates/md-template.docx`,
  and `src/md_to_docx/templates/`.

## Error Handling

- Invalid YAML syntax -> clear message with file path, `sys.exit(1)`.
- Malformed individual field (bad color/size) -> warn to stderr, use default.
- Unknown keys -> ignored (optional single warning).
- Missing config file path -> clear "not found" error, exit non-zero.

## Testing Strategy

Unit:
- `load_style`: defaults returned when no path; deep merge overrides a nested
  key while siblings keep defaults; invalid YAML raises/exits; unknown key
  ignored; malformed color falls back with warning.
- `--dump-config` output re-parses as YAML and equals `DEFAULT_STYLE`.

Direct-formatting (build a docx, read back):
- No template/config: headings render with configured size/color and carry
  `w:outlineLvl`; body paragraph uses body font/size.
- Override via config: e.g. set `headings.1.color` and assert the run color;
  set `code_block.fill` and assert the shading; set `blockquote.bar_color`
  and assert the border color; set `links.color` and assert the hyperlink run.
- Lists: bullet list renders bullets and ordered lists restart at 1 across two
  blocks (assert numbering present, second list restarts); task glyphs present.
- Table: direct `tblBorders` present with configured color; header row bold/fill.
- Page: custom margins applied to the section; content width reflected.
- Regression: all existing behavior tests pass after the template removal
  (headings 1-6, images, escapes, strike, autolinks, code language, task lists).
  Tests that currently assert template-derived specifics (e.g. the
  `Hyperlink`-style test, code-box test) are updated to the direct-formatting
  equivalents.

Remove/replace `find_template` tests.

## Migration / Compatibility Notes

- This is a breaking change to internals: `find_template`, `_find_abstract_num`,
  `_add_heading_safe`, `_ensure_hyperlink_style` (as a style injector), and the
  `PlainTable2` reference are removed or rewritten. External CLI usage stays
  compatible (same positional args) and gains `--style` / `--dump-config`.
- Visual parity is a target, not a guarantee; the default config is tuned to
  match the current look, but exact heading fonts/sizes from the old template
  must be transcribed into `DEFAULT_STYLE` (values to be confirmed during
  implementation by inspecting the current output).
