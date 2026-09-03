# Design: Generate Named Word Styles from YAML

## Overview

Replace most per-element direct formatting with **generated named Word styles**
built from the YAML config, plus a document-default font that everything
inherits. Content is assigned styles by name so the output is editable in Word.

Inline emphasis (bold/italic/strike) and inline code stay as run-level
formatting/character styles layered over the paragraph style. Lists keep the
existing code-defined numbering. Tables keep direct formatting (their header
text simply inherits the global font).

## Font inheritance model

Word resolves a run's font in this order: explicit run font → character style →
paragraph style → document defaults (`docDefaults/rPrDefault`) → theme. We use:

- **Document default** (`docDefaults`): the global `font` (default `"Aptos"`).
  Set once. Everything with no more-specific font inherits it.
- **Paragraph styles**: `Normal` inherits the default (no font set); headings
  set `"Aptos Display"` explicitly.
- **Character styles**: code style sets `"Consolas"`; `Hyperlink` sets color +
  underline (no font, so it inherits).

So "null/absent font" at any level means "inherit", exactly as requested.

## Config changes

Add a top-level `font`:

```yaml
font: "Aptos"          # global document default; everything inherits this
```

Change existing defaults:
- `body.font`: `null` (inherit the global default)
- `headings.<n>.font`: `"Aptos Display"`
- `inline_code.font`: `"Consolas"` (unchanged)
- `code_block.font`: `"Consolas"` (unchanged)
- `note` gains an implicit inherit (no font key needed)

`load_style` deep-merge and `StyleConfig` are unchanged in shape; a `null`/absent
font now means "don't set it on the style, inherit."

## Style generation

New module area: a `_build_styles(doc, sc)` routine run once at the start of
`build_docx`, before content. It creates/updates styles via python-docx's
`doc.styles` plus raw OOXML where needed.

### Document default font

`_set_default_font(doc, font)`: write `w:rFonts` (ascii/hAnsi/cs) into
`styles.xml` `docDefaults/rPrDefault/rPr`. python-docx exposes the styles
element; we get_or_add the `docDefaults` chain and set the fonts.

### Paragraph styles

Helper `_define_paragraph_style(doc, style_id, name, *, font, size_pt, color,
bold, italic, space_before_pt, space_after_pt, line_spacing, left_indent_in,
outline_level, base="Normal")`:
- Use `doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)` when absent, else
  fetch it. `Normal` already exists in a blank doc — update it in place.
- Set `style.font.*` for the run properties; only set `.name` when `font` is not
  None (so null inherits).
- Set `style.paragraph_format.*` for spacing/indent/line spacing.
- For headings, set the outline level on the style's `pPr` (`w:outlineLvl`).

Generated styles and their config sources:

| Style name          | styleId       | source keys                     |
|---------------------|---------------|---------------------------------|
| Normal              | Normal        | `body.*` (font null → inherit)  |
| Heading 1..6        | Heading1..6   | `headings.<n>.*` (font Display) |
| Md Quote            | MdQuote       | `blockquote.*` (indent/spacing) |
| Md Code             | MdCode        | `code_block.*` (font/size/fill) |
| Md Note             | MdNote        | `note.*`                        |

We use custom style ids (`MdQuote`, `MdCode`, `MdNote`) to avoid clashing with
Word's built-ins that carry extra baggage; headings reuse Word's `Heading N`
names/ids so Word treats them as real headings (nav pane, TOC field support).

### Character styles

- `MdCode` char style? No — inline code and code blocks both want the monospace
  font + size. Use one character style `MdCodeChar` (font/size, optional
  color/fill) for inline code; code-block paragraphs use the `MdCode` paragraph
  style for the box/spacing and set the code font on runs (or link the char
  style). Simplevst approach: inline code → `MdCodeChar` character style; code
  block lines → runs with the code font/size (paragraph style provides the box).
- `Hyperlink` character style: recreate the generator removed earlier
  (`_define_char_style`) with color + underline from `links.*`. Internal/
  external hyperlink runs reference `w:rStyle val="Hyperlink"` again (revert to
  the style-based approach from the direct-color approach), so link styling is
  editable.

## Applying styles in the render loop

- **Headings**: `p = doc.add_paragraph(style=f"Heading {level}")`; add the text
  run (emphasis not parsed today); bookmark as before. No per-run font/size/
  color — the style carries it. Remove `_add_heading`'s direct run formatting.
- **Body paragraph**: `doc.add_paragraph()` uses `Normal`; inline runs no longer
  get body font/size stamped in `_style_run` — they inherit from `Normal`.
  `_style_run` keeps applying bold/italic/strike and, for code, the `MdCodeChar`
  character style.
- **Blockquote**: `doc.add_paragraph(style="MdQuote")` + keep the left-bar border
  (still applied directly; a border is fine to keep direct, or move into the
  style's pPr). Move indent/spacing into the style.
- **Code block**: `doc.add_paragraph(style="MdCode")` for spacing + the shaded
  padded box (box can live in the style's pPr borders/shading, or stay direct).
  Code line runs get the code font/size (or the `MdCodeChar` char style).
- **Note**: `doc.add_paragraph(style="MdNote")`.
- **Lists**: keep code-defined numbering. List item paragraphs may use `Normal`
  (so they inherit the font) plus the numPr; or a `List Paragraph` style. Keep it
  simple: `Normal` + numPr, matching today minus the direct font.
- **Table**: unchanged (direct). Header text inherits the global font because we
  no longer stamp body font on cell runs.

### `_style_run` change

Currently `_style_run` stamps body font/size/color on every run. New behavior:
- Do NOT set font/size/color for normal runs (inherit from the paragraph style /
  docDefaults).
- Apply bold/italic/strike as today.
- For code runs, apply the `MdCodeChar` character style (or set the code font +
  size directly), so code stays monospace regardless of inheritance.

This is the key change that lets inheritance work: runs stop overriding the font.

## Hyperlinks: revert to character style

`_build_hyperlink_run` currently sets direct color + underline. Change it to
reference the generated `Hyperlink` character style (`w:rStyle`), and generate
that style from `links.*`. This makes link styling editable in Word.

## Error handling

- Missing/invalid font: a bad (non-string) font value warns and is treated as
  "inherit" (don't set it).
- Everything else follows existing coercion (colors via `_as_hex`, numbers via
  `_coerce_number`).

## Testing strategy

Update/extend tests to assert style-based output:

- **Default font**: `docDefaults` has `w:rFonts` = Aptos; a body paragraph run
  has no explicit font (inherits) and `Normal` style has no font override.
- **Headings**: `Heading 1` style exists with font `Aptos Display`, size/color
  from config, and outline level on the style; a heading paragraph's
  `style.name == "Heading 1"`.
- **Editability proxy**: heading paragraphs reference the style by name (not
  direct run color), i.e. `p.style.name` is the heading style and the run has no
  direct color.
- **Code**: inline code run uses the code char style / monospace font; code
  block paragraph uses `MdCode` and its runs are Consolas.
- **Links**: hyperlink runs reference the `Hyperlink` char style; that style
  exists with the configured color + underline.
- **Inheritance**: overriding top-level `font` in YAML changes `docDefaults`;
  body text has no direct font so it inherits (assert `docDefaults` font).
- **Emphasis + style**: bold within a heading/body still sets `run.bold`.
- **Lists/tables/regression**: existing tests pass; where a test asserted a
  direct run color/size that now lives in a style, assert the style instead.

## Migration notes

- Reverses "direct formatting only" for block elements; inline emphasis stays
  direct.
- Re-introduces a `Hyperlink` character style generator (removed in the
  yaml-styling work) and adds paragraph/char style generators.
- Tests from yaml-styling that assert direct heading run color / direct link
  color must be updated to assert the style instead.
- `_add_heading` (direct) becomes "add paragraph with Heading N style".
