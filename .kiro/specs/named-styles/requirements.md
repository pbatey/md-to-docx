# Requirements: Generate Named Word Styles from YAML

## Introduction

The converter currently styles every element with **direct formatting** pulled
from a YAML config (`DEFAULT_STYLE` + optional overrides). That is legible in
YAML but produces documents where each paragraph carries its own inline
formatting, so a user cannot change "all Heading 1s" from within Word by editing
a style.

This feature makes the converter **generate real named Word styles** (paragraph
and character styles, plus the document default font) from the same YAML config,
and apply those styles to content by name. The YAML remains the single source of
truth; the generated `.docx` becomes editable in Word the normal way (edit the
style, all instances update).

This reverses the "direct formatting only" decision from the yaml-styling spec,
but keeps its principle: styling is declared in YAML, not hidden in a binary
template. The difference is that YAML now compiles into a Word stylesheet.

## Goals

- A top-level global default font applied at the document root so everything
  inherits it unless a style overrides it.
- Named paragraph styles for body/headings/quote/code/note generated from YAML.
- Content applied by style name, so the document is editable in Word.
- YAML config remains the source of truth; defaults live in code.
- Preserve current visual output and all existing features/tests.

## Constraints

- YAML config shape stays backward compatible: existing keys keep working; new
  keys are additive.
- `pyyaml` remains the only added dependency; no others.
- Standalone `uv run` workflow preserved.
- Behavior features (headings 1-6, lists, tables, images, links, code, etc.)
  keep working; existing behavior tests pass (updated where they asserted direct
  formatting that now lives in a style).

## Requirements

### Requirement 1: Global default font (document default)

**User Story:** As a user, I want one place to set the font that everything
inherits, so most content shares a font without repeating it.

#### Acceptance Criteria

1. THEN the config SHALL have a top-level `font` key (the global default),
   defaulting to `"Aptos"`.
2. WHEN the document is built THEN `font` SHALL be written to the document
   default run properties (`docDefaults/rPrDefault`), so any run or style that
   does not specify a font inherits it.
3. WHEN an element's own `font` is null/absent THEN that element SHALL inherit
   the global default rather than being forced to a specific face.
4. THEN the global default SHALL be editable in Word (Manage Styles → Set
   Defaults) and visibly affect content that has no explicit font.

### Requirement 2: Generated named paragraph styles

**User Story:** As a user, I want headings and other blocks to use named styles,
so I can restyle all instances in Word by editing the style.

#### Acceptance Criteria

1. WHEN the document is built THEN the converter SHALL create named paragraph
   styles from the config: `Normal` (body), `Heading 1`-`Heading 6`, a
   blockquote style, a code-block style, and a note style.
2. THEN each generated style's properties (font, size, color, bold, italic,
   spacing, indent) SHALL come from the corresponding YAML keys.
3. WHEN content is rendered THEN paragraphs SHALL be assigned the matching style
   by name instead of receiving equivalent direct formatting.
4. WHEN a user edits a generated style in Word THEN all paragraphs using that
   style SHALL update accordingly.
5. WHEN a heading style is applied THEN its outline level SHALL be set in the
   style (so the navigation pane and bookmarks work) rather than stamped per
   paragraph.
6. THEN `Heading 1`-`Heading 6` SHALL default to font `"Aptos Display"` and body
   (`Normal`) SHALL inherit the global default (`"Aptos"`).

### Requirement 3: Character styles and font handling for code/links

**User Story:** As a user, I want code and links styled by named character
styles too, so they are consistent and editable.

#### Acceptance Criteria

1. THEN inline code SHALL use a generated character style (or the code paragraph
   style for code blocks) with the configured monospace font (default
   `"Consolas"`), not the global default font.
2. THEN hyperlinks SHALL use a generated `Hyperlink` character style with the
   configured color and underline.
3. WHEN code/link styles specify a font THEN that font SHALL win over the global
   default (code stays monospace even though everything else inherits Aptos).

### Requirement 4: Inline emphasis still works with styles

**User Story:** As a user, I want bold/italic/strikethrough to still render on
text that also carries a paragraph or character style.

#### Acceptance Criteria

1. WHEN inline markdown emphasis appears THEN bold/italic/strike SHALL apply as
   direct run properties layered on top of the paragraph/character style.
2. WHEN inline code appears within body text THEN it SHALL carry the code
   character style while surrounding text uses the paragraph style.
3. WHEN a heading or list item contains emphasis THEN both the style and the
   emphasis SHALL render.

### Requirement 5: Lists and tables

**User Story:** As a user, I want lists and tables to remain correct and, where
practical, editable.

#### Acceptance Criteria

1. WHEN lists render THEN they SHALL keep working numbering (bullets, restarting
   ordered numbers, one-level nesting, task glyphs) using list paragraph styles
   or the existing numbering approach.
2. WHEN list item text uses the body/list style THEN its font SHALL inherit the
   global default unless overridden.
3. WHEN tables render THEN their appearance (borders, header emphasis, margins)
   SHALL be preserved; table styling MAY remain direct or use a generated table
   style (implementation choice), but header text SHALL still inherit the global
   font.

### Requirement 6: Config, docs, example

**User Story:** As a user, I want the new `font` key and style behavior
documented.

#### Acceptance Criteria

1. THEN `DEFAULT_STYLE` SHALL include the top-level `font` and set heading fonts
   to `"Aptos Display"`, body to inherit (null), code to `"Consolas"`.
2. THEN `--dump-config` SHALL include the new `font` key.
3. THEN the README SHALL document the global `font` key, that most elements
   inherit it, and that the output uses named styles editable in Word.
4. THEN `style.example.yaml` and `example.docx` SHALL be regenerated.

## Out of Scope

- A full Word theme (`theme1.xml`) with major/minor font scheme; we set the
  document default run font, not a theme.
- Table styles as named Word table styles (may stay direct).
- Linked styles (paragraph+character pairs) beyond what is needed.
- Round-tripping edits made in Word back into YAML.
