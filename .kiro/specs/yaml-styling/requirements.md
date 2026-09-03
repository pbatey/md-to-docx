# Requirements: YAML-Based Styling (Replace the DOCX Template)

## Introduction

Today the converter styles output by loading a Word template
(`templates/md-template.docx`) as the base document: it supplies the named
styles (Heading 1-6, Normal, Quote, List Bullet/Number/Bullet 2), the list
numbering definitions, the table style (PlainTable2), and the page geometry.
This binary template is hard to keep consistent with intent and it is unclear
which parts are "styles" versus "sample content."

This feature replaces the `.docx` template with a **YAML styling config** that
declares plain, human-readable attributes (fonts, sizes, colors, indents,
spacing, borders, shading). Styling is applied as **direct formatting** on runs,
paragraphs, and tables rather than via named Word styles or a template's
numbering part. The built-in defaults live in the Python code, so **no config
file is needed unless the user wants to change something.**

## Goals

- Remove the `.docx` template dependency entirely.
- Make styling legible and diff-able: every knob is a named attribute in YAML.
- Keep working defaults in code; a YAML file is optional and overrides defaults.
- Preserve the current visual output as closely as practical (parity target).
- Preserve the single-file `uv run` standalone workflow.

## Constraints (apply to all requirements)

- Styling is applied as **direct formatting**. The output must not depend on any
  named paragraph/character/table style being present, nor on a template's
  numbering part.
- `pyyaml` is an accepted new dependency (added to both the PEP 723 inline
  metadata block and `pyproject.toml`).
- Defaults defined in code must reproduce today's appearance closely enough that
  a converted document looks materially the same (fonts, sizes, colors, indents,
  spacing, code box, blockquote bar, table look).
- YAML config is fully optional: with no config and no template, conversion
  still produces a styled document from code defaults.
- Existing features (headings 1-6, images, task lists, links, code language,
  etc.) keep working; behavior tests continue to pass.
- Malformed or partial YAML degrades gracefully (clear error for invalid YAML;
  unknown keys ignored or warned; missing keys fall back to defaults).

## Requirements

### Requirement 1: Remove the DOCX template

**User Story:** As a maintainer, I want the `.docx` template gone, so styling is
no longer hidden in a binary file.

#### Acceptance Criteria

1. WHEN the converter runs THEN it SHALL start from a blank `Document()` and
   apply all styling via direct formatting.
2. THEN `templates/md-template.docx` and `src/md_to_docx/templates/` SHALL be
   removed, along with `find_template()`, `MD_TO_DOCX_TEMPLATE`, and the
   template package-data packaging.
3. WHEN no YAML config is supplied THEN conversion SHALL still succeed using
   built-in code defaults.
4. THEN no code path SHALL require a named Word style or a template numbering
   part to exist.

### Requirement 2: Built-in default style config in code

**User Story:** As a user, I want sensible defaults baked in, so I do not need a
config file for normal use.

#### Acceptance Criteria

1. THEN the code SHALL define a complete default style configuration covering
   every element the renderer styles (see Requirement 4).
2. WHEN no YAML file is provided THEN the defaults SHALL be used unchanged.
3. THEN the default values SHALL reproduce the current appearance as closely as
   practical (body font/size, heading sizes, code box fill/padding, blockquote
   bar color/width, table look, list indents, hyperlink color).
4. THEN the default configuration SHALL be expressible as YAML (round-trippable),
   so users can dump defaults as a starting point.

### Requirement 3: Load and merge a YAML config

**User Story:** As a user, I want to point the converter at a YAML file to
override styling, so I can customize without editing code or Word.

#### Acceptance Criteria

1. WHEN a YAML config path is provided (CLI flag and/or env var) THEN its values
   SHALL override the corresponding defaults.
2. WHEN the YAML omits a key THEN the default for that key SHALL be used
   (deep merge, not wholesale replacement).
3. WHEN the YAML is syntactically invalid THEN the converter SHALL exit with a
   clear error message identifying the file and the parse problem.
4. WHEN the YAML contains an unknown key THEN the converter SHALL ignore it (and
   MAY emit a warning) rather than fail.
5. WHEN a value has the wrong type or an invalid form (e.g. a malformed color)
   THEN the converter SHALL report a clear, specific error OR fall back to the
   default for that key (behavior defined per field in the design).
6. THEN a way to emit the effective/default config as YAML SHALL be provided
   (e.g. a `--dump-config` flag) so users can see and copy the schema.

### Requirement 4: Style knobs per element (direct formatting)

**User Story:** As a user, I want to control each element's appearance through
named attributes, so styling is obvious and consistent.

#### Acceptance Criteria

1. **Page**: WHEN configured THEN page size and margins SHALL be applied to the
   document section, and downstream width calculations (image scaling, table
   width) SHALL use the configured content width.
2. **Body/Normal text**: THEN font family, size, color, and paragraph spacing
   (space before/after, line spacing) SHALL be configurable and applied to
   paragraph and inline runs by default.
3. **Headings (levels 1-6)**: THEN each level SHALL support font family, size,
   color, bold/italic, and space before/after; applied as direct formatting so
   no "Heading N" style is required. Headings SHALL keep their bookmarks and
   outline behavior needed for internal links/TOC.
4. **Inline**: THEN bold, italic, and strikethrough SHALL remain markdown-driven;
   inline `code` SHALL support a configurable monospace font family, size, and
   optional color/shading.
5. **Code block**: THEN monospace font family, code size, background fill,
   interior padding, vertical spacing, and the optional language-caption style
   (size, color, italic) SHALL be configurable.
6. **Blockquote**: THEN the left-bar color, thickness, and gap, plus indent and
   spacing, SHALL be configurable.
7. **Lists**: THEN bullet glyph(s) per level, ordered-list number format,
   per-level indent, item font/size, and inter-item spacing SHALL be
   configurable; numbering SHALL be produced without a template numbering part
   and ordered lists SHALL restart per block (parity with today).
8. **Table**: THEN border style/width/color, header row emphasis (bold and/or
   fill), cell margins, and overall width behavior SHALL be configurable; the
   proportional column-width behavior SHALL be preserved.
9. **Links**: THEN hyperlink text color and underline SHALL be configurable and
   applied as direct run formatting (no dependency on a "Hyperlink" style).
10. **Note / meta / hr / horizontal spacing**: THEN each SHALL expose the
    attributes it currently uses (note indent/italic/size; meta spacing; hr
    spacing or optional rule line).

### Requirement 5: Lists and numbering without a template

**User Story:** As a maintainer, I want list numbering to work from a blank
document, so removing the template does not break ordered/bulleted lists.

#### Acceptance Criteria

1. WHEN a bullet list renders THEN bullets SHALL appear at the configured indent
   without relying on template-provided abstractNumIds or the magic numIds
   (8/6/7) used today.
2. WHEN multiple ordered lists appear THEN each SHALL restart at 1 (current
   behavior), implemented via numbering definitions the code creates itself or
   via an equivalent direct-formatting approach.
3. WHEN a numbered list has sub-bullets THEN the one-level nesting SHALL still
   render at the configured deeper indent.
4. WHEN task-list items render THEN the checkbox glyphs SHALL still appear ahead
   of the item text.

### Requirement 6: CLI and packaging

**User Story:** As a user, I want to select a config from the command line, so I
can switch styling per run.

#### Acceptance Criteria

1. THEN a CLI option SHALL accept a YAML config path (e.g. `--style config.yaml`)
   and an env var equivalent MAY be supported.
2. THEN `--dump-config` (or similar) SHALL print the effective default config as
   YAML and exit.
3. THEN `pyproject.toml` and the PEP 723 metadata SHALL declare `pyyaml`.
4. THEN packaging SHALL no longer ship the `.docx` template as package data.
5. THEN the standalone `uv run` workflow SHALL continue to work (uv resolves
   `pyyaml` from the inline metadata).

### Requirement 7: Documentation and example

**User Story:** As a user, I want the styling attributes documented with an
example, so I know what I can change.

#### Acceptance Criteria

1. THEN the README SHALL document the config location/flag, the full set of
   style keys, their units, and value formats (colors as hex, sizes in points,
   indents in inches or points).
2. THEN an example YAML config SHALL be provided showing common overrides.
3. THEN `example.md` SHALL still convert and visually demonstrate the styled
   output under the defaults.

## Out of Scope

- Generating/compiling named Word styles or a Word theme from YAML (we apply
  direct formatting only).
- Per-document front-matter styling overrides (config is global per run).
- Multiple simultaneous themes within one document.
- Round-tripping an existing `.docx` template into YAML (the old template is
  removed, not converted).
- Nested lists beyond the existing one level.
