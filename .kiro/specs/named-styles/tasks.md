# Implementation Plan: Generate Named Word Styles from YAML

- [ ] 1. Config: global font + inherit defaults
  - [ ] 1.1 Add top-level `font: "Aptos"` to `DEFAULT_STYLE`; set
    `headings.<n>.font: "Aptos Display"`, `body.font: null`.
    - _Requirements: 1.1, 2.6, 6.1_
  - [ ] 1.2 Confirm `load_style`/`StyleConfig` treat null/absent font as inherit.
    - _Requirements: 1.3_

- [ ] 2. Document default font
  - [ ] 2.1 Add `_set_default_font(doc, font)` writing `docDefaults/rPrDefault/
    rPr/rFonts` (ascii/hAnsi/cs).
    - _Requirements: 1.2, 1.4_
  - [ ] 2.2 Call it at the start of `build_docx`.
    - _Requirements: 1.2_

- [ ] 3. Paragraph style generation
  - [ ] 3.1 Add `_define_paragraph_style(...)` (font optional=inherit, size,
    color, bold, italic, spacing, indent, outline level, base).
    - _Requirements: 2.1, 2.2, 2.5_
  - [ ] 3.2 Generate `Normal` (body, font inherit), `Heading 1`-`Heading 6`
    (Aptos Display + config), `MdQuote`, `MdCode`, `MdNote` from config in a
    `_build_styles(doc, sc)` routine.
    - _Requirements: 2.1, 2.2, 2.6_

- [ ] 4. Character styles (code + hyperlink)
  - [ ] 4.1 Add `_define_char_style(...)`; generate `MdCodeChar` (Consolas +
    size, optional color/fill) and `Hyperlink` (links color + underline).
    - _Requirements: 3.1, 3.2, 3.3_

- [ ] 5. Apply styles in the render loop
  - [ ] 5.1 Headings: use `Heading N` style; drop direct run font/size/color;
    keep bookmark. (Replace `_add_heading` direct formatting.)
    - _Requirements: 2.3, 2.4, 2.5_
  - [ ] 5.2 Body paragraph uses `Normal`; blockquote `MdQuote`; code block
    `MdCode`; note `MdNote`. Move indent/spacing into styles; keep bar/box
    borders (direct or in style pPr).
    - _Requirements: 2.3, 5.2_
  - [ ] 5.3 `_style_run`: stop setting font/size/color on normal runs (inherit);
    keep bold/italic/strike; apply `MdCodeChar` for code runs.
    - _Requirements: 3.1, 4.1, 4.2_
  - [ ] 5.4 Code block line runs use the code font/size (or `MdCodeChar`).
    - _Requirements: 3.1, 3.3_
  - [ ] 5.5 Hyperlinks reference the `Hyperlink` character style (revert direct
    color/underline to style-based).
    - _Requirements: 3.2_
  - [ ] 5.6 List item paragraphs use `Normal` + numPr (inherit font); keep
    numbering/glyphs/nesting.
    - _Requirements: 5.1, 5.2_

- [ ] 6. Tests
  - [ ] 6.1 docDefaults font = Aptos; body run has no direct font; overriding
    top-level `font` changes docDefaults.
    - _Requirements: 1.2, 1.3_
  - [ ] 6.2 `Heading 1` style exists (Aptos Display, config size/color, outline
    level); heading paragraph uses the style by name; no direct run color.
    - _Requirements: 2.2, 2.4, 2.5, 2.6_
  - [ ] 6.3 Inline code uses code char style/monospace; code block runs Consolas.
    - _Requirements: 3.1, 3.3_
  - [ ] 6.4 Hyperlink runs reference `Hyperlink` char style; style has color +
    underline.
    - _Requirements: 3.2_
  - [ ] 6.5 Emphasis still sets bold/italic on styled paragraphs; lists/tables
    regression green; update prior direct-formatting assertions to style-based.
    - _Requirements: 4.1, 4.3, 5.1_

- [ ] 7. Docs, example, verification
  - [ ] 7.1 README: document top-level `font`, inheritance, and that output uses
    named styles editable in Word.
    - _Requirements: 6.3_
  - [ ] 7.2 Regenerate `style.example.yaml` (now includes `font`) and
    `example.docx`.
    - _Requirements: 6.2, 6.4_
  - [ ] 7.3 Full suite + installed CLI + standalone `uv run` verification.
    - _Requirements: all_
