# Implementation Plan: YAML-Based Styling

- [x] 1. Add the default style config and loader
  - [x] 1.1 Define `DEFAULT_STYLE` (plain nested dict) covering page, body,
    headings 1-6, inline_code, code_block, blockquote, lists, table, note, hr,
    links. Tune values to match current output.
    - _Requirements: 2.1, 2.2, 2.3, 2.4_
  - [x] 1.2 Implement `load_style(path=None)`: deep-copy defaults, `yaml.safe_load`
    the file if given, deep-merge, ignore unknown keys, hard-error on invalid YAML.
    - _Requirements: 3.1, 3.2, 3.3, 3.4_
  - [x] 1.3 Implement a `StyleConfig` accessor with unit coercion (pt/inches,
    hex color) and per-field fallback+warn on malformed values.
    - _Requirements: 3.5, 4.*_
  - [x] 1.4 Add `pyyaml` to the PEP 723 metadata block and `pyproject.toml`.
    - _Requirements: 6.3_
  - [x] 1.5 Tests: defaults; deep merge; invalid YAML; unknown key; malformed
    field fallback.
    - _Requirements: 2.*, 3.*_

- [x] 2. Remove the DOCX template and its machinery
  - [x] 2.1 Delete `find_template`, `MD_TO_DOCX_TEMPLATE`, the template load +
    placeholder-removal block; start `build_docx` from a blank `Document()`.
    - _Requirements: 1.1, 1.2, 1.4_
  - [x] 2.2 Delete `templates/md-template.docx`, `src/md_to_docx/templates/`, and
    the wheel force-include packaging for the template.
    - _Requirements: 1.2, 6.4_
  - [x] 2.3 Remove/replace `find_template` tests.
    - _Requirements: 1.*_

- [x] 3. Direct-formatting foundation
  - [x] 3.1 Add `_apply_paragraph_format` and `_apply_run_format` helpers.
    - _Requirements: 4.2_
  - [x] 3.2 Add `_apply_page_setup(doc, style)`; make `_content_width_emu` and the
    table width use the configured content width (replace the 9360 constant).
    - _Requirements: 4.1_
  - [x] 3.3 Thread `style` through `add_formatted_text`, `_render_inline`,
    `_emit_run`, `_style_run`, `_add_image`, mirroring the `base_dir` plumbing.
    - _Requirements: 4.2, 4.4_

- [x] 4. Headings, body, note, meta, hr, links (direct)
  - [x] 4.1 Replace `_add_heading_safe` with `_add_heading` applying per-level
    font/size/color/bold/spacing and setting `w:outlineLvl`; keep bookmarks.
    - _Requirements: 4.3_
  - [x] 4.2 Apply body font/size/color/spacing/line-spacing to paragraphs and
    inline runs; apply note/meta/hr config.
    - _Requirements: 4.2, 4.10_
  - [x] 4.3 Apply link color/underline as direct run formatting; remove the
    `_ensure_hyperlink_style` style-injection dependency (links no longer need a
    named style).
    - _Requirements: 4.9_
  - [x] 4.4 Tests: heading run color/size + outlineLvl; body font; link run color.
    - _Requirements: 4.2, 4.3, 4.9_

- [x] 5. Inline code, code block, blockquote (config-driven)
  - [x] 5.1 Refactor `_style_run` to read inline-code font/size/color/fill.
    - _Requirements: 4.4_
  - [x] 5.2 Parameterize `_apply_code_block_box` and code-line/caption runs from
    `code_block.*`.
    - _Requirements: 4.5_
  - [x] 5.3 Parameterize `_apply_blockquote_bar` and blockquote indent/spacing.
    - _Requirements: 4.6_
  - [x] 5.4 Tests: code fill/padding override; blockquote bar color override;
    inline code font override.
    - _Requirements: 4.4, 4.5, 4.6_

- [x] 6. Lists and numbering without a template
  - [x] 6.1 Add `_ensure_numbering(doc, style)` creating code-defined bullet and
    decimal `abstractNum` entries at configured indents/glyphs.
    - _Requirements: 5.1, 5.2, 5.3_
  - [x] 6.2 Rewrite the `list`/`numbered_list` branches to use the code-defined
    abstracts via `_new_num_id` (delete `_find_abstract_num` and 8/6/7 magic).
    - _Requirements: 5.1, 5.2, 5.3_
  - [x] 6.3 Preserve task-list checkbox glyphs and per-block ordered restart.
    - _Requirements: 5.2, 5.4_
  - [x] 6.4 Tests: bullets render; two ordered lists both restart at 1; sub-bullet
    nesting; task glyphs present.
    - _Requirements: 5.1, 5.2, 5.3, 5.4_

- [x] 7. Table (direct borders/header/margins)
  - [x] 7.1 Rewrite `_apply_custom_table_style` to set direct `w:tblBorders`,
    header bold/fill, cell margins, width from `table.*`; keep `tblLook`,
    `tblLayout`, and `_fix_narrow_column_widths`.
    - _Requirements: 4.8_
  - [x] 7.2 Tests: direct borders present with configured color; header row
    bold/fill; column-width behavior preserved.
    - _Requirements: 4.8_

- [x] 8. CLI and config emission
  - [x] 8.1 Add `--style <path>` (and optional `MD_TO_DOCX_STYLE` env) and
    `--dump-config`; pass loaded style into `build_docx(style=...)`.
    - _Requirements: 6.1, 6.2_
  - [x] 8.2 Tests: `--dump-config` output re-parses to `DEFAULT_STYLE`; `--style`
    override reaches the output.
    - _Requirements: 6.1, 6.2_

- [x] 9. Docs, example, and full verification
  - [x] 9.1 Update README: config flag/env, full key reference with units/formats,
    remove template/`MD_TO_DOCX_TEMPLATE` docs; add an example YAML config.
    - _Requirements: 7.1, 7.2_
  - [x] 9.2 Regenerate `example.docx`; confirm it converts under defaults and looks
    materially the same as before.
    - _Requirements: 2.3, 7.3_
  - [x] 9.3 Run full test suite; verify installed CLI and standalone `uv run`
    (with pyyaml resolved from inline metadata) both convert successfully.
    - _Requirements: 6.5, all_
