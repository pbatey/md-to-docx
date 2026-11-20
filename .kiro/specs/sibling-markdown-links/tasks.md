# Implementation Plan: Sibling Markdown Links

## Overview

All code goes in `src/md_to_docx/md_to_docx.py` (standard library only, so the PEP 723 standalone run keeps working). The plan builds the pure link layer first, then bookmark naming, base URL handling, wiki links and the `LinkContext`, then label rendering and hyperlink writing, then CLI and tree conversion, then documentation. Each step is wired into the next, so no code is left unused. Hyperlinks to other documents use an external relationship without a fragment plus a `w:anchor` attribute. Property tests use Hypothesis, one file per property named `tests/test_prop_pNN_*.py`, with `max_examples` of at least 100. Each property test carries the comment `Feature: sibling-markdown-links, Property N: <title>`.

Every task that edits `src/md_to_docx/md_to_docx.py` is in its own wave of the dependency graph, so no two such tasks run in parallel.

## Tasks

- [x] 1. Set up test dependencies and the pre-feature reference
  - [x] 1.1 Add `hypothesis>=6` to the `dev` dependency group in `pyproject.toml`
    - Keep the existing `pytest>=8` entry
    - _Requirements: 13.3, 13.11_

  - [x] 1.2 Capture golden fixtures of the pre-feature converter output
    - Must run before any edit to `md_to_docx.py`
    - Create a helper script or fixture module under `tests/` that converts a fixed set of Markdown documents (headings, duplicate and long headings, internal anchors, `[[#Heading]]`, external links, sibling links, wiki links) with the current code and stores the document XML under `tests/golden/`
    - Include a loader function that Property 12 uses to read the fixtures
    - _Requirements: 7.2, 2.9, 13.19_

- [x] 2. Add the pure link layer
  - [x] 2.1 Implement `percent_decode`, `SiblingTarget`, `parse_link_target`, `to_docx_target` and `encode_path`
    - Add the regex constants `_SCHEME_RE`, `_MD_EXT_RE`, `_PCT_RUN_RE`, `_PCT_SEQ_RE` and `_SAFE_PATH`
    - `parse_link_target` follows the seven steps in the design (angle brackets, `#`, scheme, leading `/`, backslash conversion, extension check) and never raises
    - `to_docx_target` removes one leading `./`, replaces the Markdown extension with `.docx`, and returns any other input unchanged
    - _Requirements: 1.2, 1.3, 1.7, 1.8, 1.9, 2.3, 2.4, 3.1, 3.3, 3.4, 3.5, 5.1, 5.2, 5.3, 5.4, 5.5, 9.1_

  - [x] 2.2 Write property test for sibling classification
    - Create `tests/test_prop_p01_classification.py`
    - **Property 1: Sibling classification**
    - **Validates: Requirements 1.9, 3.1, 3.3, 3.4, 3.5**

  - [x] 2.3 Write property test for path rewrite
    - Create `tests/test_prop_p02_path_rewrite.py`
    - **Property 2: Path rewrite preserves everything except the extension**
    - **Validates: Requirements 1.2, 1.7, 1.8, 9.1**

  - [x] 2.4 Write property test for URL-safe output
    - Create `tests/test_prop_p07_url_safe.py`
    - **Property 7: Rewritten target contains only URL-safe characters**
    - **Validates: Requirements 5.1, 5.3, 5.4**

  - [x] 2.5 Write property test for docx path round trip
    - Create `tests/test_prop_p08_round_trip.py`
    - **Property 8: Docx path round trip**
    - **Validates: Requirements 13.9, 13.11**

  - [x] 2.6 Write property test for docx path idempotence
    - Create `tests/test_prop_p09_idempotence.py`
    - **Property 9: Docx path idempotence**
    - **Validates: Requirements 13.10, 13.11**

- [x] 3. Make heading bookmarks index-independent and Word-valid
  - [x] 3.1 Implement `BOOKMARK_MAX`, `bookmark_name` and `bookmark_for_fragment`
    - `bookmark_name`: NFKD, drop combining marks, collapse runs outside `[A-Za-z0-9]` to `_`, strip `_`; empty slug gives `section`; non-empty slug with an empty body gives `h_` plus 8 hex digits of the slug's SHA-1; prefix `_toc_`
    - Names over 40 characters become the first 31 characters, `_`, and the first 8 hex digits of the SHA-1 of the full name, giving exactly 40 characters
    - `bookmark_for_fragment` returns `bookmark_name(slugify_heading(percent_decode(fragment)))`
    - Import `hashlib` and `unicodedata`
    - _Requirements: 2.2, 2.7, 2.8_

  - [x] 3.2 Change `build_heading_anchors` to use `bookmark_name` and add the `legacy` flag
    - Number repeated slugs GitHub-style (`install`, `install-1`) and pass the numbered slug through `bookmark_name`
    - Add the `used` set and the collision loop so names stay unique within one document, including slugs that sanitize to the same name
    - With `legacy=True` keep the exact `_toc_{idx}_{slug}` naming
    - Keep `anchors` mapping both slug and lowercase raw text to the bookmark so same-file links and the TOC still resolve
    - _Requirements: 2.2, 2.7, 2.8, 2.9, 4.3, 12.9_

  - [x] 3.3 Write property test for bookmark validity
    - Create `tests/test_prop_p04_bookmark_valid.py`
    - **Property 4: Bookmark names are valid Word bookmark names**
    - Cover `bookmark_name`, `bookmark_for_fragment` and every bookmark from `build_heading_anchors`
    - **Validates: Requirements 2.7, 13.17**

  - [x] 3.4 Write property test for truncation and uniqueness
    - Create `tests/test_prop_p05_truncation_unique.py`
    - **Property 5: Truncation is deterministic and names are unique**
    - **Validates: Requirements 2.8, 13.18**

  - [x] 3.5 Write property test for fragment and bookmark equality
    - Create `tests/test_prop_p06_fragment_bookmark.py`
    - **Property 6: Fragment equals the heading bookmark**
    - **Validates: Requirements 2.2, 4.3, 12.9**

- [x] 4. Implement base URL handling
  - [x] 4.1 Implement `normalize_base_url` and `resolve_absolute`
    - `normalize_base_url` raises `ValueError` naming the value for non-http(s) or hostless input, drops query and fragment (the CLI layer warns), encodes once without touching valid `%XX`, lowercases the scheme and strips the trailing `/`
    - `resolve_absolute(base, source_dir, docx_path)` returns `(url, climbed)`, discards unmatched `..` segments, and never puts a fragment in the URL
    - _Requirements: 8.4, 8.5, 8.6, 8.8, 8.9, 8.10, 8.11, 8.12, 12.3, 12.4, 12.6_

  - [x] 4.2 Write property test for base URL validation
    - Create `tests/test_prop_p15_base_validation.py`
    - **Property 15: Base URL validation**
    - **Validates: Requirements 8.12**

  - [x] 4.3 Write property test for base URL normalization
    - Create `tests/test_prop_p16_base_idempotent.py`
    - **Property 16: Base URL normalization is idempotent and never double-encodes**
    - **Validates: Requirements 8.10, 8.11, 13.14**

  - [x] 4.4 Write property test for the path-resolution model
    - Create `tests/test_prop_p17_absolute_model.py`
    - **Property 17: Absolute link matches the path-resolution model**
    - **Validates: Requirements 8.4, 8.5, 8.6, 12.3, 12.4, 13.13**

  - [x] 4.5 Write property test for trailing slash
    - Create `tests/test_prop_p18_trailing_slash.py`
    - **Property 18: Trailing slash does not matter**
    - **Validates: Requirements 8.8, 13.12**

  - [x] 4.6 Write property test for climbing above the base
    - Create `tests/test_prop_p19_climbing.py`
    - **Property 19: Climbing above the base root**
    - **Validates: Requirements 8.9, 12.6**

- [x] 5. Implement wiki parsing and the rewrite entry points
  - [x] 5.1 Implement `WikiLink` and `wiki_parts`
    - Cover `Note`, `Note|Label`, `Note#Heading`, `Note#Heading|Label`, `#Heading`
    - Return `None` for `[[]]` and `Note|` (empty label)
    - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.7_

  - [x] 5.2 Implement `_warn_stderr`, `LinkContext`, `Rewritten(target, anchor)`, `rewrite_link` and `rewrite_wiki_link`
    - `rewrite_link` returns `None` when `ctx` is `None`, disabled, or the target is not a sibling link
    - Otherwise it builds `target` from `to_docx_target` plus `encode_path`, or from `resolve_absolute`; the target never contains `#`
    - `anchor` is `bookmark_for_fragment(fragment)` when the fragment is non-empty, else `None`
    - Emit one `WARNING: <source_path>: link target not found: <target as written>` per occurrence for a missing file, and for an existing file outside `tree_files` in recursive mode; emit the climb warning when `resolve_absolute` reports `climbed`
    - `rewrite_wiki_link` appends `.docx` to non-Markdown note names, replaces `.md`/`.markdown`, returns the `Rewritten` value and the visible label, and emits no missing-file warning
    - _Requirements: 1.1, 1.6, 2.1, 2.3, 2.4, 2.5, 2.6, 4.1, 4.2, 4.3, 4.5, 4.6, 6.1, 6.2, 6.3, 6.4, 6.5, 8.3, 8.4, 8.6, 8.9, 8.13, 9.3, 12.1, 12.2_

  - [x] 5.3 Write property test for relative links without a base URL
    - Create `tests/test_prop_p13_relative.py`
    - **Property 13: Relative links without a base URL**
    - **Validates: Requirements 1.6, 8.3, 9.3**

- [x] 6. Render formatted labels and write document hyperlinks
  - [x] 6.1 Add `RunSpec`, `parse_label_runs` and extract `_apply_run_flags`
    - Extract the flag-to-run-properties code from `_render_inline` into `_apply_run_flags(run, flags)` and call it from both paths, leaving the existing output unchanged
    - `parse_label_runs(label)` reuses the `_render_inline` tokenizer and returns `RunSpec(text, flags)` items; nested links in a label are flattened to their text
    - An unformatted label gives one run
    - _Requirements: 1.4, 1.10, 4.2_

  - [x] 6.2 Implement `_add_document_hyperlink(paragraph, href, anchor, runs)`
    - Create the external relationship with `part.relate_to(href, RT.HYPERLINK, is_external=True)` and set `r:id` on the `w:hyperlink`
    - Set `w:anchor` to the bookmark name when `anchor` is not `None`
    - Write one `w:r` per `RunSpec`, each with `_apply_run_flags` properties plus the `Hyperlink` rStyle
    - _Requirements: 1.1, 1.5, 1.10, 2.6, 8.6_

  - [x] 6.3 Thread `link_ctx` through `add_formatted_text`, `_render_list_block` and `build_docx`
    - Tokenize `[[Note...]]` as `\x00WIKI\x00<inner>\x00` before the flattening regexes; tighten the regexes so `[[]]` and `[[Note|]]` stay as written; extend the token split regex to include `WIKI`
    - In the LINK branch: internal anchor first, then `rewrite_link` rendered with `parse_label_runs` and `_add_document_hyperlink`, then the `http(s)`/`mailto:` branch, then plain label
    - With `link_ctx` of `None` or disabled, take the pre-feature paths unchanged, including labels and wiki flattening
    - Call `build_heading_anchors` with `legacy=True` when `link_ctx` is `None` or disabled
    - Add the `announce` flag to `build_docx` for the `Saved:` line
    - _Requirements: 1.1, 1.4, 1.5, 1.10, 2.6, 2.9, 3.2, 3.6, 4.1, 4.2, 4.3, 4.4, 4.7, 4.8, 5.5, 7.2, 8.13, 8.14_

  - [x] 6.4 Write property test for anchor placement
    - Create `tests/test_prop_p03_anchor_placement.py`
    - **Property 3: Anchor placement**
    - Render through `add_formatted_text`; assert no relationship target contains `#` and `w:anchor` equals `bookmark_for_fragment(fragment)` or is absent
    - **Validates: Requirements 1.3, 2.1, 2.3, 2.4, 2.5, 2.6, 4.3, 8.6, 13.16**

  - [x] 6.5 Write property test for label formatting
    - Create `tests/test_prop_p10_label_formatting.py`
    - **Property 10: Link labels keep their inline formatting**
    - **Validates: Requirements 1.1, 1.4, 1.5, 1.10, 4.2, 13.15**

  - [x] 6.6 Write property test for non-sibling links
    - Create `tests/test_prop_p11_passthrough.py`
    - **Property 11: Non-sibling links render as before**
    - **Validates: Requirements 3.6, 8.14**

  - [x] 6.7 Write property test for the opt-out output
    - Create `tests/test_prop_p12_opt_out_legacy.py`
    - **Property 12: `--no-sibling-links` reproduces the pre-feature output**
    - Compare against the golden fixtures from task 1.2 and the `link_ctx=None` path, with and without a base URL
    - **Validates: Requirements 7.2, 8.13, 2.9, 13.19**

  - [x] 6.8 Write property test for wiki link rewrite
    - Create `tests/test_prop_p14_wiki.py`
    - **Property 14: Wiki link rewrite**
    - **Validates: Requirements 4.1, 4.2, 4.3, 4.5, 4.6, 4.7**

  - [x] 6.9 Write property test for missing-target warnings
    - Create `tests/test_prop_p20_warnings.py`
    - **Property 20: One warning per missing target occurrence**
    - **Validates: Requirements 6.1, 6.2, 6.3, 6.5, 12.1**

  - [x] 6.10 Write unit tests for link rendering
    - Create `tests/test_sibling_links.py`
    - Cover every worked example in the design's Data Models table, exactly one `w:hyperlink` with the `Hyperlink` style, an external relationship target without `#`, and `w:anchor` set to the bookmark, with and without `--link-base-url` (1.1, 1.5, 2.6)
    - A label with bold, plain and code runs gives one hyperlink with the matching runs; `[[a.md|*Label*]]` gives an italic run (1.4, 1.10, 4.2)
    - A heading over 40 characters and two headings with a long common prefix give 40-character, distinct, repeatable names, and a link by full text or slug resolves to the same name (2.7, 2.8)
    - `[[#Heading]]`, `[[#Missing]]`, `[[]]`, `[[Note|]]`, unclosed `[[Note`, angle-bracket targets with and without `>`, and non-sibling links unchanged
    - Heading bookmarks with links enabled contain no index and are unique; with legacy naming they keep `_toc_{idx}_{slug}`; existing internal-link tests still pass
    - _Requirements: 1.1, 1.4, 1.5, 1.10, 2.6, 2.7, 2.8, 2.9, 3.1, 3.2, 3.3, 3.4, 3.5, 4.4, 4.7, 4.8, 5.2, 5.5, 6.1, 13.3, 13.15, 13.16, 13.17, 13.18_

- [x] 7. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 8. Add CLI options and tree conversion
  - [x] 8.1 Add the new options to `build_parser`
    - `-r/--recursive`, `-o/--outdir DIR`, `--no-sibling-links`, and `--link-base-url URL` (wrapper turns `ValueError` into `argparse.ArgumentTypeError`)
    - Change the `input` help to "file or directory" and note on `output` that it cannot be combined with `-r` or `--outdir`
    - _Requirements: 7.1, 7.4, 7.5, 8.1, 8.2, 8.12, 10.1, 10.18, 11.1, 11.12, 11.13_

  - [x] 8.2 Implement `is_skipped_dir`, `discover_tree`, `output_relpath`, `plan_outputs`, `Plan` and `Collision`
    - `discover_tree` uses `os.walk(followlinks=False)`, prunes skipped directories in place, selects `.md`/`.markdown` regular files or file symlinks case-insensitively, skips hidden files, and returns sorted posix Relative_Paths
    - `plan_outputs` maps to `outdir/<rel>.docx` or beside the source, and drops stem collisions, reporting each group with every file named
    - _Requirements: 9.5, 10.2, 10.8, 10.9, 10.10, 10.11, 11.2, 11.4, 11.10, 11.11_

  - [x] 8.3 Implement `output_lock_error`, `resolve_mode`, `convert_file`, `convert_tree` and wire them into `main`
    - Split `check_output_writable` into `output_lock_error` plus the existing exiting wrapper
    - `resolve_mode` applies the validation table in the design (directory without `-r`, `-r` with output, `--outdir` with output, unusable `--outdir`, `--dump-config` precedence, `-r` with a file)
    - `main` builds a `LinkContext` for every file from `--no-sibling-links` (`enabled=False`) and `--link-base-url`, and warns once when the base URL had a query or fragment dropped
    - `convert_tree` loads the style once, builds a `LinkContext` per file with `source_rel_dir` and `tree_files`, writes via a temp file and `os.replace`, prints one progress line per file and the summary line, continues after failures, and returns 1 if any failed
    - Single_File_Mode keeps today's flow and never touches the linked `.docx`
    - Pass `--style`, `--fetch-remote-images`, `--no-sibling-links` and `--link-base-url` to every file
    - _Requirements: 7.2, 7.3, 8.13, 9.4, 10.3, 10.4, 10.5, 10.6, 10.7, 10.12, 10.13, 10.14, 10.15, 10.16, 10.17, 10.19, 10.20, 11.3, 11.5, 11.6, 11.7, 11.8, 11.9, 12.1, 12.2, 12.3, 12.4, 12.8_

  - [x] 8.4 Write property test for tree discovery
    - Create `tests/test_prop_p21_discovery.py`
    - **Property 21: Tree discovery**
    - **Validates: Requirements 10.8, 10.9, 10.11**

  - [x] 8.5 Write property test for the mirrored output set
    - Create `tests/test_prop_p22_mirrored.py`
    - **Property 22: Mirrored output set**
    - **Validates: Requirements 9.5, 10.2, 11.2, 11.4, 13.7**

  - [x] 8.6 Write property test for relative links resolving to the linked output
    - Create `tests/test_prop_p23_link_resolves.py`
    - **Property 23: Relative links resolve to the linked output** (model tree)
    - **Validates: Requirements 12.7**

  - [x] 8.7 Write property test for stem collisions
    - Create `tests/test_prop_p24_collisions.py`
    - **Property 24: Stem collisions are detected**
    - **Validates: Requirements 11.11**

  - [x] 8.8 Write CLI and recursive-mode tests
    - Create `tests/test_cli_tree.py`, using `main()` with `tmp_path`
    - CLI: `--no-sibling-links` accepted, rejected with a value, and in `--help`; `--link-base-url` valid, invalid and in `--help`; `-r`, `-o`, `--outdir` in `--help`; `-o DIR` identical to `--outdir DIR`; `--no-sibling-links` alone and with `--link-base-url` produce the golden pre-feature document
    - Single file does not create, read or modify a sentinel linked `.docx` (9.4)
    - Tree: nested directories, each skipped directory name, a directory symlink cycle, case-insensitive extensions, sorted progress lines, directory without `-r`, empty tree, `-r` with output, `--outdir` with output, `--outdir` inside the Input_Root, `--outdir` with a single file, `--outdir` that is a file, missing directories created, `a/b/x.md` linking `../y.md` giving `<base>/a/y.docx`, `--dump-config` ignoring `-r`, stem collision, and partial failure (invalid UTF-8 file) with exit status 1, other files written and no partial output
    - _Requirements: 7.1, 7.4, 7.5, 8.1, 8.2, 8.12, 8.13, 9.4, 10.1 to 10.20, 11.1 to 11.13, 12.1 to 12.9, 13.3, 13.6, 13.8, 13.19_

- [x] 9. Documentation
  - [x] 9.1 Update `README.md`
    - Describe the link rewriting rules, `--no-sibling-links` (default behavior when omitted, and that it restores the pre-feature output including legacy bookmark names), the same-relative-layout requirement, `--link-base-url` with a SharePoint example and the statement that relative links work only in local or synced folders
    - Document `-r/--recursive`, `-o/--outdir`, the skipped directories, the mirrored layout, the exit status on partial failure, and the workflow `md-to-docx -r docs/ --outdir out/ --link-base-url https://.../docs/`
    - Update the Usage block and the standalone-options sentence
    - _Requirements: 13.1, 13.4_

  - [x] 9.2 Update `CHANGELOG.md`
    - Add entries under the unreleased section for sibling link rewriting and `--no-sibling-links`, and for `-r/--recursive` and `--outdir`
    - _Requirements: 13.2, 13.5_

- [x] 10. Final checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 11. Add the `--sharepoint` option
  - [x] 11.1 Implement `sharepoint_folder_url`, `is_sharepoint_view_url` and `LinkContext.link_suffix`
    - `sharepoint_folder_url` matches `(?i)/Forms/[^/]+\.aspx$`, returns `scheme://netloc` plus the decoded `id` path, or the library root when `id` is missing, and `None` for other URLs
    - Add `link_suffix: str = ""` to `LinkContext`; `rewrite_link` and `rewrite_wiki_link` append it to the Absolute_Link only when `base_url` is set
    - _Requirements: 14.4, 14.5, 14.6, 14.7, 14.9_

  - [x] 11.2 Add `--sharepoint`/`--sp` to `build_parser` and wire it into `resolve_mode` and `main`
    - `store_true`; `resolve_mode` calls `parser.error` when it is given without `--link-base-url`
    - `main` converts a view URL with `sharepoint_folder_url` before `normalize_base_url`, skips the dropped-query Warning for it, and sets `link_suffix="?web=1"` on every `LinkContext` in both modes
    - Without `--sharepoint`, write one hint Warning when `--link-base-url` is a SharePoint_View_URL
    - `--no-sibling-links` still wins
    - _Requirements: 14.1, 14.2, 14.3, 14.6, 14.8, 14.10_

  - [x] 11.3 Write property test for the SharePoint suffix
    - Create `tests/test_prop_p25_sharepoint.py`
    - **Property 25: SharePoint suffix**
    - **Validates: Requirements 14.4, 14.5, 14.9, 14.13**

  - [x] 11.4 Write SharePoint unit and CLI tests
    - Create `tests/test_sharepoint.py`
    - `sharepoint_folder_url`: view URL with `id` (the `Operations team/IVR Knowledge Base/TEST` example gives `https://aegonna.sharepoint.com/sites/TTWFG/Shared%20Documents/Operations%20team/IVR%20Knowledge%20Base/TEST`), without `id` (library root), other `.aspx` pages and plain folder URLs give `None`
    - CLI: `--sharepoint` and `--sp` in `--help`; `--sp` identical to `--sharepoint`; `--sharepoint=yes` and `--sharepoint` without `--link-base-url` exit non-zero with no file created; a tree run writes `.../a.docx?web=1` with the heading in `w:anchor`; no dropped-query Warning for a converted view URL; hint Warning for a view URL without `--sharepoint`; `--no-sibling-links` with `--sharepoint` produces the golden pre-feature document
    - _Requirements: 14.1, 14.2, 14.3, 14.6, 14.7, 14.8, 14.10, 14.13_

  - [x] 11.5 Update `README.md`
    - Document `--sharepoint`/`--sp`, an example with a pasted `.../Forms/AllItems.aspx?id=...` URL, and that Word for the web may open the document without moving to the heading
    - _Requirements: 14.11_

  - [x] 11.6 Update `CHANGELOG.md`
    - Add an entry under the unreleased section for `--sharepoint`
    - _Requirements: 14.12_

- [x] 12. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

## Notes

- Tasks marked with `*` are optional and can be skipped for a faster MVP
- Each task references specific requirements for traceability
- Property tests validate universal correctness properties; unit tests validate examples, CLI wiring and filesystem behavior
- Task 1.2 must run before any edit to `md_to_docx.py`, because it records the pre-feature output that Property 12 compares against
- Cross-document anchors (an external relationship plus `w:anchor`) should be verified manually by opening a converted pair and following a link to a heading in the other file in both Word desktop and Word for the web (SharePoint). This is not automated, and Word for the web may open the document without moving to the bookmark

## Task Dependency Graph

```json
{
  "waves": [
    { "id": 0, "tasks": ["1.1", "1.2"] },
    { "id": 1, "tasks": ["2.1"] },
    { "id": 2, "tasks": ["2.2", "2.3", "2.4", "2.5", "2.6", "3.1"] },
    { "id": 3, "tasks": ["3.2"] },
    { "id": 4, "tasks": ["3.3", "3.4", "3.5", "4.1"] },
    { "id": 5, "tasks": ["4.2", "4.3", "4.4", "4.5", "4.6", "5.1"] },
    { "id": 6, "tasks": ["5.2"] },
    { "id": 7, "tasks": ["5.3", "6.1"] },
    { "id": 8, "tasks": ["6.2"] },
    { "id": 9, "tasks": ["6.3"] },
    { "id": 10, "tasks": ["6.4", "6.5", "6.6", "6.7", "6.8", "6.9", "6.10", "8.1"] },
    { "id": 11, "tasks": ["8.2"] },
    { "id": 12, "tasks": ["8.3", "8.4", "8.5", "8.6", "8.7"] },
    { "id": 13, "tasks": ["8.8", "9.1", "9.2"] },
    { "id": 14, "tasks": ["11.1"] },
    { "id": 15, "tasks": ["11.2", "11.3"] },
    { "id": 16, "tasks": ["11.4", "11.5", "11.6"] }
  ]
}
```
