# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `argparse`-based command line interface with `--help` and `--version`.
- Document metadata written to the `.docx` core properties: the title is taken
  from the first level-1 heading (falling back to the input filename), the
  author from the current OS user, and the created date from the current time.
- Continuous-integration workflow renamed to `pr-validation.yml` (lint, format
  check, and tests).
- `ruff` lint configuration (`[tool.ruff.lint]`) and project-wide `ruff format`
  formatting.

### Changed

- The CLI now returns standard exit codes (argparse errors exit `2`); all
  existing flags (`--fetch-remote-images`, `--style`, `--dump-config`) and the
  `MD_TO_DOCX_STYLE` environment variable continue to work unchanged.

### Fixed

- `--dump-config` no longer crashes on non-UTF-8 consoles (e.g. Windows
  `cp1252`); it now writes UTF-8 output directly.

## [0.2.0]

### Added

- YAML-configurable styling: named Word styles generated from a set of built-in
  defaults, overridable via `--style` or the `MD_TO_DOCX_STYLE` environment
  variable (deep-merged over the defaults).
- `--dump-config` to print the effective default style as YAML.
- Release workflow that builds and publishes to PyPI via Trusted Publishing on a
  published GitHub Release, plus a manual TestPyPI dry-run workflow.
- MIT `LICENSE`.

### Changed

- Replaced the DOCX-template approach with direct, configurable formatting so no
  Word template file is required.

### Fixed

- Preserve bold/emphasis that spans a soft-wrapped line inside list items.

## [0.1.0]

### Added

- Initial Markdown-to-DOCX converter: headings, paragraphs with hard-wrap
  preservation, inline emphasis, lists, tables, fenced code blocks, blockquotes,
  images, links, and an optional table of contents.
- Installable package exposing the `md-to-docx` command, with tests and an
  example document.
