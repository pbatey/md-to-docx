"""Convert Markdown to DOCX with YAML-configurable direct-formatting styling.

The full implementation lives in :mod:`md_to_docx.md_to_docx`, a single,
self-contained module that also works as a standalone ``uv run`` script (it
carries PEP 723 inline metadata). This package simply re-exports the public
entry points so the code can be installed and invoked as ``md-to-docx``.
"""

from .md_to_docx import (
    DEFAULT_STYLE,
    StyleConfig,
    StyleError,
    build_docx,
    dump_default_style,
    load_style,
    main,
    parse_markdown,
)

__all__ = [
    "DEFAULT_STYLE",
    "StyleConfig",
    "StyleError",
    "build_docx",
    "dump_default_style",
    "load_style",
    "main",
    "parse_markdown",
]

__version__ = "0.2.0"
