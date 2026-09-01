"""Convert Markdown to DOCX with template-based styling.

The full implementation lives in :mod:`md_to_docx.md_to_docx`, a single,
self-contained module that also works as a standalone ``uv run`` script (it
carries PEP 723 inline metadata). This package simply re-exports the public
entry points so the code can be installed and invoked as ``md-to-docx``.
"""

from .md_to_docx import (
    build_docx,
    find_template,
    main,
    parse_markdown,
)

__all__ = [
    "build_docx",
    "find_template",
    "main",
    "parse_markdown",
]

__version__ = "0.1.0"
