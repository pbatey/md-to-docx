"""Tests for -L/--follow-links with -r."""

import os
import sys

import pytest

from md_to_docx.md_to_docx import discover_tree, main

pytestmark = pytest.mark.skipif(
    not hasattr(os, "symlink"), reason="symlinks unavailable"
)


def _tree(tmp_path):
    real = tmp_path / "real" / "sub"
    real.mkdir(parents=True)
    (real / "x.md").write_text("# X\n")
    root = tmp_path / "root"
    root.mkdir()
    (root / "r.md").write_text("# R\n")
    (root / "linked").symlink_to(tmp_path / "real", target_is_directory=True)
    (root / "self").symlink_to(".", target_is_directory=True)
    (tmp_path / "real" / "sub" / "loop").symlink_to(
        tmp_path / "real", target_is_directory=True
    )
    return root


def test_symlinked_dirs_ignored_by_default(tmp_path):
    assert discover_tree(_tree(tmp_path)) == ["r.md"]


def test_follow_links_enters_symlinked_dirs_and_skips_loops(tmp_path):
    assert discover_tree(_tree(tmp_path), follow_links=True) == [
        "linked/sub/x.md",
        "r.md",
    ]


def _run(args, monkeypatch, capsys):
    monkeypatch.setattr(sys, "argv", ["md-to-docx", *map(str, args)])
    code = 0
    try:
        main()
    except SystemExit as exc:
        code = exc.code if exc.code is not None else 0
    return code, capsys.readouterr().err


def test_cli_follow_links(tmp_path, monkeypatch, capsys):
    root = _tree(tmp_path)
    out = tmp_path / "out"
    code, _ = _run(["-r", "-L", "-o", out, root], monkeypatch, capsys)
    assert code == 0
    assert (out / "linked" / "sub" / "x.docx").exists()
    assert (out / "r.docx").exists()


def test_follow_links_requires_recursive(tmp_path, monkeypatch, capsys):
    code, err = _run(["-L", _tree(tmp_path)], monkeypatch, capsys)
    assert code == 2
    assert "requires -r" in err
