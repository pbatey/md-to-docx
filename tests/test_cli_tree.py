"""CLI and recursive-mode tests, driven through ``main()``.

Covers the sibling-link options, -r/-o/--outdir handling, tree discovery and
failure behaviour.
"""

import os
import re
import sys
import zipfile

import pytest
from golden_fixtures import load_golden

from md_to_docx.md_to_docx import main

GOLDEN = load_golden()


def run(args, monkeypatch, capsys):
    """Invoke main() with ``args``; return (exit_status, stdout, stderr)."""
    monkeypatch.setattr(sys, "argv", ["md-to-docx", *map(str, args)])
    code = 0
    try:
        main()
    except SystemExit as exc:
        code = exc.code if exc.code is not None else 0
    cap = capsys.readouterr()
    return code, cap.out, cap.err


def write(path, text="# T\n"):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


def part(docx, name):
    with zipfile.ZipFile(docx) as z:
        return z.read(name).decode("utf-8")


def doc_xml(docx):
    return part(docx, "word/document.xml")


def norm(xml):
    seen = {}

    def sub(m):
        return f'{m.group(1)}w:id="{seen.setdefault(m.group(2), str(len(seen) + 1))}"'

    return re.sub(r'(<w:bookmark(?:Start|End) )w:id="(\d+)"', sub, xml)


def targets(docx):
    return re.findall(r'Target="([^"]*)"', part(docx, "word/_rels/document.xml.rels"))


def docx_files(root):
    return sorted(p.relative_to(root).as_posix() for p in root.rglob("*.docx"))


# ----------------------------- option parsing ------------------------------ #


def test_help_lists_new_options(monkeypatch, capsys):
    code, out, _ = run(["--help"], monkeypatch, capsys)
    assert code == 0
    for opt in (
        "--no-sibling-links",
        "--link-base-url",
        "-r",
        "--recursive",
        "-o",
        "--outdir",
    ):
        assert opt in out


def test_no_sibling_links_accepted(tmp_path, monkeypatch, capsys):
    src = write(tmp_path / "a.md")
    code, _, _ = run(
        ["--no-sibling-links", src, tmp_path / "a.docx"], monkeypatch, capsys
    )
    assert code == 0
    assert (tmp_path / "a.docx").exists()


def test_no_sibling_links_rejects_value(tmp_path, monkeypatch, capsys):
    src = write(tmp_path / "a.md")
    code, _, err = run(
        ["--no-sibling-links=yes", src, tmp_path / "a.docx"], monkeypatch, capsys
    )
    assert code == 2
    assert "--no-sibling-links" in err
    assert not (tmp_path / "a.docx").exists()


def test_link_base_url_valid(tmp_path, monkeypatch, capsys):
    src = write(tmp_path / "a.md", "[x](other.md)\n")
    code, _, _ = run(
        ["--link-base-url", "https://example.com/docs/", src, tmp_path / "a.docx"],
        monkeypatch,
        capsys,
    )
    assert code == 0
    # other.md is not in a tree run, so it stays relative in single-file mode.
    assert "other.docx" in " ".join(targets(tmp_path / "a.docx"))


@pytest.mark.parametrize("bad", ["ftp://example.com", "example.com", "", "https://"])
def test_link_base_url_invalid(bad, tmp_path, monkeypatch, capsys):
    src = write(tmp_path / "a.md")
    code, _, err = run(
        ["--link-base-url", bad, src, tmp_path / "a.docx"], monkeypatch, capsys
    )
    assert code == 2
    assert "--link-base-url" in err
    assert not (tmp_path / "a.docx").exists()


def test_short_o_equals_outdir(tmp_path, monkeypatch, capsys):
    root = tmp_path / "in"
    write(root / "a" / "x.md", "# X\n")
    write(root / "y.md", "# Y\n")
    run(["-r", root, "-o", tmp_path / "o1"], monkeypatch, capsys)
    run(["-r", root, "--outdir", tmp_path / "o2"], monkeypatch, capsys)
    f1, f2 = docx_files(tmp_path / "o1"), docx_files(tmp_path / "o2")
    assert f1 == f2 == ["a/x.docx", "y.docx"]
    for f in f1:
        assert norm(doc_xml(tmp_path / "o1" / f)) == norm(doc_xml(tmp_path / "o2" / f))


# --------------------------- opt-out / golden ------------------------------ #


@pytest.mark.parametrize("extra", [[], ["--link-base-url", "https://example.com/docs"]])
def test_no_sibling_links_matches_golden(extra, tmp_path, monkeypatch, capsys):
    for name, (md, expected) in GOLDEN.items():
        src = write(tmp_path / f"{name}.md", md)
        out = tmp_path / f"{name}.docx"
        code, _, _ = run(["--no-sibling-links", *extra, src, out], monkeypatch, capsys)
        assert code == 0, name
        assert norm(doc_xml(out)) == norm(expected), name


def test_single_file_ignores_linked_docx(tmp_path, monkeypatch, capsys):
    src = write(tmp_path / "a.md", "[b](b.md) and [c](c.md)\n")
    sentinel = tmp_path / "b.docx"
    sentinel.write_bytes(b"SENTINEL")
    os.utime(sentinel, (1_000_000, 1_000_000))
    code, _, _ = run([src, tmp_path / "a.docx"], monkeypatch, capsys)
    assert code == 0
    assert sentinel.read_bytes() == b"SENTINEL"
    assert sentinel.stat().st_mtime == 1_000_000
    assert not (tmp_path / "c.docx").exists()


# ------------------------------- tree mode --------------------------------- #


def test_nested_tree_in_place_and_sorted_progress(tmp_path, monkeypatch, capsys):
    root = tmp_path / "in"
    for rel in ("z.md", "a/b/c.md", "a/m.md", "b.md"):
        write(root / rel)
    code, out, _ = run(["-r", root], monkeypatch, capsys)
    assert code == 0
    assert docx_files(root) == ["a/b/c.docx", "a/m.docx", "b.docx", "z.docx"]
    lines = [ln for ln in out.splitlines() if " -> " in ln]
    assert [ln.split(" -> ")[0] for ln in lines] == [
        "a/b/c.md",
        "a/m.md",
        "b.md",
        "z.md",
    ]
    assert "Converted 4 file(s), 0 failed" in out


@pytest.mark.parametrize(
    "skipped", [".git", ".venv", "node_modules", "venv", "__pycache__", ".hidden"]
)
def test_skipped_directories(skipped, tmp_path, monkeypatch, capsys):
    root = tmp_path / "in"
    write(root / "keep.md")
    write(root / skipped / "skip.md")
    write(root / "sub" / skipped / "deep.md")
    code, _, _ = run(["-r", root], monkeypatch, capsys)
    assert code == 0
    assert docx_files(root) == ["keep.docx"]


def test_directory_symlink_cycle(tmp_path, monkeypatch, capsys):
    root = tmp_path / "in"
    write(root / "sub" / "x.md")
    try:
        os.symlink(root, root / "sub" / "loop", target_is_directory=True)
    except (OSError, NotImplementedError):
        pytest.skip("symlinks unsupported")
    code, _, _ = run(["-r", root], monkeypatch, capsys)
    assert code == 0
    assert docx_files(root) == ["sub/x.docx"]


def test_case_insensitive_extensions(tmp_path, monkeypatch, capsys):
    root = tmp_path / "in"
    write(root / "A.MD")
    write(root / "b.Markdown")
    write(root / "c.txt")
    code, _, _ = run(["-r", root], monkeypatch, capsys)
    assert code == 0
    assert docx_files(root) == ["A.docx", "b.docx"]


def test_directory_without_r_is_error(tmp_path, monkeypatch, capsys):
    root = tmp_path / "in"
    write(root / "a.md")
    code, _, err = run([root], monkeypatch, capsys)
    assert code == 2
    assert "-r" in err
    assert docx_files(root) == []


def test_empty_tree(tmp_path, monkeypatch, capsys):
    root = tmp_path / "in"
    root.mkdir()
    code, _, err = run(["-r", root], monkeypatch, capsys)
    assert code == 0
    assert "no Markdown files" in err
    assert docx_files(root) == []


def test_r_with_output_rejected(tmp_path, monkeypatch, capsys):
    root = tmp_path / "in"
    write(root / "a.md")
    code, _, err = run(["-r", root, tmp_path / "o.docx"], monkeypatch, capsys)
    assert code == 2
    assert "-r" in err
    assert docx_files(tmp_path) == []


def test_outdir_with_output_rejected(tmp_path, monkeypatch, capsys):
    src = write(tmp_path / "a.md")
    code, _, err = run(
        ["--outdir", tmp_path / "o", src, tmp_path / "a.docx"], monkeypatch, capsys
    )
    assert code == 2
    assert "--outdir" in err
    assert not (tmp_path / "o").exists()
    assert not (tmp_path / "a.docx").exists()


def test_outdir_tree_mirrors_structure_and_creates_dirs(tmp_path, monkeypatch, capsys):
    root = tmp_path / "in"
    write(root / "a" / "b" / "x.md")
    out = tmp_path / "new" / "deeper" / "out"
    code, _, _ = run(["-r", root, "--outdir", out], monkeypatch, capsys)
    assert code == 0
    assert docx_files(out) == ["a/b/x.docx"]
    assert docx_files(root) == []


def test_outdir_inside_input_root(tmp_path, monkeypatch, capsys):
    root = tmp_path / "in"
    write(root / "a" / "x.md")
    write(root / "y.md")
    out = root / "build"
    for _ in range(2):  # second run must not pick up anything from the first
        code, _, _ = run(["-r", root, "-o", out], monkeypatch, capsys)
        assert code == 0
        assert docx_files(out) == ["a/x.docx", "y.docx"]
    assert sorted(p.name for p in root.glob("*.docx")) == []


def test_outdir_with_single_file(tmp_path, monkeypatch, capsys):
    src = write(tmp_path / "sub" / "x.md")
    out = tmp_path / "made" / "here"
    code, _, _ = run(["--outdir", out, src], monkeypatch, capsys)
    assert code == 0
    assert docx_files(out) == ["x.docx"]


def test_outdir_is_a_file(tmp_path, monkeypatch, capsys):
    src = write(tmp_path / "a.md")
    blocker = tmp_path / "blocker"
    blocker.write_text("x")
    code, out, err = run(["--outdir", blocker, src], monkeypatch, capsys)
    assert code == 1
    assert "--outdir" in out + err
    assert blocker.read_text() == "x"
    assert docx_files(tmp_path) == []


def test_link_base_url_resolves_outside_tree_with_climbing(
    tmp_path, monkeypatch, capsys
):
    root = tmp_path / "in"
    write(root / "a" / "b" / "x.md", "[y](../y.md)\n")
    code, _, _ = run(
        ["-r", root, "--link-base-url", "https://example.com/docs"], monkeypatch, capsys
    )
    assert code == 0
    assert "https://example.com/docs/a/y.docx" in targets(root / "a" / "b" / "x.docx")


def test_dump_config_ignores_recursive(tmp_path, monkeypatch, capsys):
    root = tmp_path / "in"
    write(root / "a.md")
    code, out, _ = run(["--dump-config", "-r", root], monkeypatch, capsys)
    assert code == 0
    assert "font:" in out
    assert docx_files(root) == []


def test_stem_collision(tmp_path, monkeypatch, capsys):
    root = tmp_path / "in"
    write(root / "a.md")
    write(root / "a.markdown")
    write(root / "ok.md")
    code, out, err = run(["-r", root], monkeypatch, capsys)
    assert code == 1
    assert "collision" in err
    assert "a.md" in err and "a.markdown" in err
    assert docx_files(root) == ["ok.docx"]


def test_partial_failure_invalid_utf8(tmp_path, monkeypatch, capsys):
    root = tmp_path / "in"
    write(root / "good.md")
    write(root / "z.md")
    bad = root / "bad.md"
    bad.write_bytes(b"# caf\xe9 \xff\xfe\n")
    code, out, err = run(["-r", root], monkeypatch, capsys)
    assert code == 1
    assert "bad.md" in err
    assert "Converted 2 file(s), 1 failed" in out
    assert docx_files(root) == ["good.docx", "z.docx"]
    assert not list(root.rglob("*.tmp"))
    assert not (root / "bad.docx").exists()
