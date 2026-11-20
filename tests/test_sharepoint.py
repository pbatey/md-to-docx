"""SharePoint support: view-URL conversion and the --sharepoint / --sp options.

Validates: Requirements 14.1, 14.2, 14.3, 14.6, 14.7, 14.8, 14.10, 14.13
"""

import re
import sys
import zipfile

import pytest

from md_to_docx.md_to_docx import main, normalize_base_url, sharepoint_folder_url

from golden_fixtures import load_golden

GOLDEN = load_golden()

LIB = "https://aegonna.sharepoint.com/sites/TTWFG/Shared%20Documents"
VIEW = (
    LIB + "/Forms/AllItems.aspx?id=%2Fsites%2FTTWFG%2FShared%20Documents"
    "%2FOperations%20team%2FIVR%20Knowledge%20Base%2FTEST"
    "&viewid=abc-123"
)
FOLDER = LIB + "/Operations%20team/IVR%20Knowledge%20Base/TEST"


def run(args, monkeypatch, capsys):
    monkeypatch.setattr(sys, "argv", ["md-to-docx", *map(str, args)])
    code = 0
    try:
        main()
    except SystemExit as exc:
        code = exc.code if exc.code is not None else 0
    cap = capsys.readouterr()
    return code, cap.out, cap.err


def part(docx, name):
    with zipfile.ZipFile(docx) as z:
        return z.read(name).decode("utf-8")


def targets(docx):
    return re.findall(r'Target="([^"]*)"', part(docx, "word/_rels/document.xml.rels"))


def norm(xml):
    seen = {}

    def sub(m):
        return f'{m.group(1)}w:id="{seen.setdefault(m.group(2), str(len(seen) + 1))}"'

    return re.sub(r'(<w:bookmark(?:Start|End) )w:id="(\d+)"', sub, xml)


# --------------------------- sharepoint_folder_url -------------------------- #


def test_view_url_with_id_gives_folder():
    result = sharepoint_folder_url(VIEW)
    assert result is not None
    assert normalize_base_url(result) == FOLDER


def test_view_url_without_id_gives_library_root():
    result = sharepoint_folder_url(LIB + "/Forms/AllItems.aspx")
    assert result is not None
    assert normalize_base_url(result) == LIB


@pytest.mark.parametrize(
    "url",
    [
        "https://aegonna.sharepoint.com/sites/TTWFG/SitePages/Home.aspx",
        "https://aegonna.sharepoint.com/sites/TTWFG/Shared%20Documents/Forms/",
        FOLDER,
        "https://example.com/docs",
    ],
)
def test_non_view_urls_give_none(url):
    assert sharepoint_folder_url(url) is None


# --------------------------------- CLI ------------------------------------- #


def make_tree(tmp_path):
    d = tmp_path / "docs"
    d.mkdir()
    (d / "a.md").write_text("# A\n\n## Setup\n\nText.\n", encoding="utf-8")
    (d / "b.md").write_text("# B\n\nSee [a](a.md#Setup).\n", encoding="utf-8")
    return d


def test_help_lists_sharepoint_and_sp(monkeypatch, capsys):
    code, out, _ = run(["--help"], monkeypatch, capsys)
    assert code == 0
    assert "--sharepoint" in out
    assert "--sp" in out


def test_sp_identical_to_sharepoint(tmp_path, monkeypatch, capsys):
    results = []
    for flag in ("--sharepoint", "--sp"):
        (tmp_path / flag.strip("-")).mkdir()
        d = make_tree(tmp_path / flag.strip("-"))
        code, out, err = run([flag, "--link-base-url", VIEW, "-r", d],
                             monkeypatch, capsys)
        assert code == 0
        results.append(
            (out.replace(flag, "F"), err, targets(d / "b.docx"),
             norm(part(d / "b.docx", "word/document.xml")))
        )
    a, b = results
    assert a[1:] == b[1:]


@pytest.mark.parametrize("flag", ["--sharepoint=yes", "--sp=yes"])
def test_flag_with_value_rejected(flag, tmp_path, monkeypatch, capsys):
    d = make_tree(tmp_path)
    code, _, _ = run([flag, "--link-base-url", VIEW, "-r", d], monkeypatch, capsys)
    assert code != 0
    assert not list(d.glob("*.docx"))


def test_sharepoint_requires_base_url(tmp_path, monkeypatch, capsys):
    d = make_tree(tmp_path)
    code, _, _ = run(["--sharepoint", "-r", d], monkeypatch, capsys)
    assert code != 0
    assert not list(d.glob("*.docx"))


def test_tree_run_writes_web1_link_with_anchor(tmp_path, monkeypatch, capsys):
    d = make_tree(tmp_path)
    code, _, _ = run(["--sharepoint", "--link-base-url", VIEW, "-r", d],
                     monkeypatch, capsys)
    assert code == 0
    links = [t for t in targets(d / "b.docx") if t.endswith("a.docx?web=1")]
    assert links, targets(d / "b.docx")
    assert links[0].startswith(FOLDER + "/")
    assert 'w:anchor="_toc_setup"' in part(d / "b.docx", "word/document.xml")


def test_no_dropped_query_warning_for_converted_view_url(tmp_path, monkeypatch, capsys):
    d = make_tree(tmp_path)
    code, _, err = run(["--sharepoint", "--link-base-url", VIEW, "-r", d],
                       monkeypatch, capsys)
    assert code == 0
    assert "dropped" not in err
    assert "WARNING" not in err


def test_hint_warning_without_sharepoint(tmp_path, monkeypatch, capsys):
    d = make_tree(tmp_path)
    code, _, err = run(["--link-base-url", VIEW, "-r", d], monkeypatch, capsys)
    assert code == 0
    assert "WARNING" in err
    assert "--sharepoint" in err


def test_no_sibling_links_with_sharepoint_matches_golden(tmp_path, monkeypatch, capsys):
    md, expected = GOLDEN["mixed"]
    src = tmp_path / "mixed.md"
    src.write_text(md, encoding="utf-8")
    out = tmp_path / "mixed.docx"
    code, _, _ = run(["--no-sibling-links", "--sharepoint", "--link-base-url", VIEW,
                      src, out], monkeypatch, capsys)
    assert code == 0
    assert norm(part(out, "word/document.xml")) == norm(expected)
