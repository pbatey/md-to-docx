"""Property test: the SharePoint ``?web=1`` suffix is appended once, to absolute links only."""

import urllib.parse
from dataclasses import replace

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import (
    LinkContext,
    normalize_base_url,
    rewrite_link,
    rewrite_wiki_link,
    sharepoint_folder_url,
)

SUFFIX = "?web=1"

_seg = st.text(
    alphabet=st.characters(
        min_codepoint=33, max_codepoint=0x2FF, blacklist_characters="/\\#%<>:|[]"
    ),
    min_size=1,
    max_size=8,
).filter(lambda s: s not in (".", ".."))

_bases = st.builds(
    lambda scheme, host, segs: normalize_base_url(
        f"{scheme}://{host}" + "".join("/" + s for s in segs)
    ),
    st.sampled_from(["http", "https"]),
    st.from_regex(r"[a-z][a-z0-9]{0,7}(\.[a-z]{2,4}){0,2}", fullmatch=True),
    st.lists(st.text(alphabet="abcXYZ09-_ ", min_size=1, max_size=6), max_size=3),
)

_dirs = st.lists(
    st.text(alphabet="abcXYZ09-_ ", min_size=1, max_size=5), max_size=3
).map(tuple)
_fragments = st.one_of(
    st.just(""), st.text(alphabet="abcXYZ 09-", min_size=1, max_size=10)
)


def _sibling(segs, ext, fragment):
    raw = "/".join(segs[:-1] + [segs[-1] + ext])
    return raw + ("#" + fragment if fragment else "")


_sibling_targets = st.builds(
    _sibling,
    st.lists(_seg, min_size=1, max_size=3),
    st.sampled_from([".md", ".markdown", ".MD"]),
    _fragments,
)

_wiki_inners = st.builds(
    lambda segs, frag, label: (
        "/".join(segs) + ("#" + frag if frag else "") + ("|" + label if label else "")
    ),
    st.lists(
        st.text(alphabet="abcXYZ09-_ ", min_size=1, max_size=6).filter(
            lambda s: s.strip()
        ),
        min_size=1,
        max_size=3,
    ),
    _fragments.map(str.strip),
    st.one_of(st.just(""), st.text(alphabet="abcXYZ09", min_size=1, max_size=6)),
)


def _ctx(base, rel_dir, suffix):
    return LinkContext(
        base_url=base,
        link_suffix=suffix,
        source_rel_dir=rel_dir,
        warn=lambda m: None,
    )


def _check_suffix(plain, suffixed, bare):
    assert plain is not None and suffixed is not None and bare is not None
    t = suffixed.target
    assert t.endswith(SUFFIX)
    assert t.count(SUFFIX) == 1
    assert t.count("?") == 1
    assert t[: -len(SUFFIX)] == plain.target
    assert "#" not in t
    assert suffixed.anchor == plain.anchor
    # Without a base URL the suffix never appears.
    assert "?" not in bare.target and SUFFIX not in bare.target
    assert "#" not in bare.target


# Feature: sibling-markdown-links, Property 25: SharePoint suffix
# **Validates: Requirements 14.4, 14.5, 14.9**
@settings(max_examples=200)
@given(base=_bases, rel_dir=_dirs, raw=_sibling_targets)
def test_sibling_link_suffix(base, rel_dir, raw):
    plain = rewrite_link(raw, _ctx(base, rel_dir, ""))
    suffixed = rewrite_link(raw, _ctx(base, rel_dir, SUFFIX))
    # A suffix set without a base URL is not applied (relative targets).
    bare = rewrite_link(raw, _ctx(None, rel_dir, SUFFIX))
    _check_suffix(plain, suffixed, bare)


# Feature: sibling-markdown-links, Property 25: SharePoint suffix
# **Validates: Requirements 14.4, 14.5, 14.9**
@settings(max_examples=200)
@given(base=_bases, rel_dir=_dirs, inner=_wiki_inners)
def test_wiki_link_suffix(base, rel_dir, inner):
    plain = rewrite_wiki_link(inner, _ctx(base, rel_dir, ""))
    suffixed = rewrite_wiki_link(inner, _ctx(base, rel_dir, SUFFIX))
    bare = rewrite_wiki_link(inner, _ctx(None, rel_dir, SUFFIX))
    assert plain is not None and suffixed is not None and bare is not None
    _check_suffix(plain[0], suffixed[0], bare[0])
    assert plain[1] == suffixed[1] == bare[1]


# Feature: sibling-markdown-links, Property 25: SharePoint suffix
# **Validates: Requirements 14.13**
@settings(max_examples=100)
@given(
    host=st.from_regex(r"[a-z][a-z0-9]{0,7}\.sharepoint\.com", fullmatch=True),
    folder=st.lists(
        st.text(alphabet="abcXYZ09-_ ", min_size=1, max_size=6), min_size=1, max_size=4
    ),
    view=st.sampled_from(["AllItems.aspx", "allitems.ASPX", "Custom.aspx"]),
    use_id=st.booleans(),
    raw=_sibling_targets,
)
def test_view_url_folder_gets_suffix(host, folder, view, use_id, raw):
    root = "/sites/team/Shared Documents"
    folder_path = "/".join([root] + folder)
    path = urllib.parse.quote(root) + "/Forms/" + view
    query = "?id=" + urllib.parse.quote(folder_path, safe="") if use_id else ""
    url = f"https://{host}{path}{query}"

    result = sharepoint_folder_url(url)
    assert result is not None
    assert "/Forms/" not in result
    expected_folder = folder_path if use_id else urllib.parse.quote(root)
    assert result == f"https://{host}{expected_folder}"

    base = normalize_base_url(result)
    ctx = _ctx(base, (), SUFFIX)
    out = rewrite_link(raw, ctx)
    plain = rewrite_link(raw, replace(ctx, link_suffix=""))
    assert out.target.startswith(base + "/")
    assert out.target.endswith(SUFFIX) and out.target.count("?") == 1
    assert out.target[: -len(SUFFIX)] == plain.target
    assert "#" not in out.target


def test_non_view_urls_are_not_folder_urls():
    assert sharepoint_folder_url("https://x.sharepoint.com/sites/a/Docs") is None
    assert sharepoint_folder_url("https://x.sharepoint.com/sites/a/Forms/x.txt") is None
