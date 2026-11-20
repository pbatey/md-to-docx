"""Property test for Base_URL normalization idempotence and no double-encoding."""
# Feature: sibling-markdown-links, Property 16: Base URL normalization is idempotent and never double-encodes
import re

from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import normalize_base_url

_VALID_PCT = re.compile(r"%[0-9A-Fa-f]{2}")

_scheme = st.sampled_from(["http", "https", "HTTP", "Https"])
_host = st.sampled_from(["example.com", "Example.COM:8080", "localhost", "a-b.org"])
_valid_pct = st.sampled_from(["%20", "%C3%A9", "%5C", "%2f", "%e6%97%a5", "%25"])
_invalid_pct = st.sampled_from(["%zz", "%C", "%", "%2", "%%"])
_piece = st.one_of(
    _valid_pct,
    _invalid_pct,
    st.sampled_from(list(" /é日本-._~")),
    st.text(
        alphabet=st.characters(
            blacklist_categories=("Cs",), blacklist_characters="?#%"
        ),
        max_size=5,
    ),
)
_path = st.lists(_piece, max_size=8).map("".join)
_url = st.builds(
    lambda scheme, host, path, slash, pad: pad[0]
    + scheme
    + "://"
    + host
    + ("/" + path if path or slash else "")
    + ("/" if slash else "")
    + pad[1],
    _scheme,
    _host,
    _path,
    st.booleans(),
    st.sampled_from([("", ""), (" ", ""), ("", " "), (" ", " ")]),
)


@settings(max_examples=200, deadline=None)
@given(_url)
def test_normalize_base_url_idempotent_no_double_encoding(u):
    """**Validates: Requirements 8.10, 8.11, 13.14**"""
    once = normalize_base_url(u)
    assert normalize_base_url(once) == once
    # Every valid %XX in the input path appears unchanged in the result.
    path = u.strip().split("://", 1)[1]
    path = path[path.find("/") :] if "/" in path else ""
    for m in _VALID_PCT.finditer(path):
        assert m.group(0) in once
    assert once.count("%25") >= path.count("%25")
    assert not once.endswith("/")
    assert " " not in once


def test_valid_escape_not_double_encoded():
    assert normalize_base_url("https://x.com/a%20b") == "https://x.com/a%20b"
    assert normalize_base_url("https://x.com/a b") == "https://x.com/a%20b"
    assert normalize_base_url("https://x.com/100%/") == "https://x.com/100%25"
