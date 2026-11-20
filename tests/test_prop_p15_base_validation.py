"""Property test for --link-base-url validation."""

import re

import pytest
from hypothesis import given, settings
from hypothesis import strategies as st

from md_to_docx.md_to_docx import normalize_base_url

_label = st.text(
    alphabet="abcdefghijklmnopqrstuvwxyz0123456789-", min_size=1, max_size=10
).filter(lambda s: not s.startswith("-") and not s.endswith("-"))
_hosts = st.lists(_label, min_size=1, max_size=3).map(".".join)
_ports = st.one_of(st.just(""), st.integers(1, 65535).map(lambda p: f":{p}"))
_paths = st.lists(st.text(alphabet="abcXYZ019-._~ é%", max_size=8), max_size=4).map(
    lambda segs: "".join("/" + s for s in segs)
)
_schemes = st.sampled_from(["http", "https", "HTTP", "HTTPS", "Http", "hTTps"])
_pad = st.sampled_from(["", " ", "\t", "\n"])

_valid = st.builds(
    lambda pre, sch, host, port, path, post: f"{pre}{sch}://{host}{port}{path}{post}",
    _pad,
    _schemes,
    _hosts,
    _ports,
    _paths,
    _pad,
)

_other_schemes = st.sampled_from(
    ["ftp", "file", "mailto", "javascript", "ws", "data", "httpx", "ssh", ""]
)
_bad_scheme = st.builds(
    lambda sch, host, path: f"{sch}://{host}{path}" if sch else f"{host}{path}",
    _other_schemes,
    _hosts,
    _paths,
)
_hostless = st.builds(
    lambda sch, path, extra: f"{sch}://{extra}{path}",
    _schemes,
    _paths,
    st.sampled_from(["", "?q=1", "#frag", ":80", "@"]),
)
_arbitrary = st.text(max_size=40).filter(
    lambda s: not re.match(r"^\s*https?://", s, re.IGNORECASE)
)
_invalid = st.one_of(_bad_scheme, _hostless, _arbitrary, st.just(""), st.just("   "))


# Feature: sibling-markdown-links, Property 15: Base URL validation
# **Validates: Requirements 8.12**
@settings(max_examples=200)
@given(_valid)
def test_valid_http_urls_are_accepted(raw):
    result = normalize_base_url(raw)
    assert result.startswith(("http://", "https://"))


# Feature: sibling-markdown-links, Property 15: Base URL validation
# **Validates: Requirements 8.12**
@settings(max_examples=200)
@given(_invalid)
def test_non_http_or_hostless_urls_raise_value_error(raw):
    with pytest.raises(ValueError) as exc:
        normalize_base_url(raw)
    assert repr(raw) in str(exc.value)
