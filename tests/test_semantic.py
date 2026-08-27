"""Semantic test (real DOCX style resolution)."""
from semantic.style_resolver import resolve_style


def test_resolve_heading_style():
    # "Heading1" is a structural heading style id -> resolved non-None.
    r = resolve_style("Heading1")
    assert r is not None
    assert r["is_heading"] is True
    assert r["level"] == 1
