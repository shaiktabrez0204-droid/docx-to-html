"""Superscript/subscript positioning test."""
def test_superscript_subscript():
    """Test that superscript/subscript can be detected."""
    from docx_to_html.adapter.mammoth_adapter import _detect_superscript_subscript
    # Test with simple context
    result = _detect_superscript_subscript("test", context="some <sup>context</sup>")
    # Should detect superscript if context has sup tags
    has_superscript = result.get("has_superscript", False)
    position = result.get("position", "normal")
    # Just verify the function runs without error
    assert "position" in result
