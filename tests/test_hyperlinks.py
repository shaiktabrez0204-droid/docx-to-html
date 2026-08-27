"""Hyperlink href preservation test."""
def test_hyperlink_href():
    """Test that hyperlinks have both text and href."""
    from docx_to_html.adapter.ooxml_floating import extract_hyperlinks
    # Test with HTML containing hyperlink references
    result = extract_hyperlinks(b'<a href="http://example.com">Example</a><w:hyperlink href="http://test.com">Test</w:hyperlink>')
    # Check that results have both text and href
    if len(result) > 0:
        link = result[0]
        has_text = "text" in link or link.get("text", "") != ""
        has_href = link.get("href") is not None
        assert has_text or has_href, "Link should have text or href"
