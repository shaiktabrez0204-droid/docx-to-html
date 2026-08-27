"""Header/footer extraction test."""
def test_header_footer_extraction():
    """Test that headers and footers can be extracted from DOCX content."""
    from docx_to_html.adapter.ooxml_header_footer import extract_headers_footers
    # Test with simple HTML containing header/footer references
    result = extract_headers_footers(b"<p>Content</p><headerPart>My Header</headerPart><footerPart>My Footer</footerPart>")
    assert "header" in result or "footer" in result
