"""OOXML header/footer handling for docx-to-html."""
import re


def extract_headers_footers(docx_content):
    """Extract header and footer content from DOCX OOXML.
    
    Returns dict with:
    - header: Header text content
    - footer: Footer text content
    - header_type: Type of header (primary, first page, etc.)
    - footer_type: Type of footer (primary, page number, etc.)
    """
    header = ""
    footer = ""
    header_type = "primary"
    footer_type = "primary"
    
    # Look for header part names in the format: header1.xml, header2.xml, etc.
    # Common patterns: header, footer, header_footer
    header_patterns = re.findall(r'[Hh]eader[_\s]?[0-9]?\.?[^\s>]+', docx_content)
    footer_patterns = re.findall(r'[Ff]oott[_\s]?[0-9]?\.?[^\s>]+', docx_content)
    
    if header_patterns:
        header = "_".join(header_patterns[:3])
    if footer_patterns:
        footer = "_".join(footer_patterns[:3])
    
    return {
        "header": header,
        "footer": footer,
        "header_type": header_type,
        "footer_type": footer_type,
    }


def render_header_html(header_data, options=None):
    """Render header as HTML."""
    if header_data.get("header"):
        return f'<div class="header">{header_data["header"]}</div>'
    return ""


def render_footer_html(footer_data, options=None):
    """Render footer as HTML."""
    if footer_data.get("footer"):
        return f'<div class="footer">{footer_data["footer"]}</div>'
    return ""
