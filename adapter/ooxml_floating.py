
def extract_hyperlinks(docx_content):
    """Extract hyperlinks from DOCX content.
    
    Returns list of dicts with:
    - text: Link text visible to reader
    - href: URL or target reference
    - target: Target frame/window if specified
    - anchor: Anchor position in document
    """
    links = []
    
    # Look for w:hyperlink elements in OOXML
    # Pattern: w:anchor -> w:r -> w:rPr -> w:hyperlink
    hyperlink_pattern = re.compile(r'w:hyperlink[^>]*>(.*?)<\/w:hyperlink>', re.DOTALL)
    matches = hyperlink_pattern.findall(docx_content)
    
    for i, match in enumerate(matches):
        # Extract the text content within the hyperlink
        text_match = re.search(r'>(.*?)<', match)
        link_text = text_match.group(1).strip() if text_match else "Link " + str(i+1)
        
        # Extract href attribute
        href_match = re.search(r'href=["\']([^"\']+)["\']', match)
        link_href = href_match.group(1) if href_match else None
        
        links.append({
            "text": link_text,
            "href": link_href,
            "target": None,
            "anchor": None
        })
    
    return links


def render_hyperlink_html(link_data, options=None):
    """Render a hyperlink as HTML."""
    href = link_data.get("href")
    text = link_data.get("text", "")
    
    if href:
        return f'<a href="{href}">{text}</a>'
    elif text:
        return f'<span class="hyperlinked">{text}</span>'
    return ""


def extract_all_hyperlinks(docx_content):
    """Extract all hyperlinks from DOCX content and return structured data."""
    links = extract_hyperlinks(docx_content)
    # Remove duplicates based on text+href
    seen = set()
    unique_links = []
    for link in links:
        key = (link.get("text", ""), link.get("href", ""))
        if key not in seen:
            seen.add(key)
            unique_links.append(link)
    return unique_links
