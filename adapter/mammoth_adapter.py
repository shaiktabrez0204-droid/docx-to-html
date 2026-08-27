"""Mammoth adapter for docx-to-html conversion."""
import os
import re
from adapter.ooxml_header_footer import extract_headers_footers, render_header_html, render_footer_html
from adapter.ooxml_floating import extract_hyperlinks, render_hyperlink_html
from core.model import Image
from semantic.style_resolver import ResolvedStyle


def mammoth_to_html(docx_content):
    """Convert DOCX content to HTML using mammoth-like interface."""
    # Extract headers and footers
    hf = extract_headers_footers(docx_content)

    # Render header and footer HTML
    header_html = render_header_html(hf)
    footer_html = render_footer_html(hf)

    full_html = "<html><head></head><body>" + header_html + "<p>Converted</p>" + footer_html + "</body></html>"

    # Extract and render hyperlinks
    links = extract_hyperlinks(docx_content)
    link_html_parts = []
    for link in links:
        link_html = render_hyperlink_html(link)
        link_html_parts.append(link_html)

    full_html = "<html><head></head><body>" + header_html + "<p>Converted</p>" + footer_html + "</body></html>"
    if link_html_parts:
        full_html = "<html><head></head><body>" + "<br>".join(link_html_parts) + "<p>Converted</p>" + footer_html + "</body></html>"

    return full_html


def extract_run_fonts(docx_path: str) -> dict:
    """Extract font information from DOCX runs.

    Returns dict with font family, size, bold, italic, etc.
    """
    from docx_to_html.adapter.ooxml_parser import OoxmlParser

    parser = OoxmlParser(docx_path)
    styles = parser.get_styles()
    default_font = parser.get_default_font()

    # Build font map from styles
    font_map = {}
    for style in styles:
        if style.font_family:
            font_map[style.style_id] = {
                'font_family': style.font_family,
                'font_size': style.font_size,
                'bold': style.bold,
                'italic': style.italic,
            }

    # Merge with default font
    result = {
        'default_font_family': default_font['font_family'],
        'default_font_size': default_font['font_size'],
        'font_map': font_map,
    }

    parser.close()
    return result


def extract_styles_from_docx(docx_path: str) -> dict:
    """Extract paragraph styles from DOCX styles.xml.

    Returns dict mapping style_id -> Style object.
    """
    from docx_to_html.adapter.ooxml_parser import OoxmlParser

    parser = OoxmlParser(docx_path)
    styles_def = parser.get_styles()

    style_objects = {}
    for style_def in styles_def:
        style_props = Style(
            name=style_def.name,
            font_family=style_def.font_family,
            font_size=style_def.font_size,
            bold=style_def.bold,
            italic=style_def.italic,
            underline=style_def.underline,
            color=style_def.font_color,
            space_before=style_def.spacing_before if style_def.spacing_before is not None else 0.0,
            space_after=style_def.spacing_after if style_def.spacing_after is not None else 0.0,
            left_indent=style_def.indent_left if style_def.indent_left is not None else 0,
            right_indent=style_def.indent_right if style_def.indent_right is not None else 0,
            first_line_indent=style_def.indent_first_line if style_def.indent_first_line is not None else 0,
            line_height=style_def.line_spacing if style_def.line_spacing is not None else 1.0,
            based_on=style_def.based_on,
            level=style_def.outline_level,
        )
        style_objects[style_def.style_id] = style_props

    parser.close()
    return style_objects


def attach_floating_images(normalized_blocks: list, image_assets: dict) -> list:
    """Attach floating images to their anchor paragraphs in the normalized document.

    Uses anchor_paragraph_index to match floating images to paragraphs.
    """
    from docx_to_html.core.anchoring import associate_floating_images

    # Run the anchoring pass to associate floating images with paragraphs
    result = associate_floating_images(normalized_blocks, image_assets)

    return result


def _extract_images(docx_content):
    """Extract images from DOCX content.

    Returns list of dicts with:
    - content_id: Internal image ID
    - alt_text: Alternative text from DOCX
    - width: Image width in EMUs
    - height: Image height in EMUs
    - source_url: Source URL if available
    """
    images = []
    # Simple placeholder - in full implementation would parse VML/SVG
    # Look for image references in the format: /images/image1.png
    img_pattern = re.compile(r'/images/([^\s"]+)[^>]*>')
    matches = img_pattern.findall(docx_content)
    for i, match in enumerate(matches):
        images.append({
            'content_id': f'image{i+1}',
            'alt_text': f'Image {i+1}',
            'width': None,
            'height': None,
            'source_url': match
        })
    return images


def _detect_superscript_subscript(run_text, context=""):
    """Detect if text has superscript or subscript formatting.

    In a full implementation, this would parse OOXML run properties.
    For now, returns default values.
    """
    # Placeholder - in full implementation, parse w:superScript/w:subScript
    # from the OOXML run properties
    has_superscript = False
    has_subscript = False
    position = "normal"  # normal, subscript, superscript

    # Simple heuristics for common patterns
    if context and context.startswith("["):
        # If context looks like markup, try to detect
        if "<sup>" in context or "</sup>" in context:
            has_superscript = True
            position = "superscript"
        elif "<sub>" in context or "</sub>" in context:
            has_subscript = True
            position = "subscript"

    return {
        "has_superscript": has_superscript,
        "has_subscript": has_subscript,
        "position": position
    }