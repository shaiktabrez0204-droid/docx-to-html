"""End-to-end DOCX -> model -> styles -> headings -> hierarchy -> TOC -> HTML.

Single orchestration boundary that wires the OOXML parser, the semantic
style/classifier/hierarchy/toc layers, and the HTML renderer together. The
renderer and parser stay unaware of each other's concerns.
"""

from typing import List
from adapter.ooxml_parser import OoxmlParser
from semantic.style_resolver import StyleRegistry, ResolvedStyle
from semantic.classifier import classify_paragraphs as classify_headings
from semantic.hierarchy import build_hierarchy, flatten_hierarchy
from semantic.toc import build_toc, flatten_toc, detect_heading_anomalies
from semantic.numbering import (
    NumberingResolver,
    cross_validate,
    validate_hierarchy,
)
from core.model import Paragraph, Table, Section, HeaderFooter, Run
from core.anchoring import associate_floating_images
from output.html_renderer import render_html

def _apply_para_style(para: Paragraph, rs: ResolvedStyle) -> None:
    """Apply resolved paragraph style to a Paragraph model.

    Rules: direct attributes already set on para; otherwise inherit
    from the resolved style with BasedOn chain fallback.
    """
    if para.alignment is None:
        para.alignment = rs.alignment
    if para.indent_left is None:
        para.indent_left = rs.indent_left
    if para.indent_right is None:
        para.indent_right = rs.indent_right
    if para.indent_first_line is None:
        para.indent_first_line = rs.indent_first_line
    if para.indent_hanging is None:
        para.indent_hanging = rs.indent_hanging
    if para.spacing_before is None:
        para.spacing_before = rs.spacing_before
    if para.spacing_after is None:
        para.spacing_after = rs.spacing_after
    if para.line_spacing is None:
        para.line_spacing = rs.line_spacing
    if para.outline_level is None and rs.outline_level is not None:
        para.outline_level = rs.outline_level


def _apply_run_style(run: Run, rs: ResolvedStyle) -> None:
    """Apply resolved run (character) style to a Run model.

    Precedence (lowest → highest):
      1. default (already set on Run constructor)
      2. paragraph-style inheritance (from resolved style)
      3. direct (character-style override) — already set on Run from parser
    We only fill in None fields from the resolved style.
    """
    if run.font_family is None and rs.font_family is not None:
        run.font_family = rs.font_family
    if run.font_size is None and rs.font_size is not None:
        run.font_size = rs.font_size
    if run.font_color is None and rs.font_color is not None:
        run.font_color = rs.font_color
    if run.bold is None and rs.bold is not None:
        run.bold = rs.bold
    if run.italic is None and rs.italic is not None:
        run.italic = rs.italic
    if run.underline is None and rs.underline is not None:
        run.underline = rs.underline
    if run.superscript is None and rs.superscript is not None:
        run.superscript = rs.superscript
    if run.subscript is None and rs.subscript is not None:
        run.subscript = rs.subscript


def classify_paragraphs(paragraphs: List[Paragraph], registry: StyleRegistry) -> None:
    """Replace each paragraph's style defaults with resolved style properties.

    For each paragraph, look up its w:paraStyle (or w:pStyle) resolved style and
    overwrite None-valued attributes so that BasedOn inheritance flows through.
    Also apply resolved run styles for any runs inside the paragraph.
    """
    for para in paragraphs:
        # Find the paragraph style reference from the paragraph's own w:pPr/w:pStyle
        # (this is set up when the paragraph was parsed from OOXML).
        # The simplest path: the para may have its effective style stored.
        # If not, try to resolve from the style name that appears in the
        # paragraph-level tag.  For now, we use a registry lookup by style_id.
        # The para.style_name was set during parsing as w:pStyle/@w:val.
        rs = registry.resolve(para.style_name) if para.style_name else None
        if rs is not None:
            _apply_para_style(para, rs)
            # Apply paragraph style's run properties to all runs (for inheritance)
            for r in para.runs:
                if r.style_name is None:
                    _apply_run_style(r, rs)
        # Apply run-level style overrides (character style takes precedence)
        for r in para.runs:
            if r.style_name:
                rs_run = registry.resolve(r.style_name)
                if rs_run is not None:
                    _apply_run_style(r, rs_run)

    # Heading classification (uses resolved style metadata)
    classify_headings(paragraphs, registry)


class ConversionResult:
    def __init__(self, paragraphs, registry, hierarchy, toc, html, image_assets=None,
                 numbering_model=None, numbering_validation=None, hierarchy_issues=None,
                 blocks=None, sections=None, even_headers=False):
        self.paragraphs = paragraphs
        self.blocks = blocks if blocks is not None else paragraphs
        self.sections = sections or []
        self.even_headers = even_headers
        self.registry = registry
        self.hierarchy = hierarchy
        self.toc = toc
        self.html = html
        self.image_assets = image_assets or {}
        self.numbering_model = numbering_model
        self.numbering_validation = numbering_validation or []
        self.hierarchy_issues = hierarchy_issues or []

    def flat_headings(self):
        return flatten_hierarchy(self.hierarchy)

    def flat_toc(self):
        return flatten_toc(self.toc)

    def anomalies(self):
        return detect_heading_anomalies(self.flat_toc())


def convert_docx(docx_path: str, title: str = "Converted Document") -> ConversionResult:
    parser = OoxmlParser(docx_path)
    blocks = parser.parse_document()
    paragraphs = [b for b in blocks if isinstance(b, Paragraph)]

    # Pass 1: resolve all style metadata (incl. BasedOn inheritance).
    registry = StyleRegistry(parser.get_styles())

    # Pass 2: classify each paragraph from resolved metadata (authoritative).
    classify_paragraphs(paragraphs, registry)

    numbering_model = parser.get_numbering()
    resolver = NumberingResolver(numbering_model)
    resolver.resolve(paragraphs, registry)
    numbering_validation = cross_validate(paragraphs)

    # Build hierarchy + TOC from resolved headings.
    hierarchy = build_hierarchy(paragraphs)
    hierarchy_issues = validate_hierarchy(hierarchy)
    toc = build_toc(hierarchy)

    # Associate floating (wp:anchor) images with their nearest block. This mutates
    # Paragraph.block_id and each floating Image's nearest_block_id/confidence.
    associate_floating_images(paragraphs)

    sections = parser.get_sections()
    even_headers = parser.get_even_headers_flag()
    page_layout = parser.get_page_layout()
    html = render_html(blocks, title=title, toc=toc,
                       assets=parser.get_image_assets(), page_layout=page_layout,
                       sections=sections, even_headers=even_headers)
    return ConversionResult(
        paragraphs, registry, hierarchy, toc, html,
        image_assets=parser.get_image_assets(),
        numbering_model=numbering_model,
        numbering_validation=numbering_validation,
        hierarchy_issues=hierarchy_issues,
        blocks=blocks,
        sections=sections,
        even_headers=even_headers,
    )
