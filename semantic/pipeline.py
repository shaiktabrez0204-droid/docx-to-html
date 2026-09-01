"""End-to-end DOCX -> model -> styles -> headings -> hierarchy -> TOC -> HTML.

Single orchestration boundary that wires the OOXML parser, the semantic
style/classifier/hierarchy/toc layers, and the HTML renderer together. The
renderer and parser stay unaware of each other's concerns.
"""

from typing import List, Optional
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
from core.model import Paragraph, Table, Section, HeaderFooter, Run, Note
from core.anchoring import associate_floating_images
from core.layout import build_pages_from_sections, resolve_layout_state, LayoutState, resolve_image_coordinates
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


def _apply_run_resolution(para: Paragraph, registry: StyleRegistry,
                          doc_defaults: Optional[ResolvedStyle]) -> None:
    """Resolve run-level (character) typography for a single paragraph.

    ``_apply_run_style`` only fills ``None`` fields, so the FIRST layer applied
    to a run wins. Effective precedence (highest -> lowest) is therefore applied
    in this order, with the direct run rPr already set on the Run during parsing
    (so it wins over everything below):

       1. direct run rPr                 (already on Run, highest priority)
       2. character style (w:rStyle)     (applied first here)
       3. paragraph style run properties (applied next)
       4. docDefaults (rPrDefault)       (applied last, lowest priority)

    The paragraph style's run properties are applied to ALL runs -- including
    those that carry a character style -- so inherited paragraph formatting is
    not lost for char-styled runs; the character style is layered on top.
    """
    rs = registry.resolve(para.style_name) if para.style_name else None
    if rs is not None:
        _apply_para_style(para, rs)

    # 1. character style (higher priority than paragraph style)
    for r in para.runs:
        if r.style_name:
            rs_run = registry.resolve(r.style_name)
            if rs_run is not None:
                _apply_run_style(r, rs_run)

    # 2. paragraph style run properties (applies to every run in the paragraph)
    if rs is not None:
        for r in para.runs:
            _apply_run_style(r, rs)

    # 3. docDefaults -- always the base layer, applied last so it only fills
    #    runs that no higher-priority layer touched.
    if doc_defaults is not None:
        for r in para.runs:
            _apply_run_style(r, doc_defaults)


def classify_paragraphs(paragraphs: List[Paragraph], registry: StyleRegistry,
                        doc_defaults: Optional[ResolvedStyle] = None) -> None:
    """Replace each paragraph's style defaults with resolved style properties.

    For each paragraph, look up its w:paraStyle (or w:pStyle) resolved style and
    overwrite None-valued attributes so that BasedOn inheritance flows through.
    Also apply resolved run styles for any runs inside the paragraph. Heading
    classification runs last, using the resolved style metadata.
    """
    for para in paragraphs:
        _apply_run_resolution(para, registry, doc_defaults)

    # Heading classification (uses resolved style metadata)
    classify_headings(paragraphs, registry)


def _build_doc_defaults(parser) -> ResolvedStyle:
    """Build a ResolvedStyle carrying the document's rPrDefault run props."""
    dd = parser.get_default_font()
    return ResolvedStyle(
        style_id="__docDefaults__",
        font_family=dd.get("font_family"),
        font_size=dd.get("font_size"),
        font_color=dd.get("font_color"),
    )


def _apply_run_resolution_to_blocks(blocks, registry: StyleRegistry,
                                    doc_defaults: ResolvedStyle) -> None:
    """Apply run resolution to paragraphs nested inside tables."""
    for b in blocks:
        if isinstance(b, Table):
            for row in b.rows:
                for cell in row.cells:
                    for p in cell.content:
                        _apply_run_resolution(p, registry, doc_defaults)


def _apply_run_resolution_to_sections(sections, registry: StyleRegistry,
                                      doc_defaults: ResolvedStyle) -> None:
    """Apply run resolution to paragraphs inside header/footer parts."""
    for section in sections:
        for hf_map in (section.headers, section.footers):
            for hf in hf_map.values():
                for blk in hf.blocks:
                    if isinstance(blk, Paragraph):
                        _apply_run_resolution(blk, registry, doc_defaults)
                    elif isinstance(blk, Table):
                        for row in blk.rows:
                            for cell in row.cells:
                                for p in cell.content:
                                    _apply_run_resolution(p, registry, doc_defaults)


def _apply_run_resolution_to_notes(notes: List[Note], registry: StyleRegistry,
                                   doc_defaults: ResolvedStyle) -> None:
    for note in notes:
        for blk in note.blocks:
            if isinstance(blk, Paragraph):
                _apply_run_resolution(blk, registry, doc_defaults)
            elif isinstance(blk, Table):
                for row in blk.rows:
                    for cell in row.cells:
                        for p in cell.content:
                            _apply_run_resolution(p, registry, doc_defaults)


class ConversionResult:
    def __init__(self, paragraphs, registry, hierarchy, toc, html, image_assets=None,
                 numbering_model=None, numbering_validation=None, hierarchy_issues=None,
                 blocks=None, sections=None, even_headers=False, footnotes=None, endnotes=None,
                 comments=None, layout_state=None, pages=None):
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
        self.footnotes: List[Note] = footnotes or []
        self.endnotes: List[Note] = endnotes or []
        self.comments: List[Note] = comments or []
        self.layout_state = layout_state
        self.pages = pages or []

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

    # Document default run properties (rPrDefault) are the lowest-priority layer
    # for font size / family / color resolution.
    doc_defaults = _build_doc_defaults(parser)

    # Pass 2: classify each paragraph from resolved metadata (authoritative).
    classify_paragraphs(paragraphs, registry, doc_defaults)

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
    footnotes = parser.get_footnotes()
    endnotes = parser.get_endnotes()
    comments = parser.get_comments()

    pages = build_pages_from_sections(sections)
    for page in pages:
        page.floating_images = []
    layout_state = resolve_layout_state(blocks, sections, pages)
    for page in pages:
        for hf in [page.header_default, page.header_first, page.header_even,
                   page.footer_default, page.footer_first, page.footer_even]:
            if hf:
                for blk in hf.blocks:
                    if isinstance(blk, Paragraph):
                        for img in blk.images:
                            if img.wrap_type == "anchor" and img not in page.floating_images:
                                page.floating_images.append(img)
                    elif isinstance(blk, Table):
                        for row in blk.rows:
                            for cell in row.cells:
                                for p in cell.content:
                                    for img in p.images:
                                        if img.wrap_type == "anchor" and img not in page.floating_images:
                                            page.floating_images.append(img)
    for page in pages:
        for img in page.floating_images:
            coord = resolve_image_coordinates(img, pages, sections, layout_state.block_page_ownership)
            layout_state.image_coordinates[img.image_id] = coord

    # Rudimentary column index assignment for column-relative images
    def _assign_column_indices():
        sec_map = {}
        for b in blocks:
            if isinstance(b, Paragraph):
                for img in b.images:
                    if img.section_index is not None:
                        sec_map.setdefault(img.section_index, []).append(img)
            elif isinstance(b, Table):
                for row in b.rows:
                    for cell in row.cells:
                        for p in cell.content:
                            for img in p.images:
                                if img.section_index is not None:
                                    sec_map.setdefault(img.section_index, []).append(img)
        for sec in sections:
            cols = sec.page_layout.cols_num if sec.page_layout else 1
            if cols <= 1:
                continue
            col_imgs = []
            for b in blocks:
                sec_idx = None
                if isinstance(b, Paragraph):
                    for img in b.images:
                        if img.relative_from_horizontal == "column" or img.relative_from_vertical == "column":
                            if img.section_index == sec.index:
                                col_imgs.append(img)
                elif isinstance(b, Table):
                    for row in b.rows:
                        for cell in row.cells:
                            for p in cell.content:
                                for img in p.images:
                                    if (img.relative_from_horizontal == "column" or img.relative_from_vertical == "column") and img.section_index == sec.index:
                                        col_imgs.append(img)
            if not col_imgs:
                continue
            for order, img in enumerate(col_imgs):
                if img.column_index is not None:
                    continue
                img.column_index = (order * cols) // len(col_imgs) if len(col_imgs) else 0

    _assign_column_indices()

    # Inherited run typography (docDefaults + styles) must also reach runs that
    # live inside tables and inside header/footer parts, not just top-level body
    # paragraphs.
    _apply_run_resolution_to_blocks(blocks, registry, doc_defaults)
    _apply_run_resolution_to_sections(sections, registry, doc_defaults)
    _apply_run_resolution_to_notes(footnotes, registry, doc_defaults)
    _apply_run_resolution_to_notes(endnotes, registry, doc_defaults)
    _apply_run_resolution_to_notes(comments, registry, doc_defaults)

    default_font_size_pt = (doc_defaults.font_size or 22) / 2.0
    html = render_html(blocks, title=title, toc=toc,
                       assets=parser.get_image_assets(), page_layout=page_layout,
                       sections=sections, even_headers=even_headers,
                       default_font_size_pt=default_font_size_pt,
                       footnotes=footnotes, endnotes=endnotes, comments=comments,
                       layout_state=layout_state, pages=pages)
    return ConversionResult(
        paragraphs, registry, hierarchy, toc, html,
        image_assets=parser.get_image_assets(),
        numbering_model=numbering_model,
        numbering_validation=numbering_validation,
        hierarchy_issues=hierarchy_issues,
        blocks=blocks,
        sections=sections,
        even_headers=even_headers,
        footnotes=footnotes,
        endnotes=endnotes,
        comments=comments,
        layout_state=layout_state,
        pages=pages,
    )
