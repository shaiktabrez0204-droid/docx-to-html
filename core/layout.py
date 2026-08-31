"""First-class page/layout state and coordinate transformation.

This module is the SINGLE authoritative source for:
1. Building pages from sections (pagination stub - one page per section for now)
2. Resolving coordinate transformations: OOXML space → section → page → column/paragraph → CSS pixels
3. Assigning explicit page ownership to every layout object

No other module should perform coordinate calculations.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Any

from core.model import (
    PageLayout,
    Page,
    Section,
    Image,
    Paragraph,
    Table,
    Block,
    HeaderFooter,
)
from core.units import emu_to_px, EMU_PER_INCH


# Supported relativeFrom values in OOXML
SUPPORTED_RELATIVE_FROM = {
    "page", "margin", "column", "character", "paragraph", "line",
}


@dataclass
class ResolvedCoordinate:
    """A coordinate resolved to page-local CSS pixels."""
    x_px: int
    y_px: int
    coordinate_space: str  # "page" | "margin" | "column" | "paragraph" | "character" | "line"
    relative_from_horizontal: Optional[str]
    relative_from_vertical: Optional[str]
    section_index: int
    page_index: int
    column_index: Optional[int]


@dataclass
class LayoutState:
    """Complete resolved layout state for a document."""
    sections: List[Section]
    pages: List[Page]
    # Maps object id -> ResolvedCoordinate
    image_coordinates: Dict[str, ResolvedCoordinate] = field(default_factory=dict)
    # Maps block_id -> (section_index, page_index)
    block_page_ownership: Dict[str, Tuple[int, int]] = field(default_factory=dict)


def build_pages_from_sections(sections: List[Section]) -> List[Page]:
    """Build Page objects from Sections.

    CURRENT IMPLEMENTATION: One page per section (pagination stub).
    FUTURE: Real pagination engine will create multiple pages per section.

    This function is the single place where page structure is created.
    """
    pages: List[Page] = []
    for sec in sections:
        page_layout = sec.page_layout or PageLayout()
        page_layout.section_index = sec.index

        # Create one page for this section (stub)
        page = Page(
            page_index=len(pages),
            section_index=sec.index,
            page_layout=page_layout,
            content_origin_x_px=page_layout.margin_left_px,
            content_origin_y_px=page_layout.margin_top_px,
        )

        # Attach headers/footers
        page.header_default = sec.get_header("default")
        page.header_first = sec.get_header("first")
        page.header_even = sec.get_header("even")
        page.footer_default = sec.get_footer("default")
        page.footer_first = sec.get_footer("first")
        page.footer_even = sec.get_footer("even")

        pages.append(page)

    return pages


def assign_blocks_to_pages(blocks: List[Block], pages: List[Page]) -> None:
    """Assign blocks to their owning page.

    CURRENT: All blocks in a section go to that section's single page.
    FUTURE: Real pagination will distribute blocks across pages.
    """
    for page in pages:
        sec_idx = page.section_index
        # Collect all blocks that belong to this section
        # For now, we don't have section assignment on blocks,
        # so we rely on the section index from the parser.
        # This will be populated by resolve_layout_state.
        pass


def resolve_horizontal_coordinate(
    img: Image,
    page: Page,
    column_index: Optional[int],
) -> Tuple[int, str]:
    """Resolve horizontal position to page-local CSS pixels.

    Returns (x_px, coordinate_space_used).
    """
    layout = page.page_layout
    if not layout:
        return 0, "page"

    rel_from = img.relative_from_horizontal or "page"
    offset_emu = img.offset_horizontal
    alignment = img.alignment_horizontal

    # Default coordinate space
    coord_space = rel_from if rel_from in SUPPORTED_RELATIVE_FROM else "page"

    if rel_from == "page":
        # Relative to page left edge
        if alignment == "center":
            return layout.page_width_px // 2, "page"
        elif alignment in ("right", "outside"):
            if offset_emu is not None:
                return layout.page_width_px - emu_to_px(offset_emu) - (img.width or 0), "page"
            return layout.page_width_px - (img.width or 0), "page"
        elif alignment in ("left", "inside"):
            if offset_emu is not None:
                return emu_to_px(offset_emu), "page"
            return 0, "page"
        else:
            # Offset-based
            return emu_to_px(offset_emu) if offset_emu is not None else 0, "page"

    elif rel_from in ("margin", "column"):
        # Relative to content box (margin box)
        content_left = layout.margin_left_px
        col_box = page.get_column_box_page_px(column_index or 0)
        if rel_from == "column" and col_box:
            col_left = col_box["left_px"]
            col_width = col_box["width_px"]
            if alignment == "center":
                return col_left + col_width // 2, "column"
            elif alignment in ("right", "outside"):
                if offset_emu is not None:
                    return col_left + col_width - emu_to_px(offset_emu) - (img.width or 0), "column"
                return col_left + col_width - (img.width or 0), "column"
            elif alignment in ("left", "inside"):
                if offset_emu is not None:
                    return col_left + emu_to_px(offset_emu), "column"
                return col_left, "column"
            else:
                return col_left + emu_to_px(offset_emu) if offset_emu is not None else col_left, "column"
        else:
            # margin-relative
            if alignment == "center":
                return content_left + layout.content_width_px // 2, "margin"
            elif alignment in ("right", "outside"):
                if offset_emu is not None:
                    return content_left + layout.content_width_px - emu_to_px(offset_emu) - (img.width or 0), "margin"
                return content_left + layout.content_width_px - (img.width or 0), "margin"
            elif alignment in ("left", "inside"):
                if offset_emu is not None:
                    return content_left + emu_to_px(offset_emu), "margin"
                return content_left, "margin"
            else:
                return content_left + emu_to_px(offset_emu) if offset_emu is not None else content_left, "margin"

    elif rel_from in ("paragraph", "character", "line"):
        # Relative to anchor paragraph - handled by paragraph renderer
        # Return 0 here; the paragraph renderer will position relative to itself
        return 0, "paragraph"

    # Unsupported relativeFrom - fall back to page
    return emu_to_px(offset_emu) if offset_emu is not None else 0, "page"


def resolve_vertical_coordinate(
    img: Image,
    page: Page,
    column_index: Optional[int],
) -> Tuple[int, str]:
    """Resolve vertical position to page-local CSS pixels.

    Returns (y_px, coordinate_space_used).
    """
    layout = page.page_layout
    if not layout:
        return 0, "page"

    rel_from = img.relative_from_vertical or "page"
    offset_emu = img.offset_vertical
    alignment = img.alignment_vertical

    if rel_from == "page":
        # Relative to page top edge
        if alignment == "center":
            return layout.page_height_px // 2, "page"
        elif alignment in ("bottom", "outside"):
            if offset_emu is not None:
                return layout.page_height_px - emu_to_px(offset_emu) - (img.height or 0), "page"
            return layout.page_height_px - (img.height or 0), "page"
        elif alignment in ("top", "inside"):
            if offset_emu is not None:
                return emu_to_px(offset_emu), "page"
            return 0, "page"
        else:
            return emu_to_px(offset_emu) if offset_emu is not None else 0, "page"

    elif rel_from in ("margin", "column"):
        # Relative to content box (margin box)
        content_top = layout.margin_top_px
        if alignment == "center":
            return content_top + layout.content_height_px // 2, "margin"
        elif alignment in ("bottom", "outside"):
            if offset_emu is not None:
                return content_top + layout.content_height_px - emu_to_px(offset_emu) - (img.height or 0), "margin"
            return content_top + layout.content_height_px - (img.height or 0), "margin"
        elif alignment in ("top", "inside"):
            if offset_emu is not None:
                return content_top + emu_to_px(offset_emu), "margin"
            return content_top, "margin"
        else:
            return content_top + emu_to_px(offset_emu) if offset_emu is not None else content_top, "margin"

    elif rel_from in ("paragraph", "character", "line"):
        # Relative to anchor paragraph - handled by paragraph renderer
        return 0, "paragraph"

    return emu_to_px(offset_emu) if offset_emu is not None else 0, "page"


def resolve_image_coordinates(
    img: Image,
    pages: List[Page],
    sections: List[Section],
) -> ResolvedCoordinate:
    """Resolve all coordinates for a floating image to page-local CSS pixels.

    This is the SINGLE authoritative coordinate transformation.
    """
    sec_idx = getattr(img, "section_index", 0)
    if sec_idx is None:
        sec_idx = 0

    # Find the page for this section (first page of section for now)
    page = None
    page_idx = 0
    for i, p in enumerate(pages):
        if p.section_index == sec_idx:
            page = p
            page_idx = i
            break
    if not page and pages:
        page = pages[0]
        page_idx = 0

    col_idx = getattr(img, "column_index", None)

    x_px, h_space = resolve_horizontal_coordinate(img, page, col_idx)
    y_px, v_space = resolve_vertical_coordinate(img, page, col_idx)

    # Determine overall coordinate space (use most outer)
    if h_space == "page" or v_space == "page":
        coord_space = "page"
    elif h_space == "column" or v_space == "column":
        coord_space = "column"
    elif h_space == "margin" or v_space == "margin":
        coord_space = "margin"
    else:
        coord_space = "paragraph"

    return ResolvedCoordinate(
        x_px=x_px,
        y_px=y_px,
        coordinate_space=coord_space,
        relative_from_horizontal=img.relative_from_horizontal,
        relative_from_vertical=img.relative_from_vertical,
        section_index=sec_idx,
        page_index=page_idx,
        column_index=col_idx,
    )


def resolve_layout_state(
    blocks: List[Block],
    sections: List[Section],
    pages: List[Page],
) -> LayoutState:
    """Resolve complete layout state for the document.

    This is the main entry point called by the pipeline.
    """
    state = LayoutState(sections=sections, pages=pages)

    # Assign blocks to pages (stub: all blocks in section -> section's page)
    for page in pages:
        sec_idx = page.section_index
        # In a real pagination engine, we'd compute which blocks fit on each page.
        # For now, we just track the section-page mapping.
        for block in blocks:
            # Block section assignment comes from parser via section_index on images
            # and paragraph order. We'll assign based on nearest image or position.
            pass

    # Resolve coordinates for all floating images
    for page in pages:
        for img in page.floating_images:
            coord = resolve_image_coordinates(img, pages, sections)
            state.image_coordinates[img.image_id] = coord

    # Also resolve for blocks (paragraphs, tables) - they inherit page from section
    for block in blocks:
        if isinstance(block, Paragraph):
            # Find section for this paragraph
            sec_idx = 0
            for img in block.images:
                if img.section_index is not None:
                    sec_idx = img.section_index
                    break
            # Find page for section
            page_idx = 0
            for i, p in enumerate(pages):
                if p.section_index == sec_idx:
                    page_idx = i
                    break
            if block.block_id:
                state.block_page_ownership[block.block_id] = (sec_idx, page_idx)

    return state


def get_page_for_section(sections: List[Section], pages: List[Page], section_index: int) -> Optional[Page]:
    """Get the first page for a given section."""
    for p in pages:
        if p.section_index == section_index:
            return p
    return None


def get_column_box_for_image(img: Image, page: Page) -> Optional[Dict[str, Any]]:
    """Get the column box for an image, handling column_index."""
    col_idx = getattr(img, "column_index", 0)
    if col_idx is None:
        col_idx = 0
    return page.get_column_box_page_px(col_idx)


def transform_coordinate(
    value_emu: int,
    relative_from: str,
    axis: str,  # "horizontal" | "vertical"
    page: Page,
    column_index: Optional[int] = None,
    alignment: Optional[str] = None,
    object_size_px: int = 0,
) -> Tuple[int, str]:
    """Low-level coordinate transformation: EMU offset + alignment → page-local px.

    This is the primitive used by resolve_horizontal_coordinate and resolve_vertical_coordinate.
    Exposed for any edge cases that need direct access.
    """
    layout = page.page_layout
    if not layout:
        return 0, "page"

    if relative_from == "page":
        origin = 0
        max_extent = layout.page_width_px if axis == "horizontal" else layout.page_height_px
    elif relative_from in ("margin", "column"):
        if relative_from == "column" and column_index is not None:
            col_box = page.get_column_box_page_px(column_index)
            if col_box:
                origin = col_box["left_px"]
                max_extent = col_box["width_px"]
            else:
                origin = layout.margin_left_px
                max_extent = layout.content_width_px
        else:
            origin = layout.margin_left_px if axis == "horizontal" else layout.margin_top_px
            max_extent = layout.content_width_px if axis == "horizontal" else layout.content_height_px
    else:
        # paragraph/character/line - origin is paragraph-relative (handled elsewhere)
        return 0, "paragraph"

    offset_px = emu_to_px(value_emu) if value_emu is not None else 0

    if alignment == "center":
        return origin + max_extent // 2, relative_from
    elif alignment in ("right", "outside", "bottom"):
        return origin + max_extent - offset_px - object_size_px, relative_from
    elif alignment in ("left", "inside", "top"):
        return origin + offset_px, relative_from
    else:
        return origin + offset_px, relative_from