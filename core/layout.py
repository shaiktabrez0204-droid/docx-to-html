"""First-class page/layout state and coordinate transformation.

This module is the SINGLE authoritative source for:
1. Building pages from sections (real multi-page pagination)
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
    Run,
)
from core.units import emu_to_px, twip_to_emu, EMU_PER_INCH

import re as _re

SUPPORTED_RELATIVE_FROM = {
    "page", "margin", "column", "character", "paragraph", "line",
}

DEFAULT_LINE_HEIGHT_PX = 15
MIN_PAGE_CONTENT_PX = 50
MAX_PAGINATION_ITERATIONS = 10000

try:
    from PIL import ImageFont as _PILImageFont  # type: ignore
    _HAS_PIL = True
except Exception:  # pragma: no cover
    _PILImageFont = None  # type: ignore
    _HAS_PIL = False

_FONT_CACHE: Dict[Tuple[str, int, int], Any] = {}
_FALLBACK_CHAR_CACHE: Dict[Tuple[float, str, bool, bool], int] = {}

_FONT_VARIANTS: Dict[str, Dict[str, str]] = {
    "arial": {"regular": "arial.ttf", "bold": "arialbd.ttf", "italic": "ariali.ttf", "bold_italic": "arialbi.ttf"},
    "calibri": {"regular": "calibri.ttf", "bold": "calibrib.ttf", "italic": "calibrii.ttf", "bold_italic": "calibriz.ttf"},
    "cambria": {"regular": "cambria.ttc", "bold": "cambriab.ttf", "italic": "cambriai.ttf", "bold_italic": "cambriaz.ttf"},
    "times": {"regular": "times.ttf", "bold": "timesbd.ttf", "italic": "timesi.ttf", "bold_italic": "timesbi.ttf"},
    "courier": {"regular": "cour.ttf", "bold": "courbd.ttf", "italic": "couri.ttf", "bold_italic": "courbi.ttf"},
    "consolas": {"regular": "consola.ttf", "bold": "consolab.ttf", "italic": "consolai.ttf", "bold_italic": "consolaz.ttf"},
    "georgia": {"regular": "georgia.ttf", "bold": "georgiab.ttf", "italic": "georgiai.ttf", "bold_italic": "georgiaz.ttf"},
    "verdana": {"regular": "verdana.ttf", "bold": "verdanab.ttf", "italic": "verdanai.ttf", "bold_italic": "verdanaz.ttf"},
    "tahoma": {"regular": "tahoma.ttf", "bold": "tahomabd.ttf", "italic": "tahoma.ttf", "bold_italic": "tahomabd.ttf"},
    "segoe": {"regular": "segoeui.ttf", "bold": "segoeuib.ttf", "italic": "segoeuii.ttf", "bold_italic": "segoeuiz.ttf"},
}


def _normalize_family(raw: Optional[str]) -> str:
    if not raw:
        return ""
    fam = raw.split(",")[0].strip().strip('"').strip("'").strip()
    return fam.lower()


def _resolve_font_filename(family: str, bold: bool, italic: bool) -> Optional[str]:
    fam = _normalize_family(family)
    group_key: Optional[str] = None
    if "arial" in fam:
        group_key = "arial"
    elif "calibri" in fam:
        group_key = "calibri"
    elif "cambria" in fam:
        group_key = "cambria"
    elif "georgia" in fam:
        group_key = "georgia"
    elif "times" in fam:
        group_key = "times"
    elif "courier" in fam or "monospace" in fam or "fixedsys" in fam:
        group_key = "courier"
    elif "consolas" in fam:
        group_key = "consolas"
    elif "verdana" in fam:
        group_key = "verdana"
    elif "tahoma" in fam:
        group_key = "tahoma"
    elif "segoe" in fam:
        group_key = "segoe"
    elif "helvetica" in fam:
        group_key = "arial"
    if "georgia" in fam:
        group_key = "georgia"
    if "times" in fam:
        group_key = "times"
    if group_key is None:
        return None
    if fam == "times new roman" or "times" in fam:
        group_key = "times"
    variant = "regular"
    if bold and italic:
        variant = "bold_italic"
    elif bold:
        variant = "bold"
    elif italic:
        variant = "italic"
    mapping = _FONT_VARIANTS.get(group_key, {})
    filename = mapping.get(variant) or mapping.get("regular")
    return filename


def _fallback_char_width(pt: float, family: str, bold: bool, italic: bool) -> int:
    key = (pt, (family or "").lower(), bold, italic)
    if key in _FALLBACK_CHAR_CACHE:
        return _FALLBACK_CHAR_CACHE[key]
    fam = (family or "").lower()
    if any(k in fam for k in ["courier", "consolas", "monospace", "fixedsys"]):
        base = 0.60
    elif any(k in fam for k in ["times", "georgia", "garamond", "cambria", "palatino", "serif"]):
        base = 0.47
    elif any(k in fam for k in ["arial", "helvetica"]):
        base = 0.52
    elif any(k in fam for k in ["calibri", "segoe", "verdana", "tahoma"]):
        base = 0.50
    else:
        base = 0.52
    if bold:
        base += 0.02
    if italic:
        base += 0.01
    w = max(1, round(pt * base * 96 / 72))
    _FALLBACK_CHAR_CACHE[key] = w
    return w


def _get_pil_font(family: Optional[str], pt: float, bold: bool, italic: bool):
    if not _HAS_PIL or _PILImageFont is None:
        return None
    size_px = max(1, round(pt * 96.0 / 72.0))
    filename = _resolve_font_filename(family or "", bold, italic)
    candidates: List[str] = []
    if filename:
        candidates.append(filename)
    raw = _normalize_family(family or "")
    if raw and raw not in (filename or ""):
        candidates.append(raw + ".ttf")
        candidates.append(raw)
    if "arial.ttf" not in candidates:
        candidates.append("arial.ttf")
    for cand in candidates:
        key = (cand.lower(), size_px, 0)
        if key in _FONT_CACHE:
            return _FONT_CACHE[key]
        try:
            font = _PILImageFont.truetype(cand, size_px)
            _FONT_CACHE[key] = font
            return font
        except Exception:
            continue
    return None


def _measure_text_width(text: str, run, default_pt: float) -> int:
    if not text:
        return 0
    pt = run.font_size / 2.0 if getattr(run, "font_size", None) is not None else default_pt
    bold = bool(getattr(run, "bold", False))
    italic = bool(getattr(run, "italic", False))
    family = getattr(run, "font_family", None) or ""
    font = _get_pil_font(family, pt, bold, italic)
    if font is not None:
        try:
            w = font.getlength(text)
            return max(1, int(round(w)))
        except Exception:
            pass
    cw = _fallback_char_width(pt, family, bold, italic)
    return max(1, len(text) * cw)


@dataclass
class ResolvedCoordinate:
    x_px: int
    y_px: int
    coordinate_space: str
    relative_from_horizontal: Optional[str]
    relative_from_vertical: Optional[str]
    section_index: int
    page_index: int
    column_index: Optional[int]


@dataclass
class LayoutState:
    sections: List[Section]
    pages: List[Page]
    image_coordinates: Dict[str, ResolvedCoordinate] = field(default_factory=dict)
    block_page_ownership: Dict[str, Tuple[int, int]] = field(default_factory=dict)


def _estimate_paragraph_height(para: Paragraph, page_layout: PageLayout, available_width_px: Optional[int] = None) -> int:
    import re
    max_font_pt = 11.0
    for run in getattr(para, 'runs', []) or []:
        if run.font_size is not None:
            pt = run.font_size / 2.0
            if pt > max_font_pt:
                max_font_pt = pt
    for item in getattr(para, 'content', []) or []:
        if isinstance(item, Run) and item.font_size is not None:
            pt = item.font_size / 2.0
            if pt > max_font_pt:
                max_font_pt = pt
    base_line_px = max(DEFAULT_LINE_HEIGHT_PX, round(max_font_pt * 1.15 * 96 / 72))
    line_height_px = base_line_px
    if para.line_spacing is not None:
        rule = para.line_spacing_rule or "auto"
        if rule == "auto":
            auto_px = max(1, round(para.line_spacing / 240.0 * 12 * 96 / 72))
            line_height_px = max(base_line_px, auto_px)
        elif rule == "exact":
            line_height_px = max(1, emu_to_px(twip_to_emu(para.line_spacing)))
        elif rule == "atLeast":
            at_least_px = emu_to_px(twip_to_emu(para.line_spacing))
            line_height_px = max(base_line_px, at_least_px)

    left_px = emu_to_px(twip_to_emu(para.indent_left)) if para.indent_left else 0
    right_px = emu_to_px(twip_to_emu(para.indent_right)) if para.indent_right else 0
    first_extra = emu_to_px(twip_to_emu(para.indent_first_line)) if para.indent_first_line is not None else 0
    hanging_px = emu_to_px(twip_to_emu(para.indent_hanging)) if para.indent_hanging is not None else 0
    if available_width_px is None:
        base_content_px = page_layout.content_width_px
    else:
        base_content_px = available_width_px
    first_avail = base_content_px - left_px - right_px - (first_extra if para.indent_first_line is not None else 0)
    other_avail = base_content_px - left_px - right_px - (hanging_px if para.indent_hanging is not None else 0)
    first_avail = max(MIN_PAGE_CONTENT_PX, first_avail)
    other_avail = max(MIN_PAGE_CONTENT_PX, other_avail)
    if first_avail < MIN_PAGE_CONTENT_PX:
        first_avail = MIN_PAGE_CONTENT_PX
    if other_avail < MIN_PAGE_CONTENT_PX:
        other_avail = MIN_PAGE_CONTENT_PX

    ordered = getattr(para, "content", None)
    if ordered and any(isinstance(c, (Run, Image)) for c in ordered):
        items = [c for c in ordered if isinstance(c, (Run, Image))]
    else:
        items = list(getattr(para, "runs", []) or [])
        for img in getattr(para, "images", []) or []:
            if getattr(img, "wrap_type", None) == "inline" and img not in items:
                items.append(img)

    has_text = False
    for it in items:
        if isinstance(it, Image) and getattr(it, "wrap_type", None) == "inline":
            has_text = True
            break
        if isinstance(it, Run) and (getattr(it, "text", "") or ""):
            has_text = True
            break
    segments_tokens: List[List[Tuple[str, int]]] = []
    cur_tokens: List[Tuple[str, int]] = []
    for it in items:
        if isinstance(it, Image) and getattr(it, "wrap_type", None) == "inline":
            w = int(getattr(it, "width", 0) or 0)
            cur_tokens.append(("image", max(0, w)))
            has_text = True
        elif isinstance(it, Run):
            text = getattr(it, "text", None) or ""
            parts = text.split("\n")
            for pi, part in enumerate(parts):
                if pi > 0:
                    segments_tokens.append(cur_tokens)
                    cur_tokens = []
                if not part:
                    continue
                tab_segs = part.split("\t")
                for ti, tseg in enumerate(tab_segs):
                    if ti > 0:
                        tab_w = _measure_text_width("    ", it, max_font_pt)
                        cur_tokens.append(("tab", tab_w))
                        has_text = True
                    if not tseg:
                        continue
                    for m in re.finditer(r' +|[^ ]+', tseg):
                        chunk = m.group(0)
                        if not chunk:
                            continue
                        w = _measure_text_width(chunk, it, max_font_pt)
                        if chunk[0] == ' ':
                            cur_tokens.append(("space", w))
                        else:
                            cur_tokens.append(("word", w))
        else:
            continue
    segments_tokens.append(cur_tokens)
    merged_segments: List[List[Tuple[str, int]]] = []
    for seg in segments_tokens:
        merged: List[Tuple[str, int]] = []
        for typ, w in seg:
            if merged and merged[-1][0] == typ and typ in ("word", "space"):
                pt, pw = merged[-1]
                merged[-1] = (pt, pw + w)
            elif merged and merged[-1][0] == "word" and typ == "word":
                pt, pw = merged[-1]
                merged[-1] = (pt, pw + w)
            else:
                if merged and merged[-1][0] == "word" and typ == "word":
                    pt, pw = merged[-1]
                    merged[-1] = (pt, pw + w)
                else:
                    merged.append((typ, w))
        coalesced: List[Tuple[str, int]] = []
        for typ, w in merged:
            if coalesced and coalesced[-1][0] == "word" and typ == "word":
                pt, pw = coalesced[-1]
                coalesced[-1] = (pt, pw + w)
            else:
                coalesced.append((typ, w))
        merged_segments.append(coalesced)

    estimated_lines = 0
    line_idx = 0
    for seg in merged_segments:
        if not seg:
            estimated_lines += 1
            line_idx += 1
            continue
        cur_width = 0
        avail = first_avail if line_idx == 0 else other_avail
        i = 0
        seg_lines = 0
        iterations = 0
        max_iter = len(seg) * 4 + 100
        while i < len(seg):
            iterations += 1
            if iterations > max_iter:
                seg_lines += 1
                break
            typ, w = seg[i]
            if typ == "space":
                if cur_width == 0:
                    i += 1
                    continue
                if cur_width + w <= avail:
                    cur_width += w
                    i += 1
                else:
                    seg_lines += 1
                    line_idx += 1
                    avail = first_avail if line_idx == 0 else other_avail
                    cur_width = 0
                    i += 1
                continue
            elif typ == "tab":
                if cur_width == 0:
                    cur_width = w
                    i += 1
                else:
                    if cur_width + w <= avail:
                        cur_width += w
                        i += 1
                    else:
                        seg_lines += 1
                        line_idx += 1
                        avail = first_avail if line_idx == 0 else other_avail
                        cur_width = 0
                        continue
                continue
            else:
                if cur_width == 0:
                    cur_width = w
                    i += 1
                else:
                    if cur_width + w <= avail:
                        cur_width += w
                        i += 1
                    else:
                        seg_lines += 1
                        line_idx += 1
                        avail = first_avail if line_idx == 0 else other_avail
                        cur_width = 0
                        continue
        if seg_lines == 0 and cur_width == 0:
            seg_lines = 1
            line_idx += 1
        elif cur_width > 0:
            seg_lines += 1
            line_idx += 1
        elif seg_lines == 0:
            seg_lines = 1
            line_idx += 1
        estimated_lines += seg_lines
    if estimated_lines == 0:
        estimated_lines = 1

    inline_imgs = [img for img in getattr(para, "images", []) or [] if getattr(img, "wrap_type", None) == "inline" and getattr(img, "height", None)]
    max_inline_h = 0
    for img in inline_imgs:
        h = img.height or 0
        if h > max_inline_h:
            max_inline_h = h
    if max_inline_h > line_height_px:
        line_height_effective = max_inline_h
        total_text_h = estimated_lines * line_height_px
        total_text_h = max(total_text_h, (estimated_lines - 1) * line_height_px + line_height_effective)
    else:
        total_text_h = estimated_lines * line_height_px

    space_before_px = emu_to_px(twip_to_emu(para.spacing_before)) if para.spacing_before else 0
    space_after_px = emu_to_px(twip_to_emu(para.spacing_after)) if para.spacing_after else 0

    total_height = space_before_px + total_text_h + space_after_px
    has_visible = has_text or bool(inline_imgs)
    if not has_visible:
        total_height = max(total_height, space_before_px + line_height_px + space_after_px)
    else:
        has_non_ws = False
        for it in items:
            if isinstance(it, Run) and (getattr(it, "text", "") or "").strip():
                has_non_ws = True
                break
            if isinstance(it, Image) and getattr(it, "wrap_type", None) == "inline":
                has_non_ws = True
                break
        if not has_non_ws and not inline_imgs:
            total_height = max(total_height, space_before_px + line_height_px + space_after_px)
    return max(total_height, line_height_px)


def _cell_available_width_px(cell: "Cell", table: "Table", page_layout: PageLayout) -> int:
    if cell.width is not None and cell.width_type not in (None, "auto", "nil", "pct"):
        try:
            w = emu_to_px(twip_to_emu(int(cell.width)))
            if w >= MIN_PAGE_CONTENT_PX:
                return max(MIN_PAGE_CONTENT_PX, w - 12)
        except Exception:
            pass
    grid = getattr(table, "grid_col_widths", None) or []
    if grid:
        valid = [v for v in grid if v is not None]
        if valid:
            avg_emu = sum(twip_to_emu(int(v)) for v in valid) // len(valid)
            span = max(1, getattr(cell, "grid_span", 1))
            span_px = emu_to_px(avg_emu * span)
            if getattr(table, "grid_col_widths", None):
                pass
            return max(MIN_PAGE_CONTENT_PX, span_px - 12)
    cols = getattr(table, "column_count", 0) or 1
    if cols < 1:
        cols = 1
    base = page_layout.content_width_px // cols
    return max(MIN_PAGE_CONTENT_PX, base - 12)


def _estimate_table_height(table: Table, page_layout: PageLayout) -> int:
    total_height = 0
    for row in table.rows:
        row_height = 0
        for cell in row.cells:
            avail = _cell_available_width_px(cell, table, page_layout)
            cell_height = 0
            for para in cell.content:
                cell_height += _estimate_paragraph_height(para, page_layout, available_width_px=avail)
            cell_height += 8
            row_height = max(row_height, cell_height)
        total_height += row_height
    total_height += len(table.rows) * 2
    return max(total_height, 20)


def _estimate_block_height(block: Block, page_layout: PageLayout) -> int:
    if isinstance(block, Paragraph):
        return _estimate_paragraph_height(block, page_layout)
    elif isinstance(block, Table):
        return _estimate_table_height(block, page_layout)
    return 0


def _ensure_block_ids(blocks: List[Block]) -> None:
    for i, b in enumerate(blocks):
        if getattr(b, "block_id", None) is None:
            b.block_id = f"blk-{i}"


def build_pages_from_sections(sections: List[Section]) -> List[Page]:
    pages: List[Page] = []
    for sec in sections:
        page_layout = sec.page_layout or PageLayout()
        page_layout.section_index = sec.index
        page = Page(
            page_index=len(pages),
            section_index=sec.index,
            page_layout=page_layout,
            content_origin_x_px=page_layout.margin_left_px,
            content_origin_y_px=page_layout.margin_top_px,
        )
        page.header_default = sec.get_header("default")
        page.header_first = sec.get_header("first")
        page.header_even = sec.get_header("even")
        page.footer_default = sec.get_footer("default")
        page.footer_first = sec.get_footer("first")
        page.footer_even = sec.get_footer("even")
        pages.append(page)
    return pages


def _block_needs_break(block: Block) -> bool:
    return bool(getattr(block, "page_break_before", False))


def _block_has_hard_break_after(block: Block) -> bool:
    return bool(getattr(block, "contains_page_break", False))


def paginate_section_blocks(
    blocks: List[Block],
    section: Section,
    existing_pages: List[Page],
    start_page_index: int,
) -> List[Page]:
    page_layout = section.page_layout or PageLayout()
    content_height_px = page_layout.content_height_px
    if content_height_px <= 0:
        content_height_px = page_layout.page_height_px - page_layout.margin_top_px - page_layout.margin_bottom_px
        if content_height_px <= 0:
            content_height_px = 800

    section_blocks = [b for b in blocks if getattr(b, "section_index", 0) == section.index]
    if not section_blocks:
        return existing_pages

    section_pages = [p for p in existing_pages if p.section_index == section.index]
    if not section_pages:
        return existing_pages

    current_page = section_pages[0]
    current_y = current_page.content_origin_y_px
    page_bottom = current_page.content_origin_y_px + content_height_px
    pending_break = False
    iteration = 0
    insert_idx = existing_pages.index(section_pages[-1]) + 1
    for block in section_blocks:
        iteration += 1
        if iteration > MAX_PAGINATION_ITERATIONS:
            break

        height = _estimate_block_height(block, page_layout)
        force_new_page = False
        if _block_needs_break(block):
            force_new_page = True
        if pending_break:
            force_new_page = True
            pending_break = False

        if force_new_page and current_y != current_page.content_origin_y_px:
            new_page = Page(
                page_index=len(existing_pages),
                section_index=section.index,
                page_layout=page_layout,
                content_origin_x_px=page_layout.margin_left_px,
                content_origin_y_px=page_layout.margin_top_px,
            )
            new_page.header_default = section.get_header("default")
            new_page.footer_default = section.get_footer("default")
            new_page.header_first = section.get_header("first")
            new_page.header_even = section.get_header("even")
            new_page.footer_first = section.get_footer("first")
            new_page.footer_even = section.get_footer("even")
            existing_pages.insert(insert_idx, new_page)
            section_pages.append(new_page)
            insert_idx += 1
            current_page = new_page
            current_y = current_page.content_origin_y_px
            page_bottom = current_y + content_height_px

        if current_y + height <= page_bottom:
            current_page.blocks.append(block)
            current_y += height
        else:
            if current_y == current_page.content_origin_y_px:
                current_page.blocks.append(block)
                current_y += height
            else:
                new_page = Page(
                    page_index=len(existing_pages),
                    section_index=section.index,
                    page_layout=page_layout,
                    content_origin_x_px=page_layout.margin_left_px,
                    content_origin_y_px=page_layout.margin_top_px,
                )
                new_page.header_default = section.get_header("default")
                new_page.footer_default = section.get_footer("default")
                new_page.header_first = section.get_header("first")
                new_page.header_even = section.get_header("even")
                new_page.footer_first = section.get_footer("first")
                new_page.footer_even = section.get_footer("even")
                existing_pages.insert(insert_idx, new_page)
                section_pages.append(new_page)
                insert_idx += 1
                current_page = new_page
                current_y = current_page.content_origin_y_px
                page_bottom = current_y + content_height_px
                current_page.blocks.append(block)
                current_y += height

        if _block_has_hard_break_after(block):
            pending_break = True

    return existing_pages


def assign_blocks_to_pages(blocks: List[Block], pages: List[Page], sections: List[Section]) -> None:
    _ensure_block_ids(blocks)
    for section in sections:
        section_pages = [p for p in pages if p.section_index == section.index]
        if not section_pages:
            continue
        for p in section_pages:
            p.blocks.clear()
        paginate_section_blocks(blocks, section, pages, section_pages[0].page_index)
    for idx, p in enumerate(pages):
        p.page_index = idx


def resolve_horizontal_coordinate(
    img: Image,
    page: Page,
    column_index: Optional[int],
) -> Tuple[int, str]:
    layout = page.page_layout
    if not layout:
        return 0, "page"
    rel_from = img.relative_from_horizontal or "page"
    offset_emu = img.offset_horizontal
    alignment = img.alignment_horizontal
    coord_space = rel_from if rel_from in SUPPORTED_RELATIVE_FROM else "page"
    if rel_from == "page":
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
            return emu_to_px(offset_emu) if offset_emu is not None else 0, "page"
    elif rel_from in ("margin", "column"):
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
        return 0, "paragraph"
    return emu_to_px(offset_emu) if offset_emu is not None else 0, "page"


def resolve_vertical_coordinate(
    img: Image,
    page: Page,
    column_index: Optional[int],
) -> Tuple[int, str]:
    layout = page.page_layout
    if not layout:
        return 0, "page"
    rel_from = img.relative_from_vertical or "page"
    offset_emu = img.offset_vertical
    alignment = img.alignment_vertical
    if rel_from == "page":
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
        return 0, "paragraph"
    return emu_to_px(offset_emu) if offset_emu is not None else 0, "page"


def _resolve_page_for_image(
    img: Image,
    pages: List[Page],
    block_page_ownership: Optional[Dict[str, Tuple[int, int]]],
    blocks_by_index: Optional[Dict[int, str]] = None,
) -> Tuple[Optional[Page], int]:
    if block_page_ownership:
        nearest = getattr(img, "nearest_block_id", None)
        if nearest and nearest in block_page_ownership:
            _, p_idx = block_page_ownership[nearest]
            for p in pages:
                if p.page_index == p_idx:
                    return p, p.page_index
        anchor_idx = getattr(img, "anchor_paragraph_index", None)
        if anchor_idx is not None and blocks_by_index is not None:
            try:
                bid = blocks_by_index.get(int(anchor_idx))
            except Exception:
                bid = None
            if bid and bid in block_page_ownership:
                _, p_idx = block_page_ownership[bid]
                for p in pages:
                    if p.page_index == p_idx:
                        return p, p.page_index
    return None, -1


def resolve_image_coordinates(
    img: Image,
    pages: List[Page],
    sections: List[Section],
    block_page_ownership: Optional[Dict[str, Tuple[int, int]]] = None,
) -> ResolvedCoordinate:
    sec_idx = getattr(img, "section_index", None)
    page = None
    page_idx = -1
    # Build blocks_by_index lazily for anchor fallback
    if block_page_ownership:
        nearest_block_id = getattr(img, "nearest_block_id", None)
        if nearest_block_id and nearest_block_id in block_page_ownership:
            sec_idx_resolved, page_idx = block_page_ownership[nearest_block_id]
            sec_idx = sec_idx_resolved
            for p in pages:
                if p.page_index == page_idx:
                    page = p
                    break
        if page is None:
            # anchor_paragraph_index fallback via ownership requires caller to supply blocks,
            # but image_coordinates call cannot; so ownership must be via nearest_block_id.
            # Keep page None to allow non-authoritative fallback only if no ownership match.
            pass
    if page is None and pages:
        # No authoritative page found; caller should have provided ownership.
        # Preserve geometry by using nearest available page only for coordinate math
        # when ownership is absent, but mark page_index as -1 so caller can detect.
        # This path is only for non-authoritative callers; authoritative path above
        # already returned when ownership existed.
        if sec_idx is not None:
            for p in pages:
                if p.section_index == sec_idx:
                    page = p
                    page_idx = p.page_index
                    break
        if page is None:
            # fallback to first page solely for geometry preservation, but do not
            # invent ownership
            page = pages[0]
            page_idx = pages[0].page_index
            if sec_idx is None:
                sec_idx = page.section_index
    if not page:
        return ResolvedCoordinate(x_px=0, y_px=0, coordinate_space="page", relative_from_horizontal=img.relative_from_horizontal, relative_from_vertical=img.relative_from_vertical, section_index=sec_idx if sec_idx is not None else 0, page_index=0, column_index=getattr(img, "column_index", None))
    col_idx = getattr(img, "column_index", None)
    x_px, h_space = resolve_horizontal_coordinate(img, page, col_idx)
    y_px, v_space = resolve_vertical_coordinate(img, page, col_idx)
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
        section_index=sec_idx if sec_idx is not None else page.section_index,
        page_index=page_idx,
        column_index=col_idx,
    )


def _collect_floating_images(blocks: List[Block]) -> List[Image]:
    out: List[Image] = []
    for b in blocks:
        if isinstance(b, Paragraph):
            for img in b.images:
                if img.wrap_type == "anchor":
                    out.append(img)
        elif isinstance(b, Table):
            for row in b.rows:
                for cell in row.cells:
                    for p in cell.content:
                        for img in p.images:
                            if img.wrap_type == "anchor":
                                out.append(img)
    return out


def resolve_layout_state(
    blocks: List[Block],
    sections: List[Section],
    pages: List[Page],
) -> LayoutState:
    _ensure_block_ids(blocks)
    state = LayoutState(sections=sections, pages=pages)
    assign_blocks_to_pages(blocks, pages, sections)
    for page in pages:
        page.floating_images = []
    blocks_by_index: Dict[int, str] = {}
    for idx, b in enumerate(blocks):
        if b.block_id:
            blocks_by_index[idx] = b.block_id
    for page in pages:
        for block in page.blocks:
            if block.block_id:
                state.block_page_ownership[block.block_id] = (page.section_index, page.page_index)
    for b in blocks:
        if isinstance(b, Paragraph):
            for img in b.images:
                if img.wrap_type != "anchor":
                    continue
                target_page = None
                target_idx = None
                bid = getattr(img, "nearest_block_id", None)
                if bid and bid in state.block_page_ownership:
                    _, target_idx = state.block_page_ownership[bid]
                else:
                    anchor_idx = getattr(img, "anchor_paragraph_index", None)
                    if anchor_idx is not None and anchor_idx in blocks_by_index:
                        abid = blocks_by_index[anchor_idx]
                        if abid in state.block_page_ownership:
                            _, target_idx = state.block_page_ownership[abid]
                    if target_idx is None and b.block_id and b.block_id in state.block_page_ownership:
                        _, target_idx = state.block_page_ownership[b.block_id]
                if target_idx is not None:
                    for p in pages:
                        if p.page_index == target_idx:
                            target_page = p
                            break
                if target_page is not None:
                    target_page.floating_images.append(img)
        elif isinstance(b, Table):
            for row in b.rows:
                for cell in row.cells:
                    for p in cell.content:
                        for img in p.images:
                            if img.wrap_type != "anchor":
                                continue
                            target_page = None
                            target_idx = None
                            bid = getattr(img, "nearest_block_id", None)
                            if bid and bid in state.block_page_ownership:
                                _, target_idx = state.block_page_ownership[bid]
                            if target_idx is None:
                                anchor_idx = getattr(img, "anchor_paragraph_index", None)
                                if anchor_idx is not None and anchor_idx in blocks_by_index:
                                    abid = blocks_by_index[anchor_idx]
                                    if abid in state.block_page_ownership:
                                        _, target_idx = state.block_page_ownership[abid]
                                if target_idx is None and b.block_id and b.block_id in state.block_page_ownership:
                                    _, target_idx = state.block_page_ownership[b.block_id]
                            if target_idx is not None:
                                for pg in pages:
                                    if pg.page_index == target_idx:
                                        target_page = pg
                                        break
                            if target_page is not None:
                                target_page.floating_images.append(img)
    for page in pages:
        for img in page.floating_images:
            coord = resolve_image_coordinates(img, pages, sections, state.block_page_ownership)
            state.image_coordinates[img.image_id] = coord
    return state


def get_page_for_section(sections: List[Section], pages: List[Page], section_index: int) -> Optional[Page]:
    for p in pages:
        if p.section_index == section_index:
            return p
    return None


def get_column_box_for_image(img: Image, page: Page) -> Optional[Dict[str, Any]]:
    col_idx = getattr(img, "column_index", 0)
    if col_idx is None:
        col_idx = 0
    return page.get_column_box_page_px(col_idx)


def transform_coordinate(
    value_emu: int,
    relative_from: str,
    axis: str,
    page: Page,
    column_index: Optional[int] = None,
    alignment: Optional[str] = None,
    object_size_px: int = 0,
) -> Tuple[int, str]:
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
