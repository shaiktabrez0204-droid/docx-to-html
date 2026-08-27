"""HTML renderer for the normalized docx-to-html model.

Consumes Paragraph/Run/Image objects (produced by the OOXML extraction boundary)
and emits HTML. No DOCX/OOXML knowledge lives here, and NO semantic decisions
are made here: heading level/id are already resolved on the Paragraph.

Floating (wp:anchor) images
---------------------------
A floating image references a DOCX coordinate system via relativeFrom
(page / margin / column / paragraph / character). The renderer maps those onto
determinate CSS pixel origins using the real page/margin geometry:

  * page-relative      -> absolute inside .docx-page (the full page box)
  * margin/column      -> absolute inside .docx-content (the margin box)
  * paragraph/character-> absolute inside the anchor paragraph (position:relative)

Wrap modes:
  * wrapNone           -> absolute overlay (text may pass under; that is the
                          DOCX intent for "no wrap")
  * wrapSquare/topAndBottom with a left/right align and a margin/paragraph
    coordinate system -> CSS `float` so text genuinely flows around the image
  * everything else    -> absolute positioning (a documented, measured
                          approximation; exact text flow is not a full layout
                          engine and is reported as such)

Positioning constants come only from the normalized model (EMU offsets via
core.units) and the PageLayout; no geometry is inferred from rendered pixels.
"""

import base64
import html
from typing import Optional

from core.model import Run, Paragraph, Table, Row, Cell, BorderEdge, HeaderFooter, Section, Image, ImageAsset, PageLayout, format_numbering_label
from core.units import emu_to_px

# Browser-safe MIME types that can be embedded via a data: URL. Others (e.g.
# EMF/WMF) are extracted but must degrade safely rather than break the page.
_BROWSER_SAFE_MIME = {
    "image/png",
    "image/jpeg",
    "image/gif",
    "image/bmp",
    "image/tiff",
    "image/webp",
    "image/svg+xml",
}

# Coordinates that resolve against the full page box (-> absolute in .docx-page).
_PAGE_COORDS = {"page"}
# Coordinates that resolve against the margin/content box (-> absolute in .docx-content).
_CONTENT_COORDS = {"margin", "column"}
# Coordinates that resolve against the anchor paragraph (-> absolute in <p>).
_PARAGRAPH_COORDS = {"paragraph", "character"}

# Alignment values that map to a horizontal float direction.
_LEFT_ALIGN = {"left", "inside"}
_RIGHT_ALIGN = {"right", "outside"}

_OOXML_BORDER_STYLE = {
    "single": "solid",
    "double": "double",
    "dotted": "dotted",
    "dashed": "dashed",
    "dotDash": "dashed",
    "dotDotDash": "dashed",
    "dashSmallGap": "dashed",
    "dashDotStroked": "dashed",
    "triple": "double",
    "thick": "solid",
    "thickThinSmallGap": "solid",
    "thinThickSmallGap": "solid",
    "wave": "solid",
    "inset": "inset",
    "outset": "outset",
    "nil": None,
    "none": None,
}

def _border_edge_to_css(edge) -> Optional[str]:
    if edge is None:
        return None
    if edge.val in ("nil", "none"):
        return "none"
    style = _OOXML_BORDER_STYLE.get(edge.val, "solid")
    if style is None:
        return "none"
    sz = edge.sz if edge.sz is not None else 4
    pt = sz / 8
    size_str = ("%gpt" % pt)
    col = edge.color
    if not col or col.lower() == "auto":
        col = "#000000"
    else:
        col = col.strip()
        if not col.startswith("#"):
            col = "#" + col
    return "%s %s %s" % (size_str, style, col)

def _effective_border(cell, table, edge: str, is_first_row: bool, is_last_row: bool, is_first_col: bool, is_last_col: bool):
    if cell.borders and edge in cell.borders:
        return cell.borders[edge]
    if table is not None and table.borders:
        if edge in ("top", "bottom") and "insideH" in table.borders:
            if edge == "top" and not is_first_row:
                return table.borders["insideH"]
            if edge == "bottom" and not is_last_row:
                return table.borders["insideH"]
        if edge in ("left", "right") and "insideV" in table.borders:
            if edge == "left" and not is_first_col:
                return table.borders["insideV"]
            if edge == "right" and not is_last_col:
                return table.borders["insideV"]
        if edge in table.borders:
            return table.borders[edge]
    return None


def _render_run_inner(run: Run) -> str:
    if not run.text:
        return ""
    content = run.text
    if run.bold is True:
        content = "<strong>%s</strong>" % content
    if run.italic is True:
        content = "<em>%s</em>" % content
    if run.underline:
        content = '<u style="text-decoration: %s">%s</u>' % (run.underline, content)
    style_parts = []
    if run.font_family and run.font_family != "Calibri":
        ff = run.font_family.replace("'", "\\'")
        style_parts.append("font-family:'%s'" % ff)
    if run.font_size is not None and run.font_size != 11:
        pt = run.font_size / 2
        pt_str = "%gpt" % pt
        style_parts.append("font-size:%s" % pt_str)
    if run.font_color and run.font_color not in ("#000000", "000000"):
        col = run.font_color.lstrip("#")
        style_parts.append("color:#%s" % col)
    if style_parts:
        content = '<span style="%s">%s</span>' % (";".join(style_parts), content)
    if run.superscript is True:
        content = "<sup>%s</sup>" % content
    elif run.subscript is True:
        content = "<sub>%s</sub>" % content
    return content


def render_run(run: Run) -> str:
    """Render a single Run, applying inline formatting and hyperlink."""
    inner = _render_run_inner(run)
    if not inner:
        return ""
    if getattr(run, "href", None):
        return '<a href="%s">%s</a>' % (html.escape(run.href, quote=True), inner)
    return inner


def _asset_src(img: Image, assets: Optional[dict]) -> Optional[str]:
    """Return a data: URL for the image's bytes, or None when it cannot render."""
    asset = (assets or {}).get(img.source_path)
    if asset is None or not asset.data or asset.media_type not in _BROWSER_SAFE_MIME:
        return None
    return "data:%s;base64,%s" % (
        asset.media_type,
        base64.b64encode(asset.data).decode("ascii"),
    )


def render_image(img: Image, assets: Optional[dict] = None) -> str:
    """Render one INLINE image placement as an <img> data URL.

    Missing/unsupported assets degrade to a labeled placeholder span (no broken
    network request). Floating images are rendered by render_float_image, not here.
    """
    alt_attr = ' alt="%s"' % html.escape(img.alt_text) if img.alt_text else ""
    aria = ' aria-label="%s"' % html.escape(img.alt_text) if img.alt_text else ""

    src = _asset_src(img, assets)
    if src is None:
        return '<span class="docx-image docx-image-missing" role="img"%s></span>' % aria

    attrs = ['src="%s"' % src]
    if img.width:
        attrs.append('width="%d"' % img.width)
    if img.height:
        attrs.append('height="%d"' % img.height)
    attrs.append(alt_attr)
    attrs.append('class="docx-image"')
    return "<img %s>" % " ".join(attrs)


# ---------------------------------------------------------------------------
# Floating image placement
# ---------------------------------------------------------------------------
def _float_container(img: Image) -> str:
    """Pick the CSS containing block for a floating image's coordinates.

    Uses the most *outer* relativeFrom present so a mixed page/margin anchor
    resolves consistently. Returns "page" | "content" | "paragraph".
    """
    coords = {img.relative_from_horizontal, img.relative_from_vertical}
    if coords & _PAGE_COORDS:
        return "page"
    if coords & _CONTENT_COORDS:
        return "content"
    if coords & _PARAGRAPH_COORDS:
        return "paragraph"
    # Default: page-relative (most common anchor intent).
    return "page"


def _is_wrap_float(img: Image) -> bool:
    """True when the wrap can use a real CSS float (genuine text flow)."""
    if img.wrap_mode not in ("square", "topAndBottom"):
        return False
    # A float needs a horizontal direction; only left/right aligns float.
    if img.alignment_horizontal in _LEFT_ALIGN:
        return True
    if img.alignment_horizontal in _RIGHT_ALIGN:
        return True
    return False


def _float_style(img: Image, page: PageLayout, container: str) -> str:
    """Build the CSS style string for a floating image (absolute or float)."""
    style: dict = {}

    # z-index derived solely from OOXML behindDoc: behind text -> -1, else 2.
    if img.behind_doc:
        style["z-index"] = "-1"
    else:
        style["z-index"] = "2"

    hoffs = emu_to_px(img.offset_horizontal)
    voffs = emu_to_px(img.offset_vertical)

    if _is_wrap_float(img):
        # Real CSS float: text flows around the rectangle. Margins reserve the
        # wrap distance so the rectangle is avoided, not overlapped.
        halign = img.alignment_horizontal
        style["float"] = "left" if halign in _LEFT_ALIGN else "right"
        d = img.wrap_distances or {}
        mt = emu_to_px(d.get("top", 0))
        mr = emu_to_px(d.get("right", 0))
        mb = emu_to_px(d.get("bottom", 0))
        ml = emu_to_px(d.get("left", 0))
        style["margin"] = "%dpx %dpx %dpx %dpx" % (mt, mr, mb, ml)
    else:
        # Absolute positioning within the chosen coordinate container.
        style["position"] = "absolute"
        # Horizontal
        halign = img.alignment_horizontal
        if halign == "center":
            style["left"] = "50%"
            tx = "-50%"
        elif halign in _RIGHT_ALIGN:
            style["right"] = str(hoffs) + "px" if img.offset_horizontal is not None else "0px"
            tx = "0"
        elif halign in _LEFT_ALIGN:
            style["left"] = str(hoffs) + "px" if img.offset_horizontal is not None else "0px"
            tx = "0"
        else:
            # Offset-based (or neither): resolve from the coordinate origin.
            # Offsets are from the container's own origin. .docx-content is
            # already inset by the page margin, so its origin IS the margin box
            # - no extra margin offset is needed here.
            if container == "page":
                style["left"] = "%dpx" % hoffs
            elif container == "content":
                style["left"] = "%dpx" % hoffs
            else:
                style["left"] = "%dpx" % hoffs
            tx = "0"
        # Vertical
        valign = img.alignment_vertical
        if valign == "center":
            style["top"] = "50%"
            ty = "-50%"
        elif valign in ("bottom", "outside"):
            style["bottom"] = str(voffs) + "px" if img.offset_vertical is not None else "0px"
            ty = "0"
        elif valign in ("top", "inside"):
            style["top"] = str(voffs) + "px" if img.offset_vertical is not None else "0px"
            ty = "0"
        else:
            if container == "page":
                style["top"] = "%dpx" % voffs
            elif container == "content":
                style["top"] = "%dpx" % voffs
            else:
                style["top"] = "%dpx" % voffs
            ty = "0"
        if tx != "0" or ty != "0":
            style["transform"] = "translate(%s, %s)" % (tx, ty)

    # Displayed geometry comes from the anchor extent (DOCX intent).
    if img.width:
        style["width"] = "%dpx" % img.width
    if img.height:
        style["height"] = "%dpx" % img.height

    return "; ".join("%s: %s" % (k, v) for k, v in style.items())


def render_float_image(img: Image, assets: Optional[dict], page: PageLayout) -> str:
    """Render a floating image placement with positioning/wrap CSS."""
    alt_attr = ' alt="%s"' % html.escape(img.alt_text) if img.alt_text else ""
    aria = ' aria-label="%s"' % html.escape(img.alt_text) if img.alt_text else ""

    src = _asset_src(img, assets)
    if src is None:
        return '<span class="docx-float docx-image-missing" role="img"%s></span>' % aria

    container = _float_container(img)
    style = _float_style(img, page, container)
    cls = "docx-float-wrapped" if _is_wrap_float(img) else "docx-float"
    attrs = [
        'src="%s"' % src,
        'class="%s"' % cls,
        'style="%s"' % style,
        alt_attr,
    ]
    return "<img %s>" % " ".join(attrs)


def _render_content(para: Paragraph, assets, page, collected):
    """Render ordered Run/Image content, splitting out container-level floats.

    Returns the inner HTML for the paragraph. Inline images and in-paragraph
    floats (paragraph-relative, or wrap floats) are rendered inline to preserve
    document order. page/content-absolute floats are appended to ``collected``
    as (container, img) so the caller can place them in the right box.

    Consecutive Runs sharing the same href are grouped into a single <a> to
    preserve the original OOXML hyperlink span.
    """
    items = para.content if para.content else para.runs
    parts = []
    needs_relative = False
    i = 0
    while i < len(items):
        item = items[i]
        if isinstance(item, Image):
            if item.wrap_type == "anchor":
                container = _float_container(item)
                if _is_wrap_float(item) or container == "paragraph":
                    if container == "paragraph":
                        needs_relative = True
                    parts.append(render_float_image(item, assets, page))
                else:
                    collected.append((container, item))
            else:
                parts.append(render_image(item, assets))
            i += 1
            continue
        href = getattr(item, "href", None)
        if href:
            group = []
            j = i
            while j < len(items) and not isinstance(items[j], Image) and getattr(items[j], "href", None) == href:
                group.append(items[j])
                j += 1
            inner = "".join(_render_run_inner(r) for r in group)
            if inner:
                parts.append('<a href="%s">%s</a>' % (html.escape(href, quote=True), inner))
            i = j
            continue
        parts.append(_render_run_inner(item))
        i += 1
    return "".join(parts), needs_relative


def _paragraph_layout_style(para: Paragraph) -> str:
    parts = []
    if para.alignment and para.alignment != "left":
        amap = {"left": "left", "center": "center", "right": "right", "justify": "justify", "both": "justify", "distribute": "justify", "start": "left", "end": "right"}
        val = amap.get(para.alignment, para.alignment)
        parts.append("text-align:%s" % val)
    left = para.indent_left or 0
    right = para.indent_right or 0
    first = para.indent_first_line
    hanging = para.indent_hanging
    if hanging is not None:
        ml = (left + hanging) / 20
        ti = -hanging / 20
        parts.append("margin-left:%gpt" % ml)
        parts.append("text-indent:%gpt" % ti)
    elif first is not None:
        if left:
            parts.append("margin-left:%gpt" % (left / 20))
        parts.append("text-indent:%gpt" % (first / 20))
    elif left:
        parts.append("margin-left:%gpt" % (left / 20))
    if right:
        parts.append("margin-right:%gpt" % (right / 20))
    if para.spacing_before is not None:
        parts.append("margin-top:%gpt" % (para.spacing_before / 20))
    if para.spacing_after is not None:
        parts.append("margin-bottom:%gpt" % (para.spacing_after / 20))
    if para.line_spacing is not None:
        rule = para.line_spacing_rule or "auto"
        if rule == "auto":
            lh = para.line_spacing / 240
            parts.append("line-height:%g" % lh)
        else:
            pt = para.line_spacing / 20
            parts.append("line-height:%gpt" % pt)
    return ";".join(parts)


def _is_list_item(para: Paragraph) -> bool:
    return para.heading_level is None and para.numbering_path is not None


def _list_label(para: Paragraph) -> Optional[str]:
    if para.numbering_format == "bullet":
        return "\u2022"
    return format_numbering_label(
        para.numbering_path, para.numbering_level_formats, para.numbering_text_pattern)


def _render_paragraph_html(para: Paragraph, assets=None, page=None):
    """Render one paragraph (or heading) including its in-paragraph floats.

    Returns (html, needs_relative, external_floats) where external_floats is a
    list of (container, Image) for page/content-absolute images that must be
    placed by the caller inside .docx-page / .docx-content.
    """
    if page is None:
        page = PageLayout()
    collected = []
    inner, needs_relative = _render_content(para, assets, page, collected)
    layout = _paragraph_layout_style(para)
    extra_styles = []
    if needs_relative:
        extra_styles.append("position: relative")
    if layout:
        extra_styles.append(layout)
    style_attr = ""
    if extra_styles:
        style_attr = ' style="%s"' % ";".join(extra_styles)

    if para.heading_level:
        n = max(1, min(para.heading_level, 6))
        hid = para.heading_id or ""
        attr = ' id="%s"' % hid if hid else ""
        label = None
        if para.numbering_path is not None and para.numbering_format != "bullet":
            label = format_numbering_label(
                para.numbering_path, para.numbering_level_formats, para.numbering_text_pattern)
        if label:
            inner = '<span class="docx-number">%s</span> %s' % (html.escape(label), inner)
        return ("<h%d%s%s>%s</h%d>" % (n, attr, style_attr, inner, n),
                needs_relative, collected)

    if _is_list_item(para):
        label = _list_label(para)
        if label:
            cls = "docx-bullet" if para.numbering_format == "bullet" else "docx-number"
            inner = '<span class="%s">%s</span> %s' % (cls, html.escape(label), inner)

    if not para.content and not para.runs:
        if style_attr:
            return ("<p%s></p>" % style_attr, needs_relative, collected)
        return "<p></p>", needs_relative, collected

    return ("<p%s>%s</p>" % (style_attr, inner), needs_relative, collected)


def _dxa_to_px(dxa: Optional[int]) -> int:
    if dxa is None:
        return 0
    return round(dxa * 635 / 9525)


def _render_cell_content(cell: Cell, assets, page) -> str:
    parts = []
    for para in cell.content:
        collected = []
        inner, needs_relative = _render_content(para, assets, page, collected)
        layout = _paragraph_layout_style(para)
        extra = []
        if needs_relative:
            extra.append("position: relative")
        if layout:
            extra.append(layout)
        style_attr = (' style="%s"' % ";".join(extra)) if extra else ""
        if para.heading_level:
            n = max(1, min(para.heading_level, 6))
            hid = para.heading_id or ""
            attr = ' id="%s"' % hid if hid else ""
            label = None
            if para.numbering_path is not None and para.numbering_format != "bullet":
                label = format_numbering_label(
                    para.numbering_path, para.numbering_level_formats, para.numbering_text_pattern)
            if label:
                inner = '<span class="docx-number">%s</span> %s' % (html.escape(label), inner)
            parts.append("<h%d%s%s>%s</h%d>" % (n, attr, style_attr, inner, n))
        else:
            if _is_list_item(para):
                label = _list_label(para)
                if label:
                    cls = "docx-bullet" if para.numbering_format == "bullet" else "docx-number"
                    inner = '<span class="%s">%s</span> %s' % (cls, html.escape(label), inner)
            if not inner.strip():
                if style_attr:
                    parts.append("<p%s></p>" % style_attr)
                else:
                    parts.append("<p></p>")
            else:
                parts.append("<p%s>%s</p>" % (style_attr, inner))
        for _container, img in collected:
            parts.append(render_float_image(img, assets, page))
    return "".join(parts)


def render_table(table: Table, assets=None, page=None) -> str:
    if page is None:
        page = PageLayout()
    style_parts = []
    if table.width and table.width_type != "auto":
        px = _dxa_to_px(table.width)
        if px:
            style_parts.append("width: %dpx" % px)
    table_style = "; ".join(style_parts)
    style_attr = ' style="%s"' % table_style if table_style else ""
    colgroup = ""
    if table.grid_col_widths:
        cols = []
        for w in table.grid_col_widths:
            if w:
                px = _dxa_to_px(w)
                cols.append('<col style="width: %dpx">' % px)
            else:
                cols.append("<col>")
        colgroup = "<colgroup>%s</colgroup>" % "".join(cols)
    has_any_table_border = table.borders is not None or any(c.borders is not None for r in table.rows for c in r.cells)
    rows_html = []
    for r_idx, row in enumerate(table.rows):
        cells_html = []
        is_first_row = r_idx == 0
        is_last_row = r_idx == len(table.rows) - 1
        for c_idx, cell in enumerate(row.cells):
            if cell.v_merge == "continue":
                continue
            is_first_col = c_idx == 0
            is_last_col = c_idx == len(row.cells) - 1
            attrs = []
            if cell.grid_span and cell.grid_span > 1:
                attrs.append('colspan="%d"' % cell.grid_span)
            if cell.row_span and cell.row_span > 1:
                attrs.append('rowspan="%d"' % cell.row_span)
            cell_style_parts = []
            if cell.vertical_align:
                va = cell.vertical_align
                if va == "center":
                    va = "middle"
                cell_style_parts.append("vertical-align: %s" % va)
            if cell.shading:
                col = cell.shading.strip()
                if not col.startswith("#"):
                    col = "#" + col
                cell_style_parts.append("background-color: %s" % col)
            if cell.width and cell.width_type != "auto":
                px = _dxa_to_px(cell.width)
                if px:
                    cell_style_parts.append("width: %dpx" % px)
            for edge in ("top", "left", "bottom", "right"):
                be = _effective_border(cell, table, edge, is_first_row, is_last_row, is_first_col, is_last_col)
                if be is not None:
                    css = _border_edge_to_css(be)
                    if css is not None:
                        if css == "none":
                            cell_style_parts.append("border-%s: none" % edge)
                        else:
                            cell_style_parts.append("border-%s: %s" % (edge, css))
                    else:
                        cell_style_parts.append("border-%s: none" % edge)
                else:
                    if not has_any_table_border:
                        cell_style_parts.append("border-%s: 1px solid #999" % edge)
                    else:
                        cell_style_parts.append("border-%s: none" % edge)
            style_a = ' style="%s"' % "; ".join(cell_style_parts) if cell_style_parts else ""
            inner = _render_cell_content(cell, assets, page)
            cells_html.append("<td%s%s>%s</td>" % (" ".join([""] + attrs) if attrs else "", style_a, inner))
        rows_html.append("<tr>%s</tr>" % "".join(cells_html))
    return '<table class="docx-table"%s>%s<tbody>%s</tbody></table>' % (style_attr, colgroup, "".join(rows_html))


def render_paragraph(para: Paragraph, assets: Optional[dict] = None) -> str:
    """Public single-paragraph render (kept for callers that render one block).

    Container-level (page/content) floats are dropped here because there is no
    page context; the full-document render_html places them correctly.
    """
    html_str, _n, _f = _render_paragraph_html(para, assets)
    return html_str


def render_toc(toc_entries: list) -> str:
    if not toc_entries:
        return ""
    items = []
    for e in toc_entries:
        text = html.escape(e["text"] or "")
        label = e.get("numbering_label")
        if label:
            text = '<span class="docx-number">%s</span> %s' % (html.escape(label), text)
        link = '<a href="#%s">%s</a>' % (e["id"], text)
        children = render_toc(e.get("children", []))
        items.append("<li>%s%s</li>" % (link, children))
    return '<ul class="toc-list">%s</ul>' % "".join(items)


def render_sidebar_toc(toc_entries: list) -> str:
    """Sidebar tree: hierarchical H1-H6, numbering preserved, no duplicate IDs."""
    if not toc_entries:
        return ""
    items = []
    for e in toc_entries:
        text = html.escape(e["text"] or "")
        label = e.get("numbering_label")
        if label:
            text = '<span class="docx-number">%s</span> %s' % (html.escape(label), text)
        lvl = int(e.get("level") or 1)
        link = '<a class="toc-link" href="#%s" data-level="%d">%s</a>' % (e["id"], lvl, text)
        children = e.get("children", [])
        if children:
            child_html = render_sidebar_toc(children)
            items.append(
                '<li class="toc-parent" role="treeitem" aria-expanded="true" data-level="%d">'
                '<div class="toc-row">'
                '<button class="toc-toggle" aria-expanded="true" aria-controls="toc-%s" aria-label="Toggle section"></button>'
                '%s'
                '</div>'
                '<ul id="toc-%s" class="toc-children" role="group">%s</ul>'
                '</li>' % (lvl, e["id"], link, e["id"], child_html)
            )
        else:
            items.append('<li class="toc-leaf" role="treeitem" data-level="%d">%s</li>' % (lvl, link))
    return '<ul class="toc-tree" role="tree">%s</ul>' % "".join(items)


def _render_hf_blocks(blocks, assets, page):
    parts = []
    for b in blocks:
        if isinstance(b, Table):
            parts.append(render_table(b, assets, page))
        elif isinstance(b, Paragraph):
            html_str, _needs, _ext = _render_paragraph_html(b, assets, page)
            parts.append(html_str)
    return "".join(parts)


def render_header_footer(hf: HeaderFooter, assets=None, page=None) -> str:
    if page is None:
        page = PageLayout()
    inner = _render_hf_blocks(hf.blocks, assets, page)
    tag = "header" if hf.kind == "header" else "footer"
    return '<%s class="docx-%s docx-%s-%s" data-type="%s">%s</%s>' % (tag, hf.kind, hf.kind, hf.hf_type, hf.hf_type, inner, tag)


def _viewer_style() -> str:
    return (
        "  <style>\n"
        "    *{box-sizing:border-box}\n"
        "    html,body{margin:0;padding:0;height:100%;}\n"
        "    body{font:14px/1.5 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;color:#1a1a1a;background:#fff;overflow:hidden;}\n"
        "    .page-frame{display:flex;flex-direction:column;height:100vh;overflow:hidden;}\n"
        "    .viewer-title{padding:10px 16px 9px;border-bottom:1px solid #e5e7eb;background:#fff;flex-shrink:0;}\n"
        "    .viewer-title .doc-title{margin:0;font-size:14px;font-weight:600;line-height:1.4;color:#111;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}\n"
        "    .viewer{display:flex;flex:1;min-height:0;overflow:hidden;}\n"
        "    .viewer-sidebar{width:280px;min-width:280px;max-width:280px;background:#fcfcfc;border-right:1px solid #e5e7eb;display:flex;flex-direction:column;overflow:hidden;flex-shrink:0;transition:width .18s ease,min-width .18s ease,opacity .18s ease,transform .18s ease;}\n"
        "    .viewer.viewer--sidebar-collapsed .viewer-sidebar{width:0;min-width:0;max-width:0;border-right-width:0;opacity:0;pointer-events:none;transform:translateX(-8px);overflow:hidden;}\n"
        "    .sidebar-header{display:flex;align-items:center;gap:8px;padding:10px 12px;border-bottom:1px solid #e5e7eb;font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#374151;white-space:nowrap;flex-shrink:0;}\n"
        "    .sidebar-title{flex:1;overflow:hidden;text-overflow:ellipsis;}\n"
        "    .toc{flex:1;overflow-y:auto;overflow-x:hidden;padding:8px;}\n"
        "    .toc-empty{color:#888;font-size:13px;padding:12px;}\n"
        "    .toc-tree{list-style:none;margin:0;padding:0;}\n"
        "    .toc-tree ul{list-style:none;margin:0;padding-left:14px;}\n"
        "    .toc-row{display:flex;align-items:center;gap:2px;}\n"
        "    .toc-link{flex:1;display:block;padding:4px 6px;font-size:13px;color:#374151;text-decoration:none;border-radius:4px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}\n"
        "    .toc-link:hover{background:#eef1f5;color:#111;}\n"
        "    .toc-link.is-active{background:#dbeafe;color:#1d4ed8;font-weight:600;}\n"
        "    .toc-link:focus-visible{outline:2px solid #2563eb;outline-offset:1px;}\n"
        "    .toc-parent{margin:1px 0;}\n"
        "    .toc-parent.is-collapsed>ul.toc-children{display:none;}\n"
        "    .toc-toggle{appearance:none;border:0;background:transparent;cursor:pointer;width:20px;height:20px;flex-shrink:0;display:inline-flex;align-items:center;justify-content:center;font-size:10px;color:#6b7280;border-radius:3px;transition:transform .15s;}\n"
        "    .toc-toggle::before{content:\"\\25BC\";display:block;font-size:9px;transition:transform .15s;}\n"
        "    .toc-parent.is-collapsed>.toc-row .toc-toggle::before{transform:rotate(-90deg);}\n"
        "    .toc-toggle:hover{background:#e5e7eb;color:#111;}\n"
        "    .toc-toggle:focus-visible{outline:2px solid #2563eb;}\n"
        "    .toc-leaf{padding-left:22px;}\n"
        "    .doc-main{flex:1;overflow-y:auto;overflow-x:hidden;background:#efefef;padding:0;display:flex;flex-direction:column;align-items:center;min-width:0;}\n"
        "    .doc-toolbar{position:sticky;top:0;z-index:5;width:100%;max-width:860px;display:flex;align-items:center;gap:8px;padding:8px 12px;background:rgba(239,239,239,.92);backdrop-filter:blur(6px);}\n"
        "    .sidebar-toggle-main{appearance:none;border:1px solid #d1d5db;background:#fff;cursor:pointer;padding:6px 10px;border-radius:6px;font-size:13px;line-height:1;display:inline-flex;align-items:center;gap:6px;color:#374151;box-shadow:0 1px 2px rgba(0,0,0,.06);}\n"
        "    .sidebar-toggle-main:hover{background:#f3f4f6;}\n"
        "    .sidebar-toggle-main:focus-visible{outline:2px solid #2563eb;outline-offset:2px;}\n"
        "    .docx-page{background:#fff;margin:16px auto 40px;box-shadow:0 1px 6px rgba(0,0,0,.08),0 0 0 1px rgba(0,0,0,.04);width:100%;max-width:860px;position:relative;flex-shrink:0;box-sizing:border-box;}\n"
        "    .docx-page[style]{max-width:min(860px,calc(100% - 32px)) !important;}\n"
        "    .docx-content{box-sizing:border-box;max-width:100%;overflow-wrap:break-word;}\n"
        "    img.docx-float{position:absolute;}\n"
        "    img.docx-float-wrapped{position:static;max-width:none;}\n"
        "    .docx-number{font:inherit;font-weight:inherit;font-style:inherit;margin-right:.4em;white-space:nowrap;}\n"
        "    .docx-bullet{margin-right:.4em;}\n"
        "    nav.toc .docx-number{margin-right:.4em;}\n"
        "    .docx-list{margin:0.6em 0 0.6em 1.8em;padding:0;}\n"
        "    .docx-list li{margin:0.15em 0;}\n"
        "    .docx-list li p{margin:0;}\n"
        "    .docx-ordered-list{list-style:none;padding-left:0;}\n"
        "    .docx-bullet-list{list-style:none;padding-left:0;}\n"
        "    table.docx-table{border-collapse:collapse;width:100%;margin:1em 0;}\n"
        "    table.docx-table td,table.docx-table th{border:1px solid #999;padding:6px 8px;vertical-align:top;word-wrap:break-word;}\n"
        "    table.docx-table td p{margin:0;}\n"
        "    header.docx-header,footer.docx-footer{border:1px dashed #bbb;padding:6px 8px;margin:6px 0;background:#fafafa;}\n"
        "    header.docx-header p,footer.docx-footer p{margin:0;}\n"
        "    .sidebar-overlay{position:fixed;inset:0;background:rgba(0,0,0,.32);z-index:15;}\n"
        "    @media (max-width:780px){.viewer-sidebar{position:fixed;left:0;top:0;bottom:0;z-index:20;width:280px;min-width:280px;max-width:280px;transform:translateX(-100%);transition:transform .22s ease;opacity:1;pointer-events:auto;border-right:1px solid #e5e7eb;}\n"
        "    .viewer.viewer--mobile-open .viewer-sidebar{transform:translateX(0);}\n"
        "    .viewer.viewer--sidebar-collapsed .viewer-sidebar{width:280px;min-width:280px;max-width:280px;opacity:1;pointer-events:auto;transform:translateX(-100%);border-right-width:1px;}\n"
        "    .viewer.viewer--mobile-open.viewer--sidebar-collapsed .viewer-sidebar{transform:translateX(0);}\n"
        "    .doc-main{padding-top:0;}}\n"
        "    @media (prefers-reduced-motion:reduce){*{transition:none!important;}}\n"
        "  </style>\n"
    )


def _viewer_script() -> str:
    return (
        "  <script>\n"
        "  (function(){\n"
        "    var viewer=document.querySelector('.viewer');\n"
        "    var sidebar=document.getElementById('viewer-sidebar');\n"
        "    var overlay=document.getElementById('sidebar-overlay');\n"
        "    var toggle=document.getElementById('sidebar-toggle');\n"
        "    var docMain=document.getElementById('doc-main');\n"
        "    if(!viewer||!sidebar||!overlay||!toggle||!docMain) return;\n"
        "    var mq=window.matchMedia('(max-width: 780px)');\n"
        "    function isMobile(){return mq.matches;}\n"
        "    function setSidebarOpen(open){\n"
        "      if(isMobile()){\n"
        "        viewer.classList.toggle('viewer--mobile-open',open);\n"
        "        overlay.hidden=!open;\n"
        "        viewer.classList.toggle('viewer--sidebar-collapsed',!open);\n"
        "      } else {\n"
        "        viewer.classList.toggle('viewer--sidebar-collapsed',!open);\n"
        "        viewer.classList.remove('viewer--mobile-open');\n"
        "        overlay.hidden=true;\n"
        "      }\n"
        "      toggle.setAttribute('aria-expanded',String(open));\n"
        "      toggle.setAttribute('aria-label',open?'Close outline':'Open outline');\n"
        "    }\n"
        "    function isSidebarOpen(){\n"
        "      if(isMobile()) return viewer.classList.contains('viewer--mobile-open');\n"
        "      return !viewer.classList.contains('viewer--sidebar-collapsed');\n"
        "    }\n"
        "    toggle.addEventListener('click',function(){setSidebarOpen(!isSidebarOpen());});\n"
        "    overlay.addEventListener('click',function(){setSidebarOpen(false);});\n"
        "    document.addEventListener('keydown',function(e){if(e.key==='Escape'&&isMobile()&&isSidebarOpen()) setSidebarOpen(false);});\n"
        "    setSidebarOpen(!isMobile());\n"
        "    var mqListener=function(){setSidebarOpen(!isMobile());};\n"
        "    if(mq.addEventListener) mq.addEventListener('change',mqListener); else if(mq.addListener) mq.addListener(mqListener);\n"
        "    var tocNav=document.querySelector('nav.toc');\n"
        "    if(tocNav){\n"
        "      tocNav.addEventListener('click',function(e){\n"
        "        var btn=e.target.closest('.toc-toggle');\n"
        "        if(!btn) return;\n"
        "        var li=btn.closest('.toc-parent');\n"
        "        if(!li) return;\n"
        "        var collapsed=li.classList.toggle('is-collapsed');\n"
        "        var expanded=!collapsed;\n"
        "        btn.setAttribute('aria-expanded',String(expanded));\n"
        "        li.setAttribute('aria-expanded',String(expanded));\n"
        "        var uid=btn.getAttribute('aria-controls');\n"
        "        var ul=document.getElementById(uid);\n"
        "        if(ul) ul.hidden=!expanded;\n"
        "      });\n"
        "    }\n"
        "    var titleEl=document.querySelector('.viewer-title .doc-title[id]');\n"
        "    var titleId=titleEl?titleEl.id:null;\n"
        "    var headings=Array.from(docMain.querySelectorAll('h1[id],h2[id],h3[id],h4[id],h5[id],h6[id]'));\n"
        "    var linkById={};\n"
        "    document.querySelectorAll('a.toc-link[href^=\"#\"]').forEach(function(a){linkById[a.getAttribute('href').slice(1)]=a;});\n"
        "    var activeId=null;\n"
        "    function setActive(id){\n"
        "      if(activeId===id) return;\n"
        "      if(activeId&&linkById[activeId]){linkById[activeId].classList.remove('is-active');linkById[activeId].removeAttribute('aria-current');}\n"
        "      activeId=id;\n"
        "      var link=linkById[id];\n"
        "      if(!link) return;\n"
        "      link.classList.add('is-active');\n"
        "      link.setAttribute('aria-current','location');\n"
        "      var el=link.closest('li');\n"
        "      while(el){\n"
        "        var parent=el.parentElement.closest('li.toc-parent');\n"
        "        if(parent){\n"
        "          parent.classList.remove('is-collapsed');\n"
        "          var b=parent.querySelector(':scope > .toc-row > .toc-toggle');\n"
        "          if(!b) b=parent.querySelector('.toc-toggle');\n"
        "          if(b){b.setAttribute('aria-expanded','true');var u=document.getElementById(b.getAttribute('aria-controls'));if(u) u.hidden=false;}\n"
        "          parent.setAttribute('aria-expanded','true');\n"
        "          el=parent;\n"
        "        } else break;\n"
        "      }\n"
        "      try{link.scrollIntoView({block:'nearest'});}catch(e){}\n"
        "    }\n"
        "    document.addEventListener('click',function(e){\n"
        "      var a=e.target.closest('a.toc-link');\n"
        "      if(!a) return;\n"
        "      var href=a.getAttribute('href');\n"
        "      if(!href||href.charAt(0)!=='#') return;\n"
        "      var id=href.slice(1);\n"
        "      var target=document.getElementById(id);\n"
        "      if(!target) return;\n"
        "      e.preventDefault();\n"
        "      if(id===titleId){\n"
        "        docMain.scrollTo({top:0,behavior:'smooth'});\n"
        "      } else {\n"
        "        target.scrollIntoView({behavior:'smooth',block:'start'});\n"
        "      }\n"
        "      history.pushState(null,'',href);\n"
        "      setActive(id);\n"
        "      if(isMobile()) setSidebarOpen(false);\n"
        "    });\n"
        "    if(headings.length && 'IntersectionObserver' in window){\n"
        "      var observer=new IntersectionObserver(function(entries){\n"
        "        var visible=entries.filter(function(en){return en.isIntersecting;});\n"
        "        if(!visible.length) return;\n"
        "        visible.sort(function(a,b){return a.boundingClientRect.top - b.boundingClientRect.top;});\n"
        "        var top=visible[0].target;\n"
        "        if(docMain.scrollTop<24 && titleId){ setActive(titleId); return; }\n"
        "        if(top&&top.id) setActive(top.id);\n"
        "      },{root:docMain,rootMargin:'-10% 0px -70% 0px',threshold:[0,1]});\n"
        "      headings.forEach(function(h){observer.observe(h);});\n"
        "      docMain.addEventListener('scroll',function(){ if(docMain.scrollTop<24 && titleId) setActive(titleId); },{passive:true});\n"
        "      var hval=location.hash.slice(1);\n"
        "      if(hval&&linkById[hval]) setActive(hval);\n"
        "      else if(titleId) setActive(titleId);\n"
        "    } else {\n"
        "      var onScroll=function(){\n"
        "        if(docMain.scrollTop<24 && titleId){ setActive(titleId); return; }\n"
        "        var cur=null;\n"
        "        var top=docMain.scrollTop+120;\n"
        "        for(var i=0;i<headings.length;i++){var hh=headings[i];if(hh.offsetTop<=top) cur=hh.id; else break;}\n"
        "        if(cur) setActive(cur);\n"
        "        else if(titleId) setActive(titleId);\n"
        "      };\n"
        "      docMain.addEventListener('scroll',onScroll,{passive:true});\n"
        "      onScroll();\n"
        "    }\n"
        "    window.__viewer={setSidebarOpen:setSidebarOpen,isSidebarOpen:isSidebarOpen,setActive:setActive};\n"
        "  })();\n"
        "  </script>\n"
        "\n"
    )


def render_html(blocks=None, title: str = "Converted Document", toc=None,
                assets=None, page_layout=None, paragraphs=None, sections=None, even_headers=False) -> str:
    """Render a full HTML document from normalized blocks.

    `blocks` is a document-order list of Paragraph | Table. Heading detection
    and TOC have already run on the paragraph subset; this renderer only lays
    out blocks in the order supplied. Floating images are placed in the correct
    coordinate container as before.

    `paragraphs` is a legacy alias for `blocks` (backward compat with tests
    that call render_html(paragraphs=...)).
    `sections` is an optional list of Section with header/footer variants.
    """
    if blocks is None:
        blocks = paragraphs if paragraphs is not None else []
    if page_layout is None:
        page_layout = PageLayout()

    first_h1 = None
    for b in blocks:
        if isinstance(b, Paragraph) and b.heading_level == 1:
            first_h1 = b
            break
    title_bar = ""
    if first_h1 is not None and first_h1.heading_id:
        src = first_h1.content if first_h1.content else first_h1.runs
        parts = []
        for it in src:
            if isinstance(it, Run):
                parts.append(it.text)
            elif isinstance(it, str):
                parts.append(it)
        heading_text = "".join(parts)
        label = None
        if first_h1.numbering_path is not None and first_h1.numbering_format != "bullet":
            label = format_numbering_label(first_h1.numbering_path, first_h1.numbering_level_formats, first_h1.numbering_text_pattern)
        if label:
            title_inner = '<span class="docx-number">%s</span> %s' % (html.escape(label), html.escape(heading_text))
        else:
            title_inner = html.escape(heading_text)
        title_bar = '<header class="viewer-title"><h1 id="%s" class="doc-title">%s</h1></header>' % (html.escape(first_h1.heading_id, quote=True), title_inner)

    content_floats = []   # (container, img) with container == "content"
    page_floats = []      # container == "page"
    block_html = []
    idx = 0
    while idx < len(blocks):
        b = blocks[idx]
        if b is first_h1:
            _, _, externals = _render_paragraph_html(b, assets, page_layout)
            for container, img in externals:
                if container == "page":
                    page_floats.append(img)
                else:
                    content_floats.append(img)
            idx += 1
            continue
        if isinstance(b, Table):
            block_html.append(render_table(b, assets, page_layout))
            idx += 1
            continue
        if isinstance(b, Paragraph) and _is_list_item(b):
            is_bullet = b.numbering_format == "bullet"
            run = []
            j = idx
            while j < len(blocks):
                cur = blocks[j]
                if cur is first_h1:
                    break
                if not isinstance(cur, Paragraph) or not _is_list_item(cur):
                    break
                if (cur.numbering_format == "bullet") != is_bullet:
                    break
                run.append(cur)
                j += 1
            tag = "ul" if is_bullet else "ol"
            cls = "docx-bullet-list" if is_bullet else "docx-ordered-list"
            lis = []
            for item in run:
                collected = []
                inner, needs_relative = _render_content(item, assets, page_layout, collected)
                label = _list_label(item)
                if label:
                    pcls = "docx-bullet" if item.numbering_format == "bullet" else "docx-number"
                    inner = '<span class="%s">%s</span> %s' % (pcls, html.escape(label), inner)
                layout = _paragraph_layout_style(item)
                extra = []
                if needs_relative:
                    extra.append("position: relative")
                if layout:
                    extra.append(layout)
                style_attr = ' style="%s"' % ";".join(extra) if extra else ""
                for container, img in collected:
                    if container == "page":
                        page_floats.append(img)
                    else:
                        content_floats.append(img)
                if not inner.strip() and not style_attr:
                    li_inner = "<p></p>"
                elif not inner.strip():
                    li_inner = "<p%s></p>" % style_attr
                else:
                    li_inner = "<p%s>%s</p>" % (style_attr, inner) if style_attr else "<p>%s</p>" % inner
                lis.append("<li>%s</li>" % li_inner)
            block_html.append('<%s class="docx-list %s">%s</%s>' % (tag, cls, "".join(lis), tag))
            idx = j
            continue
        if isinstance(b, Paragraph):
            html_str, _n, externals = _render_paragraph_html(b, assets, page_layout)
            block_html.append(html_str)
            for container, img in externals:
                if container == "page":
                    page_floats.append(img)
                else:
                    content_floats.append(img)
        idx += 1

    content_inner = "\n".join(block_html)

    header_html_parts = []
    footer_html_parts = []
    if sections:
        seen_hf = set()
        for sec in sections:
            for typ in ["default", "first", "even"]:
                hf = sec.headers.get(typ)
                if hf and hf.target not in seen_hf:
                    inner = _render_hf_blocks(hf.blocks, assets, page_layout)
                    if inner.strip():
                        header_html_parts.append('<header class="docx-header docx-header-%s" data-type="%s">%s</header>' % (typ, typ, inner))
                        seen_hf.add(hf.target)
                hf = sec.footers.get(typ)
                if hf and hf.target not in seen_hf:
                    inner = _render_hf_blocks(hf.blocks, assets, page_layout)
                    if inner.strip():
                        footer_html_parts.append('<footer class="docx-footer docx-footer-%s" data-type="%s">%s</footer>' % (typ, typ, inner))
                        seen_hf.add(hf.target)
    header_block = "\n".join(header_html_parts)
    footer_block = "\n".join(footer_html_parts)

    content_style = (
        "position: relative;"
        " margin-left: %dpx; margin-top: %dpx;"
        " width: %dpx; max-width: calc(100%% - %dpx); min-height: %dpx; box-sizing: border-box;"
        % (page_layout.margin_left_px, page_layout.margin_top_px,
           page_layout.content_width_px, page_layout.margin_left_px + 16,
           page_layout.content_height_px)
    )
    content_block = '<div class="docx-content" style="%s">\n%s\n%s\n</div>' % (
        content_style,
        content_inner,
        "\n".join(render_float_image(i, assets, page_layout) for i in content_floats),
    )

    page_inner = "\n".join([p for p in [header_block, content_block, footer_block] if p])
    page_style = "position: relative; width: %dpx; max-width: calc(100%% - 32px); min-height: %dpx; box-sizing: border-box;" % (
        page_layout.page_width_px, page_layout.page_height_px)
    page_block = '<div class="docx-page" style="%s">\n%s\n%s\n</div>' % (
        page_style,
        page_inner,
        "\n".join(render_float_image(i, assets, page_layout) for i in page_floats),
    )

    toc_tree_html = render_sidebar_toc(toc) if toc else ""
    if toc_tree_html:
        nav_inner = toc_tree_html
    else:
        nav_inner = '<p class="toc-empty">No headings</p>'
    sidebar_html = (
        '<aside id="viewer-sidebar" class="viewer-sidebar" aria-label="Document Outline">\n'
        '  <div class="sidebar-header">\n'
        '    <span class="sidebar-title">Document Outline</span>\n'
        '  </div>\n'
        '  <nav class="toc" aria-label="Table of Contents">\n'
        + nav_inner + '\n'
        '  </nav>\n'
        '</aside>\n'
        '<div id="sidebar-overlay" class="sidebar-overlay" hidden></div>'
    )
    doc_toolbar = (
        '<div class="doc-toolbar">\n'
        '  <button id="sidebar-toggle" class="sidebar-toggle-main" aria-expanded="true" aria-controls="viewer-sidebar" aria-label="Toggle outline">\n'
        '    <span aria-hidden="true">&#9776;</span> <span>Outline</span>\n'
        '  </button>\n'
        '</div>\n'
    )
    viewer_inner = (
        '<div class="viewer">\n'
        + sidebar_html + '\n'
        + '<main class="doc-main" id="doc-main">\n'
        + doc_toolbar
        + page_block + '\n'
        '</main>\n'
        '</div>'
    )
    if title_bar:
        body = '<div class="page-frame">\n' + title_bar + '\n' + viewer_inner + '\n</div>'
    else:
        body = '<div class="page-frame">\n' + viewer_inner + '\n</div>'

    style = _viewer_style()
    script = _viewer_script()

    return (
        "<!DOCTYPE html>\n"
        '<html lang="en">\n'
        "<head>\n"
        '  <meta charset="UTF-8">\n'
        '  <meta name="viewport" content="width=device-width, initial-scale=1">\n'
        "  <title>%s</title>\n"
        "%s"
        "</head>\n"
        "<body>\n%s\n%s\n</body>\n"
        "</html>\n"
    ) % (html.escape(title), style, body, script)
