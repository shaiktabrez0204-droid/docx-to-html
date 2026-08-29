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
from typing import Optional, List, Tuple

from core.model import Run, Paragraph, Table, Row, Cell, BorderEdge, HeaderFooter, Section, Image, ImageAsset, PageLayout, format_numbering_label
from core.units import emu_to_px, twip_to_emu

# The document's effective default font size, in OOXML half-points. Runs whose
# resolved size equals this default need no inline font-size (the .docx-content
# base carries it). Updated per-document by render_html() from docDefaults.
DEFAULT_FONT_HALF_POINTS = 22

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
    ft = getattr(run, "field_type", None)
    if ft in ("PAGE", "NUMPAGES", "PAGEREF"):
        code = html.escape(getattr(run, "field_code", "") or "", quote=True)
        lim = "Page number field - pagination not calculated in single-page HTML view"
        if ft == "PAGE":
            return '<span class="docx-page-number" data-field="PAGE" data-field-code="%s" title="%s" aria-label="Page number field"></span>' % (code, lim)
        elif ft == "NUMPAGES":
            return '<span class="docx-num-pages" data-field="NUMPAGES" data-field-code="%s" title="%s" aria-label="Total pages field"></span>' % (code, lim)
        else:
            return '<span class="docx-page-ref" data-field="PAGEREF" data-field-code="%s" title="%s" aria-label="Page reference field"></span>' % (code, lim)
    if run.text == "\t":
        return '<span class="docx-tab" data-tab="true"></span>'
    if run.text == "\n":
        return "<br>"
    if not run.text:
        return ""
    content = html.escape(run.text)
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
    if run.font_size is not None and run.font_size != DEFAULT_FONT_HALF_POINTS:
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
    if img.wrap_mode not in ("square", "tight", "through"):
        return False
    if img.alignment_horizontal in _LEFT_ALIGN:
        return True
    if img.alignment_horizontal in _RIGHT_ALIGN:
        return True
    return False


def _is_top_bottom(img: Image) -> bool:
    return img.wrap_mode == "topAndBottom"


def _wrap_polygon_to_css(polygon: Optional[List[Tuple[int, int]]]) -> Optional[str]:
    """Convert OOXML wrap polygon (21600x21600 space) to CSS polygon() value.

    Returns a string like "polygon(0% 0%, 100% 0%, 100% 100%, 0% 100%)"
    or None if the polygon is invalid/rectangular.
    """
    if not polygon or len(polygon) < 3:
        return None

    # Check if polygon is a simple rectangle (4 points + closed = 5 points where
    # points are (0,0), (21600,0), (21600,21600), (0,21600), (0,0)).
    # If so, return None to use margin-box instead.
    if len(polygon) == 5:
        pts = polygon
        rect_corners = {(0, 0), (21600, 0), (21600, 21600), (0, 21600)}
        if set(pts[:4]) == rect_corners and pts[4] == pts[0]:
            return None

    # Convert to percentages. OOXML wrapPolygon uses a 21600x21600 coordinate
    # space that maps to the image's extent rectangle.
    coords = []
    for x, y in polygon:
        # Skip the closing duplicate point for CSS (polygon() auto-closes).
        if coords and x == coords[0][0] and y == coords[0][1]:
            break
        x_pct = x / 21600.0 * 100.0
        y_pct = y / 21600.0 * 100.0
        coords.append(("%g%%" % x_pct, "%g%%" % y_pct))

    if len(coords) < 3:
        return None

    return "polygon(" + ", ".join("%s %s" % (x, y) for x, y in coords) + ")"


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

    if _is_top_bottom(img):
        halign = img.alignment_horizontal
        d = img.wrap_distances or {}
        mt = emu_to_px(d.get("top", 0))
        mr = emu_to_px(d.get("right", 0))
        mb = emu_to_px(d.get("bottom", 0))
        ml = emu_to_px(d.get("left", 0))
        if halign in _LEFT_ALIGN or halign in _RIGHT_ALIGN or halign == "center":
            style["display"] = "block"
            style["clear"] = "both"
            style["float"] = "none"
            if halign == "center":
                style["margin"] = "%dpx auto %dpx auto" % (mt, mb)
            elif halign in _LEFT_ALIGN:
                style["margin"] = "%dpx %dpx %dpx %dpx" % (mt, mr, mb, ml)
                style["margin-left"] = "%dpx" % ml
                style["margin-right"] = "auto"
            elif halign in _RIGHT_ALIGN:
                style["margin"] = "%dpx %dpx %dpx %dpx" % (mt, mr, mb, ml)
                style["margin-left"] = "auto"
                style["margin-right"] = "%dpx" % mr
            else:
                style["margin"] = "%dpx %dpx %dpx %dpx" % (mt, mr, mb, ml)
        else:
            style["position"] = "absolute"
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
                if container == "page":
                    style["left"] = "%dpx" % hoffs
                elif container == "content":
                    style["left"] = "%dpx" % hoffs
                else:
                    style["left"] = "%dpx" % hoffs
                tx = "0"
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
    elif _is_wrap_float(img):
        halign = img.alignment_horizontal
        style["float"] = "left" if halign in _LEFT_ALIGN else "right"
        d = img.wrap_distances or {}
        mt = emu_to_px(d.get("top", 0))
        mr = emu_to_px(d.get("right", 0))
        mb = emu_to_px(d.get("bottom", 0))
        ml = emu_to_px(d.get("left", 0))
        style["margin"] = "%dpx %dpx %dpx %dpx" % (mt, mr, mb, ml)
        if img.wrap_mode in ("tight", "through"):
            poly_css = _wrap_polygon_to_css(getattr(img, "wrap_polygon", None))
            if poly_css:
                style["shape-outside"] = poly_css
                # clip-path only if safe - polygon may not match image bounds exactly
                # For now, we don't set clip-path to avoid clipping image content
            else:
                style["shape-outside"] = "margin-box"
                style["clip-path"] = "margin-box"
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
    if _is_top_bottom(img) and (img.alignment_horizontal in _LEFT_ALIGN or img.alignment_horizontal in _RIGHT_ALIGN or img.alignment_horizontal == "center"):
        cls = "docx-float-topbottom"
    elif _is_wrap_float(img):
        cls = "docx-float-wrapped"
    else:
        cls = "docx-float"
    attrs = [
        'src="%s"' % src,
        'class="%s"' % cls,
        'style="%s"' % style,
        alt_attr,
    ]
    return "<img %s>" % " ".join(attrs)


def _render_items_segment(segment_items, assets, page, collected, is_hf=False):
    """Render a list of Run/Image items (no tab) handling floats/hyperlinks."""
    parts = []
    needs_relative = False
    i = 0
    while i < len(segment_items):
        item = segment_items[i]
        if isinstance(item, Image):
            if item.wrap_type == "anchor":
                container = _float_container(item)
                is_tb_block = _is_top_bottom(item) and (item.alignment_horizontal in _LEFT_ALIGN or item.alignment_horizontal in _RIGHT_ALIGN or item.alignment_horizontal == "center")
                if _is_wrap_float(item) or is_tb_block or container == "paragraph":
                    if is_hf:
                        if container == "paragraph":
                            needs_relative = True
                        parts.append(render_float_image(item, assets, page))
                    elif _is_wrap_float(item):
                        anc = getattr(item, "anchor_paragraph_index", -1)
                        if anc is None:
                            anc = -1
                        try:
                            anc = int(anc)
                        except Exception:
                            anc = -1
                        inner = render_float_image(item, assets, page)
                        if 'data-anchor' not in inner:
                            inner = inner.replace('<img ', '<img data-anchor="%d" ' % anc, 1)
                        parts.append(inner)
                    elif is_tb_block:
                        anc = getattr(item, "anchor_paragraph_index", -1)
                        if anc is None:
                            anc = -1
                        try:
                            anc = int(anc)
                        except Exception:
                            anc = -1
                        inner = render_float_image(item, assets, page)
                        if 'data-anchor' not in inner:
                            inner = inner.replace('<img ', '<img data-anchor="%d" ' % anc, 1)
                        parts.append(inner)
                    else:
                        anc = getattr(item, "anchor_paragraph_index", -1)
                        if anc is None:
                            anc = -1
                        try:
                            anc = int(anc)
                        except Exception:
                            anc = -1
                        inner = render_float_image(item, assets, page)
                        parts.append('<span class="docx-float-wrap docx-para-float-wrap" data-anchor="%d">%s</span>' % (anc, inner))
                        if container == "paragraph":
                            needs_relative = True
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
            while j < len(segment_items) and not isinstance(segment_items[j], Image) and getattr(segment_items[j], "href", None) == href:
                group.append(segment_items[j])
                j += 1
            inner = "".join(_render_run_inner(r) for r in group)
            if inner:
                parts.append('<a href="%s">%s</a>' % (html.escape(href, quote=True), inner))
            i = j
            continue
        parts.append(_render_run_inner(item))
        i += 1
    return "".join(parts), needs_relative

def _render_content(para: Paragraph, assets, page, collected, is_hf=False):
    """Render ordered Run/Image content, splitting out container-level floats.

    Returns the inner HTML for the paragraph. Inline images and in-paragraph
    floats (paragraph-relative, or wrap floats) are rendered inline to preserve
    document order. page/content-absolute floats are appended to ``collected``
    as (container, img) so the caller can place them in the right box.

    Consecutive Runs sharing the same href are grouped into a single <a> to
    preserve the original OOXML hyperlink span.

    Tab-stop aware: when paragraph has w:tabs, each tab Run (\\t) advances to
    the next effective tab stop (sorted by pos, clear excluded). Segments after
    a tab are absolutely positioned at that stop (left/center/right/decimal).
    """
    items = para.content if para.content else para.runs
    has_tab = any(isinstance(it, Run) and it.text == "\t" for it in items)
    if not has_tab:
        return _render_items_segment(items, assets, page, collected, is_hf=is_hf)
    segments = []
    cur = []
    for it in items:
        if isinstance(it, Run) and it.text == "\t":
            segments.append(cur)
            cur = []
        else:
            cur.append(it)
    segments.append(cur)
    tabs = getattr(para, 'tabs', None) or []
    eff = [t for t in tabs if t.val != "clear"]
    eff = sorted(eff, key=lambda t: t.pos)
    parts = []
    needs_relative = False
    has_leader = any(getattr(t, 'leader', None) and getattr(t, 'leader') != "none" for t in eff)
    seg_html, seg_need = _render_items_segment(segments[0], assets, page, collected, is_hf=is_hf)
    if has_leader and seg_html:
        seg_html = '<span class="docx-tab-segment" data-tab="0" style="display:inline-block;">%s</span>' % seg_html
        needs_relative = True
    parts.append(seg_html)
    needs_relative = needs_relative or seg_need
    for idx in range(1, len(segments)):
        seg = segments[idx]
        tab = eff[idx-1] if idx-1 < len(eff) else None
        seg_html2, seg_need2 = _render_items_segment(seg, assets, page, collected, is_hf=is_hf)
        if tab is None:
            fallback = '<span class="docx-tab" data-tab="true" style="display:inline-block; width:48px;"></span>'
            if not seg_html2:
                parts.append(fallback)
                needs_relative = needs_relative or seg_need2
            else:
                parts.append(fallback + seg_html2)
                needs_relative = needs_relative or seg_need2
        else:
            if not seg_html2:
                continue
            pos_px = round(tab.pos * 635 / 9525)
            leader = getattr(tab, 'leader', None)
            leader_html = ""
            if leader and leader != "none":
                leader_style_map = {"dot": "dotted", "hyphen": "dashed", "underscore": "solid", "heavy": "solid", "middleDot": "dotted"}
                bs = leader_style_map.get(leader, "dotted")
                bw = "2px" if leader == "heavy" else "1px"
                leader_html = '<span class="docx-tab-leader" data-leader="%s" data-val="%s" data-pos="%d" style="position:absolute; height:0; border-bottom:%s %s #999; pointer-events:none; left:0; width:0; top:50%%;"></span>' % (leader, tab.val, pos_px, bw, bs)
                needs_relative = True
                parts.append(leader_html)
            leader_style = ""
            if getattr(tab, 'leader', None) == "dot":
                leader_style = "border-bottom:1px dotted #999;"
            if tab.val == "center":
                wrapper = '<span class="docx-tab-segment" data-tab="%d" data-val="center" style="position:absolute; left:%dpx; transform:translateX(-50%%); white-space:nowrap; %s">%s</span>' % (idx, pos_px, leader_style, seg_html2)
                needs_relative = True
                parts.append(wrapper)
            elif tab.val == "right":
                wrapper = '<span class="docx-tab-segment" data-tab="%d" data-val="right" style="position:absolute; left:%dpx; transform:translateX(-100%%); white-space:nowrap; %s">%s</span>' % (idx, pos_px, leader_style, seg_html2)
                needs_relative = True
                parts.append(wrapper)
            elif tab.val == "decimal":
                wrapper = '<span class="docx-tab-segment" data-tab="%d" data-val="decimal" data-pos="%d" style="position:absolute; left:%dpx; white-space:nowrap; %s">%s</span>' % (idx, pos_px, pos_px, leader_style, seg_html2)
                needs_relative = True
                parts.append(wrapper)
            else:
                wrapper = '<span class="docx-tab-segment" data-tab="%d" data-val="left" style="position:absolute; left:%dpx; white-space:nowrap; %s">%s</span>' % (idx, pos_px, leader_style, seg_html2)
                needs_relative = True
                parts.append(wrapper)
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


def _render_paragraph_html(para: Paragraph, assets=None, page=None, is_hf=False):
    """Render one paragraph (or heading) including its in-paragraph floats.

    Returns (html, needs_relative, external_floats) where external_floats is a
    list of (container, Image) for page/content-absolute images that must be
    placed by the caller inside .docx-page / .docx-content.
    """
    if page is None:
        page = PageLayout()
    collected = []
    inner, needs_relative = _render_content(para, assets, page, collected, is_hf=is_hf)
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


def _wrap_block(inner_html: str, heading_id: Optional[str] = None,
                level: Optional[int] = None) -> str:
    """Wrap a rendered top-level block so the viewer can isolate sections.

    Every document block (heading/paragraph/table/list) is placed inside a
    ``.docx-block`` <div>. Heading blocks additionally carry ``data-heading-id``
    and ``data-level`` so the viewer JS can compute section boundaries purely
    from the existing heading hierarchy (no regex, no text inference). Non-heading
    blocks carry no heading attributes and are bounded by the nearest surrounding
    heading blocks at runtime.
    """
    attrs = 'class="docx-block"'
    if heading_id:
        attrs += ' data-heading-id="%s" data-level="%d"' % (
            html.escape(heading_id, quote=True), int(level))
    return "<div %s>%s</div>" % (attrs, inner_html)


def _render_cell_content(cell: Cell, assets, page, table_anchor=None, is_hf=False) -> str:
    parts = []
    for para in cell.content:
        if not is_hf and table_anchor is not None:
            for it in getattr(para, "content", []) or []:
                if isinstance(it, Image) and it.wrap_type == "anchor" and _float_container(it) == "paragraph":
                    try:
                        it.anchor_paragraph_index = int(table_anchor)
                    except Exception:
                        pass
        collected = []
        inner, needs_relative = _render_content(para, assets, page, collected, is_hf=is_hf)
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
            if is_hf:
                inner_float = render_float_image(img, assets, page)
                parts.append('<div class="docx-hf-float-wrap">%s</div>' % inner_float)
            else:
                try:
                    if table_anchor is not None:
                        img.anchor_paragraph_index = table_anchor
                except Exception:
                    pass
                inner_float = render_float_image(img, assets, page)
                anc = table_anchor if table_anchor is not None else -1
                parts.append('<div class="docx-float-wrap" data-anchor="%d">%s</div>' % (int(anc), inner_float))
    return "".join(parts)


def render_table(table: Table, assets=None, page=None, table_anchor=None, is_hf=False) -> str:
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
            inner = _render_cell_content(cell, assets, page, table_anchor=table_anchor, is_hf=is_hf)
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
            parts.append(render_table(b, assets, page, is_hf=True))
        elif isinstance(b, Paragraph):
            html_str, _needs, _ext = _render_paragraph_html(b, assets, page, is_hf=True)
            parts.append(html_str)
            for _container, img in _ext:
                inner = render_float_image(img, assets, page)
                parts.append('<div class="docx-hf-float-wrap">%s</div>' % inner)
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
        "    html,body{margin:0;padding:0;height:100%;overflow:hidden;}\n"
        "    body{font:14px/1.5 Inter,system-ui,-apple-system,Segoe UI,Roboto,sans-serif;color:#111827;background:#fff;overflow:hidden;-webkit-font-smoothing:antialiased;}\n"
        "    .page-frame{display:flex;flex-direction:column;height:100vh;overflow:hidden;background:#fff;}\n"
        "    .viewer-header{height:52px;display:flex;align-items:center;gap:12px;padding:0 14px;border-bottom:1px solid #e5e7eb;background:rgba(255,255,255,.92);backdrop-filter:blur(10px) saturate(1.2);flex-shrink:0;z-index:8;}\n"
        "    .viewer-brand{display:flex;align-items:center;gap:8px;font-weight:700;font-size:13px;letter-spacing:-.01em;color:#0f172a;white-space:nowrap;flex-shrink:0;}\n"
        "    .viewer-brand-icon{width:26px;height:26px;border-radius:7px;display:grid;place-items:center;background:linear-gradient(135deg,#5b5bf6,#8b5cf6);color:#fff;flex-shrink:0;}\n"
        "    .viewer-docname{flex:1;min-width:0;text-align:center;font-size:13px;font-weight:600;color:#1f2937;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;padding:0 12px;}\n"
        "    .viewer-actions{display:flex;align-items:center;gap:6px;flex-shrink:0;}\n"
        "    .viewer-action{appearance:none;border:1px solid #e5e7eb;background:#fff;color:#374151;cursor:pointer;width:32px;height:32px;border-radius:8px;display:inline-flex;align-items:center;justify-content:center;transition:all .15s;}\n"
        "    .viewer-action:hover{background:#f9fafb;border-color:#d1d5db;transform:translateY(-1px);box-shadow:0 1px 6px rgba(0,0,0,.06);}\n"
        "    .viewer-action:focus-visible{outline:2px solid #6366f1;outline-offset:2px;}\n"
        "    .viewer-action.primary{border-color:#6366f1;background:#6366f1;color:#fff;}\n"
        "    .viewer-action.primary:hover{background:#4f46e5;}\n"
        "    .viewer-title{padding:8px 16px 8px;border-bottom:1px solid #f3f4f6;background:#fff;flex-shrink:0;}\n"
        "    .viewer-title .doc-title{margin:0;font-size:14px;font-weight:600;line-height:1.4;color:#111;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}\n"
        "    .viewer{display:flex;flex:1;min-height:0;overflow:hidden;background:#f3f4f6;}\n"
        "    .viewer-sidebar{width:280px;min-width:280px;max-width:280px;background:#fff;border-right:1px solid #e5e7eb;display:flex;flex-direction:column;overflow:hidden;flex-shrink:0;transition:width .22s cubic-bezier(.2,.8,.2,1),min-width .22s cubic-bezier(.2,.8,.2,1),opacity .18s ease,transform .22s ease;}\n"
        "    .viewer.viewer--sidebar-collapsed .viewer-sidebar{width:0;min-width:0;max-width:0;border-right-width:0;opacity:0;pointer-events:none;transform:translateX(-8px);overflow:hidden;}\n"
        "    .sidebar-header{display:flex;align-items:center;gap:8px;padding:10px 12px 8px;font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#6b7280;white-space:nowrap;flex-shrink:0;}\n"
        "    .sidebar-title{flex:1;overflow:hidden;text-overflow:ellipsis;}\n"
        "    .sidebar-header-actions{display:flex;gap:4px;}\n"
        "    .toc-search-wrap{padding:0 10px 8px;flex-shrink:0;}\n"
        "    .toc-search{position:relative;}\n"
        "    .toc-search input{width:100%;padding:7px 28px 7px 30px;font-size:13px;line-height:1;border:1px solid #e5e7eb;border-radius:8px;background:#f9fafb;color:#111827;outline:none;transition:border-color .15s,box-shadow .15s,background .15s;}\n"
        "    .toc-search input::placeholder{color:#9ca3af;}\n"
        "    .toc-search input:focus{border-color:#6366f1;box-shadow:0 0 0 3px rgba(99,102,241,.12);background:#fff;}\n"
        "    .toc-search-icon{position:absolute;left:8px;top:50%;transform:translateY(-50%);color:#9ca3af;pointer-events:none;}\n"
        "    .toc-search-clear{position:absolute;right:6px;top:50%;transform:translateY(-50%);appearance:none;border:0;background:transparent;color:#9ca3af;cursor:pointer;width:20px;height:20px;border-radius:6px;display:none;align-items:center;justify-content:center;}\n"
        "    .toc-search-clear.visible{display:inline-flex;}\n"
        "    .toc-search-clear:hover{background:#f3f4f6;color:#374151;}\n"
        "    .toc-no-results{padding:14px 10px;color:#9ca3af;font-size:13px;text-align:center;display:none;}\n"
        "    .toc-no-results.visible{display:block;}\n"
        "    .toc{flex:1;overflow-y:auto;overflow-x:hidden;padding:4px 8px 12px;scrollbar-width:thin;}\n"
        "    .toc-empty{color:#888;font-size:13px;padding:12px;}\n"
        "    .toc-tree{list-style:none;margin:0;padding:0;}\n"
        "    .toc-tree ul{list-style:none;margin:0;padding-left:14px;}\n"
        "    .toc-row{display:flex;align-items:center;gap:2px;}\n"
        "    .toc-link{flex:1;display:block;padding:5px 8px;font-size:13px;line-height:1.35;color:#374151;text-decoration:none;border-radius:7px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;transition:background .12s,color .12s,transform .12s;}\n"
        "    .toc-link:hover{background:#f3f4f6;color:#111827;}\n"
        "    .toc-link.is-active{background:#eef2ff;color:#3730a3;font-weight:600;box-shadow:inset 3px 0 0 #6366f1;}\n"
        "    .toc-link:focus-visible{outline:2px solid #6366f1;outline-offset:1px;}\n"
        "    .toc-link.is-search-hidden{display:none;}\n"
        "    .toc-mark{background:#fef08a;color:#854d0e;padding:1px 2px;border-radius:3px;font-weight:600;}\n"
        "    .toc-parent{margin:1px 0;}\n"
        "    .toc-parent.is-collapsed>ul.toc-children{display:none;}\n"
        "    .toc-parent.is-search-hidden{display:none;}\n"
        "    .toc-toggle{appearance:none;border:0;background:transparent;cursor:pointer;width:22px;height:22px;flex-shrink:0;display:inline-flex;align-items:center;justify-content:center;font-size:10px;color:#9ca3af;border-radius:6px;transition:all .15s;}\n"
        "    .toc-toggle::before{content:\"▾\";display:block;font-size:10px;transition:transform .18s;}\n"
        "    .toc-parent.is-collapsed>.toc-row .toc-toggle::before{transform:rotate(-90deg);}\n"
        "    .toc-toggle:hover{background:#f3f4f6;color:#374151;}\n"
        "    .toc-toggle:focus-visible{outline:2px solid #6366f1;}\n"
        "    .toc-leaf{padding-left:24px;}\n"
        "    .toc-leaf.is-search-hidden{display:none;}\n"
        "    .doc-main{flex:1;overflow-y:auto;overflow-x:hidden;background:#f1f2f4;padding:0;display:flex;flex-direction:column;align-items:center;min-width:0;scroll-behavior:smooth;}\n"
        "    .doc-toolbar{position:sticky;top:0;z-index:5;width:100%;max-width:860px;display:flex;align-items:center;gap:8px;padding:10px 16px;background:rgba(241,242,244,.92);backdrop-filter:blur(8px);flex-shrink:0;}\n"
        "    .focus-banner{width:100%;max-width:860px;margin:0 auto;padding:10px 16px 8px;display:none;align-items:center;gap:10px;}\n"
        "    .focus-banner.visible{display:flex;animation:fadeIn .18s ease;}\n"
        "    .focus-banner-inner{flex:1;display:flex;align-items:center;gap:10px;padding:10px 12px;background:#fff;border:1px solid #e0e7ff;border-radius:12px;box-shadow:0 4px 18px rgba(99,102,241,.08);}\n"
        "    .focus-banner-kicker{font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#6366f1;white-space:nowrap;}\n"
        "    .focus-banner-title{flex:1;min-width:0;font-size:13px;font-weight:600;color:#1e293b;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}\n"
        "    .focus-banner-clear{appearance:none;border:1px solid #e5e7eb;background:#fff;color:#374151;cursor:pointer;padding:6px 10px;border-radius:999px;font-size:13px;font-weight:600;display:inline-flex;align-items:center;gap:6px;transition:all .15s;white-space:nowrap;}\n"
        "    .focus-banner-clear:hover{background:#f9fafb;border-color:#d1d5db;}\n"
        "    .focus-banner-close{appearance:none;border:0;background:transparent;color:#9ca3af;cursor:pointer;width:28px;height:28px;border-radius:8px;display:inline-flex;align-items:center;justify-content:center;}\n"
        "    .focus-banner-close:hover{background:#f3f4f6;color:#374151;}\n"
        "    .sidebar-toggle-main{appearance:none;border:1px solid #e5e7eb;background:#fff;cursor:pointer;padding:7px 12px;border-radius:8px;font-size:13px;line-height:1;display:inline-flex;align-items:center;gap:6px;color:#374151;box-shadow:0 1px 2px rgba(0,0,0,.05);transition:all .15s;}\n"
        "    .sidebar-toggle-main:hover{background:#f9fafb;transform:translateY(-1px);box-shadow:0 4px 10px rgba(0,0,0,.06);}\n"
        "    .sidebar-toggle-main:focus-visible{outline:2px solid #6366f1;outline-offset:2px;}\n"
        "    .docx-page{background:#fff;margin:12px auto 48px;box-shadow:0 8px 32px rgba(15,23,42,.08),0 0 0 1px rgba(15,23,42,.04);width:100%;max-width:860px;position:relative;flex-shrink:0;box-sizing:border-box;border-radius:10px;overflow:hidden;transition:transform .22s ease,opacity .18s ease;}\n"
        "    .docx-page[style]{max-width:min(860px,calc(100% - 32px)) !important;}\n"
        "    .docx-block{transition:opacity .18s ease,transform .18s ease;}\n"
        "    .docx-block.is-hidden{display:none !important;}\n"
        "    .docx-block.is-entering{animation:docFadeIn .22s ease both;}\n"
        "    @keyframes docFadeIn{from{opacity:0;transform:translateY(4px);}to{opacity:1;transform:translateY(0);}}\n"
        "    @keyframes fadeIn{from{opacity:0;transform:translateY(-4px);}to{opacity:1;transform:translateY(0);}}\n"
        "    .docx-float-wrap.is-hidden{display:none !important;}\n"
        "    .docx-float-wrap{position:static;}\n"
        "    .docx-para-float-wrap.is-hidden{display:none !important;}\n"
        "    .docx-para-float-wrap{position:static;}\n"
        "    img[data-anchor].is-hidden{display:none !important;}\n"
        "    img.docx-float.is-hidden, img.docx-float-wrapped.is-hidden{display:none !important;}\n"
        "    .docx-hf-float-wrap{position:static;}\n"
        "    #exit-focus{margin-left:auto;}\n"
        "    button.sidebar-toggle-main[hidden]{display:none !important;}\n"
        "    .docx-content{box-sizing:border-box;max-width:100%;overflow-wrap:break-word;font-size:" + ("%.1f" % (DEFAULT_FONT_HALF_POINTS / 2.0)) + "pt;}\n"
        "    img.docx-float{position:absolute;}\n"
        "    img.docx-float-wrapped{position:static;max-width:none;}\n"
        "    img.docx-float-topbottom{display:block; clear:both; float:none; max-width:100%; margin-left:auto; margin-right:auto;}\n"
        "    img.docx-float-topbottom.is-hidden{display:none !important;}\n"
        "    .docx-number{font:inherit;font-weight:inherit;font-style:inherit;margin-right:.4em;white-space:nowrap;}\n"
        "    .docx-bullet{margin-right:.4em;}\n"
        "    nav.toc .docx-number{margin-right:.4em;}\n"
        "    .docx-list{margin:0.6em 0 0.6em 1.8em;padding:0;}\n"
        "    .docx-list li{margin:0.15em 0;}\n"
        "    .docx-list li p{margin:0;}\n"
        "    .docx-ordered-list{list-style:none;padding-left:0;}\n"
        "    .docx-bullet-list{list-style:none;padding-left:0;}\n"
        "    table.docx-table{border-collapse:collapse;width:100%;margin:1em 0;}\n"
        "    table.docx-table td,table.docx-table th{border:1px solid #d1d5db;padding:6px 8px;vertical-align:top;word-wrap:break-word;}\n"
        "    table.docx-table td p{margin:0;}\n"
        "    header.docx-header,footer.docx-footer{border:1px dashed #cbd5e1;padding:6px 8px;margin:6px 0;background:#f8fafc;}\n"
        "    header.docx-header p,footer.docx-footer p{margin:0; position:relative; min-height:1.2em;}\n"
        "    .docx-tab{display:inline-block;width:48px;}\n"
        "    .docx-tab-leader{position:absolute; height:0; pointer-events:none;}\n"
        "    .docx-tab-segment{white-space:nowrap;}\n"
        "    .docx-page-number,.docx-num-pages,.docx-page-ref{display:inline-block;min-width:1.2em;padding:0 2px;background:#eef2ff;border:1px dashed #a5b4fc;border-radius:3px;font-size:0.85em;color:#4f46e5;vertical-align:baseline;}\n"
        "    .docx-page-number::after{content:\"[PAGE]\";}\n"
        "    .docx-num-pages::after{content:\"[NUMPAGES]\";}\n"
        "    .docx-page-ref::after{content:\"[PAGEREF]\";}\n"
        "    .sidebar-overlay{position:fixed;inset:0;background:rgba(15,23,42,.32);backdrop-filter:blur(2px);z-index:15;opacity:0;visibility:hidden;transition:opacity .18s,visibility .18s;}\n"
        "    .sidebar-overlay.visible{opacity:1;visibility:visible;}\n"
        "    @media (max-width:780px){.viewer-sidebar{position:fixed;left:0;top:0;bottom:0;z-index:20;width:280px;min-width:280px;max-width:280px;transform:translateX(-100%);transition:transform .24s cubic-bezier(.2,.8,.2,1);opacity:1;pointer-events:auto;border-right:1px solid #e5e7eb;box-shadow:12px 0 40px rgba(15,23,42,.12);}\n"
        "    .viewer.viewer--mobile-open .viewer-sidebar{transform:translateX(0);}\n"
        "    .viewer.viewer--sidebar-collapsed .viewer-sidebar{width:280px;min-width:280px;max-width:280px;opacity:1;pointer-events:auto;transform:translateX(-100%);border-right-width:1px;box-shadow:12px 0 40px rgba(15,23,42,.12);}\n"
        "    .viewer.viewer--mobile-open.viewer--sidebar-collapsed .viewer-sidebar{transform:translateX(0);}\n"
        "    .viewer-header{padding:0 10px;gap:8px;}\n"
        "    .viewer-docname{font-size:12px;padding:0 6px;}\n"
        "    .doc-main{padding-top:0;}}\n"
        "    @media (prefers-reduced-motion:reduce){*,*::before,*::after{transition:none!important;animation:none!important;scroll-behavior:auto!important;}}\n"
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
        "        if(open){ overlay.hidden=false; overlay.classList.add('visible'); } else { overlay.classList.remove('visible'); setTimeout(function(){ if(!viewer.classList.contains('viewer--mobile-open')) overlay.hidden=true; },180); }\n"
        "        viewer.classList.toggle('viewer--sidebar-collapsed',!open);\n"
        "      } else {\n"
        "        viewer.classList.toggle('viewer--sidebar-collapsed',!open);\n"
        "        viewer.classList.remove('viewer--mobile-open');\n"
        "        overlay.hidden=true; overlay.classList.remove('visible');\n"
        "      }\n"
        "      toggle.setAttribute('aria-expanded',String(open));\n"
        "      toggle.setAttribute('aria-label',open?'Close outline':'Open outline');\n"
        "      var ht=document.getElementById('header-toc'); if(ht){ ht.setAttribute('aria-expanded',String(open)); }\n"
        "    }\n"
        "    function isSidebarOpen(){\n"
        "      if(isMobile()) return viewer.classList.contains('viewer--mobile-open');\n"
        "      return !viewer.classList.contains('viewer--sidebar-collapsed');\n"
        "    }\n"
        "    toggle.addEventListener('click',function(){setSidebarOpen(!isSidebarOpen());});\n"
        "    var headerToc=document.getElementById('header-toc'); if(headerToc) headerToc.addEventListener('click',function(){setSidebarOpen(!isSidebarOpen());});\n"
        "    var sideCollapse=document.getElementById('sidebar-collapse'); if(sideCollapse) sideCollapse.addEventListener('click',function(){setSidebarOpen(false);});\n"
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
        "    var currentFocus=null;\n"
        "    var exitBtn=document.getElementById('exit-focus');\n"
        "    var focusBanner=document.getElementById('focus-banner');\n"
        "    var focusBannerTitle=document.getElementById('focus-banner-title');\n"
        "    var focusClear=document.getElementById('focus-banner-clear');\n"
        "    var focusClose=document.getElementById('focus-banner-close');\n"
        "    function showExit(on, title){\n"
        "      if(exitBtn) exitBtn.hidden=!on;\n"
        "      if(focusBanner){\n"
        "        if(on){ focusBanner.classList.add('visible'); if(focusBannerTitle && title) focusBannerTitle.textContent=title; }\n"
        "        else { focusBanner.classList.remove('visible'); }\n"
        "      }\n"
        "    }\n"
        "    function getHeadingText(id){\n"
        "      var a=linkById[id]; if(a) return a.textContent.trim();\n"
        "      var h=document.getElementById(id); if(h) return h.textContent.trim();\n"
        "      return id;\n"
        "    }\n"
        "    function clearFocus(){\n"
        "      var bs=docMain.querySelectorAll('.docx-block');\n"
        "      for(var i=0;i<bs.length;i++){ bs[i].classList.remove('is-hidden'); bs[i].classList.remove('is-entering'); }\n"
        "      var fs=docMain.querySelectorAll('.docx-float-wrap, .docx-para-float-wrap, img[data-anchor]');\n"
        "      for(var i=0;i<fs.length;i++){ fs[i].classList.remove('is-hidden'); }\n"
        "      showExit(false);\n"
        "      currentFocus=null;\n"
        "      docMain.scrollTop=0;\n"
        "    }\n"
        "    function focusHeading(id){\n"
        "      if(!id) return;\n"
        "      var link=linkById[id];\n"
        "      if(!link) return;\n"
        "      var level=parseInt(link.getAttribute('data-level')||'0',10)||1;\n"
        "      var blocks=Array.prototype.slice.call(docMain.querySelectorAll('.docx-block'));\n"
        "      var startIdx=-2;\n"
        "      if(id===titleId){ startIdx=-1; }\n"
        "      else { for(var i=0;i<blocks.length;i++){ if(blocks[i].getAttribute('data-heading-id')===id){ startIdx=i; break; } } }\n"
        "      if(startIdx===-2) return;\n"
        "      for(var i=0;i<blocks.length;i++){ blocks[i].classList.add('is-hidden'); blocks[i].classList.remove('is-entering'); }\n"
        "      var visible=[];\n"
        "      if(startIdx===-1){\n"
        "        for(var i=0;i<blocks.length;i++){\n"
        "          var l=parseInt(blocks[i].getAttribute('data-level')||'0',10);\n"
        "          if(i>0 && l && l<=level) break;\n"
        "          blocks[i].classList.remove('is-hidden'); visible.push(blocks[i]);\n"
        "        }\n"
        "      } else {\n"
        "        for(var i=startIdx;i<blocks.length;i++){\n"
        "          var l=parseInt(blocks[i].getAttribute('data-level')||'0',10);\n"
        "          if(i>startIdx && l && l<=level) break;\n"
        "          blocks[i].classList.remove('is-hidden'); visible.push(blocks[i]);\n"
        "        }\n"
        "      }\n"
        "      for(var vi=0; vi<visible.length; vi++){ (function(el,idx){ requestAnimationFrame(function(){ el.classList.add('is-entering'); }); })(visible[vi], vi); }\n"
        "      var floats=docMain.querySelectorAll('.docx-float-wrap, .docx-para-float-wrap, img[data-anchor]');\n"
        "      for(var f=0;f<floats.length;f++){\n"
        "        var aStr=floats[f].getAttribute('data-anchor');\n"
        "        var a=parseInt(aStr,10);\n"
        "        if(aStr==='-1' || a===-1){\n"
        "          if(startIdx===-1) floats[f].classList.remove('is-hidden');\n"
        "          else floats[f].classList.add('is-hidden');\n"
        "          continue;\n"
        "        }\n"
        "        if(isNaN(a) || a<0 || a>=blocks.length){ floats[f].classList.add('is-hidden'); continue; }\n"
        "        if(blocks[a].classList.contains('is-hidden')) floats[f].classList.add('is-hidden');\n"
        "        else floats[f].classList.remove('is-hidden');\n"
        "      }\n"
        "      showExit(true, getHeadingText(id));\n"
        "      docMain.scrollTop=0;\n"
        "      currentFocus=id;\n"
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
        "      focusHeading(id);\n"
        "      setActive(id);\n"
        "      history.pushState({focusId:id},'',href);\n"
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
        "      if(hval&&linkById[hval]){ focusHeading(hval); setActive(hval); }\n"
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
        "      var hv2=location.hash.slice(1);\n"
        "      if(hv2&&linkById[hv2]){ focusHeading(hv2); setActive(hv2); }\n"
        "    }\n"
        "    if(exitBtn){\n"
        "      exitBtn.addEventListener('click',function(){\n"
        "        clearFocus();\n"
        "        history.pushState({focusId:null},'',location.pathname+location.search);\n"
        "      });\n"
        "    }\n"
        "    if(focusClear) focusClear.addEventListener('click',function(){ clearFocus(); history.pushState({focusId:null},'',location.pathname+location.search); });\n"
        "    if(focusClose) focusClose.addEventListener('click',function(){ clearFocus(); history.pushState({focusId:null},'',location.pathname+location.search); });\n"
        "    window.addEventListener('popstate',function(e){\n"
        "      var st=e.state;\n"
        "      if(st && st.focusId){ focusHeading(st.focusId); setActive(st.focusId); }\n"
        "      else { clearFocus(); }\n"
        "    });\n"
        "    var tocSearch=document.getElementById('toc-search');\n"
        "    var tocClear=document.getElementById('toc-search-clear');\n"
        "    var tocNoResults=document.getElementById('toc-no-results');\n"
        "    var headerSearch=document.getElementById('header-search');\n"
        "    var viewerDownload=document.getElementById('viewer-download');\n"
        "    var tocLinksCache=null; var tocListItemsCache=null; var origLinkHTML={};\n"
        "    function ensureCache(){ if(tocLinksCache) return; tocLinksCache=Array.prototype.slice.call(document.querySelectorAll('a.toc-link')); tocListItemsCache=Array.prototype.slice.call(document.querySelectorAll('.toc-tree li')); tocLinksCache.forEach(function(a){ origLinkHTML[a.getAttribute('href')] = a.innerHTML; }); }\n"
        "    function escapeReg(s){ return s.replace(/[.*+?^${}()|[\\]\\\\]/g,'\\\\$&'); }\n"
        "    var searchTimer=null;\n"
        "    function doSearch(q){\n"
        "      ensureCache();\n"
        "      var query=(q||'').trim().toLowerCase();\n"
        "      if(tocClear) tocClear.classList.toggle('visible', !!query);\n"
        "      if(!query){\n"
        "        tocLinksCache.forEach(function(a){ a.innerHTML=origLinkHTML[a.getAttribute('href')]||a.innerHTML; a.classList.remove('is-search-hidden'); var li=a.closest('li'); if(li) li.classList.remove('is-search-hidden'); });\n"
        "        tocListItemsCache.forEach(function(li){ li.classList.remove('is-search-hidden'); });\n"
        "        document.querySelectorAll('.toc-parent.is-collapsed-search').forEach(function(el){ el.classList.remove('is-collapsed-search'); });\n"
        "        if(tocNoResults) tocNoResults.classList.remove('visible');\n"
        "        return;\n"
        "      }\n"
        "      var re=new RegExp('('+escapeReg(query)+')','gi');\n"
        "      var anyMatch=false;\n"
        "      var matchedIds={};\n"
        "      tocLinksCache.forEach(function(a){\n"
        "        var text=a.textContent||'';\n"
        "        var isMatch=text.toLowerCase().indexOf(query)!==-1;\n"
        "        if(isMatch){\n"
        "          anyMatch=true; matchedIds[a.getAttribute('href')]=true;\n"
        "          var orig=origLinkHTML[a.getAttribute('href')]||a.innerHTML;\n"
        "          a.innerHTML=orig.replace(re,'<mark class=\"toc-mark\">$1</mark>');\n"
        "        } else {\n"
        "          a.innerHTML=origLinkHTML[a.getAttribute('href')]||a.innerHTML;\n"
        "        }\n"
        "      });\n"
        "      tocLinksCache.forEach(function(a){\n"
        "        var li=a.closest('li');\n"
        "        var isDirect=!!matchedIds[a.getAttribute('href')];\n"
        "        var hasDescendant=false;\n"
        "        if(!isDirect){\n"
        "          var ul=li.querySelector('ul');\n"
        "          if(ul){ var descLinks=ul.querySelectorAll('a.toc-link'); for(var i=0;i<descLinks.length;i++){ if(matchedIds[descLinks[i].getAttribute('href')]){ hasDescendant=true; break; } } }\n"
        "        }\n"
        "        var shouldShow=isDirect || hasDescendant;\n"
        "        if(li) li.classList.toggle('is-search-hidden', !shouldShow);\n"
        "        a.classList.toggle('is-search-hidden', !isDirect && !hasDescendant);\n"
        "      });\n"
        "      if(anyMatch){\n"
        "        document.querySelectorAll('.toc-parent').forEach(function(li){\n"
        "          var ul=li.querySelector('ul.toc-children');\n"
        "          if(!ul) return;\n"
        "          var hasVisible=!!li.querySelector('a.toc-link:not(.is-search-hidden)');\n"
        "          if(hasVisible){ li.classList.remove('is-collapsed'); li.setAttribute('aria-expanded','true'); var btn=li.querySelector(':scope > .toc-row > .toc-toggle'); if(btn){ btn.setAttribute('aria-expanded','true'); } ul.hidden=false; }\n"
        "        });\n"
        "      }\n"
        "      if(tocNoResults) tocNoResults.classList.toggle('visible', !anyMatch);\n"
        "    }\n"
        "    if(tocSearch){\n"
        "      tocSearch.addEventListener('input', function(){ clearTimeout(searchTimer); var v=this.value; searchTimer=setTimeout(function(){ doSearch(v); },150); });\n"
        "      tocSearch.addEventListener('keydown', function(e){ if(e.key==='Escape'){ this.value=''; doSearch(''); this.blur(); } });\n"
        "      if(tocClear) tocClear.addEventListener('click', function(){ tocSearch.value=''; doSearch(''); tocSearch.focus(); });\n"
        "    }\n"
        "    document.addEventListener('keydown', function(e){ if((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='k'){ var ae=document.activeElement; if(ae && (ae.tagName==='INPUT' || ae.tagName==='TEXTAREA')) return; e.preventDefault(); if(!isSidebarOpen()) setSidebarOpen(true); setTimeout(function(){ if(tocSearch) tocSearch.focus(); },80); } });\n"
        "    document.addEventListener('keydown', function(e){ if(e.key==='Escape'){ var q=tocSearch?tocSearch.value:''; if(q){ tocSearch.value=''; doSearch(''); } else if(currentFocus){ clearFocus(); history.pushState({focusId:null},'',location.pathname+location.search); } } });\n"
        "    if(headerSearch) headerSearch.addEventListener('click', function(){ if(!isSidebarOpen()) setSidebarOpen(true); setTimeout(function(){ if(tocSearch) tocSearch.focus(); },80); });\n"
        "    if(viewerDownload) viewerDownload.addEventListener('click', function(){ try{ var htmlStr='<!DOCTYPE html>\\n'+document.documentElement.outerHTML; var b=new Blob([htmlStr],{type:'text/html'}); var a=document.createElement('a'); a.href=URL.createObjectURL(b); a.download=(document.title||'document')+'.html'; document.body.appendChild(a); a.click(); setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); },1000); }catch(err){ window.print(); } });\n"
"    function layoutDecimalTabs(){\n"
        "      document.querySelectorAll('.docx-tab-segment[data-val=\"decimal\"]').forEach(function(seg){\n"
        "        var pos=parseInt(seg.getAttribute('data-pos'),10);\n"
        "        if(isNaN(pos)) return;\n"
        "        var text=seg.textContent||'';\n"
        "        var dotIdx=text.indexOf('.');\n"
        "        if(dotIdx===-1){\n"
        "          seg.style.transform=\"translateX(-100%)\";\n"
        "          return;\n"
        "        }\n"
        "        // Skip if already positioned\n"
        "        if(seg.hasAttribute('data-decimal-offset')) return;\n"
        "        try{\n"
        "          // Temporarily place segment at initial tab position for measurement\n"
        "          seg.style.left = pos + 'px';\n"
        "          // Force layout\n"
        "          seg.offsetWidth;\n"
        "          \n"
        "          var walker=document.createTreeWalker(seg, NodeFilter.SHOW_TEXT, null, false);\n"
        "          var node, targetNode=null, localIdx=-1;\n"
        "          while(node=walker.nextNode()){\n"
        "            var t=node.textContent;\n"
        "            var idx=t.indexOf('.');\n"
        "            if(idx!==-1){ targetNode=node; localIdx=idx; break; }\n"
        "          }\n"
        "          if(!targetNode){ seg.style.transform=\"translateX(-100%)\"; return; }\n"
        "          var range=document.createRange();\n"
        "          range.setStart(targetNode, localIdx);\n"
        "          range.setEnd(targetNode, localIdx+1);\n"
        "          var rect=range.getBoundingClientRect();\n"
        "          \n"
        "          // Use segment's own rect as reference (more reliable for abs positioned)\n"
        "          var segRect=seg.getBoundingClientRect();\n"
        "          // decimalOffset within segment = dot position - segment left edge\n"
        "          var decimalOffsetInSeg = rect.left - segRect.left;\n"
        "          // New segment left = tabPos - decimalOffsetInSeg\n"
        "          var newLeft = pos - decimalOffsetInSeg;\n"
        "          seg.style.left = newLeft + 'px';\n"
        "          seg.style.transform = 'none';\n"
        "          seg.setAttribute('data-decimal-offset', decimalOffsetInSeg);\n"
        "        }catch(e){ seg.style.transform=\"translateX(-100%)\"; }\n"
        "      });\n"
        "    }\n"
        "    function layoutTabLeaders(){\n"
        "      document.querySelectorAll('p').forEach(function(p){\n"
        "        var leaders=p.querySelectorAll('.docx-tab-leader[data-pos]');\n"
        "        if(!leaders.length) return;\n"
        "        leaders.forEach(function(leader){\n"
        "          var next=leader.nextElementSibling;\n"
        "          while(next && !next.classList.contains('docx-tab-segment')) next=next.nextElementSibling;\n"
        "          if(!next) return;\n"
        "          var pRect=p.getBoundingClientRect();\n"
        "          var prev=leader.previousElementSibling;\n"
        "          while(prev && prev.classList.contains('docx-tab-leader')) prev=prev.previousElementSibling;\n"
        "          var prevRect=null;\n"
        "          if(prev && prev.classList.contains('docx-tab-segment')){ prevRect=prev.getBoundingClientRect(); }\n"
        "          else { var range=document.createRange(); try{ range.selectNodeContents(p); var r=range.getBoundingClientRect(); prevRect={right: pRect.left, left: pRect.left}; if(r.width) prevRect=r; }catch(e){ prevRect={right:pRect.left}; } }\n"
        "          var nextRect=next.getBoundingClientRect();\n"
        "          var left=(prevRect?prevRect.right: pRect.left) - pRect.left;\n"
        "          var right=nextRect.left - pRect.left;\n"
        "          if(right>left){ leader.style.left=left+'px'; leader.style.width=(right-left)+'px'; leader.style.top='0.7em'; }\n"
        "        });\n"
        "      });\n"
        "    }\n"
        "    function layoutAllTabs(){ layoutDecimalTabs(); layoutTabLeaders(); }\n"
        "    if(document.readyState==='complete') layoutAllTabs(); else window.addEventListener('load', layoutAllTabs);\n"
        "    window.addEventListener('resize', layoutAllTabs);\n"
        "    setTimeout(layoutAllTabs,100); setTimeout(layoutAllTabs,500);\n"
        "    window.__viewer={setSidebarOpen:setSidebarOpen,isSidebarOpen:isSidebarOpen,setActive:setActive};\n"
        "  })();\n"
        "  </script>\n"
        "\n"
    )


def render_html(blocks=None, title: str = "Converted Document", toc=None,
                assets=None, page_layout=None, paragraphs=None, sections=None, even_headers=False,
                default_font_size_pt: float = 11.0) -> str:
    """Render a full HTML document from normalized blocks.

    `blocks` is a document-order list of Paragraph | Table. Heading detection
    and TOC have already run on the paragraph subset; this renderer only lays
    out blocks in the order supplied. Floating images are placed in the correct
    coordinate container as before.

    `paragraphs` is a legacy alias for `blocks` (backward compat with tests
    that call render_html(paragraphs=...)).
    `sections` is an optional list of Section with header/footer variants.
    `default_font_size_pt` is the document's effective default font size (from
    docDefaults rPrDefault); it sets the .docx-content base so inherited runs
    render at the correct size and inline run typography always wins.
    """
    global DEFAULT_FONT_HALF_POINTS
    DEFAULT_FONT_HALF_POINTS = int(round(default_font_size_pt * 2))
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

    def _assign_float_anchor(imgs, anchor_idx):
        for _c, im in imgs:
            try:
                im.anchor_paragraph_index = anchor_idx
            except Exception:
                pass
    def _set_inline_anchors(para, anchor_idx):
        for c in getattr(para, 'content', []) or []:
            if isinstance(c, Image) and c.wrap_type == "anchor":
                try:
                    c.anchor_paragraph_index = anchor_idx
                except Exception:
                    pass

    content_floats = []   # list of (img, anchor_idx)
    page_floats = []      # list of (img, anchor_idx)
    block_html = []
    dom_idx = 0
    idx = 0
    while idx < len(blocks):
        b = blocks[idx]
        if b is first_h1:
            _, _, externals = _render_paragraph_html(b, assets, page_layout)
            _assign_float_anchor(externals, -1)
            _set_inline_anchors(b, -1)
            for container, img in externals:
                if container == "page":
                    page_floats.append((img, -1))
                else:
                    content_floats.append((img, -1))
            idx += 1
            continue
        if isinstance(b, Table):
            cur_anchor = dom_idx
            block_html.append(_wrap_block(render_table(b, assets, page_layout, table_anchor=cur_anchor)))
            dom_idx += 1
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
            # anchor for all floats inside this list block is the list's dom index
            list_anchor = dom_idx
            for item in run:
                _set_inline_anchors(item, list_anchor)
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
                _assign_float_anchor(collected, list_anchor)
                for container, img in collected:
                    if container == "page":
                        page_floats.append((img, list_anchor))
                    else:
                        content_floats.append((img, list_anchor))
                if not inner.strip() and not style_attr:
                    li_inner = "<p></p>"
                elif not inner.strip():
                    li_inner = "<p%s></p>" % style_attr
                else:
                    li_inner = "<p%s>%s</p>" % (style_attr, inner) if style_attr else "<p>%s</p>" % inner
                lis.append("<li>%s</li>" % li_inner)
            block_html.append(_wrap_block('<%s class="docx-list %s">%s</%s>' % (tag, cls, "".join(lis), tag)))
            dom_idx += 1
            idx = j
            continue
        if isinstance(b, Paragraph):
            cur_anchor = dom_idx
            _set_inline_anchors(b, cur_anchor)
            html_str, _n, externals = _render_paragraph_html(b, assets, page_layout)
            _assign_float_anchor(externals, cur_anchor)
            if b.heading_level:
                wrapped = _wrap_block(html_str, b.heading_id, max(1, min(b.heading_level, 6)))
            else:
                wrapped = _wrap_block(html_str)
            block_html.append(wrapped)
            dom_idx += 1
            for container, img in externals:
                if container == "page":
                    page_floats.append((img, cur_anchor))
                else:
                    content_floats.append((img, cur_anchor))
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
    def _float_wrap(img, anchor_idx):
        inner = render_float_image(img, assets, page_layout)
        return '<div class="docx-float-wrap" data-anchor="%d">%s</div>' % (int(anchor_idx), inner)

    content_floats_html = "\n".join(_float_wrap(im, a) for im, a in content_floats)
    page_floats_html = "\n".join(_float_wrap(im, a) for im, a in page_floats)

    content_block = '<div class="docx-content" style="%s">\n%s\n%s\n</div>' % (
        content_style,
        content_inner,
        content_floats_html,
    )

    page_inner = "\n".join([p for p in [header_block, content_block, footer_block] if p])
    page_style = "position: relative; width: %dpx; max-width: calc(100%% - 32px); min-height: %dpx; box-sizing: border-box;" % (
        page_layout.page_width_px, page_layout.page_height_px)
    page_block = '<div class="docx-page" style="%s">\n%s\n%s\n</div>' % (
        page_style,
        page_inner,
        page_floats_html,
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
        '    <span class="sidebar-header-actions"><button id="sidebar-collapse" class="viewer-action" aria-label="Collapse outline" title="Collapse"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M15 18l-6-6 6-6"/></svg></button></span>\n'
        '  </div>\n'
        '  <div class="toc-search-wrap">\n'
        '    <label class="toc-search" aria-label="Search headings">\n'
        '      <span class="toc-search-icon" aria-hidden="true"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg></span>\n'
        '      <input id="toc-search" type="search" placeholder="Search headings..." autocomplete="off" spellcheck="false" aria-label="Search headings">\n'
        '      <button id="toc-search-clear" class="toc-search-clear" aria-label="Clear search" type="button"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 6L6 18"/><path d="M6 6l12 12"/></svg></button>\n'
        '    </label>\n'
        '  </div>\n'
        '  <div id="toc-no-results" class="toc-no-results" role="status" aria-live="polite">No headings found</div>\n'
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
        '  <div id="focus-banner" class="focus-banner" role="status" aria-live="polite">\n'
        '    <div class="focus-banner-inner">\n'
        '      <span class="focus-banner-kicker">Viewing</span>\n'
        '      <span id="focus-banner-title" class="focus-banner-title"></span>\n'
        '      <button id="focus-banner-clear" class="focus-banner-clear" type="button">Show Full Document</button>\n'
        '      <button id="focus-banner-close" class="focus-banner-close" type="button" aria-label="Close focus"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 6L6 18"/><path d="M6 6l12 12"/></svg></button>\n'
        '    </div>\n'
        '  </div>\n'
        '  <button id="exit-focus" class="sidebar-toggle-main" type="button" hidden>Show Full Document</button>\n'
        '</div>\n'
    )
    viewer_header = (
        '<header class="viewer-header" role="banner">\n'
        '  <div class="viewer-brand"><span class="viewer-brand-icon" aria-hidden="true"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg></span> DOCX \u2192 HTML</div>\n'
        '  <div class="viewer-docname" id="viewer-docname" title="%s">%s</div>\n'
        '  <div class="viewer-actions">\n'
        '    <button id="header-search" class="viewer-action" aria-label="Search headings" title="Search (Ctrl+K)"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg></button>\n'
        '    <button id="header-toc" class="viewer-action" aria-label="Toggle outline" aria-controls="viewer-sidebar" aria-expanded="true" title="Toggle outline"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M3 6h18"/><path d="M3 12h18"/><path d="M3 18h18"/></svg></button>\n'
        '    <button id="viewer-download" class="viewer-action primary" aria-label="Download HTML" title="Download HTML"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M12 3v12"/><path d="M7 10l5 5 5-5"/></svg></button>\n'
        '  </div>\n'
        '</header>\n'
    ) % (html.escape(title), html.escape(title))
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
        body = '<div class="page-frame">\n' + viewer_header + '\n' + title_bar + '\n' + viewer_inner + '\n</div>'
    else:
        body = '<div class="page-frame">\n' + viewer_header + '\n' + viewer_inner + '\n</div>'

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
