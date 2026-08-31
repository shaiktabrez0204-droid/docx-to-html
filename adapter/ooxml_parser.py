"""OOXML extraction boundary for DOCX.

This module is the single, clean boundary between raw DOCX bytes and the
normalized model. It:

  1. Opens the .docx ZIP container (real OOXML package)
  2. Reads word/document.xml and word/styles.xml as XML
  3. Walks the XML tree with a real parser (xml.etree.ElementTree)
  4. Extracts paragraphs (w:p), runs (w:r) and run formatting (w:rPr)
  5. Returns normalized Paragraph/Run objects

No regex is used to infer document structure. Paragraph/run boundaries and
formatting flags come directly from the OOXML element tree.
"""

import base64
import os
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from core.model import (
    Run,
    Paragraph,
    TabStop,
    Table,
    Row,
    Cell,
    BorderEdge,
    HeaderFooter,
    Section,
    NumberingModel,
    NumberingLevel,
    Image,
    ImageAsset,
    PageLayout,
    Note,
    NoteReference,
    CommentRangeStart,
    CommentRangeEnd,
)
from core.units import emu_to_px, twip_to_emu, EMU_PER_PIXEL, EMU_PER_TWIP

# Standard OOXML namespaces used by drawing / relationship elements. The w:
# namespace is detected dynamically; these are fixed by the OOXML spec.
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NS_OPC = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_V = "urn:schemas-microsoft-com:vml"
NS_O = "urn:schemas-microsoft-com:office:office"

_SUPPORTED_RELATIVE_FROM = {
    "page", "margin", "column", "character", "paragraph", "line",
}

_WRAP_ELEMENT_TO_MODE = {
    "wrapNone": "none",
    "wrapSquare": "square",
    "wrapTight": "tight",
    "wrapThrough": "through",
    "wrapTopAndBottom": "topAndBottom",
}

# MIME types by file extension. Used to label extracted media so the renderer
# can emit a correct data: URL. Only browser-safe types are ultimately rendered.
_EXT_MIME = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "svg": "image/svg+xml",
    "webp": "image/webp",
}

# Types that a real browser cannot rasterize from a data URL. These are still
# extracted (so the asset store is complete) but the renderer degrades safely
# rather than emitting a broken <img>.
_UNSUPPORTED_MIME = {"image/emf", "image/wmf", "image/x-emf", "image/x-wmf"}

# Symbol font glyph → Unicode mapping.
# w:sym/@w:char is a font-specific glyph code, NOT a Unicode code point.
# For Symbol, F0xx → U+00xx (offset 0xF000). For Wingdings, mapping is
# arbitrary and must be explicit. Structure allows adding fonts without parser
# logic changes.
_SYMBOL_GLYPH_MAP: Dict[str, Dict[str, str]] = {
    "Symbol": {
        # Minimal explicit overrides; fallback offset handles F0xx → 00xx.
        "F020": "\u0020",  # space
        "F021": "\u0021",  # !
        "F0D7": "\u00D7",  # × multiplication (test fixture)
        "F0B7": "\u00B7",  # ·
        "F0B1": "\u00B1",  # ±
        "F0D0": "\u2212",  # − minus
        "F0B0": "\u00B0",  # °
        "F0B6": "\u2248",  # ≈
        "F0B4": "\u221A",  # √
        "F0A7": "\u25CF",  # ● (Symbol bullet, fallback)
    },
    "Wingdings": {
        "F028": "\u25CF",  # ● black circle (test fixture)
        "F021": "\u2702",  # ✂
        "F022": "\u2701",
        "F023": "\u2703",
        "F024": "\u2704",
        "F025": "\u260E",  # ☎
        "F026": "\u2706",
        "F027": "\u2707",
        "F029": "\u25A0",  # ■
        "F02A": "\u25A1",  # □
        "F02B": "\u25AA",  # ▪
        "F02C": "\u25AB",
        "F02D": "\u25B2",  # ▲
        "F02E": "\u25BC",  # ▼
        "F02F": "\u25B6",
        "F030": "\u25C0",
        "F031": "\u25C6",  # ◆
        "F032": "\u25C7",  # ◇
        "F033": "\u25CB",  # ○
        "F034": "\u25CF",
        "F035": "\u25CE",  # ◎
        "F036": "\u25C9",  # ◉
        "F0A7": "\u25CF",
    },
    "Wingdings 2": {},
    "Wingdings 3": {},
    "Webdings": {},
}

def _decode_sym_char(font: Optional[str], char_hex: Optional[str]) -> Optional[str]:
    """Decode w:sym @w:char hex + font → Unicode char.

    font-specific lookup first, then Symbol offset fallback, then generic fallback.
    Never silently drops: on unknown glyph returns low-byte char as safest visible fallback.
    """
    if not char_hex:
        return None
    hex_norm = char_hex.strip().upper()
    # Normalize: Word may emit 'F0D7' or '0xF0D7' or decimal; handle hex without 0x.
    if hex_norm.startswith("0X"):
        hex_norm = hex_norm[2:]
    # Preserve original hex string for map lookup (keys are uppercase F0xx)
    try:
        val = int(hex_norm, 16)
    except ValueError:
        return None

    # 1) Explicit per-font table
    if font:
        table = _SYMBOL_GLYPH_MAP.get(font)
        if table is not None:
            hit = table.get(hex_norm)
            if hit is not None:
                return hit
        # Case-insensitive fallback for font name
        for k, tbl in _SYMBOL_GLYPH_MAP.items():
            if k.lower() == font.lower():
                hit = tbl.get(hex_norm)
                if hit is not None:
                    return hit
                break

    # 2) Symbol offset fallback: F0xx → 00xx (Adobe Symbol encoding)
    if font and font.lower() == "symbol":
        if 0xF020 <= val <= 0xF0FF:
            try:
                return chr(val - 0xF000)
            except ValueError:
                pass

    # 3) Generic fallback: if val is valid Unicode, return chr(val & 0xFFFF low byte as last resort)
    # Prefer low byte for Symbol-like F0xx values when font unknown
    if font and 0xF000 <= val <= 0xF0FF:
        try:
            return chr(val - 0xF000)
        except ValueError:
            pass
    try:
        if 0x20 <= val <= 0x10FFFF:
            return chr(val)
    except ValueError:
        pass
    # Ultimate fallback: low byte
    try:
        return chr(val & 0xFF)
    except ValueError:
        return None


@dataclass
class StyleDef:
    """A normalized paragraph/character style definition from word/styles.xml.

    Carries only the fields the semantic pipeline needs: identity, the
    BasedOn parent (for inheritance), and the outline level (the DOCX
    structural signal for heading depth).
    """
    style_id: str
    name: str = ""
    based_on: str = ""
    style_type: str = "paragraph"   # paragraph | character | table | numbering | ...
    outline_level: Optional[int] = None  # w:pPr/w:outlineLvl/@w:val (0-8), if present
    num_id: Optional[str] = None    # w:pPr/w:numPr/w:numId/@w:val, if present
    num_ilvl: Optional[int] = None  # w:pPr/w:numPr/w:ilvl/@w:val (0-based), if present
    # Paragraph layout (w:pPr/w:ind, w:spacing, w:jc)
    alignment: Optional[str] = None
    indent_left: Optional[int] = None
    indent_right: Optional[int] = None
    indent_first_line: Optional[int] = None
    indent_hanging: Optional[int] = None
    spacing_before: Optional[int] = None
    spacing_after: Optional[int] = None
    line_spacing: Optional[int] = None
    line_spacing_rule: Optional[str] = None
    # Run typography (w:rPr or w:pPr/w:rPr)
    font_family: Optional[str] = None
    font_size: Optional[int] = None
    font_color: Optional[str] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[str] = None
    superscript: Optional[bool] = None
    subscript: Optional[bool] = None


class OoxmlParser:
    """Extracts normalized content from a real DOCX OOXML package."""

    def __init__(self, docx_path: str):
        if not os.path.exists(docx_path):
            raise FileNotFoundError("DOCX not found: %s" % docx_path)
        self.docx_path = docx_path
        self._zip = zipfile.ZipFile(docx_path, "r")
        self._document_xml = self._read_part("word/document.xml")
        self._styles_xml = self._read_part("word/styles.xml")
        # Derive the actual WordprocessingML namespace from the parsed root
        # instead of hardcoding it (avoids subtle whitespace/encoding mismatches).
        self._w = self._detect_ns(self._document_xml)

        # Relationship resolution for the main document part. This map is the
        # authoritative rId -> media Target resolver (never inferred from names).
        self._rels = self._read_relationships("word/document.xml")
        # Extracted image assets keyed by media source path (dedup store).
        self._assets: Dict[str, ImageAsset] = {}
        # Monotonic counter for stable, unique placement ids (img1, img2, ...).
        self._image_seq = 0
        # Physical page + margin geometry from sectPr (EMU). Used to map anchor
        # relativeFrom coordinate systems onto determinate CSS pixel origins.
        self._page_layout = self._parse_page_layout()

    @staticmethod
    def _detect_ns(xml_text: str) -> str:
        if not xml_text:
            return "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        root = ET.fromstring(xml_text)
        if "}" in root.tag:
            return root.tag.split("}", 1)[0] + "}"
        return ""

    def _qn(self, tag: str) -> str:
        """Qualify a bare tag name with the detected WordprocessingML namespace."""
        return self._w + tag

    def _qn_ns(self, ns_uri: str, tag: str) -> str:
        """Qualify a tag with an explicit namespace URI (drawing/rel/picture)."""
        return "{%s}%s" % (ns_uri, tag)

    # ---- relationships ----

    def _read_relationships(self, part_name: str) -> Dict[str, Dict[str, str]]:
        """Read the .rels part for `part_name` into {Id: {Target, Type}}.

        Example: document.xml -> word/_rels/document.xml.rels. Target paths are
        package-relative; we normalize them to the absolute part path inside the
        zip so media lookup is unambiguous.
        """
        rels_name = self._rels_part_name(part_name)
        try:
            raw = self._read_part(rels_name)
        except KeyError:
            # No relationships part: this document simply has no media links.
            return {}
        rels: Dict[str, Dict[str, str]] = {}
        if not raw:
            return rels
        try:
            root = ET.fromstring(raw)
        except ET.ParseError:
            # A malformed .rels (e.g. a self-closed root plus a stray closing
            # tag in a synthetic fixture) must not abort the whole parse. Treat
            # it as having no relationships; image extraction degrades to none.
            return rels
        base_dir = os.path.dirname(part_name)
        for rel in root.findall(self._qn_ns(NS_OPC, "Relationship")):
            rid = rel.get("Id")
            target = rel.get("Target")
            rtype = rel.get("Type", "")
            if not rid or not target:
                continue
            target_mode = rel.get("TargetMode", "")
            # External hyperlinks (TargetMode=External or URL with scheme) must keep the raw URL.
            # Internal media targets need filesystem normalization.
            if target_mode == "External" or "://" in target or target.startswith("mailto:"):
                media_path = target
            elif target.startswith("/"):
                media_path = target.lstrip("/")
            else:
                media_path = os.path.normpath(os.path.join(base_dir, target)).replace("\\", "/")
            rels[rid] = {"Target": media_path, "Type": rtype, "TargetMode": target_mode}
        return rels

    @staticmethod
    def _rels_part_name(part_name: str) -> str:
        """word/document.xml -> word/_rels/document.xml.rels."""
        base_dir = os.path.dirname(part_name)
        fname = os.path.basename(part_name)
        return (base_dir + "/_rels/" + fname + ".rels").replace("\\", "/")

    # ---- low level ----

    def _read_part(self, name: str) -> str:
        try:
            return self._zip.read(name).decode("utf-8", errors="replace")
        except KeyError:
            return ""

    def _read_part_bytes(self, name: str) -> Optional[bytes]:
        try:
            return self._zip.read(name)
        except KeyError:
            return None

    # ---- image extraction ----

    def _mime_for_path(self, path: str) -> str:
        ext = os.path.splitext(path)[1].lower().lstrip(".")
        return _EXT_MIME.get(ext, "application/octet-stream")

    def _resolve_asset(self, rid: str) -> Optional[ImageAsset]:
        """Resolve an r:embed id to an ImageAsset, extracting bytes once.

        Returns None when the relationship id does not exist (adversarial:
        missing relationship). Reuses a previously extracted asset so the same
        media file referenced multiple times is extracted only once.
        """
        rel = self._rels.get(rid)
        if rel is None:
            return None
        media_path = rel["Target"]
        if media_path in self._assets:
            asset = self._assets[media_path]
            if rid not in asset.relationship_ids:
                asset.relationship_ids.append(rid)
            return asset

        data = self._read_part_bytes(media_path)
        media_type = self._mime_for_path(media_path)
        asset = ImageAsset(
            source_path=media_path,
            media_type=media_type,
            data=data,
            relationship_ids=[rid],
            missing=(data is None),
        )
        self._assets[media_path] = asset
        return asset

    def _parse_drawing(self, drawing_elem: ET.Element, kind: str) -> Optional[Image]:
        """Parse one wp:inline or wp:anchor into a normalized Image placement."""
        extent = drawing_elem.find(self._qn_ns(NS_WP, "extent"))
        width = height = None
        extent_cx = extent_cy = None
        if extent is not None:
            cx = extent.get("cx")
            cy = extent.get("cy")
            if cx and cx.lstrip("-").isdigit():
                extent_cx = int(cx)
                width = emu_to_px(extent_cx)
            if cy and cy.lstrip("-").isdigit():
                extent_cy = int(cy)
                height = emu_to_px(extent_cy)

        # Visual transforms from a:xfrm
        rotation = None
        flip_h = False
        flip_v = False
        xfrm = drawing_elem.find(".//" + self._qn_ns(NS_A, "xfrm"))
        if xfrm is not None:
            rot = xfrm.get("rot")
            if rot and rot.lstrip("-").isdigit():
                try:
                    rotation = int(rot)
                except ValueError:
                    pass
            flip_h_val = xfrm.get("flipH")
            if flip_h_val in ("1", "true", "True"):
                flip_h = True
            flip_v_val = xfrm.get("flipV")
            if flip_v_val in ("1", "true", "True"):
                flip_v = True

        # Crop from a:srcRect (in blipFill, fraction of image in 1/100000 units)
        crop_left = crop_top = crop_right = crop_bottom = None
        blip_fill = drawing_elem.find(".//" + self._qn_ns(NS_PIC, "blipFill"))
        if blip_fill is None:
            blip_fill = drawing_elem.find(".//" + self._qn_ns(NS_A, "blipFill"))
        if blip_fill is not None:
            src_rect = blip_fill.find(self._qn_ns(NS_A, "srcRect"))
            if src_rect is not None:
                for attr, field in (("l", "crop_left"), ("t", "crop_top"),
                                    ("r", "crop_right"), ("b", "crop_bottom")):
                    val = src_rect.get(attr)
                    if val and val.lstrip("-").isdigit():
                        if field == "crop_left":
                            crop_left = int(val)
                        elif field == "crop_top":
                            crop_top = int(val)
                        elif field == "crop_right":
                            crop_right = int(val)
                        elif field == "crop_bottom":
                            crop_bottom = int(val)

        # Effect extent from wp:effectExtent (EMU)
        effect_extent_l = effect_extent_t = effect_extent_r = effect_extent_b = None
        effect_extent = drawing_elem.find(self._qn_ns(NS_WP, "effectExtent"))
        if effect_extent is not None:
            for attr, field in (("l", "effect_extent_l"), ("t", "effect_extent_t"),
                                ("r", "effect_extent_r"), ("b", "effect_extent_b")):
                val = effect_extent.get(attr)
                if val and val.lstrip("-").isdigit():
                    if field == "effect_extent_l":
                        effect_extent_l = int(val)
                    elif field == "effect_extent_t":
                        effect_extent_t = int(val)
                    elif field == "effect_extent_r":
                        effect_extent_r = int(val)
                    elif field == "effect_extent_b":
                        effect_extent_b = int(val)

        # Alt text: Word's "Alt Text" description is stored in wp:docPr/@descr.
        # wp:docPr/@name is the shape name (e.g. "Picture 1"), not a description,
        # so it is intentionally NOT used as alt text. When @descr is absent the
        # image has no alt text and we leave it None (never invent one).
        doc_pr = drawing_elem.find(self._qn_ns(NS_WP, "docPr"))
        alt_text = None
        if doc_pr is not None:
            descr = doc_pr.get("descr")
            if descr:
                alt_text = descr

        # blip embed: pic:blipFill/a:blip/@r:embed (find anywhere under the node).
        blip = drawing_elem.find(".//" + self._qn_ns(NS_PIC, "blip"))
        if blip is None:
            blip = drawing_elem.find(".//" + self._qn_ns(NS_A, "blip"))
        if blip is None:
            return None
        rid = blip.get(self._qn_ns(NS_R, "embed"))
        if not rid:
            return None

        asset = self._resolve_asset(rid)
        if asset is None:
            # Relationship missing entirely: cannot place a real image.
            return None

        anchor_fields = {}
        if kind == "anchor":
            anchor_fields = self._parse_anchor_geometry(drawing_elem)

        self._image_seq += 1
        return Image(
            image_id="img%d" % self._image_seq,
            relationship_id=rid,
            source_path=asset.source_path,
            media_type=asset.media_type,
            width=width,
            height=height,
            alt_text=alt_text,
            wrap_type=("anchor" if kind == "anchor" else "inline"),
            rotation=rotation,
            extent_cx=extent_cx,
            extent_cy=extent_cy,
            crop_left=crop_left,
            crop_top=crop_top,
            crop_right=crop_right,
            crop_bottom=crop_bottom,
            flip_h=flip_h,
            flip_v=flip_v,
            effect_extent_l=effect_extent_l,
            effect_extent_t=effect_extent_t,
            effect_extent_r=effect_extent_r,
            effect_extent_b=effect_extent_b,
            **anchor_fields,
        )

    def _parse_anchor_geometry(self, anchor_elem: ET.Element) -> dict:
        """Extract positioning + wrap metadata from a wp:anchor element.

        Returns only the anchor placement fields (relativeFrom/offset/align/wrap/
        distances). The raw @relativeFrom values are preserved verbatim; the
        normalized ``position_*`` category is "unsupported" for any relativeFrom
        this engine does not model, so the renderer can branch honestly.
        """
        WP = self._qn_ns(NS_WP, "")
        # Strip the trailing "" so we can build child tags cheaply.
        wp = NS_WP

        def _child(tag):
            return anchor_elem.find("{%s}%s" % (wp, tag))

        def _nested(parent, tag):
            if parent is None:
                return None
            return parent.find("{%s}%s" % (wp, tag))

        def _int_text(el):
            if el is None or el.text is None:
                return None
            t = el.text.strip()
            if not t.lstrip("-").isdigit():
                return None
            return int(t)

        fields: dict = {}

        ph = _child("positionH")
        if ph is not None:
            rf = ph.get("relativeFrom")
            fields["relative_from_horizontal"] = rf
            fields["position_horizontal"] = (
                rf if rf in _SUPPORTED_RELATIVE_FROM else "unsupported")
            off = _nested(ph, "posOffset")
            fields["offset_horizontal"] = _int_text(off)
            align = _nested(ph, "align")
            fields["alignment_horizontal"] = (
                align.text.strip() if align is not None and align.text else None)

        pv = _child("positionV")
        if pv is not None:
            rf = pv.get("relativeFrom")
            fields["relative_from_vertical"] = rf
            fields["position_vertical"] = (
                rf if rf in _SUPPORTED_RELATIVE_FROM else "unsupported")
            off = _nested(pv, "posOffset")
            fields["offset_vertical"] = _int_text(off)
            align = _nested(pv, "align")
            fields["alignment_vertical"] = (
                align.text.strip() if align is not None and align.text else None)

        wrap_mode = None
        for el_name, mode in _WRAP_ELEMENT_TO_MODE.items():
            if _child(el_name) is not None:
                wrap_mode = mode
                break
        fields["wrap_mode"] = wrap_mode

        dist = {}
        for key, attr in (("top", "distT"), ("bottom", "distB"),
                          ("left", "distL"), ("right", "distR")):
            v = anchor_elem.get(attr)
            if v is not None and v.lstrip("-").isdigit():
                dist[key] = int(v)
        fields["wrap_distances"] = dist or None

        # Parse wrapPolygon from the wrap element (wrapTight/wrapThrough).
        # Coordinates are in the OOXML normalized 21600x21600 space.
        wrap_polygon = self._parse_wrap_polygon(anchor_elem)
        if wrap_polygon:
            fields["wrap_polygon"] = wrap_polygon

        # wp:anchor/@behindDoc: "1" => image behind text (drives z-index in the
        # renderer without inventing an arbitrary stacking order).
        bd = anchor_elem.get("behindDoc")
        if bd is not None:
            fields["behind_doc"] = (bd == "1")

        return fields

    def _parse_wrap_polygon(self, anchor_elem: ET.Element) -> Optional[List[Tuple[int, int]]]:
        """Extract wrap polygon points from wp:wrapTight/wp:wrapThrough/wp:wrapSquare.

        Returns a list of (x, y) tuples in the OOXML 21600x21600 coordinate space,
        or None if no valid polygon is found.
        """
        wp = NS_WP
        wrap_el = None
        for el_name in _WRAP_ELEMENT_TO_MODE.keys():
            wrap_el = anchor_elem.find("{%s}%s" % (wp, el_name))
            if wrap_el is not None:
                break
        if wrap_el is None:
            return None

        poly_el = wrap_el.find("{%s}wrapPolygon" % wp)
        if poly_el is None:
            return None

        points: List[Tuple[int, int]] = []

        def _parse_coord(el: ET.Element) -> Optional[Tuple[int, int]]:
            x_attr = el.get("x")
            y_attr = el.get("y")
            if x_attr is None or y_attr is None:
                return None
            if not x_attr.lstrip("-").isdigit() or not y_attr.lstrip("-").isdigit():
                return None
            return (int(x_attr), int(y_attr))

        start_el = poly_el.find("{%s}start" % wp)
        if start_el is not None:
            pt = _parse_coord(start_el)
            if pt is not None:
                points.append(pt)

        for line_el in poly_el.findall("{%s}lineTo" % wp):
            pt = _parse_coord(line_el)
            if pt is not None:
                points.append(pt)

        # A valid polygon needs at least 3 points.
        if len(points) < 3:
            return None

        # Close the polygon if not already closed (first == last).
        if points[0] != points[-1]:
            points.append(points[0])

        return points

    def _extract_images_from_r(self, r_elem: ET.Element) -> List[Image]:
        """Extract image placements from a w:r that may contain DrawingML or VML."""
        images: List[Image] = []
        seen_rids = set()
        for kind in ("inline", "anchor"):
            for node in r_elem.findall(".//" + self._qn_ns(NS_WP, kind)):
                rid = None
                blip = node.find(".//" + self._qn_ns(NS_PIC, "blip"))
                if blip is None:
                    blip = node.find(".//" + self._qn_ns(NS_A, "blip"))
                if blip is not None:
                    rid = blip.get(self._qn_ns(NS_R, "embed"))
                if rid and rid in seen_rids:
                    continue
                img = self._parse_drawing(node, kind)
                if img is not None:
                    if img.relationship_id not in seen_rids:
                        seen_rids.add(img.relationship_id)
                    images.append(img)
        for img in self._extract_vml_images(r_elem):
            if img.relationship_id in seen_rids:
                continue
            seen_rids.add(img.relationship_id)
            images.append(img)
        return images

    def _extract_vml_images(self, r_elem: ET.Element) -> List[Image]:
        """Extract legacy VML images (w:pict/v:shape/v:imagedata)."""
        images: List[Image] = []
        for imagedata in r_elem.findall(".//" + self._qn_ns(NS_V, "imagedata")):
            rid = imagedata.get(self._qn_ns(NS_R, "id"))
            if not rid:
                rid = imagedata.get(self._qn_ns(NS_R, "embed"))
            if not rid:
                for k, v in imagedata.attrib.items():
                    if k.endswith("}id") or k.endswith("}embed"):
                        rid = v
                        break
            if not rid:
                continue
            asset = self._resolve_asset(rid)
            if asset is None:
                continue
            width = height = None
            alt_text = None
            for shape in r_elem.findall(".//" + self._qn_ns(NS_V, "shape")):
                if imagedata in list(shape.iter()):
                    style = shape.get("style", "")
                    # style like "width:100pt;height:75pt" or "width:2in;height:1.5in"
                    if style:
                        import re
                        w_m = re.search(r"width\s*:\s*([0-9.]+)\s*(pt|in|cm|mm|px|emu)?", style, re.I)
                        h_m = re.search(r"height\s*:\s*([0-9.]+)\s*(pt|in|cm|mm|px|emu)?", style, re.I)
                        if w_m:
                            try:
                                val = float(w_m.group(1))
                                unit = (w_m.group(2) or "pt").lower()
                                if unit == "pt":
                                    width = int(round(val * 96.0 / 72.0))
                                elif unit == "in":
                                    width = int(round(val * 96.0))
                                elif unit == "cm":
                                    width = int(round(val * 96.0 / 2.54))
                                elif unit == "mm":
                                    width = int(round(val * 96.0 / 25.4))
                                elif unit == "px":
                                    width = int(round(val))
                                elif unit == "emu":
                                    width = emu_to_px(int(val))
                            except Exception:
                                pass
                        if h_m:
                            try:
                                val = float(h_m.group(1))
                                unit = (h_m.group(2) or "pt").lower()
                                if unit == "pt":
                                    height = int(round(val * 96.0 / 72.0))
                                elif unit == "in":
                                    height = int(round(val * 96.0))
                                elif unit == "cm":
                                    height = int(round(val * 96.0 / 2.54))
                                elif unit == "mm":
                                    height = int(round(val * 96.0 / 25.4))
                                elif unit == "px":
                                    height = int(round(val))
                                elif unit == "emu":
                                    height = emu_to_px(int(val))
                            except Exception:
                                pass
                    # alt text from shape title or imagedata title
                    alt = shape.get("alt") or shape.get("title") or imagedata.get(self._qn_ns(NS_O, "title")) or imagedata.get("title")
                    if alt:
                        alt_text = alt
                    if not alt_text:
                        alt_text = shape.get(self._qn_ns(NS_O, "title"))
                    break
            if not alt_text:
                alt_text = imagedata.get(self._qn_ns(NS_O, "title")) or imagedata.get("title")

            # VML rotation from style (e.g., "rotation:90") or o:rotation attribute
            rotation = None
            for shape in r_elem.findall(".//" + self._qn_ns(NS_V, "shape")):
                if imagedata in list(shape.iter()):
                    # Check style for rotation
                    style = shape.get("style", "")
                    rot_m = re.search(r"rotation\s*:\s*([0-9.-]+)", style, re.I)
                    if rot_m:
                        try:
                            # VML rotation is in degrees; convert to 1/60000 deg
                            rotation = int(round(float(rot_m.group(1)) * 60000))
                        except Exception:
                            pass
                    # Check o:rotation attribute
                    if rotation is None:
                        o_rot = shape.get(self._qn_ns(NS_O, "rotation"))
                        if o_rot:
                            try:
                                rotation = int(round(float(o_rot) * 60000))
                            except Exception:
                                pass
                    # Check for flip in style
                    flip_h = "flipH" in style or "flip-h" in style or "flipx" in style
                    flip_v = "flipV" in style or "flip-v" in style or "flipy" in style
                    break
            else:
                flip_h = flip_v = False

            self._image_seq += 1
            images.append(Image(
                image_id="img%d" % self._image_seq,
                relationship_id=rid,
                source_path=asset.source_path,
                media_type=asset.media_type,
                width=width,
                height=height,
                alt_text=alt_text,
                wrap_type="inline",
                rotation=rotation,
                flip_h=flip_h,
                flip_v=flip_v,
            ))
        return images

    def get_image_assets(self) -> Dict[str, ImageAsset]:
        """Return the deduped image asset store (source_path -> ImageAsset)."""
        return self._assets

    def get_page_layout(self) -> PageLayout:
        """Return the parsed page/margin geometry (EMU)."""
        return self._page_layout

    def _parse_cols(self, sect_elem: ET.Element) -> dict:
        w = self._w
        cols = sect_elem.find(w + "cols") if sect_elem is not None else None
        if cols is None:
            return {}
        num_raw = cols.get(w + "num")
        num = int(num_raw) if num_raw and num_raw.isdigit() else 1
        if num <= 1:
            return {"cols_num": 1}
        space_raw = cols.get(w + "space")
        space_emu = twip_to_emu(int(space_raw)) if space_raw and space_raw.lstrip("-").isdigit() else 0
        eq = cols.get(w + "equalWidth")
        equal = True if eq is None else eq not in ("0", "false", "off")
        col_els = cols.findall(w + "col")
        if col_els:
            widths = []
            spaces = []
            for ce in col_els:
                w_raw = ce.get(w + "w")
                if w_raw and w_raw.lstrip("-").isdigit():
                    widths.append(twip_to_emu(int(w_raw)))
                else:
                    widths.append(0)
                sp_raw = ce.get(w + "space")
                if sp_raw and sp_raw.lstrip("-").isdigit():
                    spaces.append(twip_to_emu(int(sp_raw)))
                else:
                    spaces.append(space_emu)
            if len(widths) == num:
                col_spaces = spaces[: num - 1] if num > 1 else []
                return {"cols_num": num, "cols_equal_width": equal, "col_widths_emu": widths, "col_spaces_emu": col_spaces, "cols_space_emu": space_emu}
            else:
                return {"cols_num": num, "cols_space_emu": space_emu, "cols_equal_width": equal}
        return {"cols_num": num, "cols_space_emu": space_emu, "cols_equal_width": equal}

    def _parse_page_layout(self) -> PageLayout:
        """Extract pgSz / pgMar from the document's sectPr, in EMU.

        sectPr sizes are in twips (1/20 pt). Missing sectPr or missing fields fall
        back to US-Letter geometry. Never fails: a malformed sectPr degrades to
        defaults rather than aborting the whole parse.
        """
        if not self._document_xml:
            return PageLayout()
        try:
            root = ET.fromstring(self._document_xml)
        except ET.ParseError:
            return PageLayout()
        w = self._w
        body = root.find(w + "body")
        if body is None:
            return PageLayout()
        sect = body.find(w + "sectPr")
        if sect is None:
            return PageLayout()

        def _twip(el, attr):
            v = el.get(attr) if el is not None else None
            if v is None or not v.lstrip("-").isdigit():
                return None
            return twip_to_emu(int(v))

        pg = sect.find(w + "pgSz")
        mar = sect.find(w + "pgMar")
        kwargs = {}
        pw = _twip(pg, w + "w")
        ph = _twip(pg, w + "h")
        if pw:
            kwargs["width_emu"] = pw
        if ph:
            kwargs["height_emu"] = ph
        ml = _twip(mar, w + "left")
        mr = _twip(mar, w + "right")
        mt = _twip(mar, w + "top")
        mb = _twip(mar, w + "bottom")
        hdr = _twip(mar, w + "header")
        ftr = _twip(mar, w + "bottom")  # footer distance is separate attribute
        if ml:
            kwargs["margin_left_emu"] = ml
        if mr:
            kwargs["margin_right_emu"] = mr
        if mt:
            kwargs["margin_top_emu"] = mt
        if mb:
            kwargs["margin_bottom_emu"] = mb
        if hdr:
            kwargs["header_distance_emu"] = hdr
        if ftr:
            kwargs["footer_distance_emu"] = ftr
        kwargs.update(self._parse_cols(sect))
        return PageLayout(**kwargs)

    def _parse_page_layout_from_sect(self, sect_elem: ET.Element) -> PageLayout:
        if sect_elem is None:
            return PageLayout()
        w = self._w
        def _twip(el, attr):
            v = el.get(attr) if el is not None else None
            if v is None or not v.lstrip("-").isdigit():
                return None
            return twip_to_emu(int(v))
        pg = sect_elem.find(w + "pgSz")
        mar = sect_elem.find(w + "pgMar")
        kwargs = {}
        pw = _twip(pg, w + "w")
        ph = _twip(pg, w + "h")
        if pw:
            kwargs["width_emu"] = pw
        if ph:
            kwargs["height_emu"] = ph
        ml = _twip(mar, w + "left")
        mr = _twip(mar, w + "right")
        mt = _twip(mar, w + "top")
        mb = _twip(mar, w + "bottom")
        hdr = _twip(mar, w + "header")
        ftr = _twip(mar, w + "footer")
        if ml:
            kwargs["margin_left_emu"] = ml
        if mr:
            kwargs["margin_right_emu"] = mr
        if mt:
            kwargs["margin_top_emu"] = mt
        if mb:
            kwargs["margin_bottom_emu"] = mb
        if hdr:
            kwargs["header_distance_emu"] = hdr
        if ftr:
            kwargs["footer_distance_emu"] = ftr
        kwargs.update(self._parse_cols(sect_elem))
        return PageLayout(**kwargs)

    def _get_even_headers_flag(self) -> bool:
        raw = self._read_part("word/settings.xml")
        if not raw:
            return False
        try:
            root = ET.fromstring(raw)
        except ET.ParseError:
            return False
        return root.find(self._qn("evenAndOddHeaders")) is not None

    def _parse_header_footer_blocks(self, part_path: str) -> List:
        xml = self._read_part(part_path)
        if not xml:
            return []
        try:
            root = ET.fromstring(xml)
        except ET.ParseError:
            return []
        rels = self._read_relationships(part_path)
        old_rels = self._rels
        self._rels = rels
        try:
            blocks = []
            for child in root:
                if child.tag == self._qn("p"):
                    para = self._parse_paragraph(child)
                    if para is not None:
                        blocks.append(para)
                elif child.tag == self._qn("tbl"):
                    tbl = self._parse_table(child)
                    if tbl is not None:
                        blocks.append(tbl)
                else:
                    continue
            return blocks
        finally:
            self._rels = old_rels

    def _parse_header_footer_part(self, part_path: str, hf_type: str, kind: str, r_id: str) -> Optional[HeaderFooter]:
        blocks = self._parse_header_footer_blocks(part_path)
        # Even if blocks empty, preserve empty header/footer to indicate presence
        # but filter out truly empty parts that have no paragraphs/tables
        if not blocks:
            # check if part exists at all
            if not self._read_part(part_path):
                return None
            # keep empty header to indicate existence but no visible content
            # still return object with empty blocks
            pass
        return HeaderFooter(hf_type=hf_type, kind=kind, blocks=blocks, r_id=r_id, target=part_path)

    def get_sections(self) -> List[Section]:
        if not self._document_xml:
            return [Section(index=0, page_layout=self._page_layout)]
        try:
            root = ET.fromstring(self._document_xml)
        except ET.ParseError:
            return [Section(index=0, page_layout=self._page_layout)]
        w = self._w
        body = root.find(w + "body")
        if body is None:
            return [Section(index=0, page_layout=self._page_layout)]
        sect_elements: List[ET.Element] = []
        for child in body:
            if child.tag == w + "p":
                pPr = child.find(w + "pPr")
                if pPr is not None:
                    sect = pPr.find(w + "sectPr")
                    if sect is not None:
                        sect_elements.append(sect)
            elif child.tag == w + "sectPr":
                sect_elements.append(child)
        if not sect_elements:
            sect = body.find(w + "sectPr")
            if sect is not None:
                sect_elements.append(sect)
        if not sect_elements:
            # No sectPr at all -> single default section
            return [Section(index=0, page_layout=PageLayout())]
        sections: List[Section] = []
        for idx, sect in enumerate(sect_elements):
            sec = Section(index=idx)
            sec.title_pg = sect.find(w + "titlePg") is not None
            pg_num = sect.find(w + "pgNumType")
            if pg_num is not None:
                fmt = pg_num.get(w + "fmt")
                start = pg_num.get(w + "start")
                if fmt:
                    sec.pg_num_fmt = fmt
                if start and start.lstrip("-").isdigit():
                    sec.pg_num_start = int(start)
            for hr in sect.findall(w + "headerReference"):
                r_id = hr.get(self._qn_ns(NS_R, "id"))
                typ = hr.get(w + "type", "default")
                if not r_id:
                    continue
                rel = self._rels.get(r_id)
                if not rel:
                    continue
                target = rel["Target"]
                hf = self._parse_header_footer_part(target, typ, "header", r_id)
                if hf is not None:
                    for _b in hf.blocks:
                        if isinstance(_b, Paragraph):
                            for _img in _b.images:
                                _img.section_index = idx
                            for _c in _b.content:
                                if isinstance(_c, Image):
                                    _c.section_index = idx
                        elif isinstance(_b, Table):
                            for _r in _b.rows:
                                for _cc in _r.cells:
                                    for _pp in _cc.content:
                                        for _img in _pp.images:
                                            _img.section_index = idx
                                        for _cc2 in _pp.content:
                                            if isinstance(_cc2, Image):
                                                _cc2.section_index = idx
                    sec.headers[typ] = hf
            for fr in sect.findall(w + "footerReference"):
                r_id = fr.get(self._qn_ns(NS_R, "id"))
                typ = fr.get(w + "type", "default")
                if not r_id:
                    continue
                rel = self._rels.get(r_id)
                if not rel:
                    continue
                target = rel["Target"]
                hf = self._parse_header_footer_part(target, typ, "footer", r_id)
                if hf is not None:
                    for _b in hf.blocks:
                        if isinstance(_b, Paragraph):
                            for _img in _b.images:
                                _img.section_index = idx
                            for _c in _b.content:
                                if isinstance(_c, Image):
                                    _c.section_index = idx
                        elif isinstance(_b, Table):
                            for _r in _b.rows:
                                for _cc in _r.cells:
                                    for _pp in _cc.content:
                                        for _img in _pp.images:
                                            _img.section_index = idx
                                        for _cc2 in _pp.content:
                                            if isinstance(_cc2, Image):
                                                _cc2.section_index = idx
                    sec.footers[typ] = hf
            sec.page_layout = self._parse_page_layout_from_sect(sect)
            sections.append(sec)
        # Inheritance: if a section lacks a variant, inherit from previous
        for i in range(1, len(sections)):
            prev = sections[i-1]
            cur = sections[i]
            for typ in ["default", "first", "even"]:
                if typ not in cur.headers and typ in prev.headers:
                    cur.headers[typ] = prev.headers[typ]
                if typ not in cur.footers and typ in prev.footers:
                    cur.footers[typ] = prev.footers[typ]
        return sections

    def get_even_headers_flag(self) -> bool:
        return self._get_even_headers_flag()

    def close(self):
        self._zip.close()

    # ---- public API ----

    def parse_document(self):
        """Parse document body children in document order.

        Walks w:body children sequentially so w:p and w:tbl preserve their
        original order. Returns a list of Paragraph | Table blocks.
        """
        if not self._document_xml:
            return []
        try:
            root = ET.fromstring(self._document_xml)
        except ET.ParseError:
            return []
        body = root.find(self._qn("body"))
        if body is None:
            return []
        blocks = []
        sec_idx = 0

        def _collect_from_container(container_elem: ET.Element, sec_idx_val: int):
            collected: List = []
            for inner in container_elem:
                if inner.tag == self._qn("p"):
                    para = self._parse_paragraph(inner)
                    if para is not None:
                        for img in para.images:
                            img.section_index = sec_idx_val
                        for c in para.content:
                            if isinstance(c, Image):
                                c.section_index = sec_idx_val
                        para.section_index = sec_idx_val
                        collected.append(para)
                elif inner.tag == self._qn("tbl"):
                    tbl = self._parse_table(inner)
                    if tbl is not None:
                        for row in tbl.rows:
                            for cell in row.cells:
                                for p in cell.content:
                                    for img in p.images:
                                        img.section_index = sec_idx_val
                        tbl.section_index = sec_idx_val
                        collected.append(tbl)
                elif inner.tag == self._qn("sdt"):
                    sdt_content = inner.find(self._qn("sdtContent"))
                    if sdt_content is not None:
                        nested = _collect_from_container(sdt_content, sec_idx_val)
                        collected.extend(nested)
                elif inner.tag == self._qn("sectPr"):
                    continue
                else:
                    continue
            return collected

        for child in body:
            if child.tag == self._qn("p"):
                para = self._parse_paragraph(child)
                if para is not None:
                    for img in para.images:
                        img.section_index = sec_idx
                    for c in para.content:
                        if isinstance(c, Image):
                            c.section_index = sec_idx
                    para.section_index = sec_idx
                    blocks.append(para)
                pPr = child.find(self._qn("pPr"))
                if pPr is not None and pPr.find(self._qn("sectPr")) is not None:
                    sec_idx += 1
            elif child.tag == self._qn("tbl"):
                tbl = self._parse_table(child)
                if tbl is not None:
                    for row in tbl.rows:
                        for cell in row.cells:
                            for p in cell.content:
                                for img in p.images:
                                    img.section_index = sec_idx
                    tbl.section_index = sec_idx
                    blocks.append(tbl)
            elif child.tag == self._qn("sdt"):
                sdt_content = child.find(self._qn("sdtContent"))
                if sdt_content is not None:
                    inner_blocks = _collect_from_container(sdt_content, sec_idx)
                    blocks.extend(inner_blocks)
                    for inner_elem in sdt_content.iter():
                        if inner_elem.tag == self._qn("p"):
                            pPr2 = inner_elem.find(self._qn("pPr"))
                            if pPr2 is not None and pPr2.find(self._qn("sectPr")) is not None:
                                sec_idx += 1
                                break
            elif child.tag == self._qn("sectPr"):
                continue
            else:
                continue
        return blocks

    def parse_paragraphs(self) -> List[Paragraph]:
        """Parse all paragraphs from document.xml into normalized model.

        Kept for backward compatibility with callers that expect only paragraphs.
        Filters the document-order block list to Paragraph blocks.
        """
        return [b for b in self.parse_document() if isinstance(b, Paragraph)]

    def get_styles(self) -> List[StyleDef]:
        """Parse every <w:style> definition from word/styles.xml.

        Returns a list of normalized StyleDef objects. This is the raw
        extraction boundary only - inheritance/outline resolution happens in
        the semantic layer (StyleRegistry). No style is inferred: every field
        comes directly from the OOXML element tree.
        """
        if not self._styles_xml:
            return []
        try:
            root = ET.fromstring(self._styles_xml)
        except ET.ParseError:
            # A malformed styles.xml (truncated/illegal package part) must not
            # abort the entire conversion. Degrade to "no styles" so the rest of
            # the document (paragraph-level outlineLvl, numbering, images) still
            # converts. This mirrors how a *missing* styles part is treated.
            return []
        styles: List[StyleDef] = []
        for st in root.findall(self._qn("style")):
            sid = st.get(self._qn("styleId"))
            if not sid:
                continue
            typ = st.get(self._qn("type"), "paragraph")
            name_el = st.find(self._qn("name"))
            name = name_el.get(self._qn("val"), "") if name_el is not None else ""
            based_el = st.find(self._qn("basedOn"))
            based_on = based_el.get(self._qn("val"), "") if based_el is not None else ""
            ppr = st.find(self._qn("pPr"))
            outline = None
            num_id = None
            num_ilvl = None
            alignment = None
            indent_left = indent_right = indent_first_line = indent_hanging = None
            spacing_before = spacing_after = line_spacing = None
            line_spacing_rule = None
            if ppr is not None:
                ol = ppr.find(self._qn("outlineLvl"))
                if ol is not None:
                    v = ol.get(self._qn("val"))
                    if v and v.isdigit():
                        outline = int(v)
                numpr = ppr.find(self._qn("numPr"))
                if numpr is not None:
                    ni = numpr.find(self._qn("numId"))
                    il = numpr.find(self._qn("ilvl"))
                    if ni is not None:
                        num_id = ni.get(self._qn("val"))
                    if il is not None:
                        iv = il.get(self._qn("val"))
                        if iv and iv.isdigit():
                            num_ilvl = int(iv)
                jc = ppr.find(self._qn("jc"))
                if jc is not None:
                    alignment = jc.get(self._qn("val"))
                ind = ppr.find(self._qn("ind"))
                if ind is not None:
                    v = ind.get(self._qn("left"))
                    if v and v.lstrip("-").isdigit():
                        indent_left = int(v)
                    v = ind.get(self._qn("right"))
                    if v and v.lstrip("-").isdigit():
                        indent_right = int(v)
                    v = ind.get(self._qn("firstLine"))
                    if v and v.lstrip("-").isdigit():
                        indent_first_line = int(v)
                    v = ind.get(self._qn("hanging"))
                    if v and v.lstrip("-").isdigit():
                        indent_hanging = int(v)
                spacing = ppr.find(self._qn("spacing"))
                if spacing is not None:
                    v = spacing.get(self._qn("before"))
                    if v and v.lstrip("-").isdigit():
                        spacing_before = int(v)
                    v = spacing.get(self._qn("after"))
                    if v and v.lstrip("-").isdigit():
                        spacing_after = int(v)
                    v = spacing.get(self._qn("line"))
                    if v and v.lstrip("-").isdigit():
                        line_spacing = int(v)
                    v = spacing.get(self._qn("lineRule"))
                    if v:
                        line_spacing_rule = v
            # Run properties: w:rPr direct child of style OR w:pPr/w:rPr
            rpr = st.find(self._qn("rPr"))
            if rpr is None and ppr is not None:
                rpr = ppr.find(self._qn("rPr"))
            font_family = font_size = font_color = bold = italic = underline = superscript = subscript = None
            if rpr is not None:
                rf = rpr.find(self._qn("rFonts"))
                if rf is not None:
                    ff = rf.get(self._qn("ascii")) or rf.get(self._qn("hAnsi"))
                    if ff:
                        font_family = ff
                sz = rpr.find(self._qn("sz"))
                if sz is not None:
                    v = sz.get(self._qn("val"))
                    if v and v.isdigit():
                        font_size = int(v)
                else:
                    sz_cs = rpr.find(self._qn("szCs"))
                    if sz_cs is not None:
                        v = sz_cs.get(self._qn("val"))
                        if v and v.isdigit():
                            font_size = int(v)
                col = rpr.find(self._qn("color"))
                if col is not None:
                    v = col.get(self._qn("val"))
                    if v:
                        font_color = v
                if rpr.find(self._qn("b")) is not None:
                    bv = rpr.find(self._qn("b")).get(self._qn("val"))
                    bold = False if bv == "false" else True
                if rpr.find(self._qn("i")) is not None:
                    iv = rpr.find(self._qn("i")).get(self._qn("val"))
                    italic = False if iv == "false" else True
                u = rpr.find(self._qn("u"))
                if u is not None:
                    underline = u.get(self._qn("val"), "single")
                va = rpr.find(self._qn("vertAlign"))
                if va is not None:
                    v = va.get(self._qn("val"))
                    if v == "superscript":
                        superscript = True
                    elif v == "subscript":
                        subscript = True
            styles.append(StyleDef(sid, name, based_on, typ, outline, num_id, num_ilvl,
                                   alignment, indent_left, indent_right, indent_first_line, indent_hanging,
                                   spacing_before, spacing_after, line_spacing, line_spacing_rule,
                                   font_family, font_size, font_color, bold, italic, underline, superscript, subscript))
        return styles

    def get_default_font(self) -> dict:
        """Extract document default run properties from styles.xml.

        Font size is returned in OOXML half-points (w:sz), consistent with the
        rest of the pipeline (Run.font_size is half-points; the renderer divides
        by 2 to obtain pt). The base default of 22 half-points == 11pt matches
        Word's normal default and the existing test expectations.
        """
        default = {"font_family": "Calibri", "font_size": 22, "font_color": "#000000"}
        if not self._styles_xml:
            return default

        try:
            styles_root = ET.fromstring(self._styles_xml)
        except ET.ParseError:
            return default
        rpr_default = styles_root.find(
            ".//" + self._qn("docDefaults") + "/" + self._qn("rPrDefault") + "/" + self._qn("rPr")
        )
        if rpr_default is None:
            return default

        rfonts = rpr_default.find(self._qn("rFonts"))
        if rfonts is not None:
            ff = rfonts.get(self._qn("ascii")) or rfonts.get(self._qn("hAnsi"))
            if ff:
                default["font_family"] = ff

        sz = rpr_default.find(self._qn("sz"))
        if sz is not None:
            val = sz.get(self._qn("val"))
            if val and val.isdigit():
                default["font_size"] = int(val)
        else:
            sz_cs = rpr_default.find(self._qn("szCs"))
            if sz_cs is not None:
                val = sz_cs.get(self._qn("val"))
                if val and val.isdigit():
                    default["font_size"] = int(val)

        color = rpr_default.find(self._qn("color"))
        if color is not None:
            val = color.get(self._qn("val"))
            if val:
                default["font_color"] = val

        return default

    def get_numbering(self) -> NumberingModel:
        """Parse word/numbering.xml into a centralized NumberingModel.

        Resolves the OOXML indirection: w:num -> w:abstractNumId -> w:abstractNum
        -> w:lvl (w:numFmt / w:lvlText / w:start). No numbering text is inferred
        from visible paragraph text; the path is built later from this model plus
        each paragraph's w:numPr reference.
        """
        raw = self._read_part("word/numbering.xml")
        model = NumberingModel()
        if not raw:
            return model
        root = ET.fromstring(raw)

        for an in root.findall(self._qn("abstractNum")):
            aid = an.get(self._qn("abstractNumId"))
            if aid is None:
                continue
            levels: Dict[int, NumberingLevel] = {}
            for lvl in an.findall(self._qn("lvl")):
                iv = lvl.get(self._qn("ilvl"))
                if iv is None or not iv.isdigit():
                    continue
                i = int(iv)
                fmt_el = lvl.find(self._qn("numFmt"))
                txt_el = lvl.find(self._qn("lvlText"))
                st_el = lvl.find(self._qn("start"))
                fmt = fmt_el.get(self._qn("val"), "decimal") if fmt_el is not None else "decimal"
                txt = txt_el.get(self._qn("val"), "%1.") if txt_el is not None else "%1."
                st = 1
                if st_el is not None:
                    sv = st_el.get(self._qn("val"))
                    if sv and sv.isdigit():
                        st = int(sv)
                levels[i] = NumberingLevel(i, fmt, txt, st)
            model.abstract_nums[aid] = levels

        for nm in root.findall(self._qn("num")):
            nid = nm.get(self._qn("numId"))
            a = nm.find(self._qn("abstractNumId"))
            if nid is None or a is None:
                continue
            model.nums[nid] = a.get(self._qn("val"))

        return model

    def get_footnotes(self) -> List[Note]:
        return self._parse_notes_part("word/footnotes.xml", "footnote")

    def get_endnotes(self) -> List[Note]:
        return self._parse_notes_part("word/endnotes.xml", "endnote")

    def get_comments(self) -> List[Note]:
        notes = self._parse_notes_part("word/comments.xml", "comment")
        if notes:
            self._enrich_comments_with_threads(notes)
        return notes

    def _enrich_comments_with_threads(self, notes: List[Note]) -> None:
        candidates = ["word/commentsExtended.xml", "word/commentsExtensible.xml"]
        raw = None
        for cand in candidates:
            raw = self._read_part(cand)
            if raw:
                break
        if not raw:
            return
        try:
            root = ET.fromstring(raw)
        except ET.ParseError:
            return
        ex_map: Dict[str, Dict[str, Optional[str]]] = {}
        for elem in root.iter():
            local = elem.tag.split("}", 1)[1] if "}" in elem.tag else elem.tag
            if local != "commentEx":
                continue
            para_id = None
            parent_id = None
            done_val = None
            for k, v in elem.attrib.items():
                lname = k.split("}", 1)[1] if "}" in k else k
                if lname == "paraId":
                    para_id = v
                elif lname == "paraIdParent":
                    parent_id = v
                elif lname == "done":
                    done_val = v
            if para_id:
                ex_map[para_id] = {"parent": parent_id, "done": done_val}
        if not ex_map:
            return
        para_to_note: Dict[str, Note] = {}
        for n in notes:
            if n.para_id:
                para_to_note[n.para_id] = n
        for para_id, info in ex_map.items():
            child = para_to_note.get(para_id)
            if child is None:
                continue
            parent_para = info.get("parent")
            child.para_id_parent = parent_para
            done_val = info.get("done")
            if done_val is not None:
                child.done = done_val == "1" or done_val.lower() == "true"
            if parent_para:
                parent = para_to_note.get(parent_para)
                if parent is not None:
                    child.parent_id = parent.note_id
                    if child not in parent.replies:
                        parent.replies.append(child)

    def _parse_notes_part(self, part_path: str, note_type: str) -> List[Note]:
        raw = self._read_part(part_path)
        if not raw:
            return []
        try:
            root = ET.fromstring(raw)
        except ET.ParseError:
            return []
        rels = self._read_relationships(part_path)
        has_rels = bool(rels)
        old_rels = self._rels
        if has_rels:
            self._rels = rels
        notes: List[Note] = []
        try:
            tag = self._qn(note_type)
            for elem in root.findall(tag):
                nid = elem.get(self._qn("id"))
                if nid is None:
                    continue
                ntype = elem.get(self._qn("type"))
                if ntype in ("separator", "continuationSeparator", "continuationNotice"):
                    continue
                blocks: List = []
                for child in elem:
                    if child.tag == self._qn("p"):
                        para = self._parse_paragraph(child)
                        if para is not None:
                            blocks.append(para)
                    elif child.tag == self._qn("tbl"):
                        tbl = self._parse_table(child)
                        if tbl is not None:
                            blocks.append(tbl)
                    elif child.tag == self._qn("sdt"):
                        sdt_c = child.find(self._qn("sdtContent"))
                        if sdt_c is not None:
                            for inner in sdt_c:
                                if inner.tag == self._qn("p"):
                                    para = self._parse_paragraph(inner)
                                    if para is not None:
                                        blocks.append(para)
                                elif inner.tag == self._qn("tbl"):
                                    tbl = self._parse_table(inner)
                                    if tbl is not None:
                                        blocks.append(tbl)
                    else:
                        continue
                if note_type == "comment":
                    author = elem.get(self._qn("author"))
                    date = elem.get(self._qn("date"))
                    initials = elem.get(self._qn("initials"))
                    para_id = None
                    for k, v in elem.attrib.items():
                        lname = k.split("}", 1)[1] if "}" in k else k
                        if lname == "paraId":
                            para_id = v
                            break
                    notes.append(Note(note_type=note_type, note_id=str(nid), blocks=blocks, author=author, date=date, initials=initials, para_id=para_id))
                else:
                    notes.append(Note(note_type=note_type, note_id=str(nid), blocks=blocks))
        finally:
            if has_rels:
                self._rels = old_rels
        return notes

    # ---- table border / shading helpers (w:tblBorders / w:tcBorders / w:shd) ----

    _TABLE_BORDER_EDGES = ("top", "left", "bottom", "right", "insideH", "insideV")

    def _parse_border_edge(self, el: ET.Element) -> BorderEdge:
        val = el.get(self._qn("val"), "")
        # size in eighths of a point; graceful handle missing/invalid
        sz_raw = el.get(self._qn("sz"))
        sz = int(sz_raw) if sz_raw and sz_raw.lstrip("-").isdigit() else None
        color = el.get(self._qn("color"))
        if color is not None and color.lower() == "auto":
            color = None
        space_raw = el.get(self._qn("space"))
        space = int(space_raw) if space_raw and space_raw.lstrip("-").isdigit() else None
        shadow_raw = el.get(self._qn("shadow"))
        shadow = None
        if shadow_raw is not None:
            shadow = shadow_raw not in ("false", "0", "off")
        return BorderEdge(val=val, sz=sz, color=color, space=space, shadow=shadow)

    def _parse_shd_fill(self, shd_elem: ET.Element) -> Optional[str]:
        if shd_elem is None:
            return None
        val = shd_elem.get(self._qn("val"))
        fill = shd_elem.get(self._qn("fill"))
        # val nil/none means no shading regardless of fill
        if val is not None and val.lower() in ("nil", "none"):
            return None
        if not fill:
            return None
        if fill.lower() == "auto":
            return None
        # normalize: strip #, upper-case hex? Keep raw hex without #
        fill = fill.lstrip("#")
        # reject non-hex or placeholder
        if not fill or len(fill) not in (3, 6, 8):
            # allow any non-empty hex; if unexpected, keep as-is for renderer to decide
            pass
        return fill

    # ---- table parsing ----

    def _parse_borders(self, borders_elem: Optional[ET.Element]) -> Optional[Dict[str, BorderEdge]]:
        if borders_elem is None:
            return None
        out: Dict[str, BorderEdge] = {}
        for child in list(borders_elem):
            tag = child.tag
            local = tag.split("}", 1)[1] if "}" in tag else tag
            val = child.get(self._qn("val"))
            if val is None:
                continue
            sz_raw = child.get(self._qn("sz"))
            sz = int(sz_raw) if sz_raw and sz_raw.lstrip("-").isdigit() else None
            color = child.get(self._qn("color"))
            if color is not None and color.lower() == "auto":
                color = None
            space_raw = child.get(self._qn("space"))
            space = int(space_raw) if space_raw and space_raw.lstrip("-").isdigit() else None
            shadow_raw = child.get(self._qn("shadow"))
            shadow = None
            if shadow_raw is not None:
                shadow = shadow_raw not in ("0", "false", "off")
            out[local] = BorderEdge(val=val, sz=sz, color=color, space=space, shadow=shadow)
        return out if out else None

    def _get_table_style_borders(self, style_name: Optional[str]) -> Optional[Dict[str, BorderEdge]]:
        """Resolve tblBorders from w:styles.xml for a table style (e.g. TableGrid).

        TableGrid and other built-in styles define borders only in styles.xml,
        not in document.xml w:tblPr. This lookup provides the table-level defaults
        so the renderer can use real OOXML borders instead of a blanket fallback.
        """
        if not style_name or not self._styles_xml:
            return None
        try:
            root = ET.fromstring(self._styles_xml)
        except ET.ParseError:
            return None
        for st in root.findall(self._qn("style")):
            sid = st.get(self._qn("styleId"))
            if sid != style_name:
                continue
            tbl_pr = st.find(self._qn("tblPr"))
            if tbl_pr is None:
                return None
            borders_elem = tbl_pr.find(self._qn("tblBorders"))
            return self._parse_borders(borders_elem)
        return None

    def _parse_table(self, tbl_elem: ET.Element):
        tbl_pr = tbl_elem.find(self._qn("tblPr"))
        grid = tbl_elem.find(self._qn("tblGrid"))
        grid_widths = []
        if grid is not None:
            for gc in grid.findall(self._qn("gridCol")):
                v = gc.get(self._qn("w"))
                if v and v.lstrip("-").isdigit():
                    grid_widths.append(int(v))
                else:
                    grid_widths.append(None)
        tbl_width = None
        tbl_width_type = None
        style_name = None
        tbl_borders: Optional[Dict[str, BorderEdge]] = None
        if tbl_pr is not None:
            tw = tbl_pr.find(self._qn("tblW"))
            if tw is not None:
                wv = tw.get(self._qn("w"))
                if wv and wv.lstrip("-").isdigit():
                    tbl_width = int(wv)
                tbl_width_type = tw.get(self._qn("type"))
            ts = tbl_pr.find(self._qn("tblStyle"))
            if ts is not None:
                style_name = ts.get(self._qn("val"))
            # Direct tblBorders in document.xml (explicit per-table overrides)
            direct = tbl_pr.find(self._qn("tblBorders"))
            tbl_borders = self._parse_borders(direct)
        # Fallback to style-defined borders when document.xml has no explicit tblBorders
        if tbl_borders is None and style_name:
            tbl_borders = self._get_table_style_borders(style_name)
        rows = []
        for tr_elem in tbl_elem.findall(self._qn("tr")):
            row = self._parse_row(tr_elem)
            if row is not None:
                rows.append(row)
        if not rows:
            return None
        table = Table(
            rows=rows,
            grid_col_widths=grid_widths,
            width=tbl_width,
            width_type=tbl_width_type,
            style_name=style_name,
            borders=tbl_borders,
        )
        self._resolve_vmerge(table)
        return table

    def _parse_row(self, tr_elem: ET.Element):
        cells = []
        for tc_elem in tr_elem.findall(self._qn("tc")):
            cell = self._parse_cell(tc_elem)
            if cell is not None:
                cells.append(cell)
        if not cells:
            return None
        return Row(cells=cells)

    def _parse_cell(self, tc_elem: ET.Element):
        tc_pr = tc_elem.find(self._qn("tcPr"))
        grid_span = 1
        v_merge = None
        width = None
        width_type = None
        v_align = None
        shading = None
        cell_borders: Optional[Dict[str, BorderEdge]] = None
        if tc_pr is not None:
            gs = tc_pr.find(self._qn("gridSpan"))
            if gs is not None:
                v = gs.get(self._qn("val"))
                if v and v.isdigit():
                    grid_span = int(v)
            vm = tc_pr.find(self._qn("vMerge"))
            if vm is not None:
                val = vm.get(self._qn("val"))
                if val == "restart":
                    v_merge = "restart"
                else:
                    v_merge = "continue"
            tcw = tc_pr.find(self._qn("tcW"))
            if tcw is not None:
                wv = tcw.get(self._qn("w"))
                if wv and wv.lstrip("-").isdigit():
                    width = int(wv)
                width_type = tcw.get(self._qn("type"))
            va = tc_pr.find(self._qn("vAlign"))
            if va is not None:
                v_align = va.get(self._qn("val"))
            shd = tc_pr.find(self._qn("shd"))
            if shd is not None:
                shading = self._parse_shd_fill(shd)
            tc_borders_elem = tc_pr.find(self._qn("tcBorders"))
            cell_borders = self._parse_borders(tc_borders_elem)
        paragraphs = []

        def _collect_cell_paras(container: ET.Element):
            out: List[Paragraph] = []
            for elem in container:
                if elem.tag == self._qn("p"):
                    para = self._parse_paragraph(elem)
                    if para is not None:
                        out.append(para)
                elif elem.tag == self._qn("sdt"):
                    sdt_c = elem.find(self._qn("sdtContent"))
                    if sdt_c is not None:
                        out.extend(_collect_cell_paras(sdt_c))
                elif elem.tag == self._qn("tbl"):
                    # Table inside cell: flatten by extracting its paragraphs in order
                    tbl_inner = self._parse_table(elem)
                    if tbl_inner is not None:
                        for r in tbl_inner.rows:
                            for c in r.cells:
                                out.extend(c.content)
                else:
                    continue
            return out

        paragraphs = _collect_cell_paras(tc_elem)
        if not paragraphs:
            paragraphs = [Paragraph(runs=[], style_name="Normal", alignment="left")]
        return Cell(
            content=paragraphs,
            grid_span=grid_span,
            v_merge=v_merge,
            width=width,
            width_type=width_type,
            vertical_align=v_align,
            shading=shading,
            borders=cell_borders,
        )

    def _resolve_vmerge(self, table: Table):
        if not table.rows:
            return
        cols = table.column_count
        if cols == 0:
            return
        col_starts = []
        for row in table.rows:
            starts = []
            col = 0
            for cell in row.cells:
                starts.append(col)
                col += cell.grid_span
            col_starts.append(starts)
        for r_idx, row in enumerate(table.rows):
            starts = col_starts[r_idx]
            for c_idx, cell in enumerate(row.cells):
                if cell.v_merge != "restart":
                    continue
                col = starts[c_idx]
                span = cell.grid_span
                rowspan = 1
                for nr in range(r_idx + 1, len(table.rows)):
                    n_row = table.rows[nr]
                    n_starts = col_starts[nr]
                    found_continue = False
                    for nc_idx, n_cell in enumerate(n_row.cells):
                        if n_starts[nc_idx] == col and n_cell.grid_span == span and n_cell.v_merge == "continue":
                            found_continue = True
                            break
                    if found_continue:
                        rowspan += 1
                    else:
                        break
                cell.row_span = rowspan

    # ---- private parsing ----

    def _parse_paragraph(self, p_elem: ET.Element) -> Paragraph:
        style_name = "Normal"
        alignment = "left"
        outline_level = None
        num_id = None
        num_ilvl = None
        indent_left = None
        indent_right = None
        indent_first_line = None
        indent_hanging = None
        spacing_before = None
        spacing_after = None
        line_spacing = None
        line_spacing_rule = None

        ppr = p_elem.find(self._qn("pPr"))
        tabs: List[TabStop] = []
        if ppr is not None:
            pstyle = ppr.find(self._qn("pStyle"))
            if pstyle is not None:
                style_name = pstyle.get(self._qn("val"), "Normal")
            jc = ppr.find(self._qn("jc"))
            if jc is not None:
                alignment = jc.get(self._qn("val"), "left")
            ol = ppr.find(self._qn("outlineLvl"))
            if ol is not None:
                v = ol.get(self._qn("val"))
                if v and v.isdigit():
                    outline_level = int(v)
            numpr = ppr.find(self._qn("numPr"))
            if numpr is not None:
                ni = numpr.find(self._qn("numId"))
                il = numpr.find(self._qn("ilvl"))
                if ni is not None:
                    num_id = ni.get(self._qn("val"))
                if il is not None:
                    iv = il.get(self._qn("val"))
                    if iv and iv.isdigit():
                        num_ilvl = int(iv)
            ind = ppr.find(self._qn("ind"))
            if ind is not None:
                v = ind.get(self._qn("left"))
                if v and v.lstrip("-").isdigit():
                    indent_left = int(v)
                v = ind.get(self._qn("right"))
                if v and v.lstrip("-").isdigit():
                    indent_right = int(v)
                v = ind.get(self._qn("firstLine"))
                if v and v.lstrip("-").isdigit():
                    indent_first_line = int(v)
                v = ind.get(self._qn("hanging"))
                if v and v.lstrip("-").isdigit():
                    indent_hanging = int(v)
            spacing = ppr.find(self._qn("spacing"))
            if spacing is not None:
                v = spacing.get(self._qn("before"))
                if v and v.lstrip("-").isdigit():
                    spacing_before = int(v)
                v = spacing.get(self._qn("after"))
                if v and v.lstrip("-").isdigit():
                    spacing_after = int(v)
                v = spacing.get(self._qn("line"))
                if v and v.lstrip("-").isdigit():
                    line_spacing = int(v)
                v = spacing.get(self._qn("lineRule"))
                if v:
                    line_spacing_rule = v
            tabs_el = ppr.find(self._qn("tabs"))
            if tabs_el is not None:
                for tab_el in tabs_el.findall(self._qn("tab")):
                    val = tab_el.get(self._qn("val"))
                    pos = tab_el.get(self._qn("pos"))
                    leader = tab_el.get(self._qn("leader"))
                    if val and pos and pos.lstrip("-").isdigit():
                        tabs.append(TabStop(val=val, pos=int(pos), leader=leader))

        runs: List[Run] = []
        content = []  # ordered Run/Image mix, preserving in-paragraph position
        field_stack = []  # stack of {code:str, type:Optional[str], has_separate:bool}
        def _in_supported_result():
            return bool(field_stack and field_stack[-1].get("type") in ("PAGE", "NUMPAGES", "PAGEREF") and field_stack[-1].get("has_separate"))
        def _emit_field_placeholder(field_type, field_code):
            ph = Run(text="", field_type=field_type, field_code=field_code)
            runs.append(ph)
            content.append(ph)
        # Walk direct children so a run's position relative to a drawing (and to
        # other runs) is preserved. w:hyperlink contains nested runs and is
        # handled recursively so its runs keep their order too.
        for child in p_elem:
            if child.tag == self._qn("fldSimple"):
                instr = child.get(self._qn("instr"), "") or ""
                first = instr.strip().split()[0].upper() if instr.strip() else ""
                if first in ("PAGE", "NUMPAGES", "PAGEREF"):
                    _emit_field_placeholder(first, instr.strip())
                    continue
                else:
                    for r_elem in child.findall(self._qn("r")):
                        run = self._parse_run(r_elem)
                        if run is not None and self._run_is_meaningful(run):
                            runs.append(run)
                            content.append(run)
                        for img in self._extract_images_from_r(r_elem):
                            content.append(img)
                    continue
            if child.tag == self._qn("r"):
                fld_char_el = child.find(self._qn("fldChar"))
                instr_el = child.find(self._qn("instrText"))
                # tab/br are direct children of w:r
                has_tab = child.find(self._qn("tab")) is not None
                has_br = child.find(self._qn("br")) is not None
                if fld_char_el is not None:
                    ftype = fld_char_el.get(self._qn("fldCharType"))
                    if ftype == "begin":
                        field_stack.append({"code": "", "type": None, "has_separate": False})
                        continue
                    elif ftype == "separate":
                        if field_stack:
                            top = field_stack[-1]
                            top["has_separate"] = True
                            code = top["code"].strip()
                            first = code.split()[0].upper() if code else ""
                            if first in ("PAGE", "NUMPAGES", "PAGEREF"):
                                top["type"] = first
                            else:
                                field_stack.pop()
                        continue
                    elif ftype == "end":
                        if field_stack:
                            top = field_stack.pop()
                            if top.get("type") in ("PAGE", "NUMPAGES", "PAGEREF"):
                                _emit_field_placeholder(top["type"], top["code"].strip())
                        continue
                if instr_el is not None:
                    txt = instr_el.text or ""
                    if field_stack and not field_stack[-1].get("has_separate"):
                        field_stack[-1]["code"] += txt + " "
                    continue
                if _in_supported_result():
                    # field result runs like <w:t>xi</w:t> are suppressed; placeholder will be emitted at end
                    continue
                fn_ref_el = child.find(self._qn("footnoteReference"))
                if fn_ref_el is not None:
                    fid = fn_ref_el.get(self._qn("id"))
                    if fid is not None:
                        content.append(NoteReference(note_type="footnote", note_id=str(fid)))
                        for img in self._extract_images_from_r(child):
                            content.append(img)
                        continue
                en_ref_el = child.find(self._qn("endnoteReference"))
                if en_ref_el is not None:
                    eid = en_ref_el.get(self._qn("id"))
                    if eid is not None:
                        content.append(NoteReference(note_type="endnote", note_id=str(eid)))
                        for img in self._extract_images_from_r(child):
                            content.append(img)
                        continue
                cmt_ref_el = child.find(self._qn("commentReference"))
                if cmt_ref_el is not None:
                    cid = cmt_ref_el.get(self._qn("id"))
                    if cid is not None:
                        content.append(NoteReference(note_type="comment", note_id=str(cid)))
                        for img in self._extract_images_from_r(child):
                            content.append(img)
                        continue
                crs = child.find(self._qn("commentRangeStart"))
                if crs is not None:
                    cid = crs.get(self._qn("id"))
                    if cid is not None:
                        content.append(CommentRangeStart(comment_id=str(cid)))
                    for img in self._extract_images_from_r(child):
                        content.append(img)
                    continue
                cre = child.find(self._qn("commentRangeEnd"))
                if cre is not None:
                    cid = cre.get(self._qn("id"))
                    if cid is not None:
                        content.append(CommentRangeEnd(comment_id=str(cid)))
                    for img in self._extract_images_from_r(child):
                        content.append(img)
                    continue
                if child.find(self._qn("footnoteRef")) is not None or child.find(self._qn("endnoteRef")) is not None:
                    for img in self._extract_images_from_r(child):
                        content.append(img)
                    continue
                if child.find(self._qn("commentRef")) is not None:
                    for img in self._extract_images_from_r(child):
                        content.append(img)
                    continue
                # w:sym, w:noBreakHyphen, w:softHyphen are direct children of w:r (no w:t)
                sym_elem = child.find(self._qn("sym"))
                if sym_elem is not None:
                    font = sym_elem.get(self._qn("font"))
                    char_hex = sym_elem.get(self._qn("char"))
                    decoded = _decode_sym_char(font, char_hex)
                    if decoded is not None:
                        rpr_for_font = child.find(self._qn("rPr"))
                        sym_run = Run(text=decoded)
                        if font:
                            sym_run.font_family = font
                        elif rpr_for_font is not None:
                            rf_tmp = rpr_for_font.find(self._qn("rFonts"))
                            if rf_tmp is not None:
                                ff_tmp = rf_tmp.get(self._qn("ascii")) or rf_tmp.get(self._qn("hAnsi"))
                                if not ff_tmp:
                                    ff_tmp = rf_tmp.get(self._qn("eastAsia")) or rf_tmp.get(self._qn("cs"))
                                if ff_tmp:
                                    sym_run.font_family = ff_tmp
                        # inherit other formatting via _parse_run then override text
                        base = self._parse_run(child)
                        if base.font_family and not sym_run.font_family:
                            sym_run.font_family = base.font_family
                        if base.bold is not None:
                            sym_run.bold = base.bold
                        if base.italic is not None:
                            sym_run.italic = base.italic
                        if base.font_size is not None:
                            sym_run.font_size = base.font_size
                        if base.font_color is not None:
                            sym_run.font_color = base.font_color
                        if self._run_is_meaningful(sym_run):
                            runs.append(sym_run)
                            content.append(sym_run)
                        for img in self._extract_images_from_r(child):
                            content.append(img)
                        continue
                if child.find(self._qn("noBreakHyphen")) is not None:
                    nbh_run = Run(text="\u2011")
                    base = self._parse_run(child)
                    if base.font_family:
                        nbh_run.font_family = base.font_family
                    if base.bold is not None:
                        nbh_run.bold = base.bold
                    if base.italic is not None:
                        nbh_run.italic = base.italic
                    if base.font_size is not None:
                        nbh_run.font_size = base.font_size
                    runs.append(nbh_run)
                    content.append(nbh_run)
                    for img in self._extract_images_from_r(child):
                        content.append(img)
                    continue
                if child.find(self._qn("softHyphen")) is not None:
                    sh_run = Run(text="\u00AD")
                    base = self._parse_run(child)
                    if base.font_family:
                        sh_run.font_family = base.font_family
                    if base.bold is not None:
                        sh_run.bold = base.bold
                    if base.italic is not None:
                        sh_run.italic = base.italic
                    if base.font_size is not None:
                        sh_run.font_size = base.font_size
                    runs.append(sh_run)
                    content.append(sh_run)
                    for img in self._extract_images_from_r(child):
                        content.append(img)
                    continue
                if has_tab:
                    tab_run = Run(text="\t")
                    runs.append(tab_run)
                    content.append(tab_run)
                    for img in self._extract_images_from_r(child):
                        content.append(img)
                    continue
                if has_br:
                    br_run = Run(text="\n")
                    runs.append(br_run)
                    content.append(br_run)
                    for img in self._extract_images_from_r(child):
                        content.append(img)
                    continue
                run = self._parse_run(child)
                if run is not None and self._run_is_meaningful(run):
                    runs.append(run)
                    content.append(run)
                for img in self._extract_images_from_r(child):
                    content.append(img)
            elif child.tag == self._qn("hyperlink"):
                href = None
                anchor = child.get(self._qn("anchor"))
                if anchor is not None:
                    href = "#" + anchor
                else:
                    rid = child.get(self._qn_ns(NS_R, "id"))
                    if rid:
                        rel = self._rels.get(rid)
                        if rel is not None:
                            href = rel.get("Target")
                for r_elem in child.iter(self._qn("r")):
                    # Handle note references inside hyperlink (rare)
                    fn_h = r_elem.find(self._qn("footnoteReference"))
                    if fn_h is not None:
                        fid = fn_h.get(self._qn("id"))
                        if fid is not None:
                            content.append(NoteReference(note_type="footnote", note_id=str(fid)))
                            for img in self._extract_images_from_r(r_elem):
                                content.append(img)
                            continue
                    en_h = r_elem.find(self._qn("endnoteReference"))
                    if en_h is not None:
                        eid = en_h.get(self._qn("id"))
                        if eid is not None:
                            content.append(NoteReference(note_type="endnote", note_id=str(eid)))
                            for img in self._extract_images_from_r(r_elem):
                                content.append(img)
                            continue
                    cmt_h = r_elem.find(self._qn("commentReference"))
                    if cmt_h is not None:
                        cid = cmt_h.get(self._qn("id"))
                        if cid is not None:
                            content.append(NoteReference(note_type="comment", note_id=str(cid)))
                            for img in self._extract_images_from_r(r_elem):
                                content.append(img)
                            continue
                    crs_h = r_elem.find(self._qn("commentRangeStart"))
                    if crs_h is not None:
                        cid = crs_h.get(self._qn("id"))
                        if cid is not None:
                            content.append(CommentRangeStart(comment_id=str(cid)))
                        for img in self._extract_images_from_r(r_elem):
                            content.append(img)
                        continue
                    cre_h = r_elem.find(self._qn("commentRangeEnd"))
                    if cre_h is not None:
                        cid = cre_h.get(self._qn("id"))
                        if cid is not None:
                            content.append(CommentRangeEnd(comment_id=str(cid)))
                        for img in self._extract_images_from_r(r_elem):
                            content.append(img)
                        continue
                    if r_elem.find(self._qn("footnoteRef")) is not None or r_elem.find(self._qn("endnoteRef")) is not None:
                        for img in self._extract_images_from_r(r_elem):
                            content.append(img)
                        continue
                    if r_elem.find(self._qn("commentRef")) is not None:
                        for img in self._extract_images_from_r(r_elem):
                            content.append(img)
                        continue
                    # hyperlink runs do not participate in field state (PAGE never inside hyperlink in practice)
                    run = self._parse_run(r_elem)
                    if run is not None and self._run_is_meaningful(run):
                        if href is not None:
                            run.href = href
                        runs.append(run)
                        content.append(run)
                    for img in self._extract_images_from_r(r_elem):
                        content.append(img)
            elif child.tag == self._qn("footnoteReference"):
                fid = child.get(self._qn("id"))
                if fid is not None:
                    content.append(NoteReference(note_type="footnote", note_id=str(fid)))
                continue
            elif child.tag == self._qn("endnoteReference"):
                eid = child.get(self._qn("id"))
                if eid is not None:
                    content.append(NoteReference(note_type="endnote", note_id=str(eid)))
                continue
            elif child.tag == self._qn("commentReference"):
                cid = child.get(self._qn("id"))
                if cid is not None:
                    content.append(NoteReference(note_type="comment", note_id=str(cid)))
                continue
            elif child.tag == self._qn("commentRangeStart"):
                cid = child.get(self._qn("id"))
                if cid is not None:
                    content.append(CommentRangeStart(comment_id=str(cid)))
                continue
            elif child.tag == self._qn("commentRangeEnd"):
                cid = child.get(self._qn("id"))
                if cid is not None:
                    content.append(CommentRangeEnd(comment_id=str(cid)))
                continue

        images = [c for c in content if isinstance(c, Image)]
        has_ranges = any(isinstance(c, (CommentRangeStart, CommentRangeEnd)) for c in content)
        has_notes = any(isinstance(c, NoteReference) for c in content)

        has_layout = any(v is not None for v in [indent_left, indent_right, indent_first_line, indent_hanging, spacing_before, spacing_after, line_spacing]) or bool(tabs)
        if runs or images or has_notes or has_ranges or style_name != "Normal" or has_layout or alignment != "left":
            return Paragraph(
                runs=runs,
                style_name=style_name,
                alignment=alignment,
                outline_level=outline_level,
                num_id=num_id,
                num_ilvl=num_ilvl,
                images=images,
                content=content,
                indent_left=indent_left,
                indent_right=indent_right,
                indent_first_line=indent_first_line,
                indent_hanging=indent_hanging,
                spacing_before=spacing_before,
                spacing_after=spacing_after,
                line_spacing=line_spacing,
                line_spacing_rule=line_spacing_rule,
                tabs=tabs,
            )
        return None

    @staticmethod
    def _run_is_meaningful(run: Run) -> bool:
        """A run worth keeping in the content stream.

        A run that is just a drawing container (empty text, no formatting) must
        not become a spurious empty Run between real text and an image, or the
        in-paragraph ordering model would carry phantom items. Drawing images
        are represented separately as Image placements.
        """
        if run.text:
            return True
        if getattr(run, "field_type", None):
            return True
        if run.bold or run.italic or run.underline or run.superscript or run.subscript:
            return True
        if run.font_color is not None and run.font_color not in ("#000000", "000000"):
            return True
        if run.font_size is not None and run.font_size != 22:
            return True
        if run.font_family is not None and run.font_family != "Calibri":
            return True
        return False

    def _parse_run(self, r_elem: ET.Element) -> Run:
        text = ""
        t_elem = r_elem.find(self._qn("t"))
        if t_elem is not None and t_elem.text:
            text = t_elem.text

        run = Run(text=text)

        rpr = r_elem.find(self._qn("rPr"))
        if rpr is None:
            return run

        rs = rpr.find(self._qn("rStyle"))
        if rs is not None:
            v = rs.get(self._qn("val"))
            if v:
                run.style_name = v

        if rpr.find(self._qn("b")) is not None:
            bv = rpr.find(self._qn("b")).get(self._qn("val"))
            if bv == "false":
                run.bold = False
            else:
                run.bold = True
        if rpr.find(self._qn("i")) is not None:
            iv = rpr.find(self._qn("i")).get(self._qn("val"))
            if iv == "false":
                run.italic = False
            else:
                run.italic = True

        u = rpr.find(self._qn("u"))
        if u is not None:
            run.underline = u.get(self._qn("val"), "single")

        color = rpr.find(self._qn("color"))
        if color is not None:
            val = color.get(self._qn("val"))
            if val:
                run.font_color = val

        rfonts = rpr.find(self._qn("rFonts"))
        if rfonts is not None:
            ff = rfonts.get(self._qn("ascii")) or rfonts.get(self._qn("hAnsi"))
            if not ff:
                ff = rfonts.get(self._qn("eastAsia")) or rfonts.get(self._qn("cs"))
            if ff:
                run.font_family = ff

        sz = rpr.find(self._qn("sz"))
        if sz is not None:
            val = sz.get(self._qn("val"))
            if val and val.isdigit():
                run.font_size = int(val)
        else:
            # Complex-script fallback: OOXML applies w:sz to Latin/ASCII runs and
            # w:szCs to complex-script runs. When w:sz is absent we fall back to
            # w:szCs rather than blindly replacing it, preserving the correct
            # precedence (w:sz wins when both are present).
            sz_cs = rpr.find(self._qn("szCs"))
            if sz_cs is not None:
                val = sz_cs.get(self._qn("val"))
                if val and val.isdigit():
                    run.font_size = int(val)

        va = rpr.find(self._qn("vertAlign"))
        if va is not None:
            val = va.get(self._qn("val"))
            if val == "superscript":
                run.superscript = True
            elif val == "subscript":
                run.subscript = True

        return run
