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
)
from core.units import emu_to_px, twip_to_emu, EMU_PER_PIXEL, EMU_PER_TWIP

# Standard OOXML namespaces used by drawing / relationship elements. The w:
# namespace is detected dynamically; these are fixed by the OOXML spec.
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NS_OPC = "http://schemas.openxmlformats.org/package/2006/relationships"

_SUPPORTED_RELATIVE_FROM = {
    "page", "margin", "column", "character", "paragraph",
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
        # Display extent (EMU) -> CSS px. Both inline and anchor carry wp:extent.
        # This is the authoritative conversion path (core.units.emu_to_px).
        extent = drawing_elem.find(self._qn_ns(NS_WP, "extent"))
        width = height = None
        if extent is not None:
            cx = extent.get("cx")
            cy = extent.get("cy")
            if cx and cx.lstrip("-").isdigit():
                width = emu_to_px(int(cx))
            if cy and cy.lstrip("-").isdigit():
                height = emu_to_px(int(cy))

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
        """Extract image placements from a w:r that contains a w:drawing."""
        images: List[Image] = []
        drawing = r_elem.find(self._qn("drawing"))
        if drawing is None:
            return images
        for kind in ("inline", "anchor"):
            for node in drawing.findall(self._qn_ns(NS_WP, kind)):
                img = self._parse_drawing(node, kind)
                if img is not None:
                    images.append(img)
        return images

    def get_image_assets(self) -> Dict[str, ImageAsset]:
        """Return the deduped image asset store (source_path -> ImageAsset)."""
        return self._assets

    def get_page_layout(self) -> PageLayout:
        """Return the parsed page/margin geometry (EMU)."""
        return self._page_layout

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
        pw = _twip(pg, "w")
        ph = _twip(pg, "h")
        if pw:
            kwargs["width_emu"] = pw
        if ph:
            kwargs["height_emu"] = ph
        ml = _twip(mar, "left")
        mr = _twip(mar, "right")
        mt = _twip(mar, "top")
        mb = _twip(mar, "bottom")
        if ml:
            kwargs["margin_left_emu"] = ml
        if mr:
            kwargs["margin_right_emu"] = mr
        if mt:
            kwargs["margin_top_emu"] = mt
        if mb:
            kwargs["margin_bottom_emu"] = mb
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
        pw = _twip(pg, "w")
        ph = _twip(pg, "h")
        if pw:
            kwargs["width_emu"] = pw
        if ph:
            kwargs["height_emu"] = ph
        ml = _twip(mar, "left")
        mr = _twip(mar, "right")
        mt = _twip(mar, "top")
        mb = _twip(mar, "bottom")
        if ml:
            kwargs["margin_left_emu"] = ml
        if mr:
            kwargs["margin_right_emu"] = mr
        if mt:
            kwargs["margin_top_emu"] = mt
        if mb:
            kwargs["margin_bottom_emu"] = mb
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
        for child in body:
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
        for p_elem in tc_elem.findall(self._qn("p")):
            para = self._parse_paragraph(p_elem)
            if para is not None:
                paragraphs.append(para)
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
                    # hyperlink runs do not participate in field state (PAGE never inside hyperlink in practice)
                    run = self._parse_run(r_elem)
                    if run is not None and self._run_is_meaningful(run):
                        if href is not None:
                            run.href = href
                        runs.append(run)
                        content.append(run)
                    for img in self._extract_images_from_r(r_elem):
                        content.append(img)

        images = [c for c in content if isinstance(c, Image)]

        has_layout = any(v is not None for v in [indent_left, indent_right, indent_first_line, indent_hanging, spacing_before, spacing_after, line_spacing]) or bool(tabs)
        # Keep paragraph if it carries content, non-default style, or layout
        if runs or images or style_name != "Normal" or has_layout or alignment != "left":
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
