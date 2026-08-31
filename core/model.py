"""Normalized document model for docx-to-html.

This is the single source of truth for the vertical slice:
REAL DOCX -> OOXML -> normalized model -> HTML -> browser.

A Run carries the formatting extracted from <w:r>.
A Paragraph carries its runs and the resolved style name from <w:pPr>,
plus the resolved heading semantics (level / id) and any direct outline
level taken from the paragraph's own <w:pPr/w:outlineLvl>.

Heading semantics are written ONCE (by the semantic classifier) and then
consumed by the hierarchy, TOC and renderer. The renderer never makes
semantic decisions.
"""

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Union, Any, Tuple

from core.units import emu_to_px


@dataclass
class Run:
    """A run of formatted text extracted from DOCX OOXML (w:r)."""
    text: str = ""
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[str] = None          # None = none, otherwise the w:u/@w:val (e.g. "single")
    font_family: Optional[str] = None
    font_size: Optional[int] = None          # half-points (OOXML w:sz); renderer divides by 2 for pt
    font_color: Optional[str] = None         # hex, e.g. "0000FF" or "#000000"
    superscript: Optional[bool] = None
    subscript: Optional[bool] = None
    href: Optional[str] = None   # hyperlink target (external URL or #anchor) if this run is inside w:hyperlink
    style_name: Optional[str] = None  # w:rStyle/@w:val if this run uses a character style
    # Field semantics: when this run represents a PAGE/NUMPAGES/PAGEREF field placeholder,
    # field_type is the normalized first token (PAGE/NUMPAGES/PAGEREF) and field_code is
    # the raw instrText (e.g. 'PAGE \\* MERGEFORMAT'). Normal text has None.
    field_type: Optional[str] = None
    field_code: Optional[str] = None


@dataclass
class TabStop:
    """Word tab stop from w:pPr/w:tabs/w:tab."""
    val: str  # left, center, right, decimal, clear, etc.
    pos: int  # twips (1/20 pt)
    leader: Optional[str] = None  # none, dot, hyphen, underscore, heavy, middleDot

@dataclass
class Paragraph:
    """A paragraph extracted from DOCX OOXML (w:p)."""
    runs: List[Run] = field(default_factory=list)
    style_name: str = "Normal"   # resolved from w:pPr/w:pStyle/@w:val  (this is the styleId)
    alignment: str = "left"      # resolved from w:pPr/w:jc/@w:val if present
    outline_level: Optional[int] = None   # direct w:pPr/w:outlineLvl/@w:val (0-8), if present
    heading_level: Optional[int] = None   # resolved semantic level 1-9 (None = not a heading)
    heading_id: Optional[str] = None      # single authoritative id assigned by the classifier
    # Paragraph layout (w:pPr/w:ind, w:spacing, w:jc already via alignment)
    indent_left: Optional[int] = None      # w:ind/@w:left twips
    indent_right: Optional[int] = None     # w:ind/@w:right twips
    indent_first_line: Optional[int] = None  # w:ind/@w:firstLine twips
    indent_hanging: Optional[int] = None   # w:ind/@w:hanging twips
    spacing_before: Optional[int] = None   # w:spacing/@w:before twips (1/20 pt)
    spacing_after: Optional[int] = None    # w:spacing/@w:after twips
    line_spacing: Optional[int] = None     # w:spacing/@w:line
    line_spacing_rule: Optional[str] = None  # w:spacing/@w:lineRule auto/exact/atLeast

    # Structured numbering metadata (validation signal only). Sourced from
    # w:numPr + word/numbering.xml. Never mutates heading_level.
    num_id: Optional[str] = None
    num_ilvl: Optional[int] = None
    numbering_format: Optional[str] = None
    numbering_text_pattern: Optional[str] = None
    numbering_path: Optional[List[int]] = None
    numbering_consistent: Optional[bool] = None
    # Per-level numFmt for each entry in numbering_path (ilvl 0..num_ilvl),
    # resolved from the existing NumberingModel by the semantic resolver. Used
    # ONLY to format the visible label from lvlText; never re-parses XML.
    numbering_level_formats: Optional[List[str]] = None

    # Image support (VISUAL FIDELITY phase).
    # `images` holds the Image PLACEMENTS located inside this paragraph. They are
    # also represented in document order inside `content` so the renderer can
    # preserve the position of an image relative to surrounding text. `images`
    # is a convenience view (subset of `content`) for queries/tests.
    images: List["Image"] = field(default_factory=list)
    # Ordered content: a mix of Run/Image/NoteReference objects, in the exact
    # document order they appeared inside the paragraph. The renderer iterates
    # this when present; when empty (e.g. test-built paragraphs), it falls back
    # to `runs`. This single field is the source of truth for in-paragraph ordering.
    content: List[Union["Run", "Image", "NoteReference", "CommentRangeStart", "CommentRangeEnd"]] = field(default_factory=list)
    # Word tab stops from w:pPr/w:tabs (preserved verbatim for renderer)
    tabs: List["TabStop"] = field(default_factory=list)
    # Stable block id assigned by the anchoring pass (core/anchoring.py) so
    # floating images can reference their nearest containing block. None until
    # the anchoring pass runs; never set by the OOXML parser.
    block_id: Optional[str] = None


@dataclass
class BorderEdge:
    """One OOXML border edge (w:top / w:left / w:bottom / w:right / w:insideH / w:insideV).

    Parsed verbatim from w:val / w:sz / w:color / w:space / w:shadow. Size is in
    eighths of a point (OOXML w:sz), color is the raw hex without '#', or
    'auto' when the document uses automatic color. val == 'nil' or 'none'
    means no border (renderer must not invent one).
    """
    val: str                              # e.g. single, double, dashed, dotted, nil, none, etc.
    sz: Optional[int] = None              # eighths of a point (int), None when absent
    color: Optional[str] = None           # hex like 'FF0000' or 'auto', None when absent
    space: Optional[int] = None           # w:space, rarely used
    shadow: Optional[bool] = None


@dataclass
class Cell:
    """A table cell extracted from w:tc.

    Content preserves ordered paragraphs inside the cell via the existing
    paragraph parser (run formatting, images, alignment reused). Span and
    merge metadata come directly from w:tcPr.
    """
    content: List["Paragraph"] = field(default_factory=list)
    grid_span: int = 1                  # w:gridSpan/@w:val (default 1)
    v_merge: Optional[str] = None       # None | "restart" | "continue"
    row_span: int = 1                   # computed rowspan (1 = no vertical merge)
    width: Optional[int] = None         # w:tcW/@w:w (raw dxa/twip integer)
    width_type: Optional[str] = None    # w:tcW/@w:type (dxa/auto/pct/nil)
    vertical_align: Optional[str] = None  # w:vAlign/@w:val (top/center/bottom)
    shading: Optional[str] = None       # w:shd/@w:fill hex (e.g. "FF0000"), None = no fill
    borders: Optional[Dict[str, BorderEdge]] = None  # per-edge borders: top/bottom/left/right/insideH/insideV


@dataclass
class Row:
    """A table row extracted from w:tr."""
    cells: List["Cell"] = field(default_factory=list)


@dataclass
class Table:
    """A table extracted from w:tbl, preserving document order.

    This is a structural block sibling to Paragraph in the document flow.
    Heading/semantic processing ignores Table blocks.
    """
    rows: List["Row"] = field(default_factory=list)
    grid_col_widths: List[Optional[int]] = field(default_factory=list)  # w:gridCol/@w:w
    width: Optional[int] = None          # w:tblW/@w:w
    width_type: Optional[str] = None     # w:tblW/@w:type
    style_name: Optional[str] = None     # w:tblStyle/@w:val
    block_id: Optional[str] = None       # assigned by anchoring-equivalent pass if needed
    borders: Optional[Dict[str, BorderEdge]] = None  # w:tblPr/w:tblBorders per-edge defaults

    @property
    def column_count(self) -> int:
        """Logical column count derived from grid or widest row."""
        if self.grid_col_widths:
            return len(self.grid_col_widths)
        if not self.rows:
            return 0
        return max((sum(c.grid_span for c in r.cells) for r in self.rows), default=0)


# Document block is Paragraph or Table in document order.
Block = Union["Paragraph", "Table"]


@dataclass
class HeaderFooter:
    """Normalized header or footer part.

    Preserves ordered blocks (Paragraph/Table) via the same pipeline as body
    content, including hyperlinks/images/tables formatting.
    """
    hf_type: str = "default"        # default | first | even
    kind: str = "header"            # header | footer
    blocks: List[Block] = field(default_factory=list)
    r_id: Optional[str] = None      # relationship id from sectPr
    target: Optional[str] = None    # part path e.g. word/header1.xml
    is_linked_to_previous: bool = False  # not used yet, placeholder for linkToPrevious


@dataclass
class Section:
    """Physical section with associated header/footer variants."""
    index: int = 0
    title_pg: bool = False          # w:titlePg present in this sectPr
    headers: Dict[str, HeaderFooter] = field(default_factory=dict)  # type -> HeaderFooter
    footers: Dict[str, HeaderFooter] = field(default_factory=dict)
    page_layout: Optional["PageLayout"] = None
    pg_num_fmt: Optional[str] = None   # w:pgNumType/@w:fmt (e.g. lowerRoman, decimal, upperRoman)
    pg_num_start: Optional[int] = None # w:pgNumType/@w:start (1-based)

    def get_header(self, variant: str = "default") -> Optional[HeaderFooter]:
        return self.headers.get(variant)

    def get_footer(self, variant: str = "default") -> Optional[HeaderFooter]:
        return self.footers.get(variant)


@dataclass
class ImageAsset:
    """Extracted binary image asset, keyed by its media path inside the DOCX.

    Asset identity is SEPARATE from placement identity (see `Image`). The same
    media file referenced more than once yields ONE ImageAsset and multiple
    `Image` placements. Bytes are stored exactly once here, never duplicated
    into every placement object.
    """

    source_path: str                 # media path inside the package, e.g. "word/media/image1.png"
    media_type: str                  # MIME type, e.g. "image/png"
    data: Optional[bytes] = None     # raw image bytes (None if media is missing/broken)
    relationship_ids: List[str] = field(default_factory=list)  # r:embed ids that point here
    missing: bool = False            # True when the relationship resolves but bytes are absent


@dataclass
class CommentRangeStart:
    """Semantic marker for w:commentRangeStart.

    Preserves exact OOXML range boundary in document order. The renderer
    uses the ordered Paragraph.content sequence to open/close <mark> wrappers.
    """

    comment_id: str


@dataclass
class CommentRangeEnd:
    """Semantic marker for w:commentRangeEnd."""

    comment_id: str


@dataclass
class NoteReference:
    """Reference to a footnote or endnote from body content.

    Resolved via w:footnoteReference / w:endnoteReference @w:id. The reference
    itself carries no text; the renderer emits a clickable superscript link to
    the corresponding Note body. Multiple references may target the same note.
    """

    note_type: str  # "footnote" | "endnote" | "comment"
    note_id: str    # OOXML w:id as string (preserves non-contiguous / negative ids)


@dataclass
class Note:
    """Footnote or endnote definition from word/footnotes.xml or word/endnotes.xml.

    Preserves ordered blocks (Paragraph/Table) via the same pipeline as body
    content, including formatting/hyperlinks/images.

    For threaded comments (word/commentsExtended.xml / w15:commentEx) the
    parent/reply hierarchy is captured via para_id / parent_id / replies.
    Flat comments (no extended part) remain parent_id=None and replies=[].
    """

    note_type: str  # "footnote" | "endnote" | "comment"
    note_id: str
    blocks: List["Block"] = field(default_factory=list)
    # --- comment-specific metadata (None for footnote/endnote or flat comment) ---
    author: Optional[str] = None
    date: Optional[str] = None
    initials: Optional[str] = None
    # Modern Word threaded-comment identity (w15:paraId). Stable hex string.
    para_id: Optional[str] = None
    # Resolved parent comment id (w:id of parent Note), None for root.
    parent_id: Optional[str] = None
    # Raw w15:paraIdParent before resolution (debug/trace).
    para_id_parent: Optional[str] = None
    # Nested replies (direct children). For root, contains its replies; for
    # replies that themselves have children, nested recursively.
    replies: List["Note"] = field(default_factory=list)
    # Done/resolved flag from w15:done="1" (None when absent).
    done: Optional[bool] = None


@dataclass
class Image:
    """A normalized image PLACEMENT inside the document.

    References an ImageAsset by `source_path`; does NOT carry the binary bytes
    itself (those live in the asset store to avoid duplication). Holds the
    display dimensions taken from the DOCX drawing (wp:extent), so the HTML
    renders at the size the author specified, not the intrinsic pixel size.
    """

    image_id: str                   # unique placement id within the document, e.g. "img1"
    relationship_id: str            # r:embed value, e.g. "rId5"
    source_path: str                # media path; key into the ImageAsset store
    media_type: str = "image/png"
    width: Optional[int] = None     # displayed width in CSS pixels (from wp:extent)
    height: Optional[int] = None    # displayed height in CSS pixels (from wp:extent)
    alt_text: Optional[str] = None  # from wp:docPr/@descr or @name; None when absent
    wrap_type: str = "inline"       # "inline" or "anchor" (floating)

    # ---- Floating / anchored placement metadata (DOCX semantics) ----
    # All fields below are populated ONLY when wrap_type == "anchor". For inline
    # images they stay None. The values are extracted verbatim from wp:anchor so
    # the renderer can reproduce the DOCX coordinate intent; they are NOT browser
    # CSS values (the renderer converts EMU offsets via core.units).
    #
    # relative_from_* is the RAW OOXML positionH/@relativeFrom (preserved, even if
    # unsupported). position_* is the normalized/category view used for layout; it
    # is "unsupported" when relative_from is a value this engine does not model,
    # so the distinction between "we saw it" and "we can place it" is explicit.
    relative_from_horizontal: Optional[str] = None  # page|margin|column|character|paragraph
    relative_from_vertical: Optional[str] = None
    position_horizontal: Optional[str] = None       # same set, or "unsupported"
    position_vertical: Optional[str] = None
    # Exactly one of (offset_*, alignment_*) is meaningful per axis: OOXML uses
    # either <posOffset> (EMU, may be negative) or <align> (keyword). Both None
    # means the anchor supplied neither (treated as zero offset).
    offset_horizontal: Optional[int] = None         # EMU, horizontal posOffset
    offset_vertical: Optional[int] = None           # EMU, vertical posOffset
    alignment_horizontal: Optional[str] = None       # center|left|right|inside|outside
    alignment_vertical: Optional[str] = None
    # The specific wrap element present on the anchor. "inline" images have None.
    # Values: square | topAndBottom | none | tight | through
    wrap_mode: Optional[str] = None
    # Wrap distances in EMU, keyed top/bottom/left/right. Sourced from the
    # anchor's distT/distB/distL/distR attributes (and wrap sub-element where
    # applicable). None when the anchor declared no distances.
    wrap_distances: Optional[Dict[str, int]] = None
    # Wrap polygon from wp:wrapPolygon/wp:start + wp:lineTo. Coordinates are in
    # the OOXML normalized 21600x21600 coordinate space (relative to image extent).
    # List of (x, y) tuples. None when no polygon or parsing failed.
    wrap_polygon: Optional[List[Tuple[int, int]]] = None
    # Nearest-block association, filled by core.anchoring.associate_floating_images.
    # nearest_block_id references a Paragraph.block_id; confidence is 0..1.
    nearest_block_id: Optional[str] = None
    nearest_block_confidence: Optional[float] = None
    # wp:anchor/@behindDoc ("1" => image sits behind text). Drives z-index so we
    # can position the image behind the text in the rendered HTML.
    behind_doc: bool = False
    # Anchor paragraph association, set by core.anchoring.associate_floating_images.
    # Allows the renderer to identify which paragraph an anchored image is attached to.
    anchor_paragraph_index: Optional[int] = None
    anchor_paragraph_text: Optional[str] = None
    section_index: Optional[int] = None
    column_index: Optional[int] = None
    table_id: Optional[str] = None
    cell_row: Optional[int] = None
    cell_col: Optional[int] = None
    rotation: Optional[int] = None  # a:xfrm/@rot in 1/60000 deg, None when absent
    extent_cx: Optional[int] = None  # raw wp:extent cx in EMU
    extent_cy: Optional[int] = None  # raw wp:extent cy in EMU

    # ---- Visual transforms (DrawingML) ----
    # Crop rectangle from a:srcRect (fraction of image in 1/100000 units).
    # l, t, r, b represent left, top, right, bottom crop offsets.
    crop_left: Optional[int] = None
    crop_top: Optional[int] = None
    crop_right: Optional[int] = None
    crop_bottom: Optional[int] = None
    # Flip flags from a:xfrm/@flipH, @flipV
    flip_h: bool = False
    flip_v: bool = False
    # Effect extent from wp:effectExtent (EMU) - visual effect bounds
    effect_extent_l: Optional[int] = None
    effect_extent_t: Optional[int] = None
    effect_extent_r: Optional[int] = None
    effect_extent_b: Optional[int] = None

    def __getitem__(self, key: str) -> Any:
        """Allow dict-like access to Image attributes for backward compatibility."""
        return getattr(self, key)

    def __contains__(self, key: str) -> bool:
        """Check if attribute exists."""
        return hasattr(self, key)


@dataclass
class PageLayout:
    """Physical page + margin geometry extracted from sectPr (in EMU).

    DOCX has no pagination engine here, but anchored images reference page /
    margin / paragraph coordinate systems. This struct carries the real geometry
    needed to map those relativeFrom values onto determinate CSS pixel origins.
    Defaults match US-Letter when sectPr is absent.
    """

    width_emu: int = 7772400          # page width  (8.5in)
    height_emu: int = 10058400        # page height (11in)
    margin_left_emu: int = 914400     # 1in
    margin_top_emu: int = 914400      # 1in
    margin_right_emu: int = 914400    # 1in
    margin_bottom_emu: int = 914400  # 1in
    cols_num: int = 1
    cols_space_emu: int = 0
    cols_equal_width: bool = True
    col_widths_emu: Optional[List[int]] = None
    col_spaces_emu: Optional[List[int]] = None

    @property
    def page_width_px(self) -> int:
        return emu_to_px(self.width_emu)

    @property
    def page_height_px(self) -> int:
        return emu_to_px(self.height_emu)

    @property
    def content_width_px(self) -> int:
        """Usable text width = page width minus left+right margins."""
        return emu_to_px(self.width_emu - self.margin_left_emu - self.margin_right_emu)

    @property
    def content_height_px(self) -> int:
        return emu_to_px(self.height_emu - self.margin_top_emu - self.margin_bottom_emu)

    @property
    def margin_left_px(self) -> int:
        return emu_to_px(self.margin_left_emu)

    @property
    def margin_top_px(self) -> int:
        return emu_to_px(self.margin_top_emu)

    def column_boxes_px(self):
        usable_emu = self.width_emu - self.margin_left_emu - self.margin_right_emu
        if self.cols_num <= 1:
            return [{"left_emu": 0, "left_px": 0, "width_emu": usable_emu, "width_px": emu_to_px(usable_emu), "right_px": emu_to_px(usable_emu)}]
        boxes = []
        if self.col_widths_emu and len(self.col_widths_emu) == self.cols_num:
            x = 0
            for i in range(self.cols_num):
                w = self.col_widths_emu[i]
                left_px = emu_to_px(x)
                width_px = emu_to_px(w)
                boxes.append({"left_emu": x, "left_px": left_px, "width_emu": w, "width_px": width_px, "right_px": left_px + width_px})
                if i < self.cols_num - 1:
                    sp = self.col_spaces_emu[i] if self.col_spaces_emu and i < len(self.col_spaces_emu) else self.cols_space_emu
                    x += w + sp
        else:
            total_space = (self.cols_num - 1) * self.cols_space_emu
            col_w = (usable_emu - total_space) // self.cols_num if self.cols_num else usable_emu
            x = 0
            for i in range(self.cols_num):
                left_px = emu_to_px(x)
                width_px = emu_to_px(col_w)
                boxes.append({"left_emu": x, "left_px": left_px, "width_emu": col_w, "width_px": width_px, "right_px": left_px + width_px})
                x += col_w + self.cols_space_emu
        return boxes


@dataclass
class NumberingLevel:
    """One level definition from word/numbering.xml (w:abstractNum/w:lvl)."""
    ilvl: int
    num_fmt: str = "decimal"
    lvl_text: str = "%1."
    start: int = 1


@dataclass
class NumberingModel:
    """numId -> abstractNumId -> {ilvl -> NumberingLevel}, parsed from numbering.xml."""
    abstract_nums: Dict[str, Dict[int, NumberingLevel]] = field(default_factory=dict)
    nums: Dict[str, str] = field(default_factory=dict)   # numId -> abstractNumId

    def resolve_level(self, num_id: Optional[str], ilvl: Optional[int]) -> Optional[NumberingLevel]:
        if num_id is None or ilvl is None:
            return None
        aid = self.nums.get(num_id)
        if aid is None:
            return None
        return self.abstract_nums.get(aid, {}).get(ilvl)


def _format_number_component(num_fmt: str, value: int) -> str:
    """Format a single counter value using the OOXML numFmt vocabulary."""
    if num_fmt == "lowerLetter":
        return _int_to_letter(value)
    if num_fmt == "upperLetter":
        return _int_to_letter(value).upper()
    if num_fmt == "lowerRoman":
        return _int_to_roman(value).lower()
    if num_fmt == "upperRoman":
        return _int_to_roman(value)
    # decimal (and any unsupported format) fall back to the bare integer so the
    # resolved path is never lost. Bullets are handled by the caller.
    return str(value)


def _int_to_letter(n: int) -> str:
    """1 -> 'a', 26 -> 'z', 27 -> 'aa' (OOXML lowerLetter semantics)."""
    if n <= 0:
        return str(n)
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(ord("a") + rem) + s
    return s


def _int_to_roman(n: int) -> str:
    """Standard uppercase Roman numeral for a positive integer."""
    if n <= 0:
        return str(n)
    table = (
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    )
    s = ""
    for value, symbol in table:
        while n >= value:
            s += symbol
            n -= value
    return s


def format_numbering_label(
    path: Optional[List[int]],
    level_formats: Optional[List[str]],
    lvl_text: Optional[str],
) -> Optional[str]:
    """Render the visible numbering label from resolved numbering metadata.

    Single source of truth for visible numbering: it consumes the resolved
    numbering_path, the per-level numFmt list, and the OOXML lvlText pattern
    (e.g. "%1.%2.%3"). It never re-parses numbering.xml and never inspects
    visible text. Returns None when there is no numbering to render (unnumbered
    heading, or bullet numbering which is not a numeric prefix).
    """
    if not path or not lvl_text:
        return None
    if level_formats and any(f == "bullet" for f in level_formats):
        return None

    formats = list(level_formats) if level_formats else ["decimal"] * len(path)
    if len(formats) < len(path):
        formats = formats + ["decimal"] * (len(path) - len(formats))

    def _replace(match: "re.Match") -> str:
        idx = int(match.group(1)) - 1
        if idx < 0 or idx >= len(path):
            return ""
        return _format_number_component(formats[idx], path[idx])

    raw = re.sub(r"%(\d)", _replace, lvl_text)
    raw = re.sub(r"\.+", ".", raw)
    raw = raw.lstrip(" .")
    placeholder_count = len(re.findall(r"%\d", lvl_text))
    if len(path) < placeholder_count:
        raw = raw.rstrip(" .")
    else:
        raw = raw.strip()
    return raw if raw else None


@dataclass
class Style:
    """A paragraph style from DOCX styles.xml.

    Carries the resolved formatting properties for a style, used by the
    semantic renderer to apply consistent formatting across the document.
    """
    name: str = "Normal"
    font_family: Optional[str] = None
    font_size: Optional[int] = None      # half-points (OOXML w:sz); renderer divides by 2 for pt
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[str] = None      # e.g. "single", "double"
    color: Optional[str] = None          # hex, e.g. "0000FF"
    space_before: float = 0.0            # in points
    space_after: float = 0.0             # in points
    left_indent: int = 0                 # in twips (1/20 pt)
    right_indent: int = 0                # in twips
    first_line_indent: int = 0           # in twips
    line_height: float = 1.0             # unitless multiplier
    based_on: str = ""                   # styleId of the parent style
    level: Optional[int] = None          # outline level (0-9) for headings
