"""Real DOCX floating-image (wp:anchor) tests: OOXML -> model -> renderer.

Every test drives the ACTUAL pipeline:
  REAL .DOCX -> OoxmlParser -> normalized Image (anchor fields)
  -> core.anchoring -> render_html -> CSS positioning

Unit/geometry checks assert the emitted CSS is the deterministic, model-derived
positioning (no regex inference of position). Browser geometry is verified
separately in tests/test_floating_playwright.py.
"""

import os
import re
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from core.model import PageLayout
from core.units import emu_to_px
from adapter.ooxml_parser import OoxmlParser
from core.anchoring import associate_floating_images, low_confidence_associations
from output.html_renderer import render_html
from semantic.pipeline import convert_docx

FIX = os.path.join(PROJECT_ROOT, "tests", "fixtures")

PAGE = PageLayout()  # US-Letter defaults used by the renderer


def _load(name):
    path = os.path.join(FIX, name)
    assert os.path.exists(path), "missing fixture: %s" % path
    return path


def _parse(name):
    p = OoxmlParser(_load(name))
    paras = p.parse_paragraphs()
    associate_floating_images(paras)
    return p, paras


def _floats(paras):
    return [i for para in paras for i in para.images if i.wrap_type == "anchor"]


def _float_styles(html_out):
    """Extract (class, style) for every rendered floating <img> from HTML."""
    out = []
    for m in re.finditer(r'<img ([^>]*docx-float[^>]*)>', html_out):
        tag = m.group(1)
        cls = re.search(r'class="([^"]*)"', tag)
        style = re.search(r'style="([^"]*)"', tag)
        out.append((cls.group(1) if cls else "", style.group(1) if style else ""))
    return out


def _style_dict(style):
    d = {}
    for part in style.split(";"):
        part = part.strip()
        if ":" in part:
            k, v = part.split(":", 1)
            d[k.strip()] = v.strip()
    return d


# ---------------------------------------------------------------------------
# OOXML parsing: anchor metadata is extracted from real DOCX
# ---------------------------------------------------------------------------
def test_page_center_parsed():
    p, paras = _parse("flt-page-center.docx")
    imgs = _floats(paras)
    assert len(imgs) == 1
    im = imgs[0]
    assert im.wrap_type == "anchor"
    assert im.relative_from_horizontal == "page"
    assert im.alignment_horizontal == "center"
    assert im.relative_from_vertical == "page"
    assert im.alignment_vertical == "center"
    assert im.wrap_mode == "square"
    assert im.width and im.height


def test_page_offset_parsed():
    p, paras = _parse("flt-page-offset.docx")
    im = _floats(paras)[0]
    assert im.relative_from_horizontal == "page"
    assert im.offset_horizontal == 2000000
    assert im.offset_vertical == 1200000
    assert im.wrap_mode == "none"


def test_margin_parsed():
    im = _floats(_parse("flt-margin.docx")[1])[0]
    assert im.relative_from_horizontal == "margin"
    assert im.offset_horizontal == 1000000
    assert im.wrap_mode == "square"
    assert im.wrap_distances == {"top": 50000, "bottom": 50000,
                                 "left": 50000, "right": 50000}


def test_paragraph_parsed():
    im = _floats(_parse("flt-paragraph.docx")[1])[0]
    assert im.relative_from_horizontal == "paragraph"
    assert im.offset_horizontal == 500000
    assert im.wrap_mode == "square"


def test_wrapsquare_parsed():
    im = _floats(_parse("flt-wrapsquare.docx")[1])[0]
    assert im.relative_from_horizontal == "margin"
    assert im.alignment_horizontal == "left"
    assert im.wrap_mode == "square"


def test_wraptopbottom_parsed():
    im = _floats(_parse("flt-wraptopbottom.docx")[1])[0]
    assert im.wrap_mode == "topAndBottom"


def test_wrapnone_parsed():
    im = _floats(_parse("flt-wrapnone.docx")[1])[0]
    assert im.wrap_mode == "none"


def test_multi_count():
    assert len(_floats(_parse("flt-multi.docx")[1])) == 3


def test_near_heading_parsed():
    res = convert_docx(_load("flt-near-heading.docx"))
    # Heading resolved by the semantic layer and the floating image parsed.
    assert any(p.heading_level == 1 for p in res.paragraphs)
    assert len(_floats(res.paragraphs)) == 1


def test_near_paragraph_parsed():
    assert len(_floats(_parse("flt-near-paragraph.docx")[1])) == 1


def test_multi_section_parsed():
    assert len(_floats(_parse("flt-multi-section.docx")[1])) == 2


def test_sizes_parsed():
    imgs = _floats(_parse("flt-sizes.docx")[1])
    assert len(imgs) == 2
    # Distinct displayed dimensions preserved from the two anchors.
    dims = sorted((i.width, i.height) for i in imgs)
    assert dims[0] != dims[1]


# ---------------------------------------------------------------------------
# Nearest-block association + confidence
# ---------------------------------------------------------------------------
def test_nearest_block_strong():
    im = _floats(_parse("flt-paragraph.docx")[1])[0]
    assert im.nearest_block_id is not None
    assert im.nearest_block_confidence == 0.95


def test_nearest_block_low_confidence():
    p, paras = _parse("flt-lowconf.docx")
    im = _floats(paras)[0]
    # Anchor sits in an EMPTY paragraph -> associate with neighbour, low conf.
    assert im.nearest_block_confidence == 0.6
    assert im.nearest_block_id is not None
    # Surfaced by the low-confidence report helper.
    low = list(low_confidence_associations(paras))
    assert any(i is im for i, _bid, _c in low)


# ---------------------------------------------------------------------------
# Unsupported / adversarial degrade safely (no fabricated coordinates)
# ---------------------------------------------------------------------------
def test_adv_missing_positionH():
    im = _floats(_parse("flt-adv-no-posh.docx")[1])[0]
    assert im.relative_from_horizontal is None
    assert im.offset_horizontal is None
    assert im.relative_from_vertical == "page"


def test_adv_missing_extent():
    im = _floats(_parse("flt-adv-no-extent.docx")[1])[0]
    assert im.width is None and im.height is None
    # Still renders without crashing (renderer degrades to intrinsic size).
    res = convert_docx(_load("flt-adv-no-extent.docx"))
    assert "docx-float" in res.html


def test_adv_unsupported_relfrom():
    im = _floats(_parse("flt-adv-unsupported-relfrom.docx")[1])[0]
    # A relativeFrom this engine does not model is preserved verbatim...
    assert im.relative_from_horizontal == "bogusCoord"
    # ...and the normalized category is explicitly "unsupported" (not page).
    assert im.position_horizontal == "unsupported"


def test_adv_unsupported_wrap():
    im = _floats(_parse("flt-adv-unsupported-wrap.docx")[1])[0]
    assert im.wrap_mode == "tight"  # preserved, not faked as square


def test_adv_zero_dims():
    im = _floats(_parse("flt-adv-zero-dims.docx")[1])[0]
    assert im.width == 0 and im.height == 0


# ---------------------------------------------------------------------------
# Renderer CSS strategy (deterministic, model-derived)
# ---------------------------------------------------------------------------
def test_render_page_center_centered():
    res = convert_docx(_load("flt-page-center.docx"))
    cls, style = _float_styles(res.html)[0]
    d = _style_dict(style)
    assert d["position"] == "absolute"
    assert d["left"] == "50%"
    assert d["top"] == "50%"
    assert "translate(-50%, -50%)" in d["transform"]


def test_render_page_offset_absolute():
    res = convert_docx(_load("flt-page-offset.docx"))
    d = _style_dict(_float_styles(res.html)[0][1])
    assert d["position"] == "absolute"
    assert d["left"] == "%dpx" % emu_to_px(2000000)   # 210px from page origin
    assert d["top"] == "%dpx" % emu_to_px(1200000)    # 126px


def test_render_margin_uses_margin_coordinate():
    res = convert_docx(_load("flt-margin.docx"))
    d = _style_dict(_float_styles(res.html)[0][1])
    # margin-relative => offset from the margin-box origin (.docx-content is
    # already inset by the page margin, so no extra margin is added).
    assert d["left"] == "%dpx" % emu_to_px(1000000)
    assert d["top"] == "%dpx" % emu_to_px(800000)


def test_render_paragraph_absolute_and_relative_paragraph():
    res = convert_docx(_load("flt-paragraph.docx"))
    cls, style = _float_styles(res.html)[0]
    d = _style_dict(style)
    assert d["position"] == "absolute"
    assert d["left"] == "%dpx" % emu_to_px(500000)
    # The anchor paragraph must be the positioning context.
    assert 'position: relative' in res.html


def test_render_wrapsquare_real_float():
    res = convert_docx(_load("flt-wrapsquare.docx"))
    cls, style = _float_styles(res.html)[0]
    d = _style_dict(style)
    # margin-relative + left align => genuine CSS float (text wraps).
    assert "docx-float-wrapped" in cls
    assert d["float"] == "left"
    # Wrap distance reserved as margin so the rectangle is avoided.
    assert d["margin"] == "5px 5px 5px 5px"


def test_render_wrapnone_overlay():
    res = convert_docx(_load("flt-wrapnone.docx"))
    d = _style_dict(_float_styles(res.html)[0][1])
    assert d["position"] == "absolute"
    assert d["left"] == "%dpx" % emu_to_px(1200000)
    assert d["top"] == "%dpx" % emu_to_px(600000)


def test_render_wraptopbottom_absolute():
    res = convert_docx(_load("flt-wraptopbottom.docx"))
    d = _style_dict(_float_styles(res.html)[0][1])
    assert d["position"] == "absolute"
    assert d["left"] == "%dpx" % emu_to_px(1500000)
    assert d["top"] == "%dpx" % emu_to_px(1000000)


def test_render_multi_all_floats():
    res = convert_docx(_load("flt-multi.docx"))
    styles = _float_styles(res.html)
    assert len(styles) == 3
    assert all("docx-float" in c for c, _ in styles)


def test_render_dimensions_from_extent():
    res = convert_docx(_load("flt-page-center.docx"))
    d = _style_dict(_float_styles(res.html)[0][1])
    # Width/height come from the anchor extent, not intrinsic pixels.
    assert "width" in d and "height" in d


def test_render_unsupported_relfrom_still_positioned():
    res = convert_docx(_load("flt-adv-unsupported-relfrom.docx"))
    d = _style_dict(_float_styles(res.html)[0][1])
    # character-relative falls back to paragraph coordinate (left:5px), not lost.
    assert d["position"] == "absolute"
    assert d["left"] == "%dpx" % emu_to_px(50000)


def test_render_unsupported_wrap_absolute_fallback():
    res = convert_docx(_load("flt-adv-unsupported-wrap.docx"))
    cls, style = _float_styles(res.html)[0]
    d = _style_dict(style)
    # wrapTight is unsupported -> absolute fallback (not a fake float).
    assert "docx-float" == cls
    assert d["position"] == "absolute"


def test_existing_inline_image_unaffected():
    # Inline-only DOCX must still render as docx-image with ZERO float <img>.
    res = convert_docx(_load("img-inline-png.docx"))
    assert 'class="docx-image"' in res.html
    assert len(_float_styles(res.html)) == 0


def test_existing_floating_fixture_still_anchor():
    p, paras = _parse("img-floating.docx")
    im = _floats(paras)[0]
    assert im.wrap_type == "anchor"
    assert im.relative_from_horizontal == "page"
    assert im.wrap_mode == "square"
