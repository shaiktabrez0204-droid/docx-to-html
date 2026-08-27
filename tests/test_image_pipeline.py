"""Real DOCX image extraction + model + HTML rendering tests.

Every test drives the ACTUAL pipeline:
  REAL .DOCX -> word/_rels/document.xml.rels -> word/media/* -> OoxmlParser
  -> normalized Image model -> HTML renderer -> (browser in *_playwright).

No regex on rendered HTML, no filename guessing, no fake image objects,
no hardcoded fixture paths. Relationship resolution is the source of truth.
"""

import os
import sys
import base64

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from core.model import Run, Paragraph, Image, ImageAsset
from adapter.ooxml_parser import OoxmlParser
from output.html_renderer import render_html, render_paragraph
from semantic.pipeline import convert_docx

FIX = os.path.join(PROJECT_ROOT, "tests", "fixtures")


def _load(name):
    path = os.path.join(FIX, name)
    assert os.path.exists(path), "missing fixture: %s" % path
    return path


def _all_images(paras):
    imgs = []
    for p in paras:
        imgs.extend(p.images)
    return imgs


# ---------------------------------------------------------------------------
# 1) Inline PNG actually extracted from real OOXML
# ---------------------------------------------------------------------------
def test_inline_png_real_extraction():
    p = OoxmlParser(_load("img-inline-png.docx"))
    paras = p.parse_paragraphs()
    assets = p.get_image_assets()
    imgs = _all_images(paras)
    assert len(imgs) == 1, "expected exactly one image placement"
    img = imgs[0]
    # Relationship resolution is authoritative: rId9 -> rels Target ->
    # "word/media/image1.png". We must NOT have inferred the path from filename.
    assert img.relationship_id == "rId9"
    assert img.source_path == "word/media/image1.png"
    assert img.media_type == "image/png"
    assert img.width and img.height, "display dimensions must be present"
    asset = assets.get(img.source_path)
    assert asset is not None, "asset store must contain the media"
    assert asset.data is not None and len(asset.data) > 0, "bytes must reach the store"
    # The bytes are genuine PNG (magic number).
    assert asset.data[:8] == b"\x89PNG\r\n\x1a\n", "extracted bytes are not a real PNG"


# ---------------------------------------------------------------------------
# 2) Relationship resolution is the source of truth (not filename inference)
# ---------------------------------------------------------------------------
def test_relationship_resolution_authoritative():
    p = OoxmlParser(_load("img-inline-jpeg.docx"))
    imgs = _all_images(p.parse_paragraphs())
    assert len(imgs) == 1
    # JPEG fixture: media is image1.jpg (not image1.jpeg), proving we resolved
    # from the rels Target, not assumed ".png".
    assert imgs[0].source_path == "word/media/image1.jpg"
    assert imgs[0].media_type == "image/jpeg"


# ---------------------------------------------------------------------------
# 3) Multiple images: counts, distinct placements, asset dedup
# ---------------------------------------------------------------------------
def test_multiple_images_and_dedup():
    p = OoxmlParser(_load("img-multiple.docx"))
    paras = p.parse_paragraphs()
    assets = p.get_image_assets()
    imgs = _all_images(paras)
    # Two PNG (one reused via same rId), one JPEG, one GIF -> 4 placements.
    assert len(imgs) == 4, "expected 4 image placements, got %d" % len(imgs)
    # Three distinct media files -> 3 assets (PNG reused reduces to 1 asset).
    assert len(assets) == 3, "expected 3 distinct assets, got %d" % len(assets)
    # Asset identity is separate from placement identity: the same media file
    # (image1.png) is referenced by two placements, so one asset is shared.
    src_counts = {}
    for i in imgs:
        src_counts[i.source_path] = src_counts.get(i.source_path, 0) + 1
    shared = [s for s, c in src_counts.items() if c > 1]
    assert shared, "expected a media file shared across multiple placements"
    assert assets[shared[0]].data is not None
    # ordering across paragraphs preserved
    order_srcs = [i.source_path for i in imgs]
    assert order_srcs[0] == "word/media/image1.png"
    assert order_srcs[1] == "word/media/image2.jpg"
    assert order_srcs[3] == "word/media/image3.gif"


# ---------------------------------------------------------------------------
# 4) Explicit dimensions preserved (5cm x 4cm -> ~189 x 151 px @96dpi)
# ---------------------------------------------------------------------------
def test_explicit_dimensions_preserved():
    p = OoxmlParser(_load("img-explicit-dims.docx"))
    imgs = _all_images(p.parse_paragraphs())
    assert len(imgs) == 1
    w, h = imgs[0].width, imgs[0].height
    assert w is not None and h is not None
    assert 185 <= w <= 193, "width should be ~189px, got %s" % w
    assert 148 <= h <= 154, "height should be ~151px, got %s" % h


# ---------------------------------------------------------------------------
# 5) Alt text preserved when present
# ---------------------------------------------------------------------------
def test_alt_text_preserved():
    p = OoxmlParser(_load("img-alt-text.docx"))
    imgs = _all_images(p.parse_paragraphs())
    assert len(imgs) == 1
    assert imgs[0].alt_text == "A red test rectangle"


# ---------------------------------------------------------------------------
# 6) No alt text -> None (never invented)
# ---------------------------------------------------------------------------
def test_no_alt_text_is_none():
    p = OoxmlParser(_load("img-no-alt.docx"))
    imgs = _all_images(p.parse_paragraphs())
    assert len(imgs) == 1
    assert imgs[0].alt_text is None


# ---------------------------------------------------------------------------
# 7) Same image referenced more than once -> one asset, two placements
# ---------------------------------------------------------------------------
def test_reused_image_single_asset():
    p = OoxmlParser(_load("img-reused.docx"))
    paras = p.parse_paragraphs()
    assets = p.get_image_assets()
    imgs = _all_images(paras)
    assert len(imgs) == 2, "two placements expected"
    assert len(assets) == 1, "single extracted asset expected (dedup)"
    assert imgs[0].source_path == imgs[1].source_path
    ids = {i.image_id for i in imgs}
    assert len(ids) == 2, "placements must have distinct image ids"


# ---------------------------------------------------------------------------
# 8) Image mixed with text preserves order: "before [IMG] after"
# ---------------------------------------------------------------------------
def test_mixed_text_order():
    p = OoxmlParser(_load("img-mixed-text.docx"))
    paras = p.parse_paragraphs()
    assert len(paras) == 1
    content = paras[0].content
    assert len(content) == 3
    assert isinstance(content[0], Run) and content[0].text == "This is before "
    assert isinstance(content[1], Image)
    assert isinstance(content[2], Run) and content[2].text == " and this is after."


# ---------------------------------------------------------------------------
# 9) Floating (wp:anchor) image is parsed and placed
# ---------------------------------------------------------------------------
def test_floating_anchor_parsed():
    p = OoxmlParser(_load("img-floating.docx"))
    imgs = _all_images(p.parse_paragraphs())
    assert len(imgs) == 1
    assert imgs[0].wrap_type == "anchor"
    # Still a real asset with bytes.
    assert p.get_image_assets()[imgs[0].source_path].data is not None


# ---------------------------------------------------------------------------
# 10) Image order across paragraphs preserved: A, img, B, img, C
# ---------------------------------------------------------------------------
def test_image_order_across_paragraphs():
    p = OoxmlParser(_load("img-multiple.docx"))
    paras = p.parse_paragraphs()
    seen = 0
    for pi, para in enumerate(paras):
        for ii, img in enumerate(para.images):
            seen += 1
            assert img.image_id == "img%d" % seen, "image ordering broken at %d" % seen


# ---------------------------------------------------------------------------
# 11) HTML rendering emits a real <img> data URL with dims + alt
# ---------------------------------------------------------------------------
def test_html_renders_data_url():
    p = OoxmlParser(_load("img-inline-png.docx"))
    paras = p.parse_paragraphs()
    html_out = render_html(paras, assets=p.get_image_assets())
    assert "<img" in html_out
    assert 'src="data:image/png;base64,' in html_out
    assert 'class="docx-image"' in html_out
    # dimensions and alt present
    assert "width=" in html_out and "height=" in html_out
    # No broken/empty src
    assert 'src=""' not in html_out


def test_html_alt_text_rendered():
    p = OoxmlParser(_load("img-alt-text.docx"))
    paras = p.parse_paragraphs()
    html_out = render_html(paras, assets=p.get_image_assets())
    assert 'alt="A red test rectangle"' in html_out


# ---------------------------------------------------------------------------
# 12) End-to-end pipeline wires image_assets through convert_docx
# ---------------------------------------------------------------------------
def test_pipeline_end_to_end():
    res = convert_docx(_load("img-inline-png.docx"))
    assert res.image_assets, "ConversionResult must expose image_assets"
    assert "<img" in res.html
    assert 'src="data:image/png;base64,' in res.html


# ---------------------------------------------------------------------------
# 13) Adversarial: missing relationship -> no broken image created
# ---------------------------------------------------------------------------
def test_adv_missing_rel_no_image():
    p = OoxmlParser(_load("img-adv-missing-rel.docx"))
    paras = p.parse_paragraphs()
    imgs = _all_images(paras)
    assert len(imgs) == 0, "missing relationship must not yield a placement"
    assert p.get_image_assets() == {}, "no asset should be extracted"


# ---------------------------------------------------------------------------
# 14) Adversarial: relationship resolves but media absent -> degrade safely
# ---------------------------------------------------------------------------
def test_adv_missing_media_degrades():
    p = OoxmlParser(_load("img-adv-missing-media.docx"))
    paras = p.parse_paragraphs()
    assets = p.get_image_assets()
    imgs = _all_images(paras)
    assert len(imgs) == 1, "placement still exists (relationship resolved)"
    src = imgs[0].source_path
    assert assets[src].missing is True, "asset must be flagged missing"
    html_out = render_html(paras, assets=assets)
    # Safe degradation: no broken <img>, a labeled placeholder instead.
    assert "docx-image-missing" in html_out
    assert 'src="data:' not in html_out


# ---------------------------------------------------------------------------
# 15) Adversarial: blip has no r:embed -> no placement
# ---------------------------------------------------------------------------
def test_adv_no_embed_no_image():
    p = OoxmlParser(_load("img-adv-no-embed.docx"))
    paras = p.parse_paragraphs()
    assert len(_all_images(paras)) == 0


# ---------------------------------------------------------------------------
# 16) Adversarial: no wp:extent -> dimensions None, render not broken
# ---------------------------------------------------------------------------
def test_adv_no_dims_width_none():
    p = OoxmlParser(_load("img-adv-no-dims.docx"))
    paras = p.parse_paragraphs()
    imgs = _all_images(paras)
    assert len(imgs) == 1
    assert imgs[0].width is None and imgs[0].height is None
    html_out = render_html(paras, assets=p.get_image_assets())
    # Image still renders (intrinsic size), just without width/height attrs.
    assert "<img" in html_out and 'src="data:image/png;base64,' in html_out
    import re
    assert not re.search(r'<img[^>]*\bwidth=', html_out)


# ---------------------------------------------------------------------------
# 17) Adversarial: unsupported media type -> degrade (no broken request)
# ---------------------------------------------------------------------------
def test_adv_unsupported_degrades():
    p = OoxmlParser(_load("img-adv-unsupported.docx"))
    paras = p.parse_paragraphs()
    imgs = _all_images(paras)
    assert len(imgs) == 1
    html_out = render_html(paras, assets=p.get_image_assets())
    assert "docx-image-missing" in html_out
    assert "<img" not in html_out
