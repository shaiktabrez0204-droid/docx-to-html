"""Browser geometry validation for floating images (real Chromium).

This is the Phase 16 / mandatory browser evidence step: it loads the REAL
converted HTML in Chromium and measures actual boundingClientRect / computed
style to prove the floating images are positioned, wrapped, and sized per the
OOXML anchor metadata. Positioning is NEVER inferred from screenshots - we
compare measured geometry to the model-derived expected offsets in px.

Coordinate systems exercised:
  page     -> absolute inside .docx-page
  content  -> absolute inside .docx-content (margin box)
  p        -> absolute inside the anchored paragraph
  floatleft-> real CSS float (text wraps around the avoided rectangle)
"""

import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from semantic.pipeline import convert_docx
from core.units import emu_to_px

import pytest

OUT_DIR = os.path.join(PROJECT_ROOT, "tests", "test-output")
TOL = 14  # px tolerance for sub-pixel / rounding drift


def _convert(name):
    path = os.path.join(PROJECT_ROOT, "tests", "fixtures", name + ".docx")
    res = convert_docx(path)
    out = os.path.join(OUT_DIR, name + ".html")
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(res.html)
    return out, res.paragraphs


def _collect(page):
    return page.evaluate(
        """() => {
            const pr = document.querySelector('.docx-page').getBoundingClientRect();
            const cr = document.querySelector('.docx-content').getBoundingClientRect();
            const page = {x: pr.x, y: pr.y, w: pr.width, h: pr.height};
            const content = {x: cr.x, y: cr.y, w: cr.width, h: cr.height};
            const imgs = [...document.querySelectorAll(
                'img.docx-float, img.docx-float-wrapped')].map(im => {
                const r = im.getBoundingClientRect();
                const cs = getComputedStyle(im);
                const p = im.closest('p');
                const pr2 = p ? p.getBoundingClientRect() : null;
                return {
                    x: r.x, y: r.y, w: r.width, h: r.height,
                    nw: im.naturalWidth,
                    pos: cs.position,
                    float: cs.float,
                    p: p ? {x: pr2.x, y: pr2.y, w: pr2.width, h: pr2.height} : null,
                };
            });
            return {page, content, imgs};
        }"""
    )


# container -> (offH_px, offV_px); 'center'/'floatleft' handled specially.
POS = {
    "flt-page-center": ("center", None, None),
    "flt-page-offset": ("page", 210, 126),
    "flt-margin": ("content", 105, 84),
    "flt-paragraph": ("p", 52, 0),
    "flt-wrapsquare": ("floatleft", None, None),
    "flt-wrapnone": ("page", 126, 63),
    "flt-wraptopbottom": ("page", 157, 105),
    "flt-adv-no-posh": ("page", 0, 52),
    "flt-adv-unsupported-relfrom": ("page", 5, 52),
    "flt-adv-unsupported-wrap": ("page", 52, 52),
}

VISIBLE = {  # name -> expected rendered <img> count + must be visible
    "flt-multi": 3,
    "flt-near-heading": 1,
    "flt-near-paragraph": 1,
    "flt-multi-section": 2,
    "flt-lowconf": 1,
    "flt-adv-no-extent": 1,
    "flt-adv-zero-dims": 1,
    "flt-sizes": 2,
}

ALL = list(POS.keys()) + list(VISIBLE.keys())


def test_floating_geometry_in_browser():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        pytest.skip("playwright not installed")

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page(viewport={"width": 1100, "height": 1400})
            failed_reqs, console_errors = [], []

            for name in ALL:
                out, _ = _convert(name)
                page.on("requestfailed", lambda r: failed_reqs.append(r.url))
                page.on("console", lambda m: console_errors.append(m.text)
                        if m.type == "error" else None)
                failed_reqs.clear()
                console_errors.clear()
                page.goto("file://" + out)
                # attached (not "visible"): a 0x0 float (zero extent) is a
                # valid placement that is intentionally not visible. Wrapped
                # floats carry the docx-float-wrapped class (no base docx-float).
                page.wait_for_selector(
                    "img.docx-float, img.docx-float-wrapped",
                    state="attached", timeout=5000)
                data = _collect(page)

                imgs = data["imgs"]
                assert imgs, "%s: no floating <img> rendered" % name
                for im in imgs:
                    assert im["nw"] and im["nw"] > 0, \
                        "%s: image bytes did not decode (naturalWidth=%s)" % (name, im["nw"])

                if name in POS:
                    kind, off_h, off_v = POS[name]
                    im = imgs[0]
                    if kind == "center":
                        cx = data["page"]["x"] + data["page"]["w"] / 2
                        cy = data["page"]["y"] + data["page"]["h"] / 2
                        assert abs((im["x"] + im["w"] / 2) - cx) <= TOL, \
                            "%s: not horizontally centered (dx=%.1f)" % (
                                name, (im["x"] + im["w"] / 2) - cx)
                        assert abs((im["y"] + im["h"] / 2) - cy) <= TOL, \
                            "%s: not vertically centered (dy=%.1f)" % (
                                name, (im["y"] + im["h"] / 2) - cy)
                    elif kind == "floatleft":
                        assert im["float"] == "left", \
                            "%s: expected real CSS float, got %r" % (name, im["float"])
                        # Image sits on the left of the content box; text wraps.
                        c = data["content"]
                        assert im["x"] >= c["x"] - TOL and im["x"] <= c["x"] + TOL + 12, \
                            "%s: floated image not at content left" % name
                        assert im["x"] + im["w"] < c["x"] + c["w"], \
                            "%s: floated image spans full width (no wrap room)" % name
                    else:
                        base = data["page"] if kind == "page" else (
                            data["content"] if kind == "content" else im["p"])
                        assert base is not None, "%s: missing positioning context" % name
                        exp_x = base["x"] + off_h
                        exp_y = base["y"] + off_v
                        assert abs(im["x"] - exp_x) <= TOL, \
                            "%s: x off by %.1f (expected ~%.0f)" % (name, im["x"] - exp_x, exp_x)
                        assert abs(im["y"] - exp_y) <= TOL, \
                            "%s: y off by %.1f (expected ~%.0f)" % (name, im["y"] - exp_y, exp_y)
                        assert im["pos"] == "absolute", \
                            "%s: expected absolute positioning, got %r" % (name, im["pos"])

                if name in VISIBLE:
                    assert len(imgs) == VISIBLE[name], \
                        "%s: %d floats rendered, expected %d" % (
                            name, len(imgs), VISIBLE[name])

                if name == "flt-sizes":
                    widths = sorted(int(i["w"]) for i in imgs)
                    assert widths[1] - widths[0] > 10, \
                        "flt-sizes: distinct extents not reflected in render widths"

                if name == "flt-adv-zero-dims":
                    # Zero extent: spec says render WITHOUT an explicit width
                    # (no fabrication). The image falls back to its intrinsic
                    # size, so rendered width must equal natural width.
                    assert imgs[0]["w"] > 0, \
                        "flt-adv-zero-dims: must still render (intrinsic), not collapse"
                    assert abs(imgs[0]["w"] - imgs[0]["nw"]) <= 2, \
                        "flt-adv-zero-dims: zero extent must not impose explicit width"

                if name == "flt-adv-no-extent":
                    assert imgs[0]["w"] > 0, \
                        "flt-adv-no-extent: should still be visible (intrinsic size)"

                assert not failed_reqs, "%s: failed requests: %s" % (name, failed_reqs)
                assert not console_errors, "%s: console errors: %s" % (name, console_errors)

            browser.close()
    except Exception as e:
        if "executable doesn't exist" in str(e) or "launch" in str(e).lower():
            pytest.skip("chromium browser binary not installed: %s" % e)
        raise


if __name__ == "__main__":
    test_floating_geometry_in_browser()
    print("BROWSER FLOATING GEOMETRY CHECK PASSED")
