"""Image extraction smoke test against the REAL docx-to-html pipeline.

Replaces the legacy mammoth_adapter stub test: this now drives the actual
OoxmlParser -> normalized Image model -> HTML renderer chain on a real .docx
with an embedded PNG, per the visual-fidelity mandate.
"""

import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from adapter.ooxml_parser import OoxmlParser
from output.html_renderer import render_html


def test_image_extraction():
    """Real .docx with an embedded PNG yields a normalized image + <img> HTML."""
    path = os.path.join(PROJECT_ROOT, "tests", "fixtures", "img-inline-png.docx")
    assert os.path.exists(path), "missing fixture: %s" % path
    parser = OoxmlParser(path)
    paras = parser.parse_paragraphs()
    assets = parser.get_image_assets()

    imgs = []
    for p in paras:
        imgs.extend(p.images)
    assert imgs, "no image placements extracted"
    assert assets, "no image assets extracted"

    html_out = render_html(paras, assets=assets)
    assert "img-inline-png" in path  # fixture identity
    assert "<img" in html_out
    assert "data:image/png;base64," in html_out
