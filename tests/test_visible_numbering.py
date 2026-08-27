"""Visible numbering rendering tests (REAL .docx fixtures + model unit tests).

Verifies that the already-resolved numbering model is turned into VISIBLE HTML
numbering in headings and the TOC, without re-parsing numbering.xml, without
regex on visible text, and without inventing numbers for unnumbered headings.
"""

import os
import re
import sys
import html as _html

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from core.model import Paragraph, Run, format_numbering_label
from output.html_renderer import render_paragraph
from semantic.toc import flatten_toc
from semantic.pipeline import convert_docx

FIX = os.path.join(PROJECT_ROOT, "tests", "fixtures")


def _load(name):
    p = os.path.join(FIX, name)
    assert os.path.exists(p), "missing fixture: %s" % p
    return p


def _by_text(paras, text):
    return [p for p in paras if "".join(r.text for r in p.runs) == text]


# ---- Phase 2: number formats ----
def test_format_decimal():
    assert format_numbering_label([1], ["decimal"], "%1") == "1"
    assert format_numbering_label([42], ["decimal"], "%1") == "42"


def test_format_lower_letter():
    assert format_numbering_label([1], ["lowerLetter"], "%1") == "a"
    assert format_numbering_label([3], ["lowerLetter"], "%1") == "c"
    assert format_numbering_label([26], ["lowerLetter"], "%1") == "z"
    assert format_numbering_label([27], ["lowerLetter"], "%1") == "aa"


def test_format_upper_letter():
    assert format_numbering_label([1], ["upperLetter"], "%1") == "A"
    assert format_numbering_label([26], ["upperLetter"], "%1") == "Z"


def test_format_lower_roman():
    assert format_numbering_label([1], ["lowerRoman"], "%1") == "i"
    assert format_numbering_label([3], ["lowerRoman"], "%1") == "iii"
    assert format_numbering_label([4], ["lowerRoman"], "%1") == "iv"
    assert format_numbering_label([9], ["lowerRoman"], "%1") == "ix"


def test_format_upper_roman():
    assert format_numbering_label([1], ["upperRoman"], "%1") == "I"
    assert format_numbering_label([2], ["upperRoman"], "%1") == "II"
    assert format_numbering_label([4], ["upperRoman"], "%1") == "IV"


# ---- Phase 3: lvlText resolution (generic, no hardcoding) ----
def test_lvltext_multi_level_per_format():
    # Each %N uses its own level's numFmt.
    assert format_numbering_label([1, 2, 3],
                                  ["decimal", "lowerLetter", "lowerRoman"],
                                  "%1.%2.%3") == "1.b.iii"


def test_lvltext_parens_and_suffix():
    assert format_numbering_label([1], ["decimal"], "(%1)") == "(1)"
    assert format_numbering_label([1], ["decimal"], "%1)") == "1)"
    assert format_numbering_label([1, 1, 1], ["decimal", "decimal", "decimal"], "%1.%2.%3") == "1.1.1"


def test_lvltext_skipped_level_renders_zero():
    # OOXML renders skipped parent levels as 0.
    assert format_numbering_label([1, 0], ["decimal", "decimal"], "%1.%2") == "1.0"


# ---- Phase 9: bullets are not numeric prefixes ----
def test_bullet_returns_none():
    assert format_numbering_label([1], ["bullet"], "%1.") is None


def test_render_bullet_has_no_number_span():
    p = Paragraph(runs=[Run(text="Item")], heading_level=1, heading_id="h1-item",
                  numbering_path=[1], numbering_format="bullet",
                  numbering_text_pattern="%1.", numbering_level_formats=["bullet"])
    out = render_paragraph(p)
    assert "docx-number" not in out
    assert "Item" in out


def test_no_path_returns_none():
    assert format_numbering_label(None, None, "%1") is None
    assert format_numbering_label([1], ["decimal"], None) is None


# ---- Phase 4/5: headings + TOC render the resolved number ----
def test_heading_number_matches_model():
    r = convert_docx(_load("num-h1-h2-h3.docx"))
    for p in r.paragraphs:
        if p.heading_level and p.numbering_path:
            expected = format_numbering_label(
                p.numbering_path, p.numbering_level_formats, p.numbering_text_pattern)
            assert expected is not None
            assert ('<span class="docx-number">%s</span>' % _html.escape(expected)) in r.html


def test_heading_ids_preserved_with_numbering():
    r = convert_docx(_load("num-h1-h2-h3.docx"))
    for hid in (p.heading_id for p in r.paragraphs if p.heading_id):
        assert 'id="%s"' % hid in r.html


def test_toc_number_single_source():
    r = convert_docx(_load("num-h1-h2-h3.docx"))
    by_id = {p.heading_id: p for p in r.paragraphs}
    for e in r.flat_toc():
        p = by_id.get(e["id"])
        if p and p.numbering_path:
            expected = format_numbering_label(
                p.numbering_path, p.numbering_level_formats, p.numbering_text_pattern)
            # TOC label is derived from the SAME paragraph fields as the heading.
            assert e["numbering_label"] == expected
            assert ('<span class="docx-number">%s</span>' % _html.escape(expected)) in r.html


# ---- Phase 6: restarts ----
def test_restart_renders_resolved_number():
    r = convert_docx(_load("num-restart.docx"))
    by_text = {"".join(x.text for x in p.runs): p for p in r.paragraphs if p.heading_level}
    two = by_text["Chapter Two"]
    expected = format_numbering_label(
        two.numbering_path, two.numbering_level_formats, two.numbering_text_pattern)
    assert expected is not None
    assert ('<span class="docx-number">%s</span>' % _html.escape(expected)) in r.html


# ---- Phase 7: skipped levels ----
def test_skipped_levels_render_from_path():
    r = convert_docx(_load("num-skipped.docx"))
    for p in r.paragraphs:
        if p.heading_level and p.numbering_path:
            expected = format_numbering_label(
                p.numbering_path, p.numbering_level_formats, p.numbering_text_pattern)
            assert expected is not None
            assert ('<span class="docx-number">%s</span>' % _html.escape(expected)) in r.html


# ---- Phase 8 / Phase 11 adversarial: unnumbered & fake text ----
def test_unnumbered_headings_have_no_number_span():
    r = convert_docx(_load("num-unnumbered.docx"))
    for p in r.paragraphs:
        if p.heading_level:
            assert p.numbering_path is None
    # No actual numbering span element is emitted (CSS reference is fine).
    assert '<span class="docx-number"' not in r.html


def test_adversarial_visible_text_not_numbered():
    r = convert_docx(_load("num-adversarial-text.docx"))
    fake = _by_text(r.paragraphs, "9. Fake Heading")[0]
    assert fake.numbering_path is None
    m = re.search(r'<h\d id="%s"[^>]*>.*?</h\d>' % re.escape(fake.heading_id), r.html, re.S)
    assert m, "heading element not found in html"
    assert "docx-number" not in m.group(0)


if __name__ == "__main__":
    for name in sorted(globals()):
        if name.startswith("test_") and callable(globals()[name]):
            globals()[name]()
            print("PASS", name)
    print("\nALL VISIBLE NUMBERING TESTS PASSED")
