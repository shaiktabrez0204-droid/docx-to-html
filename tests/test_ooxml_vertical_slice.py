"""Vertical slice test: REAL DOCX -> OOXML -> model -> HTML -> browser.

Uses only real .docx fixtures. Asserts that paragraphs, runs and basic
run formatting are extracted from actual OOXML and rendered to HTML.
No regex on content strings; no fake/stub data.
"""

import os
import sys

# Make the project packages importable (core, adapter, output)
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from core.model import Run, Paragraph
from adapter.ooxml_parser import OoxmlParser
from output.html_renderer import render_html, render_paragraph

FIXTURES = os.path.join(PROJECT_ROOT, "tests", "fixtures")


def _load(name):
    path = os.path.join(FIXTURES, name)
    assert os.path.exists(path), "missing fixture: %s" % path
    return path


def test_headings_paragraphs_and_runs():
    """headings.docx: 3 paragraphs, each 1 run, style Heading1..3."""
    paras = OoxmlParser(_load("headings.docx")).parse_paragraphs()
    assert len(paras) == 3, "expected 3 paragraphs, got %d" % len(paras)
    assert [p.style_name for p in paras] == ["Heading1", "Heading2", "Heading3"]
    assert [p.runs[0].text for p in paras] == ["Heading 1", "Heading 2", "Heading 3"]
    for p in paras:
        assert len(p.runs) == 1


def test_basic_run_formatting():
    """formatting.docx: one paragraph, one run, bold+italic+underline."""
    paras = OoxmlParser(_load("formatting.docx")).parse_paragraphs()
    assert len(paras) == 1
    run = paras[0].runs[0]
    assert run.text == "Bold italic underline"
    assert run.bold is True
    assert run.italic is True
    assert run.underline == "single"


def test_mixed_runs_preserve_boundaries():
    """mixed-runs.docx: multiple runs per paragraph, distinct formatting."""
    paras = OoxmlParser(_load("mixed-runs.docx")).parse_paragraphs()

    # Paragraph 1: 6 runs, mixed formatting preserved per-run
    p1 = paras[0]
    texts = [r.text for r in p1.runs]
    assert texts == ["Plain text ", "bold", " and ", "italic", " and ", "bold-italic-underline"]
    # per-run flags: None = no explicit formatting (will inherit from style)
    assert p1.runs[0].bold is None
    assert p1.runs[1].bold is True and p1.runs[1].italic is None
    assert p1.runs[3].italic is True and p1.runs[3].bold is None
    assert p1.runs[5].bold is True and p1.runs[5].italic is True and p1.runs[5].underline == "single"

    # Paragraph 2: Heading1 with bold run
    p2 = paras[1]
    assert p2.style_name == "Heading1"
    assert p2.runs[0].bold is True
    assert p2.runs[0].text == "Styled Heading One"

    # Paragraph 3: red large run + normal run
    p3 = paras[2]
    assert p3.runs[0].font_color == "FF0000"
    assert p3.runs[0].font_size == 28
    assert p3.runs[1].font_color is None  # no explicit color, will inherit from style


def test_default_font_from_styles():
    """styles.xml default run props are readable."""
    parser = OoxmlParser(_load("mixed-runs.docx"))
    default = parser.get_default_font()
    assert default["font_family"] == "Calibri"
    assert default["font_size"] == 22
    assert default["font_color"] == "000000"


def test_html_rendering():
    """Model -> HTML produces expected tags for real extracted data."""
    paras = OoxmlParser(_load("formatting.docx")).parse_paragraphs()
    html = render_html(paras)
    assert "<strong>" in html
    assert "<em>" in html
    assert 'text-decoration: single' in html
    assert "Bold italic underline" in html


def test_full_pipeline_end_to_end():
    """End-to-end: real DOCX -> model -> HTML file (for browser verification)."""
    path = _load("mixed-runs.docx")
    parser = OoxmlParser(path)
    paras = parser.parse_paragraphs()
    default = parser.get_default_font()
    # apply defaults to runs that lack explicit font info
    for p in paras:
        for r in p.runs:
            if not r.font_family or r.font_family == "Calibri":
                if not (r.bold or r.italic or r.underline or r.font_color != "#000000"):
                    r.font_family = default["font_family"]
                    r.font_size = default["font_size"]
    html = render_html(paras, title="mixed-runs")
    out_path = os.path.join(PROJECT_ROOT, "tests", "output_vertical_slice.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    assert os.path.exists(out_path)
    print("WROTE HTML:", out_path)
    return out_path


if __name__ == "__main__":
    test_headings_paragraphs_and_runs()
    print("PASS: headings")
    test_basic_run_formatting()
    print("PASS: basic run formatting")
    test_mixed_runs_preserve_boundaries()
    print("PASS: mixed runs boundaries")
    test_default_font_from_styles()
    print("PASS: default font")
    test_html_rendering()
    print("PASS: html rendering")
    out = test_full_pipeline_end_to_end()
    print("PASS: end-to-end ->", out)
    print("\nALL VERTICAL SLICE TESTS PASSED")
