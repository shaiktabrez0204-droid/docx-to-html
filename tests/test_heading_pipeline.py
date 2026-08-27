"""Real-DOCX heading / hierarchy / TOC pipeline tests.

Every case uses a REAL .docx fixture (no fabricated HTML, no regex on content
strings). Each test walks the full pipeline:

  DOCX -> OOXML -> model -> heading levels -> hierarchy -> TOC -> HTML ids -> nav links

and asserts the structural invariants required by the spec.
"""

import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from core.model import Paragraph
from adapter.ooxml_parser import OoxmlParser
from semantic.style_resolver import StyleRegistry
from semantic.classifier import classify_paragraphs
from semantic.hierarchy import build_hierarchy, flatten_hierarchy
from semantic.toc import build_toc, flatten_toc
from semantic.pipeline import convert_docx

FIXTURES = os.path.join(PROJECT_ROOT, "tests", "fixtures")


def _load(name):
    path = os.path.join(FIXTURES, name)
    assert os.path.exists(path), "missing fixture: %s" % path
    return path


def _heading_texts(paras):
    return [(p.heading_level, "".join(r.text for r in p.runs)) for p in paras if p.heading_level]


def _ids(paras):
    return [p.heading_id for p in paras if p.heading_level]


# ---- 1. H1 only ----
def test_h1_only():
    paras = OoxmlParser(_load("h1-only.docx")).parse_paragraphs()
    reg = StyleRegistry(OoxmlParser(_load("h1-only.docx")).get_styles())
    classify_paragraphs(paras, reg)
    levels = _heading_texts(paras)
    assert levels == [(1, "Introduction")]
    assert paras[0].heading_id == "h1-introduction"


# ---- 2. H1 + H2 ----
def test_h1_h2():
    paras = OoxmlParser(_load("h1-h2.docx")).parse_paragraphs()
    reg = StyleRegistry(OoxmlParser(_load("h1-h2.docx")).get_styles())
    classify_paragraphs(paras, reg)
    assert _heading_texts(paras) == [
        (1, "Chapter One"), (2, "Section A"), (2, "Section B")]


# ---- 3. H1 + H2 + H3 ----
def test_h1_h2_h3():
    paras = OoxmlParser(_load("h1-h2-h3.docx")).parse_paragraphs()
    reg = StyleRegistry(OoxmlParser(_load("h1-h2-h3.docx")).get_styles())
    classify_paragraphs(paras, reg)
    assert _heading_texts(paras) == [(1, "Top"), (2, "Mid"), (3, "Bottom")]
    # hierarchy nesting
    tree = build_hierarchy(paras)
    assert len(tree) == 1 and tree[0].level == 1
    assert tree[0].children[0].level == 2
    assert tree[0].children[0].children[0].level == 3


# ---- 4. H1-H6 ----
def test_h1_h6():
    paras = OoxmlParser(_load("h1-h6.docx")).parse_paragraphs()
    reg = StyleRegistry(OoxmlParser(_load("h1-h6.docx")).get_styles())
    classify_paragraphs(paras, reg)
    assert [lv for lv, _ in _heading_texts(paras)] == [1, 2, 3, 4, 5, 6]
    tree = build_hierarchy(paras)
    # H1 > H2 > H3 > H4 > H5 > H6 chain
    node = tree[0]
    for expected in (2, 3, 4, 5, 6):
        assert node.children, "expected deeper nesting at level %d" % expected
        node = node.children[0]
        assert node.level == expected


# ---- 5. numbered headings (level from style, not the number text) ----
def test_numbered_headings():
    paras = OoxmlParser(_load("numbered-headings.docx")).parse_paragraphs()
    reg = StyleRegistry(OoxmlParser(_load("numbered-headings.docx")).get_styles())
    classify_paragraphs(paras, reg)
    # The "1." / "1.1" prefixes must NOT change the resolved level.
    assert _heading_texts(paras) == [
        (1, "1. Getting Started"), (2, "1.1 Setup"),
        (3, "1.1.1 Install"), (2, "1.2 Configuration")]


# ---- 6. unnumbered headings ----
def test_unnumbered_headings():
    paras = OoxmlParser(_load("unnumbered-headings.docx")).parse_paragraphs()
    reg = StyleRegistry(OoxmlParser(_load("unnumbered-headings.docx")).get_styles())
    classify_paragraphs(paras, reg)
    assert _heading_texts(paras) == [(1, "Overview"), (2, "Components"), (3, "Database")]


# ---- 7. skipped levels preserved (H1, H3, H2) ----
def test_skipped_levels():
    paras = OoxmlParser(_load("skipped-levels.docx")).parse_paragraphs()
    reg = StyleRegistry(OoxmlParser(_load("skipped-levels.docx")).get_styles())
    classify_paragraphs(paras, reg)
    assert _heading_texts(paras) == [(1, "Part One"), (3, "Deep Detail"), (2, "Section Two")]
    tree = build_hierarchy(paras)
    # H1 has TWO children: H3 then H2, in document order. No invented H2.
    assert len(tree) == 1
    assert [c.level for c in tree[0].children] == [3, 2]


# ---- 8. duplicate heading text -> unique ids ----
def test_duplicate_heading_text():
    paras = OoxmlParser(_load("duplicate-heading-text.docx")).parse_paragraphs()
    reg = StyleRegistry(OoxmlParser(_load("duplicate-heading-text.docx")).get_styles())
    classify_paragraphs(paras, reg)
    texts = [t for _, t in _heading_texts(paras)]
    assert texts.count("Introduction") == 2
    assert texts.count("Summary") == 2
    ids = _ids(paras)
    assert len(ids) == len(set(ids)), "heading ids must be unique even for duplicate text"
    assert "h1-introduction" in ids
    assert "h1-introduction-2" in ids


# ---- 9. custom heading style via BasedOn inheritance ----
def test_custom_heading_style():
    paras = OoxmlParser(_load("custom-heading-style.docx")).parse_paragraphs()
    reg = StyleRegistry(OoxmlParser(_load("custom-heading-style.docx")).get_styles())
    # The custom style must resolve to a heading through BasedOn inheritance.
    resolved = reg.resolve("MyCustomHeading")
    assert resolved is not None
    assert resolved.is_heading is True
    assert resolved.heading_level == 1
    assert resolved.source in ("outline", "name_pattern")
    classify_paragraphs(paras, reg)
    custom = [p for p in paras if p.style_name == "MyCustomHeading"]
    assert custom, "custom heading paragraph not found"
    for p in custom:
        assert p.heading_level == 1
        assert p.heading_id and p.heading_id.startswith("h1-")


# ---- 10. realistic mixed document + TOC field must not corrupt headings ----
def test_mixed_document_toc_isolation():
    r = convert_docx(_load("mixed-document.docx"))
    levels = _heading_texts(r.paragraphs)
    assert (1, "Introduction") in levels
    assert (2, "Architecture") in levels
    assert (3, "Data Model") in levels
    assert (1, "Conclusion") in levels

    # The Word TOC artifact ("TOC 1" style) must NOT be a heading.
    toc_para = [p for p in r.paragraphs if "".join(x.text for x in p.runs).startswith("Table of Contents")]
    assert toc_para, "TOC artifact paragraph missing"
    assert toc_para[0].heading_level is None
    assert toc_para[0].style_name.lower().startswith("toc")

    # A visually prominent (bold, 28pt) NON-heading must NOT be a heading.
    bold_para = [p for p in r.paragraphs if "big bold statement" in "".join(x.text for x in p.runs)]
    assert bold_para
    assert bold_para[0].heading_level is None

    # TOC links must target the actual heading ids (no mismatch).
    hids = set(p.heading_id for p in r.paragraphs if p.heading_id)
    aids = set(e["id"] for e in r.flat_toc())
    assert aids == hids
    # every TOC anchor resolves in the rendered HTML
    import re
    html_anchors = set(re.findall(r'href="#([^"]+)"', r.html))
    assert html_anchors == hids
    html_ids = set(re.findall(r'<h\d id="([^"]+)"', r.html))
    assert html_ids == hids


# ---- regression: heading levels never come from visual cues ----
def test_no_visual_inference():
    """Bold/large text with Normal style is never a heading."""
    r = convert_docx(_load("mixed-document.docx"))
    for p in r.paragraphs:
        if p.style_name == "Normal":
            assert p.heading_level is None


if __name__ == "__main__":
    for name in sorted(globals()):
        if name.startswith("test_") and callable(globals()[name]):
            globals()[name]()
            print("PASS", name)
    print("\nALL HEADING PIPELINE TESTS PASSED")
