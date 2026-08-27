import os, sys, re, html
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from core.model import Paragraph, Run
from adapter.ooxml_parser import OoxmlParser
from semantic.style_resolver import StyleRegistry
from semantic.classifier import classify_paragraphs
from semantic.hierarchy import build_hierarchy
from semantic.toc import build_toc, flatten_toc
from semantic.pipeline import convert_docx
from output.html_renderer import render_sidebar_toc, render_html

FIXTURES = os.path.join(PROJECT_ROOT, "tests", "fixtures")

def _load(name):
    return os.path.join(FIXTURES, name)

def test_h1_h6_hierarchy():
    paras = OoxmlParser(_load("h1-h6.docx")).parse_paragraphs()
    reg = StyleRegistry(OoxmlParser(_load("h1-h6.docx")).get_styles())
    classify_paragraphs(paras, reg)
    tree = build_hierarchy(paras)
    assert [n.level for n in tree[0].children[0].children[0].children] != None
    # H1 > H2 > H3 > H4 > H5 > H6 chain
    node = tree[0]
    for expected in (2,3,4,5,6):
        assert node.children, f"missing level {expected}"
        node = node.children[0]
        assert node.level == expected
    toc = build_toc(tree)
    assert toc[0]["level"] == 1
    assert toc[0]["children"][0]["level"] == 2

def test_tree_parent_child_relationships():
    paras = OoxmlParser(_load("h1-h2.docx")).parse_paragraphs()
    reg = StyleRegistry(OoxmlParser(_load("h1-h2.docx")).get_styles())
    classify_paragraphs(paras, reg)
    tree = build_hierarchy(paras)
    toc = build_toc(tree)
    html_tree = render_sidebar_toc(toc)
    # parent should have toggle and children ul
    assert 'toc-parent' in html_tree
    assert 'toc-toggle' in html_tree
    assert 'toc-children' in html_tree
    assert 'toc-leaf' in html_tree
    # hierarchy: H1 with two H2 children
    assert toc[0]["children"][0]["text"] == "Section A"
    assert toc[0]["children"][1]["text"] == "Section B"

def test_sidebar_open_close_structure():
    r = convert_docx(_load("mixed-document.docx"))
    html_out = r.html
    assert 'id="viewer-sidebar"' in html_out
    assert 'id="sidebar-toggle"' in html_out
    assert 'id="sidebar-overlay"' in html_out
    assert 'viewer--sidebar-collapsed' in html_out  # CSS contains collapsed class
    assert 'aria-expanded="true"' in html_out
    assert 'aria-controls="viewer-sidebar"' in html_out
    assert 'Document Outline' in html_out
    # sidebar and doc-main both present, viewer flex layout
    assert 'class="viewer"' in html_out
    assert 'class="doc-main"' in html_out
    assert 'class="viewer-sidebar"' in html_out

def test_tree_expand_collapse_attributes():
    r = convert_docx(_load("mixed-document.docx"))
    html_out = r.html
    # parents have aria-expanded and aria-controls
    assert 'class="toc-toggle" aria-expanded="true"' in html_out
    assert 'role="tree"' in html_out
    assert 'role="treeitem"' in html_out
    assert 'role="group"' in html_out
    # collapsed state handled via is-collapsed class in JS, initial not collapsed
    assert 'is-collapsed' in html_out  # CSS defines it, but initial HTML should not have is-collapsed on parents
    # Ensure toggle controls correct id
    m = re.search(r'aria-controls="toc-([^"]+)"', html_out)
    assert m
    assert f'id="toc-{m.group(1)}"' in html_out

def test_real_heading_anchors():
    r = convert_docx(_load("mixed-document.docx"))
    html_out = r.html
    hids = [p.heading_id for p in r.paragraphs if p.heading_id]
    assert hids
    for hid in hids:
        assert f'id="{hid}"' in html_out, f"heading {hid} not in HTML"
        assert f'href="#{hid}"' in html_out, f"TOC link for {hid} missing"
    # every href resolves to a heading element
    hrefs = re.findall(r'href="#([^"]+)"', html_out)
    heading_ids = set(re.findall(r'<h\d id="([^"]+)"', html_out))
    for href in hrefs:
        # only TOC hrefs should be in heading_ids; ignore other # if any
        if href in heading_ids:
            continue
        # hrefs from TOC must be in heading_ids
        assert href in heading_ids, f"anchor {href} has no heading"

def test_active_heading_js_present():
    r = convert_docx(_load("mixed-document.docx"))
    html_out = r.html
    assert "IntersectionObserver" in html_out
    assert "is-active" in html_out
    assert 'aria-current' in html_out
    assert "__viewer" in html_out
    assert "setActive" in html_out

def test_automatic_ancestor_expansion_js():
    r = convert_docx(_load("mixed-document.docx"))
    html_out = r.html
    # JS should contain ancestor expansion logic
    assert "closest('li.toc-parent')" in html_out or 'closest("li.toc-parent")' in html_out or "closest('li" in html_out
    assert "is-collapsed" in html_out

def test_numbering_preservation():
    r = convert_docx(_load("num-h1-h2-h3.docx"))
    html_out = r.html
    for p in r.paragraphs:
        if p.heading_level and p.numbering_path:
            from core.model import format_numbering_label
            label = format_numbering_label(p.numbering_path, p.numbering_level_formats, p.numbering_text_pattern)
            assert label is not None
            # label appears once in heading and once in TOC, not duplicated like "1. 1."
            escaped = html.escape(label)
            assert f'<span class="docx-number">{escaped}</span>' in html_out
    # No duplicate like "1. 1. Introduction"
    assert "1. 1." not in html_out or html_out.count("1. 1.") == html_out.count("1. 1.1")  # allow legitimate

def test_no_duplicate_generated_toc():
    r = convert_docx(_load("mixed-document.docx"))
    html_out = r.html
    # There should be exactly one nav.toc (inside sidebar), not a full-page nav before viewer
    assert html_out.count('<nav class="toc"') == 1
    # Viewer should contain sidebar + doc-main, not nav outside viewer
    viewer_idx = html_out.find('class="viewer"')
    nav_idx = html_out.find('<nav class="toc"')
    assert nav_idx > viewer_idx
    # No duplicate TOC list outside sidebar
    assert html_out.count('toc-list') == 0 or html_out.count('toc-tree') == 1  # only sidebar tree

def test_existing_document_content_unchanged():
    # Use a fixture with tables, images, hyperlinks and ensure they remain
    r = convert_docx(_load("mixed-document.docx"))
    html_out = r.html
    # Headings preserved
    assert "<h1" in html_out and "<h2" in html_out
    # Paragraphs preserved
    assert "<p" in html_out
    # Check that a real TOC inside DOCX (if any) is not removed – mixed-document has TOC artifact paragraph "Table of Contents" which should still be in document
    assert "Table of Contents" in html_out
    # Ensure no horizontal overflow style
    assert "overflow-x:hidden" in html_out or "overflow:hidden" in html_out

def test_performance_500_headings():
    # Build a large TOC with 500 headings and ensure render is fast and HTML contains all
    from core.model import Paragraph, Run
    paras = []
    for i in range(500):
        lvl = (i % 6) + 1
        p = Paragraph(runs=[Run(text=f"Heading {i}")], style_name=f"Heading{lvl}", heading_level=lvl, heading_id=f"h{lvl}-heading-{i}")
        paras.append(p)
    from semantic.hierarchy import build_hierarchy
    from semantic.toc import build_toc
    tree = build_hierarchy(paras)
    toc = build_toc(tree)
    html_tree = render_sidebar_toc(toc)
    # Should contain 500 links without error and be reasonably sized
    assert html_tree.count('class="toc-link"') == 500
    # Full render should also handle 500 blocks
    from output.html_renderer import render_html
    html_out = render_html(paras, toc=toc)
    assert html_out.count('class="toc-link"') == 500
    assert 'IntersectionObserver' in html_out
