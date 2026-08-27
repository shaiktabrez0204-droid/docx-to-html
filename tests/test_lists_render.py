import os, sys
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)
from semantic.pipeline import convert_docx
FIX = os.path.join(PROJECT_ROOT, "tests", "fixtures")

def test_ordered_lists_render_as_ol():
    r = convert_docx(os.path.join(FIX, "lists.docx"))
    assert any(p.numbering_path == [1] for p in r.paragraphs)
    assert '<ol class="docx-list docx-ordered-list">' in r.html
    assert '<span class="docx-number">1.</span> First item' in r.html
    assert r.html.count('<li>') == 3

def test_bullet_lists_render_as_ul():
    r = convert_docx(os.path.join(FIX, "unordered_lists.docx"))
    assert all(p.numbering_format == "bullet" for p in r.paragraphs if p.numbering_path)
    assert '<ul class="docx-list docx-bullet-list">' in r.html
    assert '<span class="docx-bullet">' in r.html
    assert r.html.count('<li>') == 3

def test_nested_lists_grouped():
    r = convert_docx(os.path.join(FIX, "nested-lists.docx"))
    assert '<ol class="docx-list' in r.html
    assert r.html.count('<li>') == 3
