import os, sys, threading
import pytest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from semantic.pipeline import convert_docx

FIXTURES = os.path.join(PROJECT_ROOT, "tests", "fixtures")

def _load(name):
    return os.path.join(FIXTURES, name)

def test_viewer_header_present():
    r = convert_docx(_load("mixed-document.docx"))
    html = r.html
    assert 'class="viewer-header"' in html
    assert 'viewer-brand' in html
    assert 'viewer-docname' in html
    assert 'id="header-search"' in html
    assert 'id="header-toc"' in html
    assert 'id="viewer-download"' in html
    assert 'DOCX' in html and 'HTML' in html

def test_toc_search_elements():
    r = convert_docx(_load("mixed-document.docx"))
    html = r.html
    assert 'id="toc-search"' in html
    assert 'placeholder="Search headings...' in html
    assert 'id="toc-search-clear"' in html
    assert 'id="toc-no-results"' in html
    assert 'No headings found' in html
    assert 'toc-mark' in html

def test_focus_banner_elements():
    r = convert_docx(_load("mixed-document.docx"))
    html = r.html
    assert 'id="focus-banner"' in html
    assert 'id="focus-banner-title"' in html
    assert 'id="focus-banner-clear"' in html
    assert 'Show Full Document' in html
    assert 'Viewing' in html

def test_sidebar_toggle_and_overlay():
    r = convert_docx(_load("mixed-document.docx"))
    html = r.html
    assert 'id="viewer-sidebar"' in html
    assert 'id="sidebar-toggle"' in html
    assert 'id="sidebar-collapse"' in html
    assert 'id="sidebar-overlay"' in html
    assert 'viewer--sidebar-collapsed' in html
    assert 'aria-controls="viewer-sidebar"' in html

def test_viewer_style_premium():
    r = convert_docx(_load("mixed-document.docx"))
    html = r.html
    assert 'viewer-header' in html
    assert 'focus-banner' in html
    assert 'toc-mark' in html
    assert 'docx-page' in html
    assert 'prefers-reduced-motion' in html

def test_viewer_script_premium():
    r = convert_docx(_load("mixed-document.docx"))
    html = r.html
    assert 'toc-search' in html
    assert 'doSearch' in html or 'search' in html.lower()
    assert 'focusHeading' in html
    assert 'clearFocus' in html
    assert 'is-search-hidden' in html
    assert 'viewer-download' in html

def test_viewer_accessibility_attributes():
    r = convert_docx(_load("mixed-document.docx"))
    html = r.html
    assert 'aria-live="polite"' in html
    assert 'aria-expanded' in html
    assert 'aria-controls' in html
    assert 'aria-current' in html
    assert 'role="tree"' in html
    assert 'role="treeitem"' in html
