import os, sys, re
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)
from semantic.pipeline import convert_docx

FIX = os.path.join(PROJECT_ROOT, "tests", "fixtures")

def test_complex_tables_style_borders_parsed():
    r = convert_docx(os.path.join(FIX, "complex-tables.docx"))
    tbl = next(b for b in r.blocks if hasattr(b, 'rows'))
    assert tbl.borders is not None
    assert tbl.borders['top'].val == 'single'
    assert tbl.borders['top'].sz == 4
    # renderer uses table borders
    assert 'border-top: 0.5pt solid #000000' in r.html

def test_borders_shading_varied_and_cell_override():
    r = convert_docx(os.path.join(FIX, "borders-shading.docx"))
    tbl = next(b for b in r.blocks if hasattr(b, 'rows'))
    # table varied colors
    assert tbl.borders['top'].color == 'FF0000'
    assert tbl.borders['top'].sz == 12
    assert tbl.borders['left'].val == 'dashed'
    # cell shading and per-cell border override
    c00 = tbl.rows[0].cells[0]
    assert c00.shading == 'FFFF00'
    assert c00.borders['top'].val == 'double'
    assert c00.borders['left'].val == 'nil'
    # html: yellow bg, double blue top, none left, inside borders gray
    assert 'background-color: #FFFF00' in r.html
    assert 'border-top: 2.25pt double #0000FF' in r.html
    assert 'border-left: none' in r.html
    # B2 bottom nil
    c11 = tbl.rows[1].cells[1]
    assert c11.borders['bottom'].val == 'nil'
    assert 'border-bottom: none' in r.html
    # B3 gray shading
    assert 'background-color: #C0C0C0' in r.html

def test_merged_cell_shading_and_span_preserved():
    r = convert_docx(os.path.join(FIX, "borders-shading.docx"))
    tbls = [b for b in r.blocks if hasattr(b, 'rows')]
    merged = tbls[1]
    assert merged.rows[1].cells[0].grid_span == 2
    assert merged.rows[1].cells[0].shading == 'DDEEFF'
    assert 'background-color: #DDEEFF' in r.html
    assert 'colspan="2"' in r.html

def test_fallback_when_no_borders():
    # Normal Table without explicit borders should fallback to 1px #999
    r = convert_docx(os.path.join(FIX, "borders-shading.docx"))
    tbls = [b for b in r.blocks if hasattr(b, 'rows')]
    normal_tbl = tbls[1]  # second table Normal Table no borders
    assert normal_tbl.borders is None
    # html fallback
    assert 'border-top: 1px solid #999' in r.html or 'border:1px solid #999' in r.html or '1px solid #999' in r.html

def test_merged_cells_complex_tables_still_render():
    r = convert_docx(os.path.join(FIX, "merged-cells.docx"))
    tbl = next(b for b in r.blocks if hasattr(b, 'rows'))
    assert tbl.borders is not None
    assert 'border-top: 0.5pt solid #000000' in r.html
    assert 'colspan="2"' in r.html
