import os, sys
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)
from adapter.ooxml_parser import OoxmlParser
from semantic.pipeline import convert_docx
from core.model import Table
FIX = os.path.join(PROJECT_ROOT, "tests", "fixtures")

def test_table_borders_extracted():
    p = OoxmlParser(os.path.join(FIX, "table-borders-shading.docx"))
    tbls = [b for b in p.parse_document() if isinstance(b, Table)]
    assert len(tbls) == 3
    t0 = tbls[0]
    assert t0.borders is not None and "top" in t0.borders
    assert t0.borders["top"].val == "single" and t0.borders["top"].sz == 12 and t0.borders["top"].color == "FF0000"
    assert t0.borders["insideH"].val == "dotted"
    assert t0.borders["insideV"].color == "FFA500"
    c00 = t0.rows[0].cells[0]
    assert c00.shading == "D9E1F2"
    assert c00.borders is not None and c00.borders["top"].val == "nil"
    c01 = t0.rows[0].cells[1]
    assert c01.shading == "FFF2CC"
    assert c01.borders["top"].sz == 18
    p.close()

def test_shading_rendered_as_background():
    r = convert_docx(os.path.join(FIX, "table-borders-shading.docx"))
    assert "background-color: #D9E1F2" in r.html
    assert "background-color: #FFF2CC" in r.html
    assert "background-color: #BDD7EE" in r.html

def test_border_css_rendered_with_colors_and_widths():
    r = convert_docx(os.path.join(FIX, "table-borders-shading.docx"))
    # table top FF0000, size 12 => 1.5pt solid #FF0000
    assert "1.5pt solid #FF0000" in r.html
    # insideV orange 18 => 2.25pt
    assert "2.25pt solid #FFA500" in r.html
    # cell override thick magenta
    assert "2.25pt solid #FF00FF" in r.html
    # double style
    assert "double #FF0000" in r.html or "double #0000FF" in r.html
    # nil suppresses => border-top: none
    assert "border-top: none" in r.html

def test_merged_cells_preserved_with_borders():
    r = convert_docx(os.path.join(FIX, "table-borders-shading.docx"))
    assert 'colspan="2"' in r.html
    # merged cell still has shading and borders
    assert "BDD7EE" in r.html

def test_style_fallback_table_grid_has_real_borders_not_blanket():
    r = convert_docx(os.path.join(FIX, "complex-tables.docx"))
    # complex-tables has TableGrid style borders via style lookup => 0.5pt solid #000000 (sz 4)
    assert "0.5pt solid #000000" in r.html
    # no blanket regression: tables still render
    assert '<table class="docx-table"' in r.html

def test_fallback_when_no_borders_gives_generic_via_inline():
    r = convert_docx(os.path.join(FIX, "table-borders-shading.docx"))
    # Table 1 has explicit borders, so interior cells with missing left/right should be none, not fallback
    # Ensure generic fallback not applied when table has borders (check that R0C0 left is green not #999)
    assert "0.75pt solid #00FF00" in r.html
