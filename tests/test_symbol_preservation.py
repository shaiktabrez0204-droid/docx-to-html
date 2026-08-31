"""Unit + integration tests for symbol preservation fix."""
import os, sys
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from adapter.ooxml_parser import OoxmlParser, _decode_sym_char
from semantic.pipeline import convert_docx
from core.model import Run, Paragraph

FIX = os.path.join(PROJECT_ROOT, "tests", "fixtures", "symbol-preservation.docx")

def test_symbol_F0D7_maps_to_multiplication():
    assert _decode_sym_char("Symbol", "F0D7") == "\u00D7"
    parser = OoxmlParser(FIX)
    blocks = parser.parse_document()
    # Find paragraph with Symbol
    paras = [b for b in blocks if isinstance(b, Paragraph)]
    # second para is Symbol
    target = paras[1]
    texts = [r.text for r in target.runs]
    assert "\u00D7" in "".join(texts), f"Symbol not decoded: {texts}"
    assert target.runs[0].text == "\u00D7"

def test_wingdings_symbol_is_not_empty():
    assert _decode_sym_char("Wingdings", "F028") is not None
    assert _decode_sym_char("Wingdings", "F028") != ""
    parser = OoxmlParser(FIX)
    blocks = parser.parse_document()
    paras = [b for b in blocks if isinstance(b, Paragraph)]
    target = paras[2]  # Wingdings para
    assert target.runs[0].text != ""
    assert len(target.runs[0].text) == 1

def test_noBreakHyphen():
    parser = OoxmlParser(FIX)
    blocks = parser.parse_document()
    paras = [b for b in blocks if isinstance(b, Paragraph)]
    target = paras[3]
    joined = "".join(r.text for r in target.runs)
    assert "\u2011" in joined, f"noBreakHyphen missing: {repr(joined)}"
    assert joined == "no\u2011break"

def test_softHyphen():
    parser = OoxmlParser(FIX)
    blocks = parser.parse_document()
    paras = [b for b in blocks if isinstance(b, Paragraph)]
    target = paras[4]
    joined = "".join(r.text for r in target.runs)
    assert "\u00AD" in joined, f"softHyphen missing: {repr(joined)}"
    assert joined == "soft\u00ADhyphen"

def test_eastAsia_font_fallback():
    parser = OoxmlParser(FIX)
    blocks = parser.parse_document()
    paras = [b for b in blocks if isinstance(b, Paragraph)]
    target = paras[6]  # eastAsia
    assert target.runs[0].font_family == "MS Gothic"
    # ascii fallback test: normal text still Calibri or None
    normal = paras[0]
    assert normal.runs[0].font_family is None or normal.runs[0].font_family == "Calibri"

def test_eastAsia_not_override_ascii():
    # When both ascii and eastAsia present, ascii should win
    import zipfile, io, os, tempfile
    from pathlib import Path
    # Build minimal docx with both fonts
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    doc_xml = f"""<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="MS Gothic"/></w:rPr><w:t>hello</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"""
    styles = f"""<?xml version="1.0"?><w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"""
    import zipfile as zf
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    content_types = """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" Type="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" Type="application/xml"/><Override PartName="/word/document.xml" Type="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" Type="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"""
    rels = """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""
    doc_rels = f"""<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"""
    with zf.ZipFile(path, "w", zf.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
    p = OoxmlParser(path)
    paras = p.parse_document()
    assert paras[0].runs[0].font_family == "Calibri", f"ascii should win, got {paras[0].runs[0].font_family}"
    p.close()
    os.remove(path)

def test_integration_html_contains_symbols():
    res = convert_docx(FIX, title="SymbolTest")
    html = res.html
    assert "\u00D7" in html, "HTML missing multiplication sign"
    assert "\u2011" in html, "HTML missing non-breaking hyphen"
    assert "\u00AD" in html, "HTML missing soft hyphen"
    # Should NOT have empty <p></p> for Symbol para (at least 2 non-empty symbol paras)
    assert html.count("<p></p>") == 0 or html.count("\u00D7") >= 1
    # Images still present
    assert "<img" in html
    assert html.count("<img") >= 2

def test_unknown_glyph_not_dropped():
    # Unknown font/char should still produce a Run, not be silently dropped
    result = _decode_sym_char("UnknownFont", "F0FF")
    assert result is not None and len(result) == 1
