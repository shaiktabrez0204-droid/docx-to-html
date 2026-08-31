"""Regression test for w:sdt / w:sdtContent picture placeholder (final-technical-report)."""
import os, sys, zipfile
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from adapter.ooxml_parser import OoxmlParser
from semantic.pipeline import convert_docx
from core.model import Paragraph, Table

FIX = os.path.join(PROJECT_ROOT, "benchmark_doc", "final-technical-report-template.docx")

def test_sdt_image_extracted():
    assert os.path.exists(FIX), f"missing fixture {FIX}"
    parser = OoxmlParser(FIX)
    blocks = parser.parse_document()
    assets = parser.get_image_assets()
    placements = []
    for b in blocks:
        if isinstance(b, Paragraph):
            placements.extend(b.images)
        elif isinstance(b, Table):
            for row in b.rows:
                for cell in row.cells:
                    for p in cell.content:
                        placements.extend(p.images)
    assert len(placements) == 1, f"expected 1 placement, got {len(placements)}"
    img = placements[0]
    assert img.relationship_id == "rId13"
    assert img.source_path == "word/media/image1.png"
    assert img.media_type == "image/png"
    assert img.width == 133, f"width {img.width}"
    assert img.height == 133
    assert img.extent_cx == 1266825
    assert len(assets) == 1
    assert "word/media/image1.png" in assets
    assert assets["word/media/image1.png"].data is not None
    parser.close()

def test_sdt_html_has_one_img():
    res = convert_docx(FIX, title="SDT")
    html = res.html
    assert html.count("<img") == 1, f"expected 1 <img>, got {html.count('<img')}"
    assert "data:image/png;base64," in html
    assert html.count("docx-image-missing") == 0
    # width attr or style
    assert 'width="133"' in html or "133px" in html

def test_sdt_no_duplicate():
    parser = OoxmlParser(FIX)
    blocks = parser.parse_document()
    placements = []
    for b in blocks:
        if isinstance(b, Paragraph):
            placements.extend(b.images)
    # Ensure no duplicate image_id
    ids = [p.image_id for p in placements]
    assert len(ids) == len(set(ids))
    parser.close()

def test_sdt_document_order_preserved():
    parser = OoxmlParser(FIX)
    blocks = parser.parse_document()
    # The sdt image should appear after the cover title block (first block is title)
    # and before the later table (if any) — simply verify it's not at 0
    idx = None
    for i, b in enumerate(blocks):
        if isinstance(b, Paragraph) and b.images:
            idx = i
            break
    assert idx is not None and idx > 0, f"image block index {idx} should be >0"
    parser.close()

def test_sdt_table_cell_nested():
    # Build minimal docx with w:sdt inside w:tc
    import tempfile, zipfile
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    doc_xml = f"""<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>
    <w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>cell before</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>nested sdt in cell</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"""
    styles = f"""<?xml version="1.0"?><w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"""
    content_types = """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" Type="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" Type="application/xml"/><Override PartName="/word/document.xml" Type="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" Type="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"""
    rels = """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""
    doc_rels = f"""<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"""
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", content_types)
            z.writestr("_rels/.rels", rels)
            z.writestr("word/document.xml", doc_xml)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/_rels/document.xml.rels", doc_rels)
        p = OoxmlParser(path)
        blocks = p.parse_document()
        assert len(blocks) == 1 and isinstance(blocks[0], Table)
        tbl = blocks[0]
        assert len(tbl.rows[0].cells) == 2
        assert tbl.rows[0].cells[0].content[0].runs[0].text == "cell before"
        assert tbl.rows[0].cells[1].content[0].runs[0].text == "nested sdt in cell"
        p.close()
    finally:
        os.remove(path)
