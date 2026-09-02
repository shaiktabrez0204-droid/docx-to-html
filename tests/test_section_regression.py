"""Focused regression for section-index / section-boundary invariant.

Covers the blank-preview root cause: parse_document() and get_sections()
must share ONE document-order traversal. Every block.section_index must be
< len(sections), section breaks inside wrappers (w:sdt, w:sdtContent,
customXml, smartTag) are handled identically by both APIs, and normal
documents remain single-section.
"""
import os
import sys
import tempfile
import zipfile

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from adapter.ooxml_parser import OoxmlParser
from core.layout import build_pages_from_sections, resolve_layout_state, assign_blocks_to_pages
from semantic.pipeline import convert_docx

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
FIX_FINAL = os.path.join(PROJECT_ROOT, "benchmark_doc", "final-technical-report-template.docx")
FIXTURES = os.path.join(PROJECT_ROOT, "tests", "fixtures")


def _make_docx(doc_xml: str) -> str:
    styles = f'''<?xml version="1.0"?><w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'''
    content_types = '''<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" Type="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" Type="application/xml"/><Override PartName="/word/document.xml" Type="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" Type="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'''
    rels = '''<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
    doc_rels = '''<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'''
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
    return path


def _hello_sdt_wrap2_xml() -> str:
    """Double-wrapped sdt with section break inside inner sdt (issue repro)."""
    return f'''<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>
<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent>
  <w:p><w:r><w:t>first in outer sdt</w:t></w:r></w:p>
  <w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr><w:sdtContent>
    <w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>break inside inner sdt</w:t></w:r></w:p>
  </w:sdtContent></w:sdt>
  <w:p><w:r><w:t>after break still outer sdt</w:t></w:r></w:p>
</w:sdtContent></w:sdt>
<w:p><w:r><w:t>after outer sdt</w:t></w:r></w:p>
<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>second break</w:t></w:r></w:p>
<w:p><w:r><w:t>final para</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>'''


def test_invariant_every_block_has_section():
    """Invariant: every block.section_index < len(sections) for all fixtures."""
    for name in ["final-technical-report-template.docx", "mixed-runs.docx", "headings.docx"]:
        if name == "final-technical-report-template.docx":
            path = FIX_FINAL
        else:
            path = os.path.join(FIXTURES, name)
        if not os.path.exists(path):
            continue
        parser = OoxmlParser(path)
        blocks = parser.parse_document()
        sections = parser.get_sections()
        parser.close()
        assert len(sections) >= 1
        for b in blocks:
            assert b.section_index < len(sections), f"{name}: block idx {b.section_index} >= len(sections) {len(sections)}"
            assert b.section_index >= 0


def test_section_indices_contiguous_and_correct():
    """Section indices must be contiguous from 0..n-1 and reflect document order."""
    xml = _hello_sdt_wrap2_xml()
    docx = _make_docx(xml)
    try:
        parser = OoxmlParser(docx)
        blocks = parser.parse_document()
        sections = parser.get_sections()
        parser.close()
        assert len(sections) == 3, f"expected 3 sections got {len(sections)}"
        assert len(blocks) == 6, f"expected 6 blocks got {len(blocks)}"
        # Contiguous: all indices 0..n-1 appear, none skipped
        indices = [b.section_index for b in blocks]
        assert indices == [0, 0, 1, 1, 1, 2], f"unexpected contiguity {indices}"
        # Monotonic non-decreasing reflects document order
        for i in range(1, len(indices)):
            assert indices[i] >= indices[i - 1], f"indices must be non-decreasing, got {indices}"
            assert indices[i] - indices[i - 1] <= 1, f"index jump >1 at {i}: {indices}"
    finally:
        os.remove(docx)


def test_hello_sdt_wrap2_no_loss():
    """hello_sdt_wrap2.docx must not lose blocks; layout must place every block."""
    xml = _hello_sdt_wrap2_xml()
    docx = _make_docx(xml)
    try:
        parser = OoxmlParser(docx)
        blocks = parser.parse_document()
        sections = parser.get_sections()
        parser.close()
        # Every block has a section
        for b in blocks:
            assert b.section_index < len(sections)
        # Layout: every block should be assigned to some page (sections drive pages)
        pages = build_pages_from_sections(sections)
        # assign_blocks_to_pages mutates pages; verify total assigned == total blocks
        assign_blocks_to_pages(blocks, pages, sections)
        assigned = sum(len(p.blocks) for p in pages)
        assert assigned == len(blocks), f"layout lost blocks: assigned {assigned} != {len(blocks)}"
        # Ensure html would render all blocks (no blank preview)
        res = convert_docx(docx, title="hello_sdt_wrap2")
        # All 6 blocks should be represented in html (at least 6 docx-block wrappers)
        assert res.html.count("docx-block") >= 6, f"html missing blocks {res.html.count('docx-block')}"
    finally:
        os.remove(docx)


def test_final_technical_no_orphans():
    """final-technical-report-template.docx must produce zero orphan blocks."""
    assert os.path.exists(FIX_FINAL), f"missing {FIX_FINAL}"
    parser = OoxmlParser(FIX_FINAL)
    blocks = parser.parse_document()
    sections = parser.get_sections()
    parser.close()
    orphans = [b for b in blocks if b.section_index >= len(sections)]
    assert len(orphans) == 0, f"found {len(orphans)} orphan blocks, sections={len(sections)}"
    # Contiguous check
    indices = sorted(set(b.section_index for b in blocks))
    assert indices == list(range(len(sections))), f"non-contiguous {indices} vs {len(sections)}"
    # Layout must place all blocks
    res = convert_docx(FIX_FINAL)
    assigned = sum(len(p.blocks) for p in res.pages)
    assert assigned == len(res.blocks), f"layout lost blocks {assigned} vs {len(res.blocks)}"
    # HTML must not be blank - at least 10 docx-blocks
    assert res.html.count("docx-block") >= 10


def test_normal_documents_single_section():
    """Documents without nested section boundaries remain single-section identically."""
    for name in ["mixed-runs.docx", "headings.docx", "formatting.docx"]:
        path = os.path.join(FIXTURES, name)
        assert os.path.exists(path), f"missing {path}"
        parser = OoxmlParser(path)
        blocks = parser.parse_document()
        sections = parser.get_sections()
        parser.close()
        # These fixtures have no explicit section breaks except the default body sectPr -> single section
        assert len(sections) == 1, f"{name}: expected 1 section got {len(sections)}"
        for b in blocks:
            assert b.section_index == 0, f"{name}: expected section 0 got {b.section_index}"


def test_customxml_wrapper_section():
    """Section break inside customXml wrapper must be discovered by both APIs."""
    xml = f'''<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>
<w:customXml w:uri="urn:test" w:element="x"><w:p><w:r><w:t>before</w:t></w:r></w:p>
<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>break inside customXml</w:t></w:r></w:p>
<w:p><w:r><w:t>after inside customXml</w:t></w:r></w:p></w:customXml>
<w:p><w:r><w:t>after wrapper</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>'''
    docx = _make_docx(xml)
    try:
        parser = OoxmlParser(docx)
        blocks = parser.parse_document()
        sections = parser.get_sections()
        parser.close()
        assert len(sections) == 2, f"customXml: expected 2 sections got {len(sections)}"
        assert len(blocks) == 4
        indices = [b.section_index for b in blocks]
        assert indices == [0, 0, 1, 1], f"customXml indices {indices}"
        for b in blocks:
            assert b.section_index < len(sections)
    finally:
        os.remove(docx)
