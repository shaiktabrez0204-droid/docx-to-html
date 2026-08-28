"""Focused tests for PAGE/NUMPAGES/PAGEREF field parsing and rendering."""
import os
import zipfile
import tempfile
from adapter.ooxml_parser import OoxmlParser
from core.model import Run
from semantic.pipeline import convert_docx

def _build_minimal_docx(path, body_xml, footer_xml=None, header_xml=None):
    content_types = '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
{footer_override}
{header_override}
</Types>'''.format(
        footer_override='<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' if footer_xml else '',
        header_override='<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' if header_xml else '',
    )
    rels = '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    doc_rels = '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{footer_rel}
{header_rel}
</Relationships>'''.format(
        footer_rel='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' if footer_xml else '',
        header_rel='<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' if header_xml else '',
    )
    styles = '''<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>'''
    document = body_xml
    # If footer/header provided, inject sectPr with reference using proper namespaces
    if footer_xml or header_xml:
        # body_xml already contains proper namespaces, we inject sectPr with r namespace
        sect_inner = ''
        if footer_xml:
            sect_inner += '<w:footerReference w:type="default" r:id="rId2" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />'
        if header_xml:
            sect_inner += '<w:headerReference w:type="default" r:id="rId3" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />'
        # Ensure body has r namespace in document element
        if 'xmlns:r' not in document:
            document = document.replace('xmlns:w=', 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w=')
        document = document.replace('</w:body>', '<w:sectPr>{0}</w:sectPr></w:body>'.format(sect_inner))
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('word/document.xml', document)
        z.writestr('word/styles.xml', styles)
        if footer_xml:
            z.writestr('word/footer1.xml', footer_xml)
        if header_xml:
            z.writestr('word/header1.xml', header_xml)

def _field_para(field_code, result="1"):
    return '''<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">{code}  </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{res}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'''.format(code=field_code, res=result)

def _temp_path():
    fd, path = tempfile.mkstemp(suffix='.docx')
    os.close(fd)
    return path

def test_page_field_parsing():
    path = _temp_path()
    try:
        body = '''<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{p}<w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>'''.format(p=_field_para("PAGE"))
        _build_minimal_docx(path, body)
        parser = OoxmlParser(path)
        blocks = parser.parse_document()
        parser.close()
        assert len(blocks) >= 1
        para = blocks[0]
        assert any(getattr(r, 'field_type', None) == 'PAGE' for r in para.runs), "PAGE field not parsed"
        assert not any(r.text == '1' for r in para.runs), "PAGE result should not be emitted as text"
    finally:
        try: os.unlink(path)
        except: pass

def test_numpages_field_parsing():
    path = _temp_path()
    try:
        body = '''<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{p}</w:body></w:document>'''.format(p=_field_para("NUMPAGES", "5"))
        _build_minimal_docx(path, body)
        parser = OoxmlParser(path)
        blocks = parser.parse_document()
        parser.close()
        para = blocks[0]
        assert any(getattr(r, 'field_type', None) == 'NUMPAGES' for r in para.runs)
        assert not any(r.text == '5' for r in para.runs)
    finally:
        try: os.unlink(path)
        except: pass

def test_pageref_field_parsing():
    path = _temp_path()
    try:
        body = '''<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{p}</w:body></w:document>'''.format(p=_field_para("PAGEREF _Toc123 \\h", "3"))
        _build_minimal_docx(path, body)
        parser = OoxmlParser(path)
        blocks = parser.parse_document()
        parser.close()
        para = blocks[0]
        assert any(getattr(r, 'field_type', None) == 'PAGEREF' for r in para.runs)
    finally:
        try: os.unlink(path)
        except: pass

def test_field_text_not_concatenated():
    path = _temp_path()
    try:
        footer = '''<?xml version="1.0" encoding="UTF-8"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:r><w:t>February 2015</w:t></w:r><w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">PAGE  </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>13</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'''
        body = '''<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>'''
        _build_minimal_docx(path, body, footer_xml=footer)
        result = convert_docx(path)
        html = result.html
        assert 'February 201513' not in html
        assert 'February 2015' in html
        assert 'docx-page-number' in html
    finally:
        try: os.unlink(path)
        except: pass

def test_ordinary_footer_text_unchanged():
    path = _temp_path()
    try:
        footer = '''<?xml version="1.0" encoding="UTF-8"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:r><w:t>Confidential</w:t></w:r></w:p></w:ftr>'''
        body = '''<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>'''
        _build_minimal_docx(path, body, footer_xml=footer)
        result = convert_docx(path)
        html = result.html
        assert 'Confidential' in html
    finally:
        try: os.unlink(path)
        except: pass

def test_roman_numbering_metadata_preserved():
    result = convert_docx('benchmark_doc/csd-thesis-template-9th-draft.docx')
    assert any(s.pg_num_fmt == 'lowerRoman' for s in result.sections)
    assert any(s.pg_num_start == 1 for s in result.sections)

def test_unsupported_fields_degrade_safely():
    path = _temp_path()
    try:
        body = '''<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">DATE  \\@ "MMMM yyyy"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>February 2015</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>'''
        _build_minimal_docx(path, body)
        parser = OoxmlParser(path)
        blocks = parser.parse_document()
        parser.close()
        para = blocks[0]
        assert not any(getattr(r, 'field_type', None) == 'DATE' for r in para.runs)
        assert any(r.text == 'February 2015' for r in para.runs)
        result = convert_docx(path)
        assert 'February 2015' in result.html
        assert '<span class="docx-page-number"' not in result.html
    finally:
        try: os.unlink(path)
        except: pass

def test_existing_headers_footers_still_render():
    result = convert_docx('benchmark_doc/csd-thesis-template-9th-draft.docx')
    html = result.html
    assert html.count('docx-header') >= 1 or html.count('docx-footer') >= 1
    assert len([b for b in result.blocks if getattr(b, 'heading_level', None)]) == 67
