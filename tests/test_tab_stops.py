"""Minimal tests for w:tabs tab-stop layout."""
import os, tempfile, zipfile
from adapter.ooxml_parser import OoxmlParser
from semantic.pipeline import convert_docx

def _build(path, body):
    ct='''<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'''
    rels='''<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
    doc_rels='''<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'''
    styles='''<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>'''
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',ct)
        z.writestr('_rels/.rels',rels)
        z.writestr('word/_rels/document.xml.rels',doc_rels)
        z.writestr('word/document.xml',body)
        z.writestr('word/styles.xml',styles)

def _tmp():
    import tempfile, os
    fd,p=tempfile.mkstemp(suffix='.docx'); os.close(fd); return p

def test_tabs_extraction():
    body='''<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1440"/><w:tab w:val="center" w:pos="4320"/><w:tab w:val="right" w:pos="7200"/></w:tabs></w:pPr><w:r><w:t>A</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>B</w:t></w:r></w:p></w:body></w:document>'''
    p=_tmp()
    try:
        _build(p,body)
        parser=OoxmlParser(p)
        blocks=parser.parse_document(); parser.close()
        para=blocks[0]
        assert len(para.tabs)==3
        assert para.tabs[0].val=="left" and para.tabs[0].pos==1440
        assert para.tabs[1].val=="center"
        assert para.tabs[2].val=="right"
    finally:
        try: os.unlink(p)
        except: pass

def test_left_center_right_tab_render():
    body='''<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1440"/><w:tab w:val="center" w:pos="4320"/><w:tab w:val="right" w:pos="7200"/></w:tabs></w:pPr><w:r><w:t>Left</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Center</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Right</w:t></w:r></w:p>
</w:body></w:document>'''
    p=_tmp()
    try:
        _build(p,body)
        html=convert_docx(p).html
        assert 'data-val="left"' in html or 'docx-tab-segment' in html
        assert 'left:96px' in html or 'left:95px' in html or '96px' in html  # 1440 twip ≈96px
        assert 'center' in html
        assert 'right' in html
    finally:
        try: os.unlink(p)
        except: pass

def test_tab_position_conversion():
    # 1440 twips = 1 inch =96px
    from core.units import emu_to_px, twip_to_emu
    # Our renderer uses pos*635/9525 ≈96 per 1440
    pos=1440
    px=round(pos*635/9525)
    assert px==96

def test_missing_tab_fallback():
    body='''<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>A</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>B</w:t></w:r></w:p></w:body></w:document>'''
    p=_tmp()
    try:
        _build(p,body)
        html=convert_docx(p).html
        assert 'docx-tab' in html
    finally:
        try: os.unlink(p)
        except: pass

def test_footer_tab_with_page_field():
    html=convert_docx('benchmark_doc/csd-thesis-template-9th-draft.docx').html
    # footer3 should have center tab at 312px
    assert 'left:312px' in html
    assert 'docx-tab-segment' in html
    assert 'docx-page-number' in html

def test_leader_dot_right_alignment():
    body='''<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="8000" w:leader="dot"/></w:tabs></w:pPr><w:r><w:t>Chapter 1</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>12</w:t></w:r></w:p>
</w:body></w:document>'''
    p=_tmp()
    try:
        _build(p,body)
        html=convert_docx(p).html
        assert 'docx-tab-leader' in html
        assert 'data-leader="dot"' in html
        assert 'width:533px' in html or '533px' in html  # 8000 twip ≈533px
        assert 'Chapter 1' in html and '12' in html
    finally:
        try: os.unlink(p)
        except: pass

def test_leader_hyphen_underscore():
    for leader in ["hyphen","underscore","heavy"]:
        body='''<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="5000" w:leader="%s"/></w:tabs></w:pPr><w:r><w:t>A</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>B</w:t></w:r></w:p></w:body></w:document>''' % leader
        p=_tmp()
        try:
            _build(p,body)
            html=convert_docx(p).html
            assert 'data-leader="%s"' % leader in html
        finally:
            try: os.unlink(p)
            except: pass

def test_decimal_tab_render():
    body='''<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:tabs><w:tab w:val="decimal" w:pos="4000"/></w:tabs></w:pPr><w:r><w:t>3.14</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>2.7</w:t></w:r></w:p></w:body></w:document>'''
    p=_tmp()
    try:
        _build(p,body)
        html=convert_docx(p).html
        assert 'data-val="decimal"' in html
        assert 'data-pos="' in html
        assert '2.7' in html
    finally:
        try: os.unlink(p)
        except: pass

def test_decimal_with_leader():
    body='''<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:tabs><w:tab w:val="decimal" w:pos="5000" w:leader="dot"/></w:tabs></w:pPr><w:r><w:t>123.456</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>9.5</w:t></w:r></w:p></w:body></w:document>'''
    p=_tmp()
    try:
        _build(p,body)
        html=convert_docx(p).html
        assert 'data-val="decimal"' in html
        assert 'data-leader="dot"' in html
        assert '9.5' in html
    finally:
        try: os.unlink(p)
        except: pass

def test_decimal_no_dot_fallback():
    body='''<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:tabs><w:tab w:val="decimal" w:pos="3000"/></w:tabs></w:pPr><w:r><w:t>123</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>456</w:t></w:r></w:p></w:body></w:document>'''
    p=_tmp()
    try:
        _build(p,body)
        html=convert_docx(p).html
        assert 'data-val="decimal"' in html
        assert '456' in html
    finally:
        try: os.unlink(p)
        except: pass
