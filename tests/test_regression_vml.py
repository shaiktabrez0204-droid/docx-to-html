import os, sys, re, zipfile, base64
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from adapter.ooxml_parser import OoxmlParser
from semantic.pipeline import convert_docx
from core.model import Paragraph, Table

FIX = os.path.join(PROJECT_ROOT, "tests", "fixtures", "regression-vml.docx")

def test_regression_media_discovered():
    assert os.path.exists(FIX)
    with zipfile.ZipFile(FIX) as z:
        media=[n for n in z.namelist() if 'word/media/' in n]
        assert len(media)==4, f"expected 4 media, got {len(media)} {media}"

def test_regression_relationships_resolve():
    p=OoxmlParser(FIX)
    # document rels + header/footer rels should all resolve
    blocks=p.parse_document()
    paras=[b for b in blocks if isinstance(b, Paragraph)]
    tables=[b for b in blocks if isinstance(b, Table)]
    secs=p.get_sections()
    assets=p.get_image_assets()
    # 4 distinct media files
    assert len(assets)==4
    for a in assets.values():
        assert a.data is not None and len(a.data)>0
        assert not a.missing

def test_regression_image_objects():
    p=OoxmlParser(FIX)
    blocks=p.parse_document()
    paras=[b for b in blocks if isinstance(b, Paragraph)]
    tables=[b for b in blocks if isinstance(b, Table)]
    secs=p.get_sections()
    total=0
    # paras
    for pa in paras:
        total+=len(pa.images)
    # table cells
    for t in tables:
        for row in t.rows:
            for cell in row.cells:
                for pa in cell.content:
                    total+=len(pa.images)
    # header/footer
    for sec in secs:
        for mp in (sec.headers, sec.footers):
            for hf in mp.values():
                for b in hf.blocks:
                    if isinstance(b, Paragraph):
                        total+=len(b.images)
                    elif isinstance(b, Table):
                        for row in b.rows:
                            for cell in row.cells:
                                for pa in cell.content:
                                    total+=len(pa.images)
    # Expected: header 1, footer 1, inline 1, floating 1, table 1, VML 1 = 6
    assert total==6, f"expected 6 image placements, got {total}"

def test_regression_html_img_count():
    res=convert_docx(FIX)
    assert res.html.count('<img')==6, f"html img count {res.html.count('<img')}"
    assert res.html.count('data:image')==6
    assert 'docx-image-missing' not in res.html
    # src valid and mime correct
    srcs=re.findall(r'src="([^"]+)"', res.html)
    data_srcs=[s for s in srcs if s.startswith('data:image')]
    assert len(data_srcs)==6
    for s in data_srcs:
        assert s.startswith('data:image/png;base64,') or s.startswith('data:image/jpeg;base64,') or s.startswith('data:image/gif;base64,')
        b64=s.split(',',1)[1]
        raw=base64.b64decode(b64)
        assert len(raw)>0
    # no duplicate ids
    ids=re.findall(r'image_id|img\d+', res.html)
    # check Image ids uniqueness via parser
    p=OoxmlParser(FIX)
    # collect ids
    blocks=p.parse_document()
    seen=set()
    from core.model import Paragraph, Table
    secs=p.get_sections()
    def collect():
        lst=[]
        for b in blocks:
            if isinstance(b, Paragraph):
                for im in b.images:
                    lst.append(im)
            elif isinstance(b, Table):
                for row in b.rows:
                    for cell in row.cells:
                        for pa in cell.content:
                            for im in pa.images:
                                lst.append(im)
        for sec in secs:
            for mp in (sec.headers, sec.footers):
                for hf in mp.values():
                    for b in hf.blocks:
                        if isinstance(b, Paragraph):
                            for im in b.images:
                                lst.append(im)
        return lst
    imgs=collect()
    ids=[i.image_id for i in imgs]
    assert len(ids)==len(set(ids)), "duplicate image ids"

def test_regression_dimensions_and_floating():
    p=OoxmlParser(FIX)
    blocks=p.parse_document()
    paras=[b for b in blocks if isinstance(b, Paragraph)]
    # find floating
    floats=[]
    tables=[b for b in blocks if isinstance(b, Table)]
    for pa in paras:
        for im in pa.images:
            if im.wrap_type=='anchor':
                floats.append(im)
    assert len(floats)==1
    fl=floats[0]
    assert fl.width is not None and fl.height is not None
    assert fl.wrap_mode=='square'
    # VML should have dimensions from style
    # find VML image by source_path image4
    all_imgs=[]
    for pa in paras:
        all_imgs.extend(pa.images)
    for t in tables:
        for row in t.rows:
            for cell in row.cells:
                for pa in cell.content:
                    all_imgs.extend(pa.images)
    vml=[i for i in all_imgs if 'image4' in i.source_path]
    assert len(vml)==1
    assert vml[0].width is not None and vml[0].height is not None

def test_regression_table_and_header_footer():
    p=OoxmlParser(FIX)
    blocks=p.parse_document()
    tables=[b for b in blocks if isinstance(b, Table)]
    assert len(tables)==1
    # table cell image remains inside table html
    res=convert_docx(FIX)
    # header/footer images are rendered inside header/footer tags
    assert 'docx-header' in res.html
    assert 'docx-footer' in res.html
    # table html contains images
    assert '<table' in res.html
    # ensure no horizontal overflow due to large images is not tested here, but html valid
