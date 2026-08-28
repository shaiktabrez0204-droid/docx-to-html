import os, re, sys
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)
from semantic.pipeline import convert_docx
FIX = os.path.join(PROJECT_ROOT, "tests", "fixtures", "inline-float-isolation.docx")

def _render():
    return convert_docx(FIX).html

def _blocks(html):
    out=[]
    for m in re.finditer(r'<div class="docx-block"([^>]*)>', html):
        attrs=m.group(1)
        hid=re.search(r'data-heading-id="([^"]+)"', attrs)
        lvl=re.search(r'data-level="(\d+)"', attrs)
        out.append((hid.group(1) if hid else None, int(lvl.group(1)) if lvl else 0))
    return out

def _title_id(html):
    m=re.search(r'<h1 id="([^"]+)" class="doc-title"', html)
    return m.group(1) if m else None

def _inline_anchors(html):
    body=html.split("<body>")[1]
    anchors=[]
    for m in re.finditer(r'<img[^>]*data-anchor="([^"]+)"[^>]*>', body):
        anchors.append(int(m.group(1)))
    return anchors

def _visible_blocks(blocks, focus_id, title_id):
    start_idx=-2
    level=1
    if focus_id==title_id:
        start_idx=-1
    else:
        for i,(hid,lvl) in enumerate(blocks):
            if hid==focus_id:
                start_idx=i
                level=lvl or 1
                break
    if start_idx==-2:
        return set()
    vis=set()
    if start_idx==-1:
        for i,(hid,lvl) in enumerate(blocks):
            if i>0 and lvl and lvl<=level:
                break
            vis.add(i)
    else:
        for i in range(start_idx, len(blocks)):
            hid,lvl=blocks[i]
            if i>start_idx and lvl and lvl<=level:
                break
            vis.add(i)
    return vis

def test_inline_wrappers_present():
    html=_render()
    anchors=_inline_anchors(html)
    assert len(anchors)==5, f"expected 5 inline floats, got {anchors}"
    body=html.split("<body>")[1]
    assert body.count('data-anchor=')>=5
    assert not re.search(r'<p[^>]*>.*?<div class="docx-float-wrap', body, re.DOTALL)
    # each inline float must be img with float left/right
    for m in re.finditer(r'<img[^>]*data-anchor="[^"]+"[^>]*>', body):
        tag=m.group(0)
        assert 'float: left' in tag or 'float: right' in tag
        assert 'docx-float-wrapped' in tag

def test_inline_ownership_distinct():
    html=_render()
    anchors=_inline_anchors(html)
    assert len(set(anchors))==5
    assert all(0 <= a < 14 for a in anchors)
    body=html.split("<body>")[1]
    block_starts=[m.start() for m in re.finditer(r'<div class="docx-block"', body)]
    wrap_pos=[(m.start(), int(m.group(1))) for m in re.finditer(r'data-anchor="(\d+)"', body) if 'docx-float-wrapped' in body[max(0,m.start()-200):m.start()+200] or 'data-anchor' in body[m.start()-50:m.start()+50]]
    for pos, anc in wrap_pos:
        idx=-1
        for i, bs in enumerate(block_starts):
            if bs <= pos:
                idx=i
            else:
                break
        assert idx==anc, f"inline anchor {anc} != block {idx}"

def test_architecture_inline_visibility():
    html=_render()
    blocks=_blocks(html)
    tid=_title_id(html)
    vis=_visible_blocks(blocks, "h2-architecture", tid)
    assert 4 in vis and 7 in vis
    anchors=_inline_anchors(html)
    vis_anchors=[a for a in anchors if a in vis]
    assert set(vis_anchors)=={4,7}

def test_data_model_inline_visibility():
    html=_render()
    blocks=_blocks(html)
    tid=_title_id(html)
    vis=_visible_blocks(blocks, "h3-data-model", tid)
    assert vis=={5,6,7}
    anchors=_inline_anchors(html)
    vis_anchors=[a for a in anchors if a in vis]
    assert vis_anchors==[7]

def test_implementation_inline_visibility():
    html=_render()
    blocks=_blocks(html)
    tid=_title_id(html)
    vis=_visible_blocks(blocks, "h2-implementation", tid)
    assert vis=={8,9,10}
    anchors=_inline_anchors(html)
    vis_anchors=[a for a in anchors if a in vis]
    assert vis_anchors==[10]

def test_clear_focus_restores_all():
    html=_render()
    anchors=_inline_anchors(html)
    assert len(anchors)==5
    vis=set(range(len(_blocks(html))))
    for a in anchors:
        assert a in vis

def test_positioning_unchanged():
    html=_render()
    body=html.split("<body>")[1]
    for m in re.finditer(r'<img[^>]*data-anchor="[^"]+"[^>]*>', body):
        tag=m.group(0)
        assert 'float: left' in tag or 'float: right' in tag
        assert 'width:' in tag and 'height:' in tag

def test_no_invalid_html_wrapper():
    html=_render()
    body=html.split("<body>")[1]
    assert '<div class="docx-para-float-wrap"' not in body
    # inline floats are img with data-anchor, not span wrapper
    assert body.count('data-anchor=')==5
