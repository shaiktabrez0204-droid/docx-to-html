import os, re, sys
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)
from semantic.pipeline import convert_docx
FIX = os.path.join(PROJECT_ROOT, "tests", "fixtures", "wraptight-isolation.docx")

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

def _anchors(html):
    body=html.split("<body>")[1]
    return [int(m.group(1)) for m in re.finditer(r'data-anchor="(\d+)"', body)]

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

def test_wraptight_wrappers_present():
    html=_render()
    anchors=_anchors(html)
    assert len(anchors)==5, f"expected 5 wraps, got {anchors}"
    body=html.split("<body>")[1]
    assert body.count('data-anchor=')==5
    # tight/through should be float with shape-outside, topAndBottom should be block
    assert 'float: left' in body or 'float: right' in body
    assert 'clear: both' in body
    assert 'shape-outside' in body

def test_wraptight_ownership():
    html=_render()
    anchors=_anchors(html)
    assert len(set(anchors))==5
    body=html.split("<body>")[1]
    block_starts=[m.start() for m in re.finditer(r'<div class="docx-block"', body)]
    for pos, anc in [(m.start(), int(m.group(1))) for m in re.finditer(r'data-anchor="(\d+)"', body)]:
        idx=-1
        for i, bs in enumerate(block_starts):
            if bs <= pos:
                idx=i
            else:
                break
        assert idx==anc, f"anchor {anc} != block {idx}"

def test_architecture_visibility():
    html=_render()
    blocks=_blocks(html)
    tid=_title_id(html)
    vis=_visible_blocks(blocks, "h2-architecture", tid)
    assert 4 in vis and 7 in vis
    anchors=_anchors(html)
    assert set(a for a in anchors if a in vis)=={4,7}

def test_data_model_visibility():
    html=_render()
    blocks=_blocks(html)
    tid=_title_id(html)
    vis=_visible_blocks(blocks, "h3-data-model", tid)
    assert vis=={5,6,7}
    anchors=_anchors(html)
    assert [a for a in anchors if a in vis]==[7]

def test_implementation_visibility():
    html=_render()
    blocks=_blocks(html)
    tid=_title_id(html)
    vis=_visible_blocks(blocks, "h2-implementation", tid)
    assert vis=={8,9,10}
    anchors=_anchors(html)
    assert [a for a in anchors if a in vis]==[10]

def test_topbottom_block():
    html=_render()
    # topAndBottom center should be block with clear both
    assert 'display: block' in html
    assert 'clear: both' in html
    assert 'float: none' in html

def test_dist_spacing():
    html=_render()
    # dist 24px should appear for one image
    assert '24px' in html
    assert '12px' in html

def test_tight_through_float():
    html=_render()
    # tight/through left/right should be float
    assert html.count('float: left')>=2
    assert html.count('float: right')>=2
