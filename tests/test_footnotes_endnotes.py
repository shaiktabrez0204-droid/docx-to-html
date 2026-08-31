"""Footnote / endnote parsing and rendering tests.

Covers OOXML -> parser -> model -> HTML -> browser invariants.
"""

import os
import re
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from core.model import Paragraph, Note, NoteReference
from adapter.ooxml_parser import OoxmlParser
from semantic.pipeline import convert_docx

FIXTURES = os.path.join(PROJECT_ROOT, "tests", "fixtures")
BENCH = os.path.join(PROJECT_ROOT, "benchmark_doc")


def _fixture(name):
    return os.path.join(FIXTURES, name)


def _bench(name):
    return os.path.join(BENCH, name)


# ---------------------------------------------------------------------------
# 1. footnote XML parsing
# ---------------------------------------------------------------------------
def test_footnote_xml_parsing():
    path = _fixture("footnotes-endnotes.docx")
    parser = OoxmlParser(path)
    notes = parser.get_footnotes()
    assert len(notes) == 3, f"expected 3 footnotes, got {len(notes)} ids={[n.note_id for n in notes]}"
    ids = sorted(n.note_id for n in notes)
    assert ids == ["2", "4", "9"]
    # content checks
    by_id = {n.note_id: n for n in notes}
    # id 2 plain
    assert any("First footnote" in (c.text if hasattr(c, "text") else "") for c in by_id["2"].blocks[0].content)
    # id 4 bold
    runs_4 = by_id["4"].blocks[0].content
    assert any(r.bold for r in runs_4 if hasattr(r, "bold"))
    # id 9 hyperlink
    runs_9 = by_id["9"].blocks[0].content
    assert any(getattr(r, "href", None) == "https://example.com" for r in runs_9)


def test_endnote_xml_parsing():
    path = _fixture("footnotes-endnotes.docx")
    parser = OoxmlParser(path)
    notes = parser.get_endnotes()
    assert len(notes) == 2, f"expected 2 endnotes, got {notes}"
    ids = sorted(n.note_id for n in notes)
    assert ids == ["1", "5"]
    by_id = {n.note_id: n for n in notes}
    assert len(by_id["5"].blocks) == 2  # second paragraph
    # hyperlink in second paragraph
    second_para_runs = by_id["5"].blocks[1].content
    assert any(getattr(r, "href", None) == "https://example.org" for r in second_para_runs)


def test_body_reference_parsing():
    path = _fixture("footnotes-endnotes.docx")
    parser = OoxmlParser(path)
    blocks = parser.parse_document()
    fn_refs = []
    en_refs = []
    for b in blocks:
        if isinstance(b, Paragraph):
            for c in b.content:
                if isinstance(c, NoteReference):
                    if c.note_type == "footnote":
                        fn_refs.append(c.note_id)
                    else:
                        en_refs.append(c.note_id)
        else:  # Table
            for row in b.rows:
                for cell in row.cells:
                    for para in cell.content:
                        for c in para.content:
                            if isinstance(c, NoteReference):
                                if c.note_type == "footnote":
                                    fn_refs.append(c.note_id)
                                else:
                                    en_refs.append(c.note_id)
    assert sorted(fn_refs) == ["2", "2", "4", "9", "9"]
    assert sorted(en_refs) == ["1", "1", "5"]


def test_reference_note_mapping():
    path = _fixture("footnotes-endnotes.docx")
    res = convert_docx(path)
    fn_ids = set(n.note_id for n in res.footnotes)
    en_ids = set(n.note_id for n in res.endnotes)
    for b in res.blocks:
        if isinstance(b, Paragraph):
            for c in b.content:
                if isinstance(c, NoteReference):
                    if c.note_type == "footnote":
                        assert c.note_id in fn_ids, f"footnote ref {c.note_id} has no body"
                    else:
                        assert c.note_id in en_ids, f"endnote ref {c.note_id} has no body"
        elif hasattr(b, "rows"):
            for row in b.rows:
                for cell in row.cells:
                    for para in cell.content:
                        for c in para.content:
                            if isinstance(c, NoteReference):
                                if c.note_type == "footnote":
                                    assert c.note_id in fn_ids
                                else:
                                    assert c.note_id in en_ids


def test_multiple_references_to_same_note():
    path = _fixture("footnotes-endnotes.docx")
    res = convert_docx(path)
    html = res.html
    # footnote 2 referenced twice -> two refs with distinct ids, one body, two backrefs
    refs_2 = re.findall(r'id="footnote-ref-2[^"]*"', html)
    assert len(refs_2) == 2, f"expected 2 refs for footnote 2, got {refs_2}"
    assert 'id="footnote-ref-2-1"' in html
    assert 'id="footnote-ref-2-2"' in html
    bodies_2 = re.findall(r'id="footnote-2"', html)
    assert len(bodies_2) == 1
    # backrefs to each reference
    assert 'href="#footnote-ref-2-1"' in html
    assert 'href="#footnote-ref-2-2"' in html
    # endnote 1 similarly
    refs_e1 = re.findall(r'id="endnote-ref-1[^"]*"', html)
    assert len(refs_e1) == 2
    assert len(re.findall(r'id="endnote-1"', html)) == 1


def test_non_contiguous_note_ids():
    path = _fixture("footnotes-endnotes.docx")
    res = convert_docx(path)
    fn_ids = [n.note_id for n in res.footnotes]
    # must preserve exact OOXML ids, not reindexed
    assert fn_ids == ["2", "4", "9"] or sorted(fn_ids) == ["2", "4", "9"]
    en_ids = [n.note_id for n in res.endnotes]
    assert sorted(en_ids) == ["1", "5"]
    html = res.html
    for nid in ["2", "4", "9"]:
        assert f'id="footnote-{nid}"' in html
    for nid in ["1", "5"]:
        assert f'id="endnote-{nid}"' in html


def test_special_separator_ids():
    # footnotes.xml contains -1,0,1 special; parser must skip them
    path = _fixture("footnotes-endnotes.docx")
    parser = OoxmlParser(path)
    fns = parser.get_footnotes()
    for n in fns:
        assert n.note_id not in ("-1", "0", "1"), "separator/continuation should be excluded"
    # thesis template has 3 specials + 1 real
    thesis = _bench("csd-thesis-template-9th-draft.docx")
    if os.path.exists(thesis):
        parser2 = OoxmlParser(thesis)
        fns2 = parser2.get_footnotes()
        # raw xml has 4 footnotes, but after filtering only 1 real
        assert len(fns2) == 1 and fns2[0].note_id == "2"
        ens2 = parser2.get_endnotes()
        assert len(ens2) == 1 and ens2[0].note_id == "1"


def test_html_reference_generation():
    path = _fixture("footnotes-endnotes.docx")
    res = convert_docx(path)
    html = res.html
    # body references must be sup with link
    for nid in ["2", "4", "9"]:
        pattern = rf'<sup class="docx-footnote-ref"><a href="#footnote-{nid}"[^>]*>{re.escape(nid)}</a></sup>'
        # For duplicates, href same but id differs; check at least one matches
        assert re.search(pattern, html), f"missing footnote ref {nid}"
    for nid in ["1", "5"]:
        pattern = rf'<sup class="docx-endnote-ref"><a href="#endnote-{nid}"[^>]*>{re.escape(nid)}</a></sup>'
        assert re.search(pattern, html), f"missing endnote ref {nid}"


def test_html_note_generation():
    path = _fixture("footnotes-endnotes.docx")
    res = convert_docx(path)
    html = res.html
    assert '<section class="docx-footnotes"><h2>Footnotes</h2>' in html
    assert '<section class="docx-endnotes"><h2>Endnotes</h2>' in html
    assert 'class="docx-footnote" id="footnote-2"' in html
    assert 'class="docx-footnote" id="footnote-9"' in html
    assert 'class="docx-endnote" id="endnote-5"' in html
    # note ordering preserved
    idx2 = html.find('id="footnote-2"')
    idx4 = html.find('id="footnote-4"')
    idx9 = html.find('id="footnote-9"')
    assert idx2 < idx4 < idx9


def test_back_reference_generation():
    path = _fixture("footnotes-endnotes.docx")
    res = convert_docx(path)
    html = res.html
    # every footnote body must have backref(s) to its reference(s)
    # footnote 2 has two backrefs
    foot2_section = html[html.find('id="footnote-2"'): html.find('id="footnote-4"')]
    assert foot2_section.count('docx-footnote-backref') == 2
    assert '#footnote-ref-2-1' in foot2_section
    assert '#footnote-ref-2-2' in foot2_section
    # footnote 4 single
    foot4_section = html[html.find('id="footnote-4"'): html.find('id="footnote-9"')]
    assert foot4_section.count('docx-footnote-backref') == 1
    # backref clickable: href matches ref id
    assert re.search(r'<a href="#footnote-ref-2-1"[^>]*>.*?</a>', foot2_section)


def test_formatting_preservation():
    path = _fixture("footnotes-endnotes.docx")
    res = convert_docx(path)
    html = res.html
    # footnote 4 bold-italic
    assert "<strong>" in html and "<em>" in html
    # hyperlink in footnote 9 and endnote 5 second para
    assert 'href="https://example.com"' in html
    assert 'href="https://example.org"' in html
    # tab and br in footnote 9: tab span, br tag
    foot9_html = html[html.find('id="footnote-9"'): html.find('</section>', html.find('id="footnote-9"'))]
    assert 'docx-tab' in foot9_html
    assert '<br>' in foot9_html


def test_real_docx_integration():
    # thesis template
    thesis = _bench("csd-thesis-template-9th-draft.docx")
    if os.path.exists(thesis):
        res = convert_docx(thesis)
        # forensic counts: 1 footnote ref, 1 endnote ref, 1 footnote body, 1 endnote body
        fn_refs = sum(1 for b in res.blocks if isinstance(b, Paragraph) for c in b.content if isinstance(c, NoteReference) and c.note_type == "footnote")
        en_refs = sum(1 for b in res.blocks if isinstance(b, Paragraph) for c in b.content if isinstance(c, NoteReference) and c.note_type == "endnote")
        assert fn_refs == 1
        assert en_refs == 1
        assert len(res.footnotes) == 1
        assert len(res.endnotes) == 1
        html = res.html
        assert 'docx-footnote-ref' in html
        assert 'docx-endnote-ref' in html
        assert 'docx-footnote' in html
        assert 'docx-endnote' in html
    # synthetic fixture
    path = _fixture("footnotes-endnotes.docx")
    res2 = convert_docx(path)
    assert len(res2.footnotes) == 3
    assert len(res2.endnotes) == 2
    html2 = res2.html
    assert html2.count('<sup class="docx-footnote-ref">') == 5
    assert html2.count('<sup class="docx-endnote-ref">') == 3
    assert html2.count('class="docx-footnote"') == 3
    assert html2.count('class="docx-endnote"') == 2


def test_no_duplicate_ids():
    path = _fixture("footnotes-endnotes.docx")
    res = convert_docx(path)
    html = res.html
    ids = re.findall(r'id="([^"]+)"', html)
    # footnote and endnote ids must be unique among themselves; heading duplicates pre-existing but footnote ids must not duplicate
    footnote_ids = [i for i in ids if i.startswith("footnote-")]
    endnote_ids = [i for i in ids if i.startswith("endnote-")]
    assert len(footnote_ids) == len(set(footnote_ids)), f"duplicate footnote ids: {footnote_ids}"
    assert len(endnote_ids) == len(set(endnote_ids)), f"duplicate endnote ids: {endnote_ids}"


def test_chromium_dom_structure():
    # Validate that generated HTML is well-formed for browser: anchors have matching targets
    path = _fixture("footnotes-endnotes.docx")
    res = convert_docx(path)
    html = res.html
    # every href="#footnote-X" must have a corresponding id="footnote-X"
    for href in re.findall(r'href="#(footnote-[^"]+)"', html):
        assert f'id="{href}"' in html, f"missing target for {href}"
    for href in re.findall(r'href="#(endnote-[^"]+)"', html):
        assert f'id="{href}"' in html
    for href in re.findall(r'href="#(footnote-ref-[^"]+)"', html):
        assert f'id="{href}"' in html
    for href in re.findall(r'href="#(endnote-ref-[^"]+)"', html):
        assert f'id="{href}"' in html
