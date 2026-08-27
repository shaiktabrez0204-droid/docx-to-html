"""Structured numbering pipeline tests (REAL .docx fixtures only).

Verifies the SECOND, independent structural signal:
  numbering.xml -> numId/ilvl -> numbering path -> heading cross-validation
and that it never overrides style-derived heading levels or the hierarchy.
"""

import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from adapter.ooxml_parser import OoxmlParser
from semantic.style_resolver import StyleRegistry
from semantic.classifier import classify_paragraphs
from semantic.hierarchy import build_hierarchy, flatten_hierarchy
from semantic.numbering import (
    NumberingResolver,
    cross_validate,
    validate_hierarchy,
)
from semantic.pipeline import convert_docx
from core.model import NumberingModel, NumberingLevel, Paragraph, Run

FIX = os.path.join(PROJECT_ROOT, "tests", "fixtures")


def _load(name):
    p = os.path.join(FIX, name)
    assert os.path.exists(p), "missing fixture: %s" % p
    return p


def _by_text(paras, text):
    return [p for p in paras if "".join(r.text for r in p.runs) == text]


# ---- Phase 2/3/4: numbering.xml parsed, numPr parsed, path resolved ----
def test_numbering_xml_parsed():
    parser = OoxmlParser(_load("num-h1-h2-h3.docx"))
    model = parser.get_numbering()
    assert "1" in model.nums
    aid = model.nums["1"]
    assert aid in model.abstract_nums
    assert model.abstract_nums[aid][0].num_fmt == "decimal"
    assert model.abstract_nums[aid][2].lvl_text == "%1.%2.%3."


def test_numpr_parsed_from_paragraph():
    parser = OoxmlParser(_load("num-h1-h2-h3.docx"))
    paras = parser.parse_paragraphs()
    heads = [p for p in paras if p.num_id is not None]
    assert heads[0].num_id == "1" and heads[0].num_ilvl == 0
    assert heads[1].num_ilvl == 1
    assert heads[2].num_ilvl == 2


def test_numbering_path_resolved():
    r = convert_docx(_load("num-h1-h2-h3.docx"))
    texts = {("".join(x.text for x in p.runs)): p for p in r.paragraphs if p.heading_level}
    assert texts["Top"].numbering_path == [1]
    assert texts["Mid"].numbering_path == [1, 1]
    assert texts["Bottom"].numbering_path == [1, 1, 1]


# ---- Phase 5: cross-validation, level not changed ----
def test_consistent_numbering_validated():
    r = convert_docx(_load("num-h1-h2-h3.docx"))
    for v in r.numbering_validation:
        if v.has_numbering:
            assert v.consistent is True


def test_inconsistent_not_mutated():
    r = convert_docx(_load("num-inconsistent.docx"))
    p = _by_text(r.paragraphs, "Looks Like Chapter")[0]
    assert p.heading_level == 2
    assert p.num_ilvl == 0
    assert p.numbering_consistent is False
    issues = [v for v in r.numbering_validation if v.has_numbering and not v.consistent]
    assert issues, "expected a flagged inconsistency"


# ---- Phase 6: hierarchy validation ----
def test_hierarchy_consistent_fixtures_no_issues():
    for fn in ["num-h1.docx", "num-h1-h2.docx", "num-h1-h2-h3.docx",
               "num-restart.docx", "num-skipped.docx", "num-mixed.docx",
               "num-custom-style.docx", "num-adversarial-metadata.docx"]:
        r = convert_docx(_load(fn))
        assert r.hierarchy_issues == [], "%s produced unexpected issues: %s" % (fn, r.hierarchy_issues)


def test_duplicate_numbering_path_flagged():
    from semantic.hierarchy import HeadingNode
    p1 = _mk_para(1, "A", path=[1])
    p2 = _mk_para(1, "B", path=[1])
    p3 = _mk_para(2, "C", path=[1, 1])
    nodes = [HeadingNode(p1, 1, "A", "h1-a", 0),
             HeadingNode(p2, 1, "B", "h1-b", 1),
             HeadingNode(p3, 2, "C", "h2-c", 2)]
    issues = validate_hierarchy(nodes)
    assert any(i["type"] == "duplicate_numbering_path" for i in issues)


def test_numbering_not_nested_flagged():
    from semantic.hierarchy import HeadingNode
    p_parent = _mk_para(1, "A", path=[1])
    p_child = _mk_para(2, "B", path=[2, 3])
    parent_node = HeadingNode(p_parent, 1, "A", "h1-a", 0)
    child_node = HeadingNode(p_child, 2, "B", "h2-b", 1)
    parent_node.children.append(child_node)
    issues = validate_hierarchy([parent_node])
    assert any(i["type"] == "numbering_not_nested" for i in issues)


# ---- Phase 7: skipped levels preserved ----
def test_skipped_levels_preserved():
    r = convert_docx(_load("num-skipped.docx"))
    tree = r.hierarchy
    assert len(tree) == 1
    assert [c.level for c in tree[0].children] == [3, 2]


# ---- Phase 8: restarts not flagged ----
def test_restart_not_flagged():
    r = convert_docx(_load("num-restart.docx"))
    texts = {("".join(x.text for x in p.runs)): p for p in r.paragraphs if p.heading_level}
    assert texts["Chapter Two"].numbering_path == [2]
    assert texts["Design"].numbering_path == [2, 1]
    assert texts["Build"].numbering_path == [2, 2]
    assert r.hierarchy_issues == []


# ---- Phase 9: number formats ----
def test_number_formats_parsed():
    model = NumberingModel(
        abstract_nums={
            "0": {
                0: NumberingLevel(0, "lowerLetter", "%1.", 1),
                1: NumberingLevel(1, "upperRoman", "%1.%2.", 1),
                2: NumberingLevel(2, "decimal", "%1.%2.%3.", 1),
            }
        },
        nums={"1": "0"},
    )
    assert model.resolve_level("1", 0).num_fmt == "lowerLetter"
    assert model.resolve_level("1", 1).num_fmt == "upperRoman"
    assert model.resolve_level("1", 2).num_fmt == "decimal"
    assert model.resolve_level("9", 0) is None
    assert model.resolve_level("1", 9) is None


# ---- Phase 11: fixtures validate full pipeline ----
def test_fixture_h1_numbered():
    r = convert_docx(_load("num-h1.docx"))
    h = [p for p in r.paragraphs if p.heading_level][0]
    assert h.heading_level == 1 and h.numbering_path == [1]


def test_fixture_h1_h2_numbered():
    r = convert_docx(_load("num-h1-h2.docx"))
    assert [p.heading_level for p in r.paragraphs if p.heading_level] == [1, 2, 2]
    assert _by_text(r.paragraphs, "Section B")[0].numbering_path == [1, 2]


def test_fixture_h1_h2_h3_numbered():
    r = convert_docx(_load("num-h1-h2-h3.docx"))
    assert [p.heading_level for p in r.paragraphs if p.heading_level] == [1, 2, 3]


def test_fixture_duplicate_text_with_numbering():
    r = convert_docx(_load("num-duplicate.docx"))
    intro = _by_text(r.paragraphs, "Introduction")
    assert len(intro) == 2
    paths = sorted(p.numbering_path[0] for p in intro)
    assert paths == [1, 2]


def test_fixture_unnumbered_headings():
    r = convert_docx(_load("num-unnumbered.docx"))
    for p in r.paragraphs:
        if p.heading_level:
            assert p.num_ilvl is None and p.numbering_consistent is None


def test_fixture_custom_style_plus_numbering():
    r = convert_docx(_load("num-custom-style.docx"))
    custom = [p for p in r.paragraphs if p.style_name == "MyNumberedHeading"]
    assert custom
    assert custom[0].heading_level == 1
    assert custom[0].numbering_path == [1]


def test_fixture_mixed_numbered_unnumbered():
    r = convert_docx(_load("num-mixed.docx"))
    numbered = [p for p in r.paragraphs if p.heading_level and p.num_ilvl is not None]
    unnumbered = [p for p in r.paragraphs if p.heading_level and p.num_ilvl is None]
    assert len(numbered) == 2 and len(unnumbered) == 2


# ---- Phase 12: adversarial ----
def test_adversarial_visible_text_not_authoritative():
    r = convert_docx(_load("num-adversarial-text.docx"))
    fake = _by_text(r.paragraphs, "9. Fake Heading")[0]
    assert fake.heading_level == 2
    assert fake.num_ilvl is None
    assert fake.numbering_consistent is None


def test_adversarial_metadata_overrides_text():
    r = convert_docx(_load("num-adversarial-metadata.docx"))
    ch = _by_text(r.paragraphs, "Chapter Alpha")[0]
    assert ch.num_ilvl == 0 and ch.numbering_path == [1]


# ---- Phase 10: model design minimal (no per-paragraph numbering defs) ----
def test_model_no_full_defs_per_paragraph():
    r = convert_docx(_load("num-h1-h2-h3.docx"))
    for p in r.paragraphs:
        if p.heading_level:
            assert not isinstance(p.numbering_format, NumberingLevel)
            assert isinstance(p.numbering_path, (list, type(None)))


def _mk_para(level, text, path):
    return Paragraph(runs=[Run(text=text)], heading_level=level,
                     heading_id="x-%s" % text, numbering_path=path,
                     num_ilvl=len(path) - 1)


if __name__ == "__main__":
    for name in sorted(globals()):
        if name.startswith("test_") and callable(globals()[name]):
            globals()[name]()
            print("PASS", name)
    print("\nALL NUMBERING TESTS PASSED")
