"""Generate a real DOCX fixture for font-size fidelity verification.

Uses raw OOXML (not python-docx) so we have exact, inspectable control over
w:sz / w:szCs / w:rStyle / w:pStyle / w:hyperlink / tables / numbering. This is
the ground-truth document the browser + model tests assert against.

Every run carries a unique text marker so tests can locate it and reconstruct
the effective size from the rendered computed px.

Effective-size matrix (OOXML w:sz is half-points):
  SZ19=9.5pt  SZ21=10.5pt SZ22=11pt  SZ23=11.5pt SZ24=12pt  SZ27=13.5pt
  SZ28=14pt   SZ32=16pt   SZ36=18pt  SZ48=24pt  SZ56=28pt
  PLAIN       -> docDefaults 11pt (no direct, no style)
  PARA12      -> paragraph style "Para12" 12pt (inherited)
  OVERRIDE16  -> direct 16pt inside Para12 paragraph (direct wins)
  CHAR135     -> character style "Char135" 13.5pt (inherited)
  CHARBOLD12  -> char style "CharBold" (no size) inside Para12 -> inherits 12pt
  HEADING18   -> heading style "Heading1" 18pt (inherited)
  LIST11      -> list item, docDefaults 11pt
  TABLE11     -> table cell run, docDefaults 11pt
  HYPER11     -> hyperlink run, docDefaults 11pt
  MIXED       -> one paragraph mixing SZ19 + SZ28 + SZ48
"""

import os
import zipfile

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")
OUT = os.path.join(BASE, "font-fidelity.docx")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# ---- styles.xml -----------------------------------------------------------
STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="%s">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:color w:val="000000"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr><w:spacing w:after="160" w:before="160" w:line="259" w:lineRule="auto"/></w:pPr>
    </w:pPrDefault>
  </w:docDefaults>

  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>

  <w:style w:type="paragraph" w:styleId="Para12">
    <w:name w:val="Para 12"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading1Char"/>
    <w:uiPriority w:val="9"/>
    <w:outlineLvl w:val="0"/>
    <w:rPr><w:sz w:val="36"/><w:szCs w:val="36"/><w:b/></w:rPr>
  </w:style>

  <w:style w:type="character" w:styleId="Char135">
    <w:name w:val="Char 135"/>
    <w:rPr><w:sz w:val="27"/><w:szCs w:val="27"/></w:rPr>
  </w:style>

  <w:style w:type="character" w:styleId="CharBold">
    <w:name w:val="Char Bold"/>
    <w:rPr><w:b/></w:rPr>
  </w:style>
</w:styles>
""" % W

# ---- numbering.xml (one decimal list) --------------------------------------
NUMBERING = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="%s">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%%1."/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
""" % W

# ---- document.xml ----------------------------------------------------------
def run(text, sz=None, rstyle=None, bold=False, hyperlink=False):
    rpr = ""
    if rstyle:
        rpr += '<w:rStyle w:val="%s"/>' % rstyle
    if bold:
        rpr += "<w:b/>"
    if sz is not None:
        rpr += '<w:sz w:val="%d"/>' % sz
    rpr_xml = "<w:rPr>%s</w:rPr>" % rpr if rpr else ""
    r_xml = "<w:r>%s<w:t xml:space=\"preserve\">%s</w:t></w:r>" % (rpr_xml, text)
    return r_xml

def para(runs_xml, pstyle=None, numpr=False):
    ppr = ""
    if pstyle or numpr:
        parts = ""
        if pstyle:
            parts += '<w:pStyle w:val="%s"/>' % pstyle
        if numpr:
            parts += '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
        ppr = "<w:pPr>%s</w:pPr>" % parts
    return "<w:p>%s%s</w:p>" % (ppr, runs_xml)

# Direct-size runs (one per size) for explicit verification.
direct = "".join(
    para(run("SZ%d " % sz, sz=sz)) for sz in (19, 21, 22, 23, 24, 27, 28, 32, 36, 48, 56)
)

# Plain run -> docDefaults 11pt.
plain = para(run("PLAIN "))

# Paragraph-style inheritance (Para12 = 12pt).
para12 = para(run("PARA12 "), pstyle="Para12")

# Direct override of style (Para12 paragraph, run direct 16pt -> 16pt wins).
override = para(run("OVERRIDE16 ", sz=32), pstyle="Para12")

# Character style inheritance (Char135 = 13.5pt).
char135 = para(run("CHAR135 ", rstyle="Char135"))

# Char style with NO size (CharBold) inside Para12 -> inherits 12pt.
charbold = para(run("CHARBOLD12 ", rstyle="CharBold"), pstyle="Para12")

# Heading with inherited size (Heading1 = 18pt). Must NOT be the first heading
# or it would be consumed by the UI title bar and skipped from the body.
heading = para(run("HEADING18 "), pstyle="Heading1")

# List item with inherited size (docDefaults 11pt).
list_item = para(run("LIST11 "), numpr=True)

# Mixed sizes in one paragraph.
mixed = para(run("MX19 ", sz=19) + run("MX28 ", sz=28) + run("MX48", sz=48))

# Table with an inherited-size cell run (docDefaults 11pt).
table = (
    "<w:tbl>"
    "<w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblBorders>"
    "<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    "<w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    "<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    "<w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    "</w:tblBorders></w:tblPr>"
    "<w:tblGrid><w:gridCol w:w=\"4620\"/></w:tblGrid>"
    "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"4620\" w:type=\"dxa\"/></w:tcPr>"
    "<w:p>%s</w:p>"
    "</w:tc></w:tr>"
    "</w:tbl>"
) % run("TABLE11 ")

# Hyperlink run with inherited size (docDefaults 11pt).
hyperlink = para('<w:hyperlink w:rel="rId3"><w:r><w:t xml:space="preserve">HYPER11 </w:t></w:r></w:hyperlink>')

# Leading heading (becomes the UI title bar, rendered outside .docx-content).
lead = para(run("DOCTYPE "), pstyle="Heading1")

body = (
    lead + direct + plain + para12 + override + char135 + charbold + heading
    + list_item + mixed + table + hyperlink
)

DOCUMENT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="%s" xmlns:r="%s">
  <w:body>
    %s
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>
""" % (W, R, body)

# ---- package plumbing ------------------------------------------------------
CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" Type="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" Type="application/xml"/>
  <Override PartName="/word/document.xml" Type="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" Type="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" Type="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"""


def build(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", DOCUMENT)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", NUMBERING)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
    print("wrote", path)


if __name__ == "__main__":
    os.makedirs(BASE, exist_ok=True)
    build(OUT)
