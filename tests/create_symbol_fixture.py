"""Create focused fixture for symbol preservation fix."""
import os, zipfile, io
from pathlib import Path
from PIL import Image as PILImage

PROJECT_ROOT = Path(__file__).parent.parent
FIX = PROJECT_ROOT / "tests" / "fixtures"
OUT = FIX / "symbol-preservation.docx"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def make_png():
    img = PILImage.new("RGB", (120,80), (200,30,30))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

PNG_BYTES = make_png()

styles = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:outlineLvl w:val="0"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
</w:styles>"""

# Build document body
# Helper to create image inline xml
image_inline_xml = """<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1143000" cy="762000"/><wp:docPr id="1" name="Picture 1" descr="test inline"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId_img1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="762000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"""

floating_anchor_xml = """<w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" distT="0" distB="0" distL="0" distR="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" simplePos="0" relativeHeight="251658240"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>1000000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>1000000</wp:posOffset></wp:positionV><wp:extent cx="1143000" cy="762000"/><wp:docPr id="2" name="Picture 2" descr="floating"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="image2.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId_img2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="762000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic><wp:wrapSquare wrapText="bothSides"/></wp:anchor></w:drawing>"""

body_parts = []
body_parts.append('<w:p><w:r><w:t xml:space="preserve">Normal text paragraph.</w:t></w:r></w:p>')
body_parts.append('<w:p><w:r><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr><w:sym w:font="Symbol" w:char="F0D7"/></w:r><w:r><w:t xml:space="preserve"> after Symbol</w:t></w:r></w:p>')
body_parts.append('<w:p><w:r><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr><w:sym w:font="Wingdings" w:char="F028"/></w:r><w:r><w:t xml:space="preserve"> after Wingdings</w:t></w:r></w:p>')
body_parts.append('<w:p><w:r><w:t>no</w:t></w:r><w:r><w:noBreakHyphen/></w:r><w:r><w:t>break</w:t></w:r></w:p>')
body_parts.append('<w:p><w:r><w:t>soft</w:t></w:r><w:r><w:softHyphen/></w:r><w:r><w:t>hyphen</w:t></w:r></w:p>')
body_parts.append('<w:p><w:r><w:t xml:space="preserve">Unicode symbols: \u2192 \u2764 \u263A \u00D7</w:t></w:r></w:p>')
# eastAsia font fallback test: only eastAsia, no ascii
body_parts.append('<w:p><w:r><w:rPr><w:rFonts w:eastAsia="MS Gothic"/></w:rPr><w:t>eastAsia font test</w:t></w:r></w:p>')
# inline image
body_parts.append(f'<w:p><w:r><w:t>Before image </w:t></w:r><w:r>{image_inline_xml}</w:r><w:r><w:t> after image.</w:t></w:r></w:p>')
# floating image
body_parts.append(f'<w:p><w:r>{floating_anchor_xml}</w:r><w:r><w:t>Paragraph with floating image.</w:t></w:r></w:p>')
# another normal
body_parts.append('<w:p><w:r><w:t>Final paragraph after images.</w:t></w:r></w:p>')

body_inner = "".join(body_parts)

document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    {body_inner}
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>"""

content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" Type="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" Type="application/xml"/>
  <Default Extension="png" Type="image/png"/>
  <Override PartName="/word/document.xml" Type="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" Type="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId_img1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId_img2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
</Relationships>"""

with zipfile.ZipFile(str(OUT), "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/_rels/document.xml.rels", doc_rels)
    z.writestr("word/media/image1.png", PNG_BYTES)
    z.writestr("word/media/image2.png", PNG_BYTES)

print("WROTE", OUT)
# verify
with zipfile.ZipFile(str(OUT)) as z:
    xml = z.read("word/document.xml").decode()
    assert '<w:sym w:font="Symbol" w:char="F0D7"/>' in xml
    assert '<w:noBreakHyphen/>' in xml
    assert '<w:softHyphen/>' in xml
    assert 'w:eastAsia="MS Gothic"' in xml
    print("Verified symbols present in XML")
