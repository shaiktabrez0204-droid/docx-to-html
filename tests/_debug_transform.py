import os, zipfile, xml.etree.ElementTree as ET
FIX = "tests/fixtures"
out = os.path.join(FIX, "img-floating.docx")
with zipfile.ZipFile(out, "r") as z:
    xml = z.read("word/document.xml").decode("utf-8")
ns = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
root = ET.fromstring(xml)
inline = root.find(".//wp:inline", ns)
print("inline is not None:", inline is not None)
if inline is not None:
    print("inline tag:", inline.tag)
    drawing_parent = None
    for d_ in root.iter(ns["w"] + "drawing"):
        if inline in list(d_):
            drawing_parent = d_
            break
    print("drawing_parent found:", drawing_parent is not None)
