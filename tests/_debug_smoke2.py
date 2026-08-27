import sys, zipfile, re
sys.path.insert(0, ".")
from adapter.ooxml_parser import OoxmlParser

for name in ["img-reused", "img-floating"]:
    p = OoxmlParser("tests/fixtures/%s.docx" % name)
    paras = p.parse_paragraphs()
    assets = p.get_image_assets()
    total = sum(len(pr.images) for pr in paras)
    print("==", name, "paras=%d placements=%d assets=%d" % (len(paras), total, len(assets)))
    for pr in paras:
        for img in pr.images:
            a = assets.get(img.source_path)
            ok = a is not None and a.data is not None
            print("   %s rid=%s src=%s %dx%d alt=%r wrap=%s bytes=%s" % (
                img.image_id, img.relationship_id, img.source_path,
                img.width, img.height, img.alt_text, img.wrap_type,
                (len(a.data) if ok else None)))

print("\n--- floating document.xml structure ---")
z = zipfile.ZipFile("tests/fixtures/img-floating.docx")
xml = z.read("word/document.xml").decode("utf-8")
print("contains 'anchor':", "anchor" in xml)
print("contains 'wp:inline':", "wp:inline" in xml)
print("contains 'ns0:inline':", "ns0:inline" in xml)
print("contains 'blip':", "blip" in xml)
m = re.search(r"<(.*?:)?(inline|anchor)[^>]*>", xml)
print("first drawing root tag:", m.group(0)[:120] if m else None)
