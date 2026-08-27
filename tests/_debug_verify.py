import zipfile, re
for n in ["img-floating", "img-reused"]:
    z = zipfile.ZipFile("tests/fixtures/%s.docx" % n)
    xml = z.read("word/document.xml").decode("utf-8")
    print(n, "anchor=", "wp:anchor" in xml, "inline=", "wp:inline" in xml)
    rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
    print("   rels targets:", re.findall(r'Target="([^"]+)"', rels))
