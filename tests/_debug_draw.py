import zipfile, re
z = zipfile.ZipFile("tests/fixtures/img-floating.docx")
xml = z.read("word/document.xml").decode("utf-8")
i = xml.find("blip")
print("around blip:", xml[i-400:i+80])
print("\n--- all xmlns decls ---")
for m in re.finditer(r'xmlns:([^=]+)="([^"]+)"', xml):
    print(m.group(1), "=", m.group(2))
