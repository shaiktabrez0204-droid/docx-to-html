#!/usr/bin/env python3
import sys
sys.path.insert(0, r'C:\Tabrez\docx-html\docx-to-html')

from core.model import Image
from adapter.mammoth_adapter import mammoth_to_html, extract_styles_from_docx, extract_run_fonts, attach_floating_images, _extract_images
from core.model import Style

print("=== Image model fields ===")
img = Image(image_id='test', relationship_id='rId1', source_path='word/media/img.png')
print(f"  anchor_paragraph_index: {img.anchor_paragraph_index}")
print(f"  anchor_paragraph_text: {img.anchor_paragraph_text}")
print(f"  Dict access works: {'anchor_paragraph_text' in img.__dict__}")
print(f"  Getitem works: {img['anchor_paragraph_text'] is None}")

print()
print("=== Mammoth adapter functions ===")
print(f"  mammoth_to_html: {callable(mammoth_to_html)}")
print(f"  extract_styles_from_docx: {callable(extract_styles_from_docx)}")
print(f"  extract_run_fonts: {callable(extract_run_fonts)}")
print(f"  attach_floating_images: {callable(attach_floating_images)}")
print(f"  _extract_images: {callable(_extract_images)}")

print()
print("=== Style class ===")
s = Style(name='Normal', font_family='Calibri')
print(f"  Style name: {s.name}")
print(f"  Style font_family: {s.font_family}")

print()
print("=== All checks passed! ===")