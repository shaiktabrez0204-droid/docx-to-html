# Plan: Dynamic Page Fields in Headers/Footers

## TODOs

- [x] 1. Trace exact flow footer.xml/header.xml → adapter/ooxml_parser.py → HeaderFooter model → semantic/pipeline.py → output/html_renderer.py and identify discard point for fldChar/instrText
- [x] 2. Parse field sequences (w:fldChar begin / w:instrText / w:fldChar separate / result runs / w:fldChar end) supporting PAGE, NUMPAGES, PAGEREF with graceful fallback for unsupported fields
- [x] 3. Extend smallest appropriate model representation so fields survive OOXML → model → renderer (preserve pgNumType/start/fmt section metadata)
- [x] 4. Render semantic field spans in output/html_renderer.py (docx-page-number / docx-num-pages / docx-page-ref) without faking numbers, document placeholder limitation, fix concatenation ("February 201513") while preserving header/footer typography/alignment/links/images
- [x] 5. Add focused tests for PAGE/NUMPAGES/PAGEREF parsing, non-concatenation, ordinary footer text unchanged, roman metadata, unsupported fallback, existing headers/footers still render
- [x] 6. Real benchmark verification via ui/web.py → Playwright upload → /preview/<id> iframe on benchmark_doc/csd-thesis-template-9th-draft.docx (check: header/footer count, separated text, semantic fields, no concatenation, roman/arabic metadata, 67 headings 1.1.1, 95 lists, 1 table, 8 images, 62 links, heading-focus, console 0, overflow 0) + full pytest regression separation

## Final Verification Wave

- [x] F1. Field Parsing & Model Integrity Review
- [x] F2. Rendering & Semantic Placeholders Review (no fake numbers, no concatenation, graceful fallback)
- [x] F3. Header/Footer & Viewer Preservation Review (first/even/default, section selection, typography, images, links, heading-focus)
- [x] F4. Real CSD Browser & Regression Evidence Review

## Context
Benchmark: benchmark_doc/csd-thesis-template-9th-draft.docx (6 sections, roman front matter, arabic body)
Current bug: footer contains PAGE fields but renders as static concatenated "xi" / "February 201513" (fldChar/instrText treated as ordinary text)
Evidence: word/footer*.xml has w:fldChar + w:instrText PAGE/NUMPAGES
Product is single scrolling HTML, NOT true pagination engine - must not invent fake numbers; render semantic placeholder if no deterministic layout calculation exists.
Must reuse existing OOXML relationship/header-footer architecture, not build separate parser/pipeline.
