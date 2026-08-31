# Image Extraction Failure — Debug & Fix Plan

## Objective
Fix real DOCX image extraction where images are missing in HTML preview despite being present in word/media/*.

## Success Condition
- Real DOCX with inline, floating, table, header, footer, multiple, PNG/JPEG/GIF → all images extracted as Image objects → data: URLs → <img> visible in Chromium
- No regression on existing 18 image tests

## Critical Unknowns (resolved via reproduction)
- ZIP media exists but Image objects 0 for VML and AlternateContent paths
- Reproduction shows:
  - ZIP media 2, XML pict present, Image objects 1 (VML missing)
  - ZIP media 2, AlternateContent present, Image objects 1 (Choice drawing missing)
  => Breakpoint = B (XML image detection) + fallback to C

## Classification
- A ZIP/media extraction: PASS (reads bytes correctly)
- B XML image detection: FAIL for w:pict/v:shape/v:imagedata and for mc:AlternateContent wrapped wp:inline/anchor
- C relationship resolution: PASS when B passes
- D Image model: PASS
- E pipeline propagation: PASS (but B blocks it)
- F HTML rendering: PASS when assets present
- G data URL: PASS
- H browser: PASS when data URL valid

## Root Cause (confirmed)
`adapter/ooxml_parser.py::_extract_images_from_r` only handles:
```
w:r -> w:drawing -> wp:inline or wp:anchor -> a:blip/@r:embed
```
Missing:
1. `w:pict` → `v:shape`/`v:imagedata/@r:id` (VML legacy path, still emitted by Word for some images, watermarks, compat headers/footers, and fallback inside AlternateContent)
2. `mc:AlternateContent` wrappers: `w:drawing` inside `mc:Choice` or fallback `w:pict` — direct `find(w:drawing)` returns None when drawing is deeper than one level

Existing fixtures are all pure `w:drawing` without AlternateContent/VML, so tests pass but real docs fail.

## Fix (smallest production change)
Modify single file: `adapter/ooxml_parser.py`

### Change 1: Deep search for drawings
- Replace `r_elem.find(w:drawing)` with recursive search for all `wp:inline`/`wp:anchor` under the run (`.//wp:inline`, `.//wp:anchor`) regardless of AlternateContent depth.
- Deduplicate by not double-counting if same inline already handled via drawing path.
- Keep existing `_parse_drawing` for extent/alt/blip logic unchanged.

### Change 2: Add VML pict extraction
- Add namespace `NS_V = "urn:schemas-microsoft-com:vml"` and `NS_WPict = already via W`
- Add helper `_extract_vml_images(r_elem) -> List[Image]` that finds all `v:imagedata` under `w:pict`:
  - `r:id` (and fallback `r:embed`) → rid
  - style width/height (e.g. `width:100pt;height:50pt`) → px via pt->px (pt*96/72)
  - Also handle `o:imagedata`? Check but primary is v:imagedata
  - Resolve asset via existing `_resolve_asset`
  - Create Image with wrap_type="inline", preserve alt from shape title if available
- Call this helper from `_extract_images_from_r` and extend results.
- Ensure VML images get valid ImageAsset dedup same as drawing images.

### No changes to
- `core/model.py` (reuse Image)
- `semantic/pipeline.py` (already passes assets)
- `output/html_renderer.py` (already renders inline Image; no css change)
- UI

## Preservation
- Inline, floating, paragraph-relative, table-cell, header/footer, wrapSquare/Tight/Through/topAndBottom, polygons, sizing, positioning, section isolation, data-anchor must not regress.
- All existing image tests must still pass; floating geometry unchanged.

## Test Cases (must verify)
1. Normal inline `<w:drawing><wp:inline>` → PASS before, stays PASS
2. Floating `<wp:anchor>` → PASS
3. Table cell image → PASS (via same _parse_paragraph)
4. Header image → PASS
5. Footer image → PASS
6. Multiple images → PASS
7. PNG/JPEG/GIF → PASS
8. **New** VML pict image → should now PASS (was FAIL)
9. **New** AlternateContent wrapped drawing → should now PASS (was FAIL)
10. **New** VML inside AlternateContent fallback → PASS

## Regression Test
Add focused test using real DOCX fixture containing VML + AlternateContent:
- Verify media discovered, rel resolved, Image objects produced, HTML <img> count correct, src valid, mime correct, no duplicates.

## Browser Verification
- Use Chromium via existing `test_image_pipeline_playwright.py` pattern + new fixture preview to verify naturalWidth>0.

## Files Changed
- `adapter/ooxml_parser.py` only (plus new test fixture creation if needed)

## Execution Steps
1. Patch `_extract_images_from_r` and add VML helper
2. Re-run reproduction scripts (tmp_vml, tmp_alt, complex_real) to confirm 2→2 and 1→2 fixes
3. Run full pytest image suite
4. Run browser verification via Playwright on complex doc
5. Add regression test
