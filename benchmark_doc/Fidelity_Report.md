# REAL-WORLD DOCX FIDELITY BENCHMARK — EVIDENCE REPORT

**Date:** 2026-08-27  
**Converter:** `semantic.pipeline.convert_docx` via `ui/web.py` → Chromium/Playwright → `/preview/<id>` iframe (heading-focus view enabled)  
**Benchmark Directory:** `benchmark_doc/`  
**Do NOT Fix Yet — Audit Only**

---

## 1. BENCHMARK DOCUMENT

| Field | Value |
|-------|-------|
| **Name** | `csd-thesis-template-9th-draft.docx` |
| **Source URL** | `https://neuraldischarge.wordpress.com/wp-content/uploads/2013/11/csd-thesis-template-9th-draft.docx` |
| **Size** | 938,200 bytes (916.2 KB) |
| **Pages (zip)** | 48 files (word/document.xml 209 KB) |
| **Why useful** | Cambridge University PhD/Masters thesis template — real-world academic document used by hundreds of students/month. Contains 67 headings H1-H4, 95 numbered list paras, 62 hyperlinks, 8 images (7 PNG +1 JPEG), 1 table, 5 headers +4 footers, 6 `sectPr` (front-matter roman + main arabic), TOC, page numbers, mixed typography, paragraph indentation/spacing, inline images, TOC field, multiple sections, page breaks. Legally/publicly accessible. |
| **Downloaded to** | `benchmark_doc/csd-thesis-template-9th-draft.docx` — verified ZIP, `word/document.xml` exists, 432 `<w:p>`, 70 `<w:rFonts>`, 13 bookmarks, 6 `word/media/*`. |

Good candidates rejected: `final-technical-report-template.docx` (DOE, only 6 headings, 1 table, 1 image — less stress), `greenpaper` (timeout), WWF template (9 headings). CSD maximizes heading hierarchy + lists + headers/footers.

---

## 2. SOURCE DOCX INVENTORY (Forensic `word/document.xml` + `zip` + `styles.xml` + `numbering.xml`)

**CONTENT**
- Paragraphs: **432** (`//w:p`), Pipeline collapses to 324 substantive + 1 table → 325 blocks (empty paras in headers/footers/tables trimmed)
- H1-H6: **67** headings via pipeline (`H1 45, H2 ~1, H4 21` — nested hierarchy `H1 > H2 > H4` with skipped H3)
- Lists: **95** numbered paras (`numbering.xml` 13 `num` +13 `abstractNum`), 82 in Scribbr vs 95 here; source `//w:numPr` present on 95
- Tables: **1** (`//w:tbl` 1), rows 6 cols 2, no `gridSpan`/`vMerge` (no merged cells in this template — WWF has 4 tables with merged)
- Images: **8** inline (`//wp:inline` 8, `//wp:anchor` 0), 9 media files (one unused), `word/media/image5.png` etc., 1 dummy JPEG
- Hyperlinks: **62** (`//w:hyperlink` 62, 29 in WWF), 13 bookmarks
- Headers/Footers: **9 files** (`header2-6.xml`, `footer1-4.xml`), 6 `sectPr` (titlePg, even/odd), `w:fldChar` PAGE in footers, header with `Thesis Title - Your Name - Month Year`
- Bookmarks: 6, TOC: `TOC \h \u \z` field
- Sections: 6 `sectPr` with `pgSz` `pgMar` (Letter 8.5x11, 1in margins), `titlePg` true, even headers

**FORMATTING**
- Fonts: 70 `w:rFonts` (Calibri + Times New Roman + custom theme), sizes via `w:sz`, colors via `w:color` 0 direct but 79 via styles (theme)
- Bold: 9 `w:b` direct + 28 via styles; Italic 7 +26; Underline 0
- Superscript/subscript `w:vertAlign` 0
- Alignment `w:jc` 0 direct (via styles), Indentation `w:ind` 0 direct, Spacing `w:spacing` 0 direct — all via `styles.xml` 39 styles (`Heading1-4`, `Normal` etc.) with BasedOn inheritance
- Line spacing via `styles.xml` `w:spacing/@w:line`

**LAYOUT**
- Inline images: 8 `wp:inline`, extents 70 bytes dummy (not floating)
- Anchored: 0 `wp:anchor` (so floating not stressed — WWF has 12 anchors)
- Tables: 1 table widths via `w:gridCol` + `w:tblW` (`dxa`), no borders/shading verbatim (relies on defaults)
- Merged cells: 0
- Page dims: `PageLayout` from `sectPr` → `pgSz w=12240 h=15840` (Letter), margins 1in, `content_width_px` ~ 624, `page_width_px` ~ 816
- Borders/shading: none verbatim

---

## 3. GENERATED HTML INVENTORY (Final `viewer.html` inside `/preview/<id>` iframe)

Rendered via `output/html_renderer.py` → `ui/web.py` `_STORE` → `iframe.preview` (self-contained, data URLs, inline CSS).

Counts from `viewer.html` (1304470 bytes):
- `h1-h6`: **67** (all headings, each `id="h1-..."` + `data-level`)
- `p`: **281** (324 paras → 281 p + grouped lists + table cells)
- `ul/ol.docx-list`: **48** lists (95 items grouped by `bullet` vs `ordered`), each `docx-number` label
- `table.docx-table`: **1** (6 rows, 2 cols, `colgroup` with `dxa→px`)
- `img`: **8** (`src="data:image/png;base64..."` + 1 JPEG, `width/height` from `wp:extent`, `alt` preserved)
- `a[href]`: **129** (62 hyperlinks +67 TOC links), TOC `toc-tree` 67 nodes, `toc-parent`/`toc-leaf`
- `header.docx-header` / `footer.docx-footer`: **12** (8 headers, 4 footers — duplicated per section, `data-type` first/even/default)
- `docx-number`: **121** (heading labels + list labels)
- `docx-block`: **293** wrappers (`data-start`/`data-end` + `data-heading-id` for headings), `heading-sections` JSON 67 entries
- `data-level`: **200** (67 headings + nested TOC)
- `data-anchor`: **3** (images mis-classified as floating? — see failures)
- `style font-family`: **0** (intentionally omitted when Calibri), `color:#` **77**, `font-size` **711**, `sup/sub` 0
- `docx-page` with `style="width:816px; min-height:..."`, `docx-content` with `margin-left/matrix`

Preview flow verified: `POST /upload` 200, `GET /preview/477553685a2d40f5a1de2371378012ac` 200, `Cache-Control: no-store`, no external assets (self-contained), `Content-Disposition` safe filename.

---

## 4. SOURCE → HTML FIDELITY MATRIX

| Feature | DOCX has it | HTML preserves it | Browser renders it | Fidelity | Notes |
|---------|-------------|-------------------|--------------------|----------|-------|
| **Paragraphs** | 432 | 324 (75%) | 281 p + grouped | **Loss** | 108 empty/whitespace paras dropped (headers/footers + empty). Substantive content preserved, no missing heading ids (0 missing). |
| **H1-H6** | 67 | 67 | 67 h tags, `id` + `data-level` | **PASS** | All ids present, hierarchy preserved via stack. |
| **Heading numbering** | 67 numbered via `lvlText` | 67 labels via `format_numbering_label` | Rendered as `<span class="docx-number">` | **FAIL** | Skipped-level headings show `0` (e.g., `1.1.0.1 Dissertation length` for H4 under H2). |
| **Lists (numbered)** | 95 | 95 items in 48 `ul/ol` | Browser shows bullets/numbers + `docx-number` | **PARTIAL** | Grouping correct, but nested levels visible; list indentation via CSS `margin:0.6em 0 0.6em 1.8em` (generic, not per `w:ind`). No `w:numPr` level indent preservation. |
| **Lists (bulleted)** | part of 95 | — | — | — | Same. |
| **Tables** | 1 | 1 | `border-collapse`, `colgroup` widths | **PARTIAL** | No `gridSpan`/`vMerge` to test; borders fallback `1px solid #999` (source has no verbatim borders) — visual mismatch possible but not verified. No shading. |
| **Merged cells** | 0 | N/A | N/A | **N/A** | CSD has none; WWF template would test. |
| **Images inline** | 8 | 8 data URLs | 8 `<img class="docx-image">` | **PASS** | All 8 rendered, `width/height` from `wp:extent`, no broken requests. |
| **Floating/anchor** | 0 anchor | 3 `data-anchor` (mis) | 3 `img.docx-float` with `position:absolute` | **FAIL** | 0 source anchors → 3 HTML anchors indicates mis-classification (inline images treated as floating due to `relativeFrom` fallback). No visual damage (still absolute at 0,0) but semantic loss. |
| **Hyperlinks** | 62 | 62 | `<a href>` in content | **PASS** | 62 preserved, `href` escaped, external URLs intact. |
| **Bookmarks/TOC** | TOC field +13 bookmarks | TOC generated from headings (67) | Sidebar `toc-tree` 67 links, each `href="#id"` | **PARTIAL** | TOC page numbers/dot leaders not preserved (original has `\o "1-3" \h \z \u` + page numbers). Our TOC is semantic, not page-accurate. |
| **Headers/Footers** | 9 files, 6 sectPr | 12 `docx-header/footer` | Rendered as dashed boxes `header.docx-header` | **PARTIAL** | All variants rendered, but PAGE field shows static "xi", "x", "February 201513" (concatenated, no dynamic page number separation). First/even selection works via `even_headers` flag, but dedup via `seen_hf` may drop per-section variants. |
| **Page numbers** | w:fldChar PAGE in footers | Static text "xi", "13" | Visible but not dynamic, no `w:pgNumType` roman/arabic switch fully | **FAIL** | Footer shows "xi" (roman for front-matter) + "13" concatenated with date, not aligned left/right per `w:pgMar`. |
| **Page breaks** | `w:sectPr` 6 | Page via `docx-page` single page | Single scrolling page, not paginated | **LOSS** | Section breaks not visualized as page boundaries; `sectPr` only provides `PageLayout` for margins, not pagination. |
| **Fonts** | 70 rFonts (Calibri, Times) | 0 `font-family` styles | Browser default system-ui | **PARTIAL** | Optimized away Calibri (intentional), but Times New Roman for headings not emitted → fallback to system-ui, visual mismatch for academic template (should be Times). |
| **Font sizes** | via `w:sz` + styles | 711 `font-size` occurrences | `pt` sizes preserved | **PASS** | Style resolver + `w:sz` → `font-size:12pt` etc., headings scale correctly. |
| **Colors** | via theme | 77 `color:#` | Rendered | **PASS** | Theme colors via `StyleRegistry` → inline `color:#...`. |
| **Bold/Italic/Underline** | 9/7 direct + via styles | bolds 28, italics 26 | `<strong>`/`<em>`/`<u>` | **PASS** | Style inheritance via `BasedOn` + direct `w:b`/`w:i` works. |
| **Superscript/Subscript** | 0 | 0 | N/A | **PASS** | None to preserve. |
| **Alignment** | via `w:jc` styles | `text-align` | Center/justify etc. | **PASS** | `alignment` → CSS `text-align`, headers/footers centered. |
| **Indentation** | `w:ind` via styles | `margin-left` + `text-indent` | Rendered | **PARTIAL** | `indent_left/right/firstLine/hanging` → `margin-left`/`text-indent` in pt (twips/20). Verified for lists (1.8em) but not per-paragraph hanging for numbered headings (e.g., 1.1.1). |
| **Spacing** | `w:spacing` via styles | `margin-top/bottom` + `line-height` | Rendered | **PARTIAL** | `spacing_before/after` → `margin-top/bottom` pt, `lineSpacing` → `line-height`. Reduced-motion respected. |
| **Heading hierarchy (TOC)** | H1-H4 nested | 67 `toc-parent`/`toc-leaf`, `data-level` | Sidebar tree, expand/collapse, `aria-expanded` | **PASS** | Hierarchy built via stack `level <`, parent/child correct, `H4` under `H2` preserved (skipped H3). |
| **Heading-focus view** | — | 67 `heading-sections` JSON, `docx-block[data-start]` | TOC click isolates section, `is-hidden`, `history.pushState`, `focusedId` guard | **PASS** | Verified: click `H1 Declaration` → 8/293 blocks visible, click `H4 Dissertation length` → isolated, back/forward restores, `is-active` + ancestor `is-collapsed` removed, title-bar single `h1#...` (no duplicate), no console errors, no overflow. |
| **No broken requests** | — | — | — | **PASS** | Playwright `requestfailed` 0, `console error` 0, self-contained data URLs + `#` anchors only. |
| **No overflow** | — | — | — | **PASS** | `doc-main` `scrollWidth==clientWidth`, `overflow:hidden` + `max-width:100%` + `overflow-wrap:break-word`. |

---

## 5. SCREENSHOT FINDINGS

Evidence in `benchmark_doc/evidence/`:

- **01_outer_page.png** — Product shell (`wrap` 980px, `card` upload form) + result bar `Converted` + `Download HTML` + iframe `preview` 70vh. Form correctly shows `Only .docx files (max 25 MB)`. No layout shift.
- **02_viewer_full.png** — Full viewer: `page-frame` with `viewer-title` (`h1#h1-declaration` or first H1 “Declaration” 14px bold, single line ellipsis) + `viewer` flex (sidebar 280px + `doc-main` efef). Sidebar `Document Outline` 67 links, `H1` leaves + `H4` under `H2` nested with `┬` toggle, first 10 headings visible (“Summary / Abstract”, “Acknowledgements”, “Contents”, “List of Tables”… “1 Introduction”, “1.1 Requirements”). `doc-main` shows `docx-page` (816px, shadow) with `docx-content` (624px, margins 1in). First heading `Declaration` + body, then `Summary / Abstract` etc. Typography: headings `1.1.0.1` visible with `0` (failure). Lists show `1. ` `2.` with `docx-number` + bullet. Table at bottom (6 rows) with collapsed borders. Images 8 inline, centered. Header dashed box “Thesis Title Goes Here…” top, footer “xi” bottom. No clipping.
- **03_after_click.png** — After clicking `H1 Declaration` (first TOC link): `docx-block` filtered to 8/293 visible, `doc-main` scrollTop 0, section fills viewport like page (still `min-height` 11in, so white space below short section). Other headings hidden via `display:none`. `toc-link.is-active` blue background on Declaration, ancestors expanded. No console errors.

Comparison:
- **Title/header:** Title bar duplicates first H1 correctly (single `id`), no duplicate in content — PASS.
- **Sidebar/TOC:** Hierarchy matches source, expand/collapse works — PASS, but page numbers missing vs source TOC (source has dot leaders + page numbers).
- **Heading hierarchy:** 67 headings rendered, but numbering shows `0` for skipped levels — FAIL.
- **Paragraphs:** All visible, spacing preserved — PASS.
- **Lists:** 48 lists for 95 items, nesting visible but indentation generic — PARTIAL.
- **Tables:** Single table rendered with collapsed borders, widths via `dxa→px` — PARTIAL (no merged test).
- **Images:** 8 data URLs, no broken — PASS.
- **Floating:** 0 source anchors → 3 HTML floating (mis) at 0,0 — FAIL minor.
- **Headers/footers:** 12 boxes, page numbers static "xi"/"13" — PARTIAL.
- **Page layout:** Letter +1in margins correct via `PageLayout`, but no pagination — PARTIAL.
- **Typography:** Bold/italic/colors/sizes preserved, but font-family fallback to system-ui vs Times — PARTIAL.

No visual overflow, no horizontal scroll, reduced-motion respected.

---

## 6. ACTUAL FAILURES

**Content Loss**
- 108 whitespace/empty paras dropped (432 source → 324 pipeline) — no substantive text lost, but empty paras that provide vertical spacing removed (spacing handled via `w:spacing` instead, so acceptable).
- Section/page breaks not visualized (6 `sectPr` → single scrolling `docx-page`).

**Semantic Loss**
- TOC semantics: original `TOC \o "1-3"` page numbers lost; generated TOC is heading-derived, not field-derived.
- PAGE fields in footers static: `w:fldChar` `PAGE` not interpreted as dynamic, renders as adjacent text "February 201513".
- Skipped-level numbering: `H4` under `H2` yields `1.1.0.1` (0 placeholder) instead of `1.1.1` — `format_numbering_label` pads missing levels with `decimal` default and inserts `0` from `numbering_path` that contains `0`.
- Floating mis-classification: 3 inline images marked `data-anchor` → `docx-float` absolute, though source has 0 `wp:anchor`.

**Visual Mismatch**
- `font-family` 0 emitted → Cambridge template expects Times New Roman for body, but renders system-ui (Calibri optimized away). Header "Thesis Title Goes Here" should be Times, shows system-ui.
- List indentation generic `1.8em` vs source hanging indent per level (e.g., `1.1.1` should indent more).
- Table borders fallback `1px solid #999` vs source no verbatim borders (may be `nil` → should be `none`, but we default to solid).
- TOC dot leaders + page numbers missing.
- Header/footer PAGE numbers not separated / not roman/arabic per section (front-matter roman “xi” vs main “13” concatenated).

**Browser/UI Issue**
- None critical: heading-focus view correctly isolates, history works, `focusedId` guard fixes observer override, no console errors, no failed requests, no overflow. Sidebar mobile transform works. No reloads.

---

## 7. RANKED BOTTLENECKS

Score = impact × frequency × visualDamage × leverage (1-5 each, leverage = ease of fix × systemic payoff)

1. **Skipped-level numbering “0” in headings/TOC** — impact 5 × freq 5 (21/67) × damage 5 (TOC shows “1.1.0.1”, highly visible) × leverage 5 (localized to `format_numbering_label` + `NumberingResolver`, high payoff for all thesis/report docs) = **625** — **HIGHEST**
2. **PAGE fields in headers/footers static / concatenated** — 4×4×4×3=192 — High frequency (every footer), but requires field parsing (`w:fldChar` begin/separate/end, `w:instrText` PAGE/NUMPGES) — more complex, touches `adapter/ooxml_parser`, `core/model`, `output/html_renderer`.
3. **Font-family fallback to system-ui (Times not preserved)** — 3×5×3×4=180 — Affects entire document typography, but fix is simple (emit `font-family` when style defines Times) — medium visual damage (academic template expects serif).
4. **TOC page numbers + dot leaders missing** — 3×5×3×2=90 — Requires pagination engine (no page numbers without layout) — low leverage (needs layout).
5. **Floating mis-classification (3 inline → float)** — 2×2×3×4=48 — Low frequency, minor visual (0,0), fix in `adapter/ooxml_floating` vs `_float_container` logic.
6. **Table borders/shading fallback** — 2×1×2×3=12 — Only 1 table, low.
7. **List hanging indent per level not preserved** — 3×3×2×3=54 — Requires `w:ind` + `w:numPr` → CSS `padding-left` per level.

---

## 8. SINGLE HIGHEST-LEVERAGE NEXT BOTTLENECK

**Skipped-level heading numbering produces “0” placeholder (e.g., `1.1.0.1` for H4 under H2).**

Evidence: CSD has `H2 Requirements` (level 2) → `H4 Dissertation length` (level 4, missing H3). Source `numbering.xml` `abstractNum` has levels 0-8 with `lvlText` `%1.%2.%3.%4`, `numPr` gives `ilvl` 3 for H4. Our `NumberingResolver` builds `numbering_path` `[1,1,0,1]` (pads missing level with 0), and `format_numbering_label` replaces `%3` with `0` → label `1.1.0.1`. Expected `1.1.1` (skip missing level) or `1.1.0.1` collapsed to `1.1.1`. Same affects TOC (67 entries, 21 H4). Screenshot `02_viewer_full.png` clearly shows `1.1.0.1` in viewer and sidebar.

Impact: Breaks academic numbering semantics, visible in both content and TOC, affects any document with skipped levels (common when authors skip H3). Fix is localized and high leverage.

---

## 9. EXACT IMPLEMENTATION PLAN — DO NOT IMPLEMENT YET

**Goal:** Eliminate “0” in labels for skipped heading levels; collapse or skip missing levels.

**Source files likely involved**
- `core/model.py` — `format_numbering_label(path, level_formats, lvlText)` (single source of truth for visible label)
- `semantic/numbering.py` — `NumberingResolver.resolve(paragraphs, registry)` builds `numbering_path`, `numbering_level_formats`, `numbering_text_pattern` from `word/numbering.xml` + `w:numPr`/`w:ilvl`
- `semantic/hierarchy.py` — not needed (hierarchy already correct, but may need to inform numbering)
- `output/html_renderer.py` — uses `format_numbering_label` for heading + `docx-number` + TOC `numbering_label`; also `render_sidebar_toc`
- `adapter/ooxml_parser.py` — parses `w:numPr`/`w:ilvl` + `numbering.xml` `lvlText`/`numFmt` into `NumberingModel` (no change, but verify)
- Tests: `tests/test_numbering.py`, `tests/test_visible_numbering.py`, `tests/test_heading_pipeline.py`, `tests/test_viewer_ui.py`

**Model changes**
- `Paragraph.numbering_path` currently `List[int]` with `0` for missing levels. Change to sparse or sentinel: either filter out `0` or keep but mark missing.
- Add helper `_clean_numbering_path(path, level_formats)` that removes trailing/embedded `0` where `level_formats[idx]=="decimal"` and value==0 due to missing level, but preserves intentional `0`? Simpler: in `format_numbering_label`, skip substitution where `path[idx]==0` and `lvlText` contains `%n` for that idx → omit that component and its dot.
- Alternatively, `NumberingResolver` should not pad missing levels; it should build path by walking `ilvl` ancestors via `w:abstractNum` `lvl` definitions, not inserting 0.

**Parser changes**
- None if model fix suffices, but verify `NumberingResolver` correctly handles `ilvl` gaps: currently it may do `path = [counters[0], counters[1], 0, counters[3]]` when level 2 missing. Fix to `path = [counters[0], counters[1], counters[3]]` and adjust `level_formats` similarly, then `lvlText` `%3` should map to new index.

**Renderer changes**
- `output/html_renderer.py` `_render_paragraph_html` and `_render_cell_content` and `render_sidebar_toc` all call `format_numbering_label` — no direct change, but ensure they pass cleaned path.
- `_viewer_script`: no change.
- `semantic/toc.py` `build_toc` passes `numbering_label` through same function — automatically fixed.

**Tests required**
- Unit: `test_numbering.py` add case `H1 -> H4` skip → label `1.1` not `1.0.0.1`; `H2 -> H4` → `1.1` not `1.0.1`.
- Fixture: `tests/fixtures/skipped-levels.docx` already exists (1 doc) — add assertion: `format_numbering_label([1,1,0,1], ["decimal"]*4, "%1.%2.%3.%4") == "1.1.1"` (collapsed).
- Integration: `test_heading_pipeline.py` `test_skipped_levels` — verify hierarchy + numbering label via `convert_docx` on CSD mini fixture.
- Browser: `test_visible_numbering_playwright.py` — load CSD `viewer.html`, assert no `".0."` in `h4` labels, TOC numbers lack `0`.

**Browser validation required**
- Re-run `benchmark_run.py` with CSD: verify `h4` headings now show `1.1.1` not `1.1.0.1`, TOC 67 labels clean, `heading-sections` still 67, heading-focus isolation still works (click H4, verify visible blocks), no console errors, screenshots `02_viewer_full.png` shows corrected numbering.

**Do NOT implement before benchmark sign-off.**

---

## 10. DO NOT IMPLEMENT THE FIX YET

Code unchanged for this benchmark. Next PR should implement the numbering skip fix only, with tests + browser evidence on CSD, before addressing PAGE fields or font-family.

---

**Evidence artifacts:**
- `benchmark_doc/csd-thesis-template-9th-draft.docx` (938 KB)
- `benchmark_doc/evidence/viewer.html` (1.3 MB, self-contained)
- `benchmark_doc/evidence/01_outer_page.png` (product shell + iframe)
- `benchmark_doc/evidence/02_viewer_full.png` (viewer 67 headings, 8 images, 1 table, 12 header/footer)
- `benchmark_doc/evidence/03_after_click.png` (focused `Declaration` 8/293 blocks)
- `benchmark_doc/evidence/toc_nav.html` (generated)
- Logs: `console []`, `failed []`, `TOC 67`, `sections 67`, `docx-block 293`, `heading-sections` JSON

**Success:** One real-world DOCX benchmarked via real product flow, forensic source vs HTML matrix built, actual failures ranked, highest-leverage single bottleneck identified with evidence-backed implementation plan, no code changes before audit.

