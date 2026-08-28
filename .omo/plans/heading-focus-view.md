# Plan: Heading-Focus View

## TODOs

- [x] 1. Implement section-boundary logic and renderer annotations in output/html_renderer.py
- [x] 2. Implement focused-viewer JS/CSS (TOC click isolation, history, active state, viewport filling) in output/html_renderer.py
- [x] 3. Verify with real DOCX (mixed-document.docx, num-h1-h2-h3.docx) - manual browser flow and automated checks
- [x] 4. Final Verification Wave - all reviewers APPROVE

## Context
Goal: TOC click shows ONLY that heading's content as focused viewer page, bounded until next same/higher level heading.
Constraints: Reuse Paragraph.heading_level/heading_id/TOC hierarchy, no regex, no duplicate conversion, no OOXML modification, keep sidebar/expand/collapse, keep typography/tables/images/hyperlinks/floating images, no duplicate heading, first H1 title bar intact, history back/forward restores section, avoid reload, actually isolate viewport not just scroll, HTML/model is source of truth, no regeneration on click.

## Final Verification Wave
- [x] F1. Section Boundary Logic Review
- [x] F2. Viewer Isolation & Navigation Review
- [x] F3. Fidelity & Regression Review (typography, tables, images, hyperlinks, floating, first-H1, no overflow)
- [x] F4. Real DOCX Browser Evidence (mixed-document.docx, num-h1-h2-h3.docx)

