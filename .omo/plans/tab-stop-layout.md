# Plan: Word Tab-Stop Layout

## TODOs

- [x] 1. Model - add TabStop representation for w:tabs
- [x] 2. Parser - parse w:pPr/w:tabs/w:tab (val/pos/leader)
- [x] 3. Renderer - tab-stop-aware rendering (left/center/right/clear, fallback, leader)
- [x] 4. Tests - minimal locking for extraction, position conversion, fallback, footer rendering
- [x] 5. Real product validation - CSD footer positions via chromium, screenshot, no overflow/errors
- [x] 6. Regression - numbering, lists, viewer, PAGE fields preserved

## Final Verification Wave

- [x] F1. Model/Parser Integrity
- [x] F2. Rendering Positions
- [x] F3. Footer Alignment & PAGE Preservation
- [x] F4. Browser & Regression Evidence
