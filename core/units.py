"""Authoritative EMU / OOXML unit conversion for docx-to-html.

DOCX stores all drawing geometry in EMU (English Metric Units):
  1 inch   = 914400 EMU
  1 pixel  = 9525   EMU   (at 96 DPI, the CSS reference resolution)
  1 pt     = 12700  EMU   (1/72 inch)
  1 twip   = 635    EMU   (1/1440 inch = 1/20 pt; sectPr pgSz/pgMar, ind, spacing)

  px = emu / 914400 * 96
  px = twip / 1440 * 96
  px = pt / 72 * 96

This module is the SINGLE source of truth for converting between EMU and
browser pixels. The parser stores geometry in EMU on the normalized model;
the renderer converts to CSS pixels through these helpers so the conversion
constants live in exactly one place.
"""

# 1 inch = 914400 EMU.
EMU_PER_INCH = 914400
# 96 DPI reference: 1 CSS pixel = 914400 / 96 = 9525 EMU.
EMU_PER_PIXEL = 9525
# 1 pt = 12700 EMU.
EMU_PER_POINT = 12700
# 1 twip = 635 EMU.
EMU_PER_TWIP = 635

# US-Letter default page geometry (EMU), used when sectPr is absent.
DEFAULT_PAGE_WIDTH_EMU = 8.5 * EMU_PER_INCH    # 7772400
DEFAULT_PAGE_HEIGHT_EMU = 11.0 * EMU_PER_INCH   # 10058400
DEFAULT_MARGIN_EMU = 1.0 * EMU_PER_INCH         # 914400 (1 inch)


def emu_to_px(emu) -> int:
    """Convert an EMU integer to a CSS pixel integer (96 DPI).

    Returns 0 for None / empty input and rounds to the nearest pixel. Never
    returns None so callers can use the value directly in arithmetic/layout.
    """
    if emu is None:
        return 0
    return round(emu / EMU_PER_PIXEL)


def px_to_emu(px) -> int:
    """Convert a CSS pixel value back to EMU (inverse of emu_to_px)."""
    if px is None:
        return 0
    return round(px * EMU_PER_PIXEL)


def twip_to_emu(twip) -> int:
    """Convert a twip (used by sectPr pgSz/pgMar) to EMU."""
    if twip is None:
        return 0
    return round(twip * EMU_PER_TWIP)
