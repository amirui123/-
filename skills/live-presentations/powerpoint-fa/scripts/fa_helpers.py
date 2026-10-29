# -*- coding: utf-8 -*-
"""Shared Persian-deck helpers: theme tokens + RTL primitives.

Single source for brand skin (colors/type) so builders and checkers agree.
Extracted from build_showcase.py visual system; safe to import from any
builder script (requires python-pptx only).
"""
from pptx.dml.color import RGBColor
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ----------------------------------------------------------------------------
# Canvas
# ----------------------------------------------------------------------------
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.6)
EMU_W, EMU_H = 12192000, 6858000
FOOTER_Y = Inches(7.06)  # body must end above this (see slide-rules-fa §4)

# ----------------------------------------------------------------------------
# Semantic color tokens (warm default; navy alt lives in comments)
# ----------------------------------------------------------------------------
SAND = RGBColor(0xF7, 0xF2, 0xE9)
INK = RGBColor(0x21, 0x2B, 0x26)
SECONDARY = RGBColor(0x5A, 0x6B, 0x63)
ACCENT = RGBColor(0xD9, 0x8E, 0x2B)
TEAL = RGBColor(0x2E, 0x7D, 0x5B)
POSITIVE = RGBColor(0x2E, 0x7D, 0x5B)
WARNING = RGBColor(0xD9, 0x8E, 0x2B)
NEGATIVE = RGBColor(0xB3, 0x3A, 0x2E)
CARD = RGBColor(0xFF, 0xFF, 0xFF)
CARDBORDER = RGBColor(0xE3, 0xDA, 0xC8)
FAINT = RGBColor(0x8A, 0x97, 0x90)
AMBER_SOFT = RGBColor(0xF7, 0xE6, 0xC8)
GREEN_SOFT = RGBColor(0xE3, 0xEF, 0xE7)
RED_SOFT = RGBColor(0xF6, 0xE0, 0xDC)
DARKPANEL = RGBColor(0x1B, 0x2A, 0x24)

FONT = "Vazirmatn"
FARSI = MSO_LANGUAGE_ID.FARSI
EN = MSO_LANGUAGE_ID.ENGLISH_US


# ----------------------------------------------------------------------------
# RTL primitives
# ----------------------------------------------------------------------------
def set_persian_font(run, name=FONT):
    """Set latin+ea+cs typefaces so Persian glyphs truly render in `name`.

    `run.font.name` alone only sets `a:latin`; Persian/Arabic glyphs render
    through the complex-script (`a:cs`) slot and silently fall back without
    this. Covers `a:ea` too for safety.
    """
    run.font.name = name  # a:latin
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)


def set_run_font(run, size=18, bold=False, color=INK, font=FONT, lang=FARSI):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    f.name = font
    f.language_id = lang
    set_persian_font(run, font)


def set_para_rtl(paragraph, rtl=True):
    pPr = paragraph._p.get_or_add_pPr()
    if rtl:
        pPr.set("rtl", "1")
    elif "rtl" in pPr.attrib:
        del pPr.attrib["rtl"]


def set_bullet(paragraph, level=0, char="•"):
    """Real list bullet via `buChar` (never typed `• ` into run text).

    slide-rules-fa §7.3: typed markers force receivers to reformat every
    slide. `char=None` + level=None clears to `buNone` for single-statement
    blocks that must carry no bullet.
    """
    paragraph.level = level
    pPr = paragraph._p.get_or_add_pPr()
    # clear existing bullet props
    for tag in ("a:buChar", "a:buAutoNum", "a:buBlip", "a:buNone"):
        el = pPr.find(qn(tag))
        if el is not None:
            pPr.remove(el)
    if char is None:
        none = pPr.makeelement(qn("a:buNone"), {})
        pPr.append(none)
        return
    bu = pPr.makeelement(qn("a:buChar"), {"char": char})
    pPr.append(bu)


def fw_len(text):
    """Full-width-equivalent length for Persian decks.

    Japanese-source heuristic (ord < 0x3000 -> half) miscounts Persian:
    Arabic-script chars (U+0600-U+06FF) are below U+3000 but render
    full-width. Count them (plus presentation forms) as 1.0, ZWNJ/ZWJ as
    0.0, other sub-U+3000 chars (Latin/digits) as 0.5.
    """
    total = 0.0
    for ch in text:
        o = ord(ch)
        if o in (0x200C, 0x200D):
            continue
        if o >= 0x3000 or 0x0600 <= o <= 0x06FF or 0xFB50 <= o <= 0xFEFF:
            total += 1.0
        else:
            total += 0.5
    return total
