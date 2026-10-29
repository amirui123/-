# -*- coding: utf-8 -*-
"""deckkit — minimal, measurable layout kit for Persian RTL decks (python-pptx).

Design intent: a dark cinematic "muscle car" canvas where every text block is
measured with the real Vazirmatn metrics before it is placed, so overflow is
caught at build time instead of after the fact.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field

from PIL import Image, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# --------------------------------------------------------------------------
# Canvas / theme
# --------------------------------------------------------------------------
W, H = 13.333, 7.5
M = 0.85                      # side margin
CW = W - 2 * M                # content width


def C(hexstr: str) -> RGBColor:
    hexstr = hexstr.lstrip("#")
    return RGBColor(int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16))


BG      = C("0A0B0D")
PANEL   = C("15181D")
PANEL2  = C("1B1F26")
HAIR    = C("2A2F37")
INK     = C("F4F5F7")
MUTED   = C("9AA2AE")
DIM     = C("6B7480")
RED     = C("E11D2E")
RED_DK  = C("8E121D")
ORANGE  = C("FF7A1A")
SILVER  = C("C8CDD6")

FONT = "Vazirmatn"
FONT_DIR = os.environ.get("VAZIR_DIR", "/tmp/fonts")
_FONTS = {}


def _font_path(bold: bool, black: bool = False, light: bool = False) -> str:
    if black:
        name = "Vazirmatn-Black.ttf"
    elif bold:
        name = "Vazirmatn-Bold.ttf"
    elif light:
        name = "Vazirmatn-Light.ttf"
    else:
        name = "Vazirmatn-Regular.ttf"
    return os.path.join(FONT_DIR, name)


def _pil(bold: bool, black: bool, light: bool, px: int = 100):
    key = (bold, black, light, px)
    if key not in _FONTS:
        _FONTS[key] = ImageFont.truetype(_font_path(bold, black, light), px)
    return _FONTS[key]


# --------------------------------------------------------------------------
# Text measurement + wrapping (real font metrics, in inches)
# --------------------------------------------------------------------------
@dataclass
class Run:
    text: str
    size: float = 18
    bold: bool = False
    black: bool = False
    light: bool = False
    color: RGBColor = INK


@dataclass
class Line:
    runs: list[Run] = field(default_factory=list)
    width: float = 0.0   # inches
    height: float = 0.0  # inches (line box)


def _advance(text: str, run: Run, px: int = 100) -> float:
    """Advance width in inches for `text` rendered with `run`'s font at its size."""
    f = _pil(run.bold, run.black, run.light, px)
    return f.getlength(text) / px * run.size / 72.0


def wrap(paragraph_runs: list[Run], max_w: float, line_spacing: float = 1.35) -> list[Line]:
    """Greedy word wrap across runs; returns positioned-independent lines."""
    tokens: list[tuple[str, Run]] = []   # (word_without_space, run)
    for run in paragraph_runs:
        for i, word in enumerate(run.text.split(" ")):
            if word == "":
                continue
            tokens.append((word, run))
    lines: list[Line] = []
    cur = Line()
    for idx, (word, run) in enumerate(tokens):
        space = 0.0 if not cur.runs else _advance(" ", run)
        w_word = _advance(word, run)
        if cur.width + space + w_word > max_w and cur.runs:
            lines.append(cur)
            cur = Line()
            space = 0.0
        # merge into the last run of the line when the style matches
        if cur.runs and cur.runs[-1].size == run.size and cur.runs[-1].bold == run.bold \
           and cur.runs[-1].color == run.color and cur.runs[-1].black == run.black:
            cur.runs[-1] = Run(cur.runs[-1].text + " " + word,
                               run.size, run.bold, run.black, run.light, run.color)
        else:
            cur.runs.append(Run(word, run.size, run.bold, run.black, run.light, run.color))
        cur.width += space + w_word
        cur.height = max(cur.height, run.size * line_spacing / 72.0)
    if cur.runs:
        lines.append(cur)
    return lines


def text_height(paragraphs: list[list[Run]], max_w: float, line_spacing: float,
                space_after: float) -> tuple[float, list[list[Line]]]:
    out = []
    total = 0.0
    for i, para in enumerate(paragraphs):
        ls = wrap(para, max_w, line_spacing)
        out.append(ls)
        total += sum(l.height for l in ls)
        if i < len(paragraphs) - 1:
            total += space_after / 72.0
    return total, out


# --------------------------------------------------------------------------
# XML helpers
# --------------------------------------------------------------------------
def _set_alpha(fill, pct: int):
    """pct = opacity 0..100"""
    srgb = fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
    alpha = srgb.makeelement(qn("a:alpha"), {"val": str(int(pct * 1000))})
    srgb.append(alpha)


def _set_font(run, size: float, bold: bool, color: RGBColor, black: bool = False,
              light: bool = False, letter_spacing: float | None = None):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    f.name = FONT
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", FONT)
    if letter_spacing:
        rPr.set("spc", str(int(letter_spacing * 100)))
    if black or light:
        # cosmetic weight via bold flag only (Vazirmatn Black/Light not embedded)
        pass


def _rtl(paragraph, align):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("rtl", "1")
    paragraph.alignment = align


# --------------------------------------------------------------------------
# Deck
# --------------------------------------------------------------------------
class Deck:
    def __init__(self):
        self.prs = Presentation()
        # exact 16:9 canvas (Inches(13.333) rounds off 12192000 EMU)
        self.prs.slide_width = Emu(12192000)
        self.prs.slide_height = Emu(6858000)
        self.warnings: list[str] = []
        self.n = 0

    # -- primitives ---------------------------------------------------
    def slide(self, bg: RGBColor = BG):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = bg
        self.n += 1
        return s

    def rect(self, s, x, y, w, h, fill=None, opacity=None, line=None, lw=1.0,
             shape=MSO_SHAPE.RECTANGLE, name=None):
        sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.shadow.inherit = False
        if fill is None:
            sh.fill.background()
        else:
            sh.fill.solid()
            sh.fill.fore_color.rgb = fill
            if opacity is not None:
                _set_alpha(sh.fill, opacity)
        if line is None:
            sh.line.fill.background()
        else:
            sh.line.color.rgb = line
            sh.line.width = Pt(lw)
        if sh.has_text_frame:
            sh.text_frame.text = ""
        if name:
            sh.name = name
        return sh

    def grad(self, s, x, y, w, h, stops, angle=90, name="scrim"):
        """Linear-gradient scrim. stops = [(pos0..1, hex, opacity%), ...].

        Full-bleed scrims (0,0,W,H) keep layout checkers from mistaking the
        darkening layer for content, and read better than a hard band.
        """
        sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                Inches(w), Inches(h))
        sh.shadow.inherit = False
        sh.line.fill.background()
        sh.name = name
        sh.fill.gradient()
        gs_lst = sh.fill.gradient_stops
        # rebuild the stop list to match `stops`
        gradFill = sh.fill._xPr.find(qn("a:gradFill"))
        gsLst = gradFill.find(qn("a:gsLst"))
        for el in list(gsLst):
            gsLst.remove(el)
        for pos, hexcol, alpha in stops:
            gs = gsLst.makeelement(qn("a:gs"), {"pos": str(int(pos * 100000))})
            clr = gs.makeelement(qn("a:srgbClr"), {"val": hexcol.lstrip("#")})
            a = clr.makeelement(qn("a:alpha"), {"val": str(int(alpha * 1000))})
            clr.append(a)
            gs.append(clr)
            gsLst.append(gs)
        gradFill.set("rotWithShape", "1")
        lin = gradFill.find(qn("a:lin"))
        if lin is None:
            lin = gradFill.makeelement(qn("a:lin"), {})
            gradFill.append(lin)
        lin.set("ang", str(int(angle * 60000)))
        lin.set("scaled", "0")
        return sh

    def hline(self, s, x, y, w, color=HAIR, th=0.012, opacity=None):
        return self.rect(s, x, y, w, th, fill=color, opacity=opacity, name="rule")

    def text(self, s, x, y, w, h, paragraphs, align=PP_ALIGN.RIGHT,
             anchor=MSO_ANCHOR.TOP, line_spacing=1.35, space_after=0,
             wrap_text=True, check=True, label=""):
        """paragraphs: list of list[Run] (each inner list = one paragraph)."""
        if check:
            need, _ = text_height(paragraphs, w - 0.02, line_spacing, space_after)
            if need > h + 0.02:
                self.warnings.append(
                    f"S{self.n:02d}: text overflow {label or ''} need {need:.2f}in > box {h:.2f}in")
        box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = wrap_text
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = anchor
        tf.auto_size = None
        for i, para in enumerate(paragraphs):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            _rtl(p, align)
            p.line_spacing = line_spacing
            if i < len(paragraphs) - 1:
                p.space_after = Pt(space_after)
            for run in para:
                r = p.add_run()
                r.text = run.text
                _set_font(r, run.size, run.bold, run.color, run.black, run.light)
        return box

    def pic(self, s, path, x, y, w, h):
        iw, ih = Image.open(path).size
        target = w / h
        src = iw / ih
        pic = s.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
        if src > target:          # crop left/right
            frac = (1 - target / src) / 2
            pic.crop_left = frac
            pic.crop_right = frac
        else:                     # crop top/bottom
            frac = (1 - src / target) / 2
            pic.crop_top = frac
            pic.crop_bottom = frac
        return pic

    # -- furniture ----------------------------------------------------
    def kicker(self, s, text, color=RED, x=None, y=0.55, w=4.0):
        x = (W - M - w) if x is None else x
        return self.text(s, x, y, w, 0.3, [[Run(text, 13.5, True, False, False, color)]],
                         align=PP_ALIGN.RIGHT, line_spacing=1.1, label="kicker")

    def title(self, s, text, size=37, y=0.95, w=CW, h=None, color=INK,
              align=PP_ALIGN.RIGHT, line_spacing=1.26):
        paras = [[Run(text, size, True, False, False, color)]]
        need, _ = text_height(paras, w - 0.02, line_spacing, 0)
        h = need if h is None else h
        return self.text(s, M, y, w, h, paras, align=align, line_spacing=line_spacing,
                         label="title")

    def footer(self, s, source_text=None, page=True, idx=None, src_x=None, src_w=7.2):
        self.hline(s, M, 6.94, CW, HAIR, 0.01)
        if source_text:
            x0 = M if src_x is None else src_x
            src_w = min(src_w, 11.62 - x0)
            self.text(s, x0, 7.09, src_w, 0.26,
                      [[Run(source_text, 10.5, False, False, True, DIM)]],
                      align=PP_ALIGN.LEFT, line_spacing=1.1, label="src")
        if page and idx is not None:
            self.text(s, 11.9, 7.09, 0.58, 0.26,
                      [[Run(fa_num(idx), 11, True, False, False, DIM)]],
                      align=PP_ALIGN.RIGHT, line_spacing=1.1, label="pnum")

    def save(self, path):
        self.prs.save(path)


# --------------------------------------------------------------------------
# Persian numerals
# --------------------------------------------------------------------------
_FA = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")


def fa_num(n) -> str:
    return str(n).translate(_FA)


def fa(s: str) -> str:
    """Convert ASCII digits in an already-Persian string to Persian digits."""
    return s.translate(_FA)
