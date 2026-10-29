# -*- coding: utf-8 -*-
"""True Arabic/Persian shaping + glyph outlines (preview + measurement).

uharfbuzz does the shaping (same engine class as PowerPoint/browsers), fontTools
supplies outlines. All geometry is returned in EM units; callers scale by
size_pt/72/upem to get inches. No presentation-form hackery, so dots survive.
"""
from __future__ import annotations

import functools
import os

import uharfbuzz as hb
from fontTools.pens.basePen import BasePen
from fontTools.ttLib import TTFont

FONT_DIR = os.environ.get("VAZIR_DIR", "/tmp/fonts")


ARABIC_LETTER_RANGES = (
    ("\u0620", "\u064a"), ("\u0671", "\u06d3"), ("\u06fa", "\u06ff"),
    ("\ufb50", "\ufdff"), ("\ufe70", "\ufeff"),
)


def is_arabic_letter(ch: str) -> bool:
    """Arabic-script *letters* only. Persian/Latin digits are NOT included:
    digits are laid out left-to-right even inside RTL text."""
    return any(lo <= ch <= hi for lo, hi in ARABIC_LETTER_RANGES)


def font_path(bold: bool) -> str:
    return os.path.join(FONT_DIR, "Vazirmatn-Bold.ttf" if bold else "Vazirmatn-Regular.ttf")


@functools.lru_cache(maxsize=None)
def _hb(path: str):
    with open(path, "rb") as f:
        face = hb.Face(f.read())
    return face, hb.Font(face)


@functools.lru_cache(maxsize=None)
def _tt(path: str) -> TTFont:
    return TTFont(path)


@functools.lru_cache(maxsize=None)
def shape_em(text: str, bold: bool = False):
    """Shape text. Returns ([(glyph_name, dx_em, dy_em)], advance_em)."""
    path = font_path(bold)
    face, font = _hb(path)
    buf = hb.Buffer()
    buf.add_str(text)
    buf.guess_segment_properties()
    # Arabic-script text shapes RTL; Latin words / ASCII numbers are pure LTR.
    has_arabic = any(is_arabic_letter(ch) for ch in text)
    buf.direction = "rtl" if has_arabic else "ltr"
    if has_arabic:
        buf.script = "Arab"
        buf.language = "fa"
    hb.shape(font, buf, {"kern": True, "liga": True})
    upem = face.upem
    order = _tt(path).getGlyphOrder()
    pen_x = 0.0
    out = []
    for info, pos in zip(buf.glyph_infos, buf.glyph_positions):
        out.append((order[info.codepoint], (pen_x + pos.x_offset) / upem, -pos.y_offset / upem))
        pen_x += pos.x_advance
    return out, pen_x / upem


def measure_in(text: str, bold: bool = False, size_pt: float = 18.0) -> float:
    return shape_em(text, bold)[1] * size_pt / 72.0


class _Pen(BasePen):
    """Collects matplotlib-ready contours (verts, codes) in inches."""

    def __init__(self, glyphSet, scale, ox, baseline):
        super().__init__(glyphSet)
        self.s = scale
        self.ox = ox
        self.base = baseline
        self.verts = []
        self.codes = []
        self._start = None

    def _p(self, pt):
        return (self.ox + pt[0] * self.s, self.base - pt[1] * self.s)

    def _moveTo(self, pt):
        self._start = self._p(pt)
        self.verts.append(self._start)
        self.codes.append(1)

    def _lineTo(self, pt):
        self.verts.append(self._p(pt))
        self.codes.append(2)

    def _qCurveToOne(self, p1, p2):
        self.verts.append(self._p(p1))
        self.codes.append(3)
        self.verts.append(self._p(p2))
        self.codes.append(2)

    def _curveToOne(self, p1, p2, p3):
        self.verts.append(self._p(p1))
        self.codes.append(4)
        self.verts.append(self._p(p2))
        self.codes.append(4)
        self.verts.append(self._p(p3))
        self.codes.append(2)

    def _closePath(self):
        self.verts.append(self._start or (self.ox, self.base))
        self.codes.append(79)


def text_paths_in(text: str, bold: bool, size_pt: float, x_in: float, baseline_in: float):
    """Glyph outlines in inches, logical start (left edge) at x_in."""
    glyphs, _adv = shape_em(text, bold)
    path = font_path(bold)
    tt = _tt(path)
    gs = tt.getGlyphSet()
    upem = tt["head"].unitsPerEm
    scale = size_pt / 72.0 / upem      # em units -> inches
    out = []
    for gname, dx, dy in glyphs:
        pen = _Pen(gs, scale, x_in + dx * size_pt / 72.0, baseline_in + dy * size_pt / 72.0)
        gs[gname].draw(pen)
        if pen.verts:
            out.append((pen.verts, pen.codes))
    return out
