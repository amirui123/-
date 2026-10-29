# -*- coding: utf-8 -*-
"""Post-build text fitting with HarfBuzz metrics.

deckkit measures before placing, but PPTX re-wraps with its own layout engine.
This pass re-measures every paragraph with real shaping and, where a paragraph
would wrap into more lines than the box can hold, shrinks that paragraph's runs
uniformly (never below 85% of design size) until it fits.
"""
from __future__ import annotations

import os
import sys

from pptx import Presentation
from pptx.util import Emu, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import shaping

SAFETY = 0.97        # head-room for engine differences
MIN_FACTOR = 0.85


def _wrap_count(words, max_w):
    lines, cur_w = 1, 0.0
    for text, size, bold in words:
        w = shaping.measure_in(text, bold, size)
        sp = shaping.measure_in(" ", bold, size) if cur_w else 0.0
        if cur_w and cur_w + sp + w > max_w:
            lines += 1
            cur_w = w
        else:
            cur_w += sp + w
    return lines


def fit_presentation(path, verbose=True):
    prs = Presentation(path)
    fixes = []
    for si, slide in enumerate(prs.slides, start=1):
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            tf = sh.text_frame
            box_w = Emu(sh.width).inches
            box_h = Emu(sh.height).inches
            if box_w <= 0.1:
                continue
            for pi, p in enumerate(tf.paragraphs):
                runs = [(r, r.text) for r in p.runs if r.text]
                if not runs:
                    continue
                sizes = [(r.font.size.pt if r.font.size else 18) for r, _ in runs]
                bols = [bool(r.font.bold) for r, _ in runs]
                ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.2
                factor = 1.0
                for _ in range(14):
                    words = []
                    for (r, t), sz, b in zip(runs, sizes, bols):
                        for w in t.split(" "):
                            if w:
                                words.append((w, sz * factor, b))
                    n = _wrap_count(words, box_w * SAFETY)
                    need = n * max(sizes) * factor * ls / 72.0
                    if need <= box_h + 0.01 or factor <= MIN_FACTOR:
                        break
                    factor *= 0.96
                if factor < 0.999:
                    for (r, _), sz in zip(runs, sizes):
                        r.font.size = Pt(round(sz * factor, 1))
                    if factor <= MIN_FACTOR + 0.001:
                        fixes.append(f"S{si:02d} para{pi}: shrunk to floor "
                                     f"({factor:.2f}) box {box_w:.2f}x{box_h:.2f}in "
                                     f"'{runs[0][1][:28]}'")
    if fixes:
        for f in fixes:
            print("  FIT-WARN", f)
    prs.save(path)
    if verbose:
        print(f"fit pass done ({len(fixes)} floor warnings)")
    return fixes
