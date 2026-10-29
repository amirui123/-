---
name: powerpoint-fa
description: Create Persian RTL .pptx decks with python-pptx.
version: 0.1.0
author: Eledah (eledah), Hermes Agent
license: MIT
platforms: [linux, macos, windows]
metadata:
  hermes:
    tags: [pptx, powerpoint, persian, farsi, rtl, python-pptx, presentations]
    category: productivity
---

# Persian PowerPoint Skill

Create, inspect, edit, and verify Persian-first PowerPoint presentations with
`python-pptx`. Use this skill for Persian, bilingual, or RTL `.pptx` work. It
supports a portable public-API workflow and clearly isolates renderer- and
OOXML-dependent behavior.

This is a `.pptx` skill, not a legacy `.ppt` skill. Convert legacy files before
working on them. Do not claim that a deck is visually correct without rendering
and reviewing it when a renderer is available.

## When to Use

Use for:

- creating Persian, Farsi, Arabic-script, bilingual, or mixed RTL/LTR decks;
- reading slide text, tables, charts, images, notes, layouts, and metadata;
- editing text, charts, images, backgrounds, hyperlinks, notes, or slide order;
- generating report, pitch, teaching, proposal, or executive decks;
- filling a branded template while preserving its layouts and theme;
- checking Persian directionality, fonts, overflow, contrast, and visual quality.

Do not use this skill for legacy `.ppt` files without conversion, PDF-only
editing, or pixel-perfect PowerPoint features unsupported by `python-pptx`.

## Prerequisites

Check the environment with `terminal` before writing code:

```bash
python3 --version
python3 -c "import pptx; print(pptx.__version__)"
command -v soffice || true
command -v pdftoppm || command -v pdftocairo || true
```

If `python-pptx` is unavailable, create or use a project virtual environment
and install it. Pin the tested version when reproducibility matters:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/pip install python-pptx==1.0.2
```

LibreOffice and Poppler are optional but required for the strongest visual QA
loop. Rendering with LibreOffice is an approximation of PowerPoint, not proof
of identical output. A Persian-capable font must be installed on the rendering
machine; naming a font does not embed it in the presentation.

## Quick Reference

```python
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.util import Inches, Pt

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])
box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(2))
p = box.text_frame.paragraphs[0]
p.alignment = PP_ALIGN.RIGHT
p.text = "عنوان فارسی"
p._p.get_or_add_pPr().set("rtl", "1")
run = p.runs[0]
run.font.name = "Vazirmatn"
run.font.size = Pt(30)
run.font.bold = True
run.font.language_id = MSO_LANGUAGE_ID.FARSI
run.font.color.rgb = RGBColor(0x15, 0x23, 0x2F)
set_persian_font(run, "Vazirmatn")  # sets latin + ea + cs so Persian truly renders
prs.save("deck.pptx")
```

Note: `run.font.name` alone only sets the `a:latin` typeface. To make Persian
glyphs render in the chosen font, also set the east-asian and complex-script
typefaces on the run's `rPr`:

```python
from pptx.oxml.ns import qn

def set_persian_font(run, name):
    run.font.name = name  # a:latin
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)
```

Common public APIs:

- `Presentation()`, `Presentation(path)`, `prs.save(path)`
- `prs.slides.add_slide(layout)`, `prs.slide_layouts`, `prs.slide_masters`
- `slide.shapes.add_textbox`, `add_picture`, `add_shape`, `add_table`,
  `add_chart`, `add_connector`
- `shape.text_frame`, `TextFrame.paragraphs`, `Paragraph.runs`, `Run.font`
- `Inches`, `Cm`, `Pt`, `Emu`, `RGBColor`
- `CategoryChartData`, `XyChartData`, `BubbleChartData`
- `slide.notes_slide.notes_text_frame`
- `prs.core_properties`

## Procedure

### 1. Define the deck contract

Before implementation, record:

- audience, purpose, delivery context, duration, and desired action;
- one central message and one claim-style takeaway per slide;
- language/script policy: Persian, bilingual, or mixed RTL/LTR;
- digit policy: Persian, Arabic-Indic, or European digits;
- approved Persian font and fallback policy;
- slide size, safe margins, theme colors, and renderer targets;
- source URLs or citations for factual claims.

Do not begin with arbitrary shapes. Start from the audience message and choose
a layout that supports it.

Follow Define-before-Produce (canon: `references/slide-rules-fa.md` §0):
storyline first (one line per future slide, reading as one speech), then
pattern choice per line from `references/archetype-catalog-fa.md` (patterns
are an idea book, never a target — discard on misfit), then fact/format
check, AI-tic pass (`references/ai-smell-fa.md`), and blind fresh-eye
review (`references/content-review-fa.md`). For drafts, assemble with
`python3 scripts/pptx_new_fa.py --parts cover,map,axis,chart,decision,back
--title "نام سند" -o mydeck.pptx` (`--list` shows pattern IDs), then replace
every placeholder with the real storyline — zero placeholders survive.

### 2. Design the visual system

Use 16:9 unless the user or template requires another size. Define semantic
tokens for background, primary text, secondary text, accent, positive,
warning, negative, and categorical data colors. Use a small type scale for
claim, heading, body, label, source, and footnote roles.

Prefer one dominant idea per slide. Use whitespace, alignment, grouping, and
contrast to establish hierarchy. Do not shrink text to rescue an overloaded
slide; shorten, split, or restructure it. Do not encode meaning by color alone.
Use meaningful titles, simple tables, direct chart labels, and progressive
disclosure for complex explanations.

### 3. Prepare Persian text

Keep text in logical Unicode order. Never reverse a Persian string or pre-shape
it into presentation-form glyphs. Preserve meaningful ZWNJ (`U+200C`) and ZWJ
(`U+200D`). Apply the project's explicit policy for Persian `ی` and `ک`,
combining marks, punctuation, dates, URLs, and digits.

Separate Persian, English, URLs, numbers, and code-like fragments into runs when
they need different fonts, languages, emphasis, or directional behavior. Use
Unicode directional isolates only around genuinely ambiguous fragments; do not
scatter invisible direction marks through ordinary prose.

For each Persian paragraph:

1. keep logical text order;
2. set `paragraph.alignment = PP_ALIGN.RIGHT`;
3. set the DrawingML paragraph `rtl="1"` attribute through a small tested XML
   helper or the compatible private element API;
4. set `run.font.language_id = MSO_LANGUAGE_ID.FARSI`;
5. set an approved Persian-capable font explicitly — for every run use the
   `set_persian_font(run, name)` helper above (sets `a:latin`, `a:ea`, and
   `a:cs`), because `run.font.name` alone leaves Persian glyphs in the default
   complex-script font;
6. use explicit margins, wrapping, and conservative dimensions.

Alignment is not the same as bidi direction. Language metadata is not a bidi
switch. `python-pptx` does not itself perform shaping, fallback, or renderer
layout.

### 4. Build slides with semantic geometry

Keep content, geometry, and styling separate. Use named semantic zones such as
`title`, `claim`, `evidence`, `chart`, `source`, and `footer`. Mirror visual
relationships deliberately for RTL: arrows, timelines, process flows, chart
axes, and comparison columns do not mirror automatically when text becomes RTL.

Prefer explicit textboxes when exact Persian geometry matters. Use placeholders
when template inheritance, master styling, or accessibility structure is more
important. Check `shape.has_text_frame` before accessing text and remember that
blank layouts may have no `slide.shapes.title`.

For body text, prefer short phrases and labels over paragraphs. Set line
spacing, paragraph spacing, margins, and text-frame height explicitly. Treat
`fit_text()` and automatic sizing as fallback tools, not as a Persian layout
strategy.

### 5. Add data, tables, notes, and sources

Choose a chart for the question: bars/columns for comparisons, lines for
trends, scatter for relationships, and pie only for a small part-to-whole
composition. Keep underlying chart values numeric; format displayed labels
separately. State the chart's conclusion in its title or annotation.

Treat RTL tables as two separate problems: RTL paragraph text and visual/logical
column order. Choose and document whether columns are preserved logically or
mirrored visually. Test headers, totals, merged cells, numeric columns, borders,
and long wrapped text.

Use `slide.notes_slide.notes_text_frame` for speaking cues, source details,
alt-text drafts, and content that should not clutter the slide. Add meaningful
alt text and titles where the target PowerPoint/OOXML workflow supports them.

### 6. Validate structurally

After every create or edit operation, save to a new output path when possible
and reopen it with `Presentation(output_path)`. Check:

- intended slide count, order, layouts, and slide size;
- all required titles, claims, labels, sources, and notes;
- Persian text round-trip, ZWNJ/ZWJ, digit, and mixed-script integrity;
- `PP_ALIGN.RIGHT` and tested `rtl="1"` on Persian paragraphs;
- approved fonts and `FARSI` language metadata;
- table logical order and chart categories/series/values;
- image paths, hyperlinks, and metadata;
- no empty required placeholders or missing text frames.

If a repository-provided validator is available, run it and treat errors as
blockers. Otherwise write a small deterministic read-back check with
`python-pptx` rather than trusting that `prs.save()` succeeded.

Mandatory gate — mechanical check FAIL 0 before delivery:

```bash
python3 scripts/check_deck_fa.py deck.pptx        # titles, placeholders,
                                                  # counts, tables, RTL, AI-smell
python3 scripts/check_layout_fa.py deck.pptx      # footer overlap / overflow
```

`check_deck_fa.py` prints the title list: read it top to bottom and confirm
one story. `check_layout_fa.py` ports the `consulting-pptx-skill`
`check_layout.mjs` taxonomy (footer/body overlap, right/bottom overflow)
to static `python-pptx` geometry with no extra dependencies. Large
`roundRect` is WARN by default (`--strict` promotes to FAIL); typed `•`
bullets in run text are FAIL (use `buChar` via `set_bullet()` in
`scripts/fa_helpers.py`, `buNone` for single-statement blocks).

### 7. Render and inspect visually

When `soffice` and Poppler are available:

```bash
mkdir -p render
soffice --headless --convert-to pdf --outdir render deck.pptx
pdftoppm -png render/deck.pdf render/slide
```

Review every rendered slide with `vision_analyze` when available. Look for
clipped or overlapping objects, incorrect RTL order, broken shaping or marks,
font substitution, bad line breaks, unreadable contrast, dense text, chart
labels, table columns, and inconsistent margins. When rendering tools are
missing, report that limitation and complete structural validation instead;
do not claim visual verification.

Then run the blind fresh-eye review (`references/content-review-fa.md`):
a second agent with the method hidden reads the deck file, findings go into
an adopt/reject/hold triage table, adopted fixes land, and both mechanical
checks re-run. One rally = mechanical → eye-QA → fresh-eye → fixes →
re-check. **Stop after 1 rally, show product + table, and ask before rally
2. Never self-run rally 2+ unasked.** Agree scope up front: speed
(mechanical + self-eye only) / standard (1 rally) / quality (N rallies).

## Capability Boundaries

`python-pptx` supports presentations, slides, layouts, masters, placeholders,
text, images, shapes, tables, common charts, notes, backgrounds, hyperlinks,
and core properties. It does not provide complete public APIs for paragraph
bidi, Kashida, OpenType shaping, arbitrary theme/master XML, comments,
animations, transitions, SmartArt, slide deletion/reordering/duplication, or
cross-presentation slide copying. Some operations require narrowly scoped
OOXML edits and post-save tests. Chart updates generally replace the dataset;
they are not cell-level workbook edits.

Unsupported Open XML parts may survive a load/save round trip, but preservation
does not mean they are safe to understand or edit. Media and OLE behavior is
client-dependent. LibreOffice is useful for Linux regression rendering but is
not pixel-identical to Microsoft PowerPoint.

## Pitfalls

- Do not manually reverse Persian text or pre-shape Arabic-script glyphs.
- Right alignment alone does not establish RTL paragraph direction.
- Font names are not font embedding; fallback changes metrics and wrapping.
- `run.font.name` only sets the `a:latin` typeface. Persian/Arabic glyphs render
  through the complex-script (`a:cs`) typeface, so if you only set the latin
  font the Persian text silently falls back to a default font. Set `a:ea` and
  `a:cs` typefaces explicitly on the run's `rPr` so Persian actually renders in
  the chosen font (e.g. Vazirmatn), not just the Latin/digits:
- RTL text does not automatically mirror geometry, tables, or charts. For RTL decks, sequence/step blocks (1,2,3,4 / مرحله ۱ تا ۴) must be laid out RIGHT-to-LEFT: the first item goes at the highest x (right side), then decreasing. Placing them left-to-right (x ascending) reverses the reading order for a Persian audience. Mirror any numbered cards, timeline, or process steps explicitly.
- RTL text does not automatically mirror geometry, tables, or charts.
- Do not assume Persian bullets, indents, chart axes, or legends behave like
  ordinary textboxes; test each object type.
- LibreOffice ignores the `a:tblPr` `rtl="1"` attribute when rendering tables,
  so a table that relies on it for column order renders in LTR order. For a
  visual RTL table, place the columns explicitly (right-most column = the
  RTL-leading column) instead of depending on the table-level rtl attribute.
- Charts carry their own default theme font and do NOT inherit the deck's
  font, so Persian chart titles/axis labels/tick labels render in a fallback
  font unless you set it. Force the font on the chart XML by setting
  `typeface` on every `a:latin`/`a:ea`/`a:cs` node and injecting them into
  every `a:defRPr`/`a:endParaRPr` in `chart._chartSpace`.
- Decorative accent bars/strips/shapes must be added to the slide BEFORE the
  text box so the color sits behind the text, and placed on the RTL-leading
  (right) side — a bar left detached from right-aligned Persian text looks
  like it is "in front" and on the wrong side.
- Corner footer anchors: a footer element only hugs the slide edge if its BOX is sized/positioned to end AT that edge — setting text alignment alone is not enough. Right-aligning a Persian brand line inside a WIDE mid-slide box (e.g. L0.5" W9.0" on a 13.33" slide) pushes the text to the box's right edge (~9.5"), so the brand floats toward the CENTER of the footer instead of the corner. For a left-corner brand use a NARROW right-aligned box sitting at the left edge (e.g. L0.5" W2.2" → text ends ~2.7"); for a right-corner page number use a NARROW right-aligned box at the right edge (e.g. L12.8" W0.3" → digit ends ~13.1"; widen to ~0.4" for two-digit numbers). Verify each footer element lands in its corner after rendering — never trust that "alignment set" equals "at the edge".
- For stacked card/list layouts, budget the vertical space explicitly: compute
  card height, gap, and start so the last card ends above the footer
  (e.g. y < ~7.06" on a 7.5" slide) — the last item silently collides with the
  footer otherwise.
- Do not use arbitrary XML patches. Keep them namespace-qualified, minimal,
  version-aware, and covered by regression fixtures.
- Do not copy slides between decks by cloning XML without relationship tests.
- Do not trust an outline check as a substitute for rendered visual review.
- Do not silently normalize identifiers, URLs, code, or user-provided tokens.
- Do not claim PowerPoint-equivalent output from a successful LibreOffice render.

## Verification

A production-ready deck has:

1. a documented audience, message, language, font, digit, and renderer contract;
2. a successful package save/reopen round trip;
3. passing semantic, Unicode, RTL XML, font, table, and chart checks;
4. a meaningful title and intended reading order on every slide;
5. contrast and non-color-cue checks;
6. rendered inspection of every slide when a renderer is available;
7. an honest final report listing renderer, installed fonts, warnings, and
   unsupported features.

For detailed API, RTL, design, OOXML, schema, and visual-QA material, use the
reference files shipped alongside this skill and load only the one needed
for the current task:

- `references/slide-rules-fa.md` — rule canon (wins on conflict; RTL mirror
  of `consulting-pptx-skill` slide-rules). **Read before building.**
- `references/archetype-catalog-fa.md` — pattern catalog (idea book).
  Read when choosing the showing-how per storyline line.
- `references/ai-smell-fa.md` — Persian AI-tic lexicon + 30-second
  self-check. Read at wording polish.
- `references/content-review-fa.md` — blind fresh-eye prompt + triage
  table. Read after mechanical FAIL 0, before delivery.
- `scripts/check_deck_fa.py` — mechanical gate (stdlib + python-pptx).
- `scripts/check_layout_fa.py` — static geometry audit (no Node needed).
- `scripts/pptx_new_fa.py` — pattern draft assembler (`--list`).
- `scripts/fa_helpers.py` — theme tokens + RTL primitives
  (`set_persian_font`, `set_bullet`, `fw_len`).

Methodology credit: storyline→pattern→adjust→FAIL-0→fresh-eye→render loop,
triage table, and accumulation practice (adopted feedback appends one rule
line) are adapted from `carnot-tech/consulting-pptx-skill` (MIT).
