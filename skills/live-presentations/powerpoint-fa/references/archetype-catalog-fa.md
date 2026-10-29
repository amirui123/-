# Pattern Catalog (idea book) — Persian RTL edition

> Adapted from `consulting-pptx-skill` (`references/archetype-catalog.md`,
> 62 types, MIT). **Patterns are "layout idea books", not targets.**
> The canon is `references/slide-rules-fa.md`; patterns only serve it faster.
> Pick for the story; discard on misfit; freely recombine inside the rules.
>
> Read timing: when writing the "showing-how" per storyline line. For
> eye-shopping, render `assets/SlideCatalog-fa.pdf` (planned) instead of
> reading this file end to end.

Assemble drafts with:

```bash
python3 scripts/pptx_new_fa.py --list
python3 scripts/pptx_new_fa.py --parts cover,map,chevron,premise,axis,chart,decision,back --title "نام سند" -o mydeck.pptx
```

Page numbers are renumbered in build order. Cover / section doors / back
cover are unnumbered (`slide-rules-fa.md` §4.45).

## 1. Core patterns (build first — cover ~80% of decks)

| ID | Pattern | When to use | Notes (RTL) |
|---|---|---|---|
| `cover` | Cover | Title + context + hero visual | Dark panel + framed image (showcase slide 1). Latin only for program names. |
| `map` | Whole-map overview | Body >10 slides: all elements + `→ P.n` refs | Each row links forward; later slides chip back (§4.13). |
| `toc` | Table of contents | Long decks needing a 1-column TOC | One column only (§4.14). |
| `door` | Section door | Chapters in >10-page decks | Same ground as cover, `SECTION n / N`, current bold, unnumbered. |
| `chevron` | Process chevron (RTL) | Standalone stages/transitions | **Step 1 rightmost**, running right→left. Single recolor only for the pivotal stage. |
| `vsteps` | Vertical steps + side comment | Process + commentary on one page | RIGHT vertical step-flow, LEFT comment (RTL mirror of §4.26). Never stack content under a horizontal chevron. |
| `premise` | Premise → conclusion | Strong causal pair | RIGHT premise (table/fact), center single filled triangle, LEFT consequence with standalone noun-phrase heading (§4.27). |
| `axis` | Axis table | Item × viewpoint organization | Rows = items, cols = viewpoints; header thick rule, no fill/zebra (§6). RTL-leading column rightmost. |
| `compare` | Choice comparison | Options × criteria | Highlight the winning side, include an axis rivals win (§5.5–5.6). |
| `risk` | Risk list | Risks + signs + countermeasures | Never risks alone; different-nature rows split slides (§4.29). |
| `stat` | Large-number readout | Scale/impact numbers | One big number + short reading; Persian digits per deck policy. |
| `chart` | Single chart + insight | Prove one claim with one figure | RIGHT chart, LEFT insight; So-what mandatory; graph never table-dumped (§5.11). |
| `bars` | Ranked horizontal bars | Item comparison | Descending; single-brand emphasis bar; base→target arrows for ratios (§4.40). Mirror value axis deliberately for RTL. |
| `matrix` | 2×2 matrix | Two-axis positioning | Neutral series gray; focus cell washed + legend (§4.57). |
| `harvey` | Harvey-ball / heatmap table | Sufficiency/status overview | `○◐●` density marks; N/A `—` gray-washed. |
| `roadmap` | Roadmap | Execution timeline | Chevron rail thin + full-width; explanations in ruled columns below (§4.48). RTL time flows right→left. |
| `gantt` | Gantt | Schedules | Rows = roles/work, cols = points; counts first. |
| `decision` | Decision page | Ask for a decision | What to decide + premises + asks (noun-phrase headings). |
| `case` | Issue / solution | Case pages | RIGHT `چالش` bullets / LEFT `راه‌حل` with row axis (RTL mirror §4.52). |
| `back` | Back cover | Close | Logo + company + contact **only**. No message, no CTA (§4.24). Credit line (`consulting-pptx-skill` reuse / sources) lives here bottom-left when needed. |

## 2. Extended patterns (add on demand)

Executive summary (1:1 to body titles, `1 main + 2–3 details` per line,
§7.16) · evidence basis · KPI dashboard · stacked bars · waterfall/bridge
(strict) · small multiples · scenario table · horizontal-axis table ·
process×viewpoint matrix · nested-row matrix · time matrix · theme cards ·
recommendation pillars · numbered imperatives · SCR (situation/complication/
resolution) · issue→cause→solution flow · issue tree · current→target state ·
calc flow · cycle · value chain · decision fork · research basis.

## 3. Usage rules

- Every `h1`/title is replaced from the storyline (`slide-rules-fa.md`
  §2.8). Zero placeholders survive (`متن نمونه`, `Text N`, `YYYY` FAIL).
- Colors/type come from `scripts/fa_helpers.py` theme tokens (single
  source). Change once, apply deck-wide.
- Never force a slide into a type. Split dense pages into staged build-ups
  (same page, added elements, **zero layout shift**), or change the type.
- Review feedback that generalizes becomes one appended rule line in
  `slide-rules-fa.md`, same accumulation practice as the source skill.

## 4. Coverage map vs. current showcase

The 9-slide cheetah showcase (`build_showcase.py`) demonstrates
`cover`, `stat`, `axis`, `chart`, image+alt, warning/positive card stacks,
sources/metadata. It does not yet demonstrate `map`, `chevron` (RTL),
`premise`+triangle, `compare`/`risk`, `matrix`, `roadmap`/`gantt`,
`decision`, structured `door`/`toc`. Build those next as `python-pptx`
builder functions — do not import the source's HTML templates as runtime
(the source itself keeps PPTX out of its pipeline; its `.pptx` is a
hand-copy sourcebook, which is exactly the role for a future
`templates/SuperTemplate-fa.pptx` here).
