# Fresh-Eye Review (blind second agent) — fa prompt

> Adapted from `consulting-pptx-skill`
> (`references/content-review-prompt.md`, MIT). **Mandatory final step**
> after mechanical FAIL 0, before delivery. Catches what the maker cannot
> see: subtle Persian wording, logic jumps/twists, title↔figure mismatches,
> unevidenced evaluators, re-skinned earlier pages.

## How to run

1. **Hand over the deck file as-is** (`.pptx` path; PDF renders as
   supplement only when page-balance/overflow must be seen — e.g.
   `pdftoppm -png -r 60 deck.pdf pages/p`). The reviewer reads the file
   with Read/text extraction. No image-ification needed.
2. Pass the instruction below to **a fresh agent with the method hidden**
   (new Agent session; never reveal the skill, rules, templates, or what
   you cared about). Replace `{FILE}` and `{N}`.
3. Turn findings into a **triage table** (below). Fix one by one —
   adopt / reject / hold — never fix-all. Meaning-changing rewordings and
   restructures go to the requester for the call.
4. Fix adopted, re-run `check_deck_fa.py` + layout check. Rejects keep a
   one-line reason (`taste`, `obvious to readers`, `factually true`, …).
5. Ship the table with the delivery report (what was flagged, fixed, not
   fixed — at a glance).

### Triage table format

| # | Page | Finding (verbatim) | Kind | Call | Action / reason |
|---|---|---|---|---|---|
| 1 | P.7 | «پوشش یافته است» — completive for ongoing coverage | wording | adopt | → «پوشش یافته / تحت پوشش است» |
| 2 | P.9 | conclusion not derivable from premise table | logic | adopt | rewrite LEFT heading as standalone claim |
| 3 | P.3 | «بازیگران خارجی» vague | wording | reject | deliberate hypernym, defined P.2 |

Kinds: `wording` / `logic` / `breakage (numbers/figures/values)` / `craft`.

---

## Instruction (copy from here)

You are a presentation reviewer. Read one deck ({N} pages) and flag
**breakage the maker cannot see**. You are told nothing about how it was
made. Do not guess the method.

### Target

File `{FILE}` ({N}-page deck, PPTX/PDF). **Read the whole file, page by
page in order.** Do not skim-and-guess. Page images, if attached, are for
page-balance only. Do not read any other files (rules/skills/templates).
Judge by yourself.

### Check first (2 points — report with page numbers)

- **Persian wording:** anything a Persian speaker trips on. Completive vs.
  progressive mixups, slang (`به‌صرفه نیست`), subject/verb-dropped
  compressions (`ساخت دستی`, `افزایش دقت`), un-unified bullet endings,
  unreadable long noun-stack modifiers, undefined abbreviations/inventions,
  same-concept spelling drift (ی/ک Arabic-vs-Persian, ZWNJ vs. space,
  digit-system drift). **Flag unnatural Persian even when content is right.**
- **Logic:** read titles alone top-to-bottom — jumps, returns, repeats,
  conclusions without premises. Within a page: does RIGHT (premise/reason)
  → LEFT (consequence) read? Does the LEFT heading derive from the RIGHT?
  Are comparison bases (what-vs-what) explicit?

### Always check breakage (report with page numbers)

1. **Title↔body number mismatch:** `۳ مرحله` title over 4 chevrons,
   `۴ محور` over 3 items, …
2. **Title↔figure conclusion clash:** title limits/denies what the figure
   or final box on the same page asserts.
3. **Unevidenced evaluators:** `محدود`, `کافی`, `بدون مشکل`, `حداقلی`
   sitting on a page with no evidence for them.
4. **Re-skinned earlier pages:** late pages repeating early content with no
   new point.
5. **Map/TOC vs. body drift:** announced chapters/point-counts/`→ P.n`
   refs vs. actual body.
6. **Number inconsistency:** same quantity with different values across
   pages, bad totals, unit/period drift mid-deck.
7. **Missing sources/notes:** numbers without source rows; estimates/trials
   without `برآورد` disclaimers where needed.

### Score axes (10 each · 70 total)

1. Story flow: titles alone as one claim — does it hold?
2. One message per page: clear claims; no stuffing, no thinness.
3. Information structure: comparisons/causes/parallels in fitting vessels
   (tables/flows/charts/2-col)? Graph-shaped data not dumped as numbers?
4. Page craft: whitespace balance, alignment, type size, overflow.
5. Persuasiveness: usable as-is for judging/objecting/acting?
6. Consistency: craft, terms, labels aligned deck-wide?
7. Honesty: source/limit disclaimers fitting? Non-misleading?

### Output format

- Wording + logic findings as **page no. · quoted span · what's wrong ·
  rewording proposal** each (many welcome, 10+ fine).
- Breakage 1–7 hits as **page no. · quoted span · what clashes** each.
  Write `none` for clean items.
- Per-axis scores + 1–2 sentence reasons each.
- Total.
- **Fix-ordered finding list** (top 5). 1–2 sentences each ending in how.

Write unflattered. List at least 3 defects. No deck is spotless.
