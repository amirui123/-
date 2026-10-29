# Slide Rules for Persian PPTX (fa) — Canon

> Adapted for `powerpoint-fa` from `carnot-tech/consulting-pptx-skill`
> (`references/slide-rules.md`, MIT). The consulting skill is the source of the
> methodology; this file is the Persian/`python-pptx` port. Where this file and
> any other doc conflict, **this file wins**. Section numbers are stable
> (`check_deck_fa.py` and other docs reference them) — **never renumber,
> only append at the end of a section**.
>
> Reading order: §0 start → §1 canvas → §2 titles → §3 subtitle ban →
> §4 layout (index at top) → §5 boxes/color → §6 tables → §7 wording
> (starts with "3 AI factors") → §8 mechanical check → §9 pitch (stage decks only).
>
> Core rule taken from the source: **don't write a new builder — reuse an
> existing builder that follows these rules, and pass the mechanical check
> (§8) before delivery.**
>
> **RTL mirror note (applies to every §4 rule that says left/right):**
> this canon is RTL-first. Persian readers enter from the **right**.
> Upstream (premise, definition, reason, fact, chart) goes **RIGHT**;
> downstream (example, result, player, meaning, So-what) goes **LEFT**.
> Step 1 sits at the highest x (rightmost); steps run right→left.
> The Japanese source says left=premise/right=conclusion for LTR — that is
> the same principle, mirrored. Never mix both orders in one deck.

## 0. Before starting (Define-before-Produce)

Standard flow — do not start from a type (steps ①→② in order):

1. **Story design:** break the problem down (issue tree), then write a
   storyline — one line per future slide. Lines include premises, facts,
   questions, not only answers. The lines must read as one speech top to
   bottom. Slide count and types come after.
2. **Type adaptation:** pick a showing-pattern per claim from
   `references/archetype-catalog-fa.md`. If none fits, compose freely.
3. **Fact & format check:** verify numbers, dates, proper nouns against
   primary sources (open the source for negative facts; record measurement
   conditions for measured values). Format = `check_deck_fa.py` FAIL 0 +
   full visual read of every rendered slide.
4. **AI-tic correction:** self-check with `references/ai-smell-fa.md` and
   clear the checker's AI-smell WARNs.
5. **Fresh-eye review:** a second agent with no context reads the deck file,
   you triage into adopt/reject/hold, fix adopted, re-run checks (§8).

- Read theme/knowledge files and primary sources first. Never invent.
- State purpose, deliverable definition, scope IN/OUT in 3–5 lines before
  building (skip only for trivial requests).
- Substance over styling. If the material is thin, say it is thin.
- Confirm background/expectation first: who asked, draft vs. submittable,
  which meeting, what decision is needed.
- **If the background is unknown, do not build a deck.** Ship one issue
  note or plain text. More slides = more AI-tic and error surface. One
  precise page beats mass production.

## 1. Canvas

| Item | Value |
|---|---|
| Ratio | **Always 16:9.** PPTX `13.333 × 7.5 in` (`12192000 × 6858000 EMU`). |
| Fonts | Brand fonts only, max 2 typefaces per deck (heading + body). Persian-capable (e.g. Vazirmatn); naming a font does not embed it. |
| Color | Brand palette, via theme tokens in `scripts/fa_helpers.py`. No off-palette colors. Only neutral/gray series may use gray. |
| Page numbers | Cover, section dividers, back cover are **not** numbered. Body pages are sequential; `→ P.n` refs are renumbered after inserts. |

Brand skin = color + fonts (+ size only when reproducing a third-party
format). It lives in theme tokens, not in this file.

## 2. Titles (message line)

1. **The conclusion goes in the title. One line by default; two lines max**
   when the message is long. ~40 full-width chars per line is the guide
   (Latin/digits count half). **Two-line conditions:** break at a meaning
   boundary / no 1–3 char orphan on the last line / never shrink type to
   squeeze into one line (go to two lines instead) / re-check body bottom
   after growing to two lines.
2. **One register, no polite padding.** Complete sentence with subject +
   verb as the default; noun-ending is allowed but not required
   (bullet-ending unity is §7.4). No slogan compression, no dropped
   subject/object. **Test: the title alone must make sense to a stranger.**
3. **No cross-slide references** (`این`, `آن`, `همان`, `تا اینجا` at title
   start). One slide must stand alone.
4. **Do not put the page's item-count in the title** ("5 steps", "3 reasons",
   "۷ بخش"). Counts are not claims. Write the structure/conclusion instead
   ("اعتبارسنجی پیش از انتخاب نوع، و بازبینی کور در پایان"). Only numbers
   that *are* the message belong in titles (amounts, %, `×8`, `2.9× در ۴ سال`).
5. **Do not color-split titles.** No accent runs inside the title.
6. **Titles alone, read top to bottom, must form one story.** Dump the title
   list before delivery (the checker prints it) and read it aloud.
7. Do not concatenate tags into titles (`Step1 …`, `مرحله ۱…` — use the
   top-corner tag chip instead).
8. **Write titles from the storyline (§0-1), never by filling the template's
   sample title.** Sample titles only show the granularity a type expects.
   One repeated sentence mold across the deck (`… است` × 8) means the
   template was copied — checker WARNs at ≥60% same mold. Placeholders
   (`متن نمونه`, `عنوان نمونه`, `Text N`, `ラベル`, `YYYY`, `◯◯`,
   `Source N`, bare `۱۲۳`) left in a delivery deck are **FAIL**.
9. **A number in the title must match the body.** "۳ مرحله" with 4 chevrons
   is the first breakage a reader spots. Checker FAILs title-count vs. max
   body ordinal.
10. **Chart pages must carry the So-what, never just the fact figure.**
    Default home is the title; if the title stays factual, the side column
    heading becomes the claim. Whether to add "so what is needed" depends
    on the topic — not a general rule.
11. **Never claim stronger than the facts, especially about the reader.**
    Shrink claims to what the evidence table supports; widen examples that
    are drawn too narrowly. Claim width must match evidence width.
12. **No `label:` prefix on titles** (`وضعیت و چالش: …`, `نقشه راه: …`).
    The topic chip (top corner) owns the topic; the title owns the claim
    from its subject onward. If deleting the prefix leaves a claim, delete it.
13. Two-line titles break at meaning boundaries with no orphan (checker
    estimates line capacity from box width × type size and WARNs).
14. **Relative words name their frame** (بیرونی/داخلی/میدانی/شرکت = from
    whose view? Name the actor).
15. **Self-talk stays modest and evidenced.** Records go in panels; titles
    stay factual. Founding motives appear once, briefly.
16. Company-intro / section entries may be bare labels (exception to the
    claim-title rule) — they present material, not arguments.

## 3. Subtitles and label rows (banned)

- **No explainer line under the title. No figure/table label rows above or
  below figures.** "What table is this" belongs in the slide title or the
  column/card heading band.
- **The single exception is the axis-title:** one line directly above a
  figure/table stating what + unit/period (e.g. `تقاضای برق داخل، GW،
  ۲۰۲۴–۲۰۳۰`), full-width with a thick bottom rule. That is the only
  permitted label row.

## 4. Layout structure

> Index (numbers are append-order; search from here):
> one-theme-per-slide 4.1, 4.29, 4.34 · two-column craft (headings, bottoms,
> heights) 4.2–4.7, 4.20, 4.21, 4.23, 4.26, 4.27 · bottom-band/frame/float
> bans 4.8, 4.9, 4.24, 4.28 · axis-tables over cards 4.10, 4.11, 4.15–4.17,
> 4.25 · process/chevron/steps 4.17, 4.18, 4.26 · deck structure
> (summary↔deep-dive, axis continuity, numbering, problem→solution,
> chapter flow) 4.13, 4.31, 4.33, 4.35, 4.43–4.45 · figures/icons/scope
> 4.4, 4.32, 4.36 · TOC/whitespace/template-unity/price-stat/review points
> 4.12, 4.14, 4.19, 4.22, 4.30 · fill/symmetry/issue-solution/scope-row/
> one-logo/focus-row 4.50–4.57 · pitch-only §9 · **reading order & placement
> (premise RIGHT, figure/table direction, anchoring, comparison arrows,
> chapter shape) 4.37–4.46** · source home / chevron-rail craft 4.47, 4.48.

1. **1 slide = 1 message = 1 theme.** Two themes → two slides.
2. **Never read top-to-bottom.** Split unrelated content **side by side**
   with a heading per section. Vertical stacking is allowed only inside one
   ordered chart (vertical step flow, §4.26).
3. **One table per slide. Never stack a table under cards.** Main
   fact/analysis table goes **RIGHT** (upstream), meaning comment with
   heading goes **LEFT** — RTL mirror of the source rule. If an upstream
   premise paragraph feeds the table, premise goes RIGHT, table LEFT
   (§4.6/§4.37 win; judge by "which is more upstream").
4. **Figure slides are two-column: figure RIGHT, comment LEFT.** No comment
   strip under the figure. **Exception:** So-what already concluded in the
   title (§2.10) *and* the figure reads alone (legend + direct values) →
   full-width figure with no side comment is allowed. Figures needing
   reading help keep the side column.
5. **Align column bottoms** (no short-side dangling whitespace; stretch/`flex`
   equivalent: equal-height boxes in `python-pptx` geometry).
6. **Columns get relationship-telling headings** (e.g. RIGHT `پیکربندی کلی
   سامانه` / LEFT `چرا این پیکربندی`). Premise+table → prose RIGHT, table LEFT.
7. **Only strong causal premise→conclusion pairs** get the horizontal filled
   triangle between columns (§4.27 shape). Plain table+comment gets none.
8. **No message band at the slide bottom** (`GOAL/KEY/POINT` strips). Bottom
   bands are allowed only for enumerating contents (logo rows, deliverable
   chips); the source line (§4.47) and detail refs (§4.31) are one-liners,
   not bands.
9. **Do not abuse frames.** Body is ground text; frame only the 1–2 elements
   that must stand.
10. **Card grids → axis-tables** (rows = item axis, cols = viewpoint axis).
    Bundle the left axis into meaningful coarse groups.
11. **Never compare two boxes side by side.** A-box/B-box → rows=viewpoints,
    cols=compared table.
12. **Thin slides center vertically.** Big-figure slides top-align. **Never
    stretch vessels/leading to fill the page** (§5.13: fill-rate comes from
    content volume, not stretching). Order: (a) add info or change type,
    (b) else center and split whitespace top/bottom.
13. **Separate "summary" and "deep dive."** Chapter head: whole-map summary
    (each element with `→ P.n`); deep dives carry chapter tag chips linking
    back. **Mandatory for body >10 slides:** open whole-map, chip every
    later slide, connect titles (`بر این بستر…`, `… پیش‌گفته`), read as a
    speech script. **≤10 slides or single-issue decks get no whole-map**
    (§0 one-issue note, §4.45).
14. **TOCs are one column.**
15. **Parallel cards/boxes align widths and gaps exactly** (compute geometry;
    uneven lengths/gaps are instant review failures).
16. **Parallels also align granularity + subject.** Same-viewpoint sentence
    shape (`subject + what-it-does`) on every card.
17. **Step×enabler goes in a table,** not stacked boxes. Left col = stage,
    right col = enabler. Unreadable top-4/bottom-3 box piles are banned.
18. **Standalone processes default to chevron (RTL: step 1 rightmost).**
    Process+comment pages use vertical step-flow + side comment (§4.26),
    never a second content block under a horizontal chevron.
19. **One template per deck.** Different headers/margins/type systems across
    halves are banned; normalize merged sources to one template.
20. **So-what lives in the side section, never the bottom band.**
    RIGHT = fact/chart, LEFT = meaning/implication.
21. **Side-column headings share one top line.** With a table on the right,
    the left heading matches the `th` top padding, size, weight, and thick
    underline so baselines and rule-y match. Headings stay one line (shorten
    wording instead of wrapping).
22. **Price/cost stat boxes sit at the bottom of the side column.**
    Top = heading + points, bottom = price.
23. **One heading per column.** Two stacked heading bands in one column =
    the same axis twice. Inner structure uses L2 bullets (indent), not
    second headings.
24. **The last slide carries no message.** Close with a **back cover**
    (logo + company + contact only). CTAs/conclusions (`۳ تصمیم موردنظر`)
    are normal content slides *before* it. No 3-dark-card CTA page.
25. **No 2+-sentence prose inside table/row cells.** Split into `•` bullets.
    One sentence (+ short complement) may stay inline.
26. **No second content block under a horizontal chevron.** Process+comment
    = LEFT-comment / RIGHT-vertical-step-flow two-column (RTL mirror).
    Horizontal chevrons only when the chevron alone *is* the slide's point.
    Axis-titles crossing both columns use one full-width thick rule; the
    crossed side then carries bullets without its own heading (only case
    where a side column has no heading).
27. **Premise→conclusion pairs with strong causality get the horizontal
    triangle** (RIGHT premise, center triangle, LEFT `consequence`
    consequence). Craft: equal columns/heights, one triangle centered in the
    gutter at vertical middle, never one triangle per row, never clipped
    into the left element, filled triangle only (no block arrows). Forgetting
    the triangle on a premise/conclusion pair is the most common violation.
28. **No floating objects.** Labels/chips/mini-elements must live inside a
    table/column/chart structure, never drift on the page (classification
    labels become row/column headers).
29. **Risk slides always pair concern with countermeasure.** Never list risks
    alone. Different-nature items (`why risk is low` vs. `risk response`)
    split into separate slides (§4.1).
30. **Review comments/homework never stay on body slides.** Collect on a
    closing "review points" slide as the answer TOC. Client-review summaries
    alone go to Appendix.
31. **Forward-reference details with a footnote** (`مدیر A/B — جزئیات صفحه
    بعد`), the small form of §4.13 summary↔dive.
32. **People/documents in figures are monochrome single icons** (human/doc
    SVG + label, singular actors). Speech bubbles get tails and complete
    exchanges (request → response naming the mechanism).
33. **Carry declared axes through later slides.** After declaring
    topics/divisions in a map/tree/summary: (a) number branches
    (`论点1` → `محور ۱`), (b) pin all-division tag chips top-corner on later
    slides with the current one inverted, (c) reuse identical numbers/names
    in chips. Free-text chips after a declared axis lose readers.
34. **Dense single pages split into staged build-ups (2–3 slides).** Same
    page, elements added — **layout must not move 1mm** on page 2 or the
    effect dies.
35. **Chapter details carry the index page's numbers** (number+division+that
    page's claim in a side panel, or `تقاضا:` prefixes on titles).
    Numbers/names match the index character-for-character.
36. **Narrowed scope is said on the narrowed figure:** show all components,
    frame the one handled with a `محدوده این سند` note. Shorter and less
    ambiguous than prose disclaimers.
37. **Premise/definition/reason RIGHT; example/result/player LEFT.**
    Readers move right→left as why→what→who. Three columns:
    RIGHT reason / MIDDLE example table / LEFT conclusion. **When lost, the
    more upstream item goes RIGHT.**
38. **Align table and chart order/direction.** Same content as table +
    stacked chart → match row order to stack order, defining side first
    (RIGHT in RTL). Mismatched eye-shuttling kills the claim.
39. **Anchor right-column points onto the left figure with bands+labels.**
    The debated rows get thin bands ending in the same point names, bolded
    on the right too. Missing bars get added (no claim without its bar).
40. **Ratio/delta comparisons draw the arrow** (dot on base, arrow to target,
    legend `پیکان: مبنا ← مقایسه` in RTL order). `۴–۵ برابر` alone never
    states base vs. target.
41. **Ordered rows get a direction axis** (wedge thickening downward +
    `مقدار لازم بیشتر می‌شود`). Ordered-but-axis-less rows read as plain lists.
42. **Comparisons reuse the earlier figure's axis as a checkable table**
    (same value-chain layers as rows, `○△—`). Same-axis tabling exposes
    over-claims and shrinks claims to evidence.
43. **Problem pages pair with solving pages titled to be crossable**
    (`گلوگاه … با … قابل عبور است`). Index rows get `گلوگاه n P.xx` tags;
    deep sides get `گلوگاه ۲/۴` chips for round-trips (§4.13/§4.35).
44. **Chapter stories advance one step per page** (`تقاضا بالاست` →
    `عرضه‌کننده محدود است` → `هزینه سنگین است` → `بی‌عملی رقابت را می‌برد`).
    Never two steps per page.
45. **Section doors:** same ground as cover, `SECTION n / N` + one-column
    chapter list, current chapter bold. Doors are unnumbered. **But never
    force chapters:** ~10-page decks need no doors.
46. **Figure examples are plural, real-named, reader-relevant first.**
    Value-chain columns list ~7 known names; reader-related firms align to
    the top row. Coverage bands stretch to true coverage.
47. **Sources home: bottom-left** (one line above the footer). Never in
    titles/table names. Leave the line empty on numbered/quote pages.
48. **Value-chain/process rails are thin full-width chevron bands;**
    explanations sit in ruled columns below, never crammed inside the band.
    First segment is a pentagon (no clipped left end in RTL).
49. **Two-content column headings must read alone as noun phrases.**
    No conjunction-led headings (`پس، …`, `بنابراین …`) — causality lives
    in the triangle + body. **This ban is headings-only;** table cells/cols
    read along a row axis may carry `پس …` consequents.
50. **Filled cards only when the color means something** (legend-sayable:
    one emphasized step, grayed non-applicable). Explanation vessels are
    white ground + heading + thick underline + bullets (§5.2). No dark
    decoration cards, no fill-framed comparisons.
51. **Comparisons are same-width, same-structure.** Mirrored dark-1 vs.
    3-tier halves are "shapely-wrong". Same width/height, heading bands,
    row counts, row structures (label+explanation); row-mapped comparisons
    become tables (§4.11).
52. **Issue→solution case pages are two columns `چالش | راه‌حل` (RIGHT→LEFT
    in RTL), never Before/After timelines.** LEFT solution column carries a
    row axis (who/what/how-fixed/who-uses); the challenge side is bullets
    without a forced axis. Headings are noun phrases (§4.49).
53. **The pivotal stage stands alone in one recolored chevron.** The turning
    segment (updates stop after delivery, …) splits out as its own chevron,
    the only recolored one, meaning clear via legend or in-chevron heading.
54. **Bottom-half whitespace is fixed by shrinking vessels or adding
    content — with more points and more characters.** Priority: (a) align
    table/card cells to 3 points each (§6; §7.19 varied-angles前提),
    (b) raise body toward 12pt, (c) else §4.12 centering. Filling with
    conclusion bands/decor is banned (§4.28).
55. **Lower elements must show their relation to the upper.** Unrelated lower
    blocks (`position of the lower element unclear`) join the card structure
    as labeled rows (vertical flow §4.26) or the upper columns extend down
    and the lower block disappears.
56. **One logo per page.** No header mini-logo + body giant logo together.
    Brand display = footer line (§8) or one body instance.
57. **Focus rows get highlight wash.** In 2-row contrasts where the proposal
    handles one side only, wash that row thin and state the meaning
    (legend/footnote) — the table form of §4.36 scoping.

## 5. Boxes, color, legends, rules

1. **Rectangles by default.** No `roundRect` / `border-radius` except small
   status pills (<0.4 in tall). Ported decks replace them with `RECTANGLE`.
   (Current showcase cards are grandfathered; new patterns use rect.)
2. **Recolored boxes need an on-slide legend** (color chip + meaning).
   Meaningless "emphasis" dark cards are banned.
3. **Filled boxes get no border.** Fill already marks the region; borders
   belong on unfilled (white) elements only. PPTX: filled cards carry no
   `a:ln`. HTML: `.card{border:0}`.
4. **Rules only where they divide something.** No rule under the last table
   row, none at the slide bottom (footer rule excepted). `枠線＋塗り`
   (border+fill) and destination-less rules read as AI-made decor.
5. **Comparison highlights go on the winning side,** never always-ours.
6. **Competitor comparisons always include an axis rivals win.**
7. **Measured in-house values carry method notes** (use case, n, verifier).
8. Meaning colors (✕ red, add green) survive rebranding. Product UI
   screenshots stay unedited.
9. **Palettes include monochrome (gray).** Never all-series chromatic;
   neutral/base series take gray, chromatic marks only the stood-up series.
10. **Bounded-metric charts draw full-scale** (correlation, %, attainment:
    0–max with reference lines as needed). Floating bars that erase scale
    context hide the claim (weak correlation, low progress).
11. **Trends/compositions/distributions are graphs.** Never substitute
    stacked numbers or tables. Trend/ratio reads at a glance only as a
    figure. Needing a graph means abandoning the template and drawing it.
12. **Charts come from base forms, not freehand.** Base: ① table-chart
    ② premise-LEFT… (RTL: premise-RIGHT) + meaning chart
    ③ contrasted sides ④ analysis-LEFT… (RTL: analysis-RIGHT) + commentary.
    Then applied forms (vertical step-flow, roadmap, Before→After, 2×2,
    pyramid, value-chain, org, Harvey-ball table, waterfall, funnel).
    **Fit ①–④ first; applied forms only on misfit.** One chart = one
    message; no creative charting; body type in charts matches body size.
    Data-chart shape guide: composition = pie, item compare = horizontal bar
    (RTL: value axis mirrored deliberately), time = vertical bar/line,
    distribution = vertical bar, correlation = scatter.
13. **Fill-rate: body uses ≥55% of the title-to-source vertical band.**
    Flow/decision/现状-vs-ideal types hollow the bottom easily. Fixes:
    (a) add So-what/premise/notes to the side column, (b) size figures to
    the page, (c) change to a fitting type. Checker WARNs <55%.
    **Never stretch few characters/rows to kill the WARN**; accept centering
    (§4.12) when content cannot grow. **Decor to fill space is banned**
    (§4.28, §7.13).

## 6. Tables

| Item | Rule |
|---|---|
| Header | **≥2pt larger than body (guide +3–4pt; body 11.5pt → header 14.5pt), bold, same ink as body** (never thin gray). No background fill. **Axes are rules, not fills:** thick ink rule under header row. |
| Rows | No zebra, no faint emphasis rows. Hairline thin-gray row rules (never base-color rules). |
| Rules | Header = thick ink; rows = 1px thin; **no rule under the last row** (§5.4). |
| Columns | No abbreviations/invented words. Each cell reads alone as natural language. No orphans (widen/shorten/break at meaning). |
| Body size | **Table/row-list body ≥11.5pt, th = body+2pt, row-axis (edge column) = body+4pt. Diagram body ≥9pt.** Early-10s pt is "too small" projected. Footnotes and node labels only may be smaller. Re-render everything after raising type. |
| Short-table stretch | Stretching short tables vertically → `vertical-anchor middle`. Top-pinned stretch shows dangling whitespace. |
| Detail-cell thickness | **Evidence/practice/countermeasure cells with 2+ contents list 2–3 `•` bullets.** Rows of single-sentence cells read thin. Never pad by paraphrase (needs §7.19 varied angles first; §4.25 pair). |
| Sources | Never in titles/table names. Bottom-left home (§4.47). |
| N/A cells | **`—` gray-washed de-emphasis.** Skippable regions fall back; readable cells stand. |
| Key cells | Cells driving judgment/action (published rules, approval gates) go **bold**. Never uniform density. |
| Few-item definitions | ~3 legend-like definitions stay **`label: explanation` one-liners**, not nested bullets (§7.12 is for repeated joins). |
| Axis display | **Row-axis (edge column) shows the axis: heavy rule + body+4pt type.** Axis labels are self-identifying words (no vague `کلی`). |
| Rows obey axis | **Rows contain only members of the edge axis set.** Axis `demand sources (industries)` → industry rows only. Macro facts, our claims, conclusions, foreign categories move to side bullets/another slide. **Self-check every row against the axis aloud before shipping** (`is national balance a demand source? → no → out`). |
| Columns minimal | **No claim-irrelevant columns.** A responsibility table's `track-record` column starts record-comparisons; records move to logos/another page. Droppable layers drop with a source-line note. |
| Headings concrete | **No abstract column headings** (`fact`, `assets` → `held assets / scale & content`: name what the cell holds; §7.2 pair). |
| N/A glyph | Non-applicable layers are `—`, never blank. Same-meaning washes keep one meaning deck-wide with paged legends (`unmanned layer — evidence P.8`). |
| Column granularity | Defined columns keep one granularity/viewpoint; **never mix categories** (`how-to-use` col must not smuggle `how-to-update`; §6 row-axis pair). |
| No invariant facts | Delta tables carry **only what changed**; invariants compress to one note line. |
| Staffing tables | Rows = roles, cols = points (initial/now/next FY); cells `N名: name…` **count first**; close with `total (延べ)` and `actual heads (deduped, exclusions named)` as two rows. |
| Conclusion rows/cols | **No off-grain `پس …` conclusion row as the last row.** Same-grain rows only. A consequent col/cell reading along a rowed axis is allowed (row = viewpoint, cols = issue/solution/what-becomes-needless). Same ban applies to floating two-content headings (§4.49) but not row-fixed tables. |
| No claim headings | Column headings are **noun phrases** (`چالش`, `راه‌حل`, `دلیل ناممکن‌بودن`). Conclusions go to titles. |

RTL table policy (Persian addition, normative): RTL paragraph text and
visual/logical column order are **two separate problems**. LibreOffice
ignores table-level `rtl`, so visual RTL order is built by **explicit column
placement** (RTL-leading column rightmost). Document per table whether columns
are logically preserved or visually mirrored.

## 7. Wording and labels

> **3 AI factors** (eye-check these three before shipping):
> 1. **Floating boxes** — drifting labels/chips/minis outside any
>    table/column/chart structure (→ §4.28)
> 2. **Verbose prose-dumps** — paragraphs where bullets belong; enumerate as
>    `•`, one claim per line (→ §7.3, §4.25)
> 3. **Odd compressions** — subject/verb-dropped crunch-words
>    (`手動作成` → `دستی‌سازی`, `精度向上` → `افزایش دقت`). Titles and body
>    must read who-does-what (→ §2.3, §7.2, §7.14)

1. **Few parens.** Complements via `／` or sentences. Label separators use
   `:` (never ` — ` joins; §7.9/§7.12).
2. **No self-made shortenings/inventions** — one-read natural words only.
3. **Bullets:** shred prose paragraphs. Kill few-char orphans. Markers are
   L1 `•` + hanging, L2 `–` + hanging (`・` never a marker).
   **Never type markers as text** (`•  `, `–  ` inside runs forces receivers
   to reformat every slide; PPTX uses `buChar`, non-list uses `buNone`;
   checker FAILs typed `•`). **One-sentence blocks get no bullet**
   (exec-summary lines, Before/After singles, standalone statements —
   bullets only where 2+ items enumerate).
4. **Same-level endings unify. No polite-form drift.** All noun-ending or
   all verb-ending per level; mixed `…توصیه` + `…ببینید` in one level is
   banned; unfinished endings (`…به‌صورت سه‌بعدی`) banned. Per-level typing
   (L1 verb-stop / L2 noun-stop) is allowed.
5. Integrate footnote-class micro-notes into the body where possible.
6. **One doc, one term (no drift):** scan full text before shipping
   (checker WARNs representative pairs). **Abbreviations defined at first
   use** (`workspace (WS)`), then short form allowed. Never undefined
   in-house inventions. **Same-sound different-meaning words stay distinct.**
   Never ship raw AI text; re-inspect per §7.15 + §8 fresh-eye.
   Persian scope adds: ی/ک Arabic-vs-Persian, ZWNJ vs. space, and digit
   system (§3 procedure) obey the deck's declared policy deck-wide.
7. **Metaphorical/invented terms take `「」` + a definition** somewhere
   (same handling as §7.6 abbreviation definitions).
8. **Tell the final audience what to check.** Review decks are not feature
   enumerations; they are check-me/decide-me stories. Judge when a meeting
   explanation + question base suffices — exhaustive builds explode effort.
9. **Remove AI-smell words/phrasings.** Empty intensifiers, katakana-pileups
   (Persian: buzzword piles), fixed closings, three-set syntaxes,
   emoji/traffic-light decor are classic AI tells. Full lists + 30-second
   self-check: `references/ai-smell-fa.md`. Checker WARNs high-precision words.
10. **Function/process bullets start with the subject (user/AI).** Not
    `دستی بساز` but `ساخت دستی توسط کاربر`, `تولید خودکار توسط AI`. For
    risk/governance audiences, **AI automatic moves always pair with human
    checkpoints** (`AI proposes, user confirms and registers`). Sentences
    reading as AI registering/updating alone always get flagged in regulated
    reviews.
11. **No internal metaphors/jargon in external decks.** In-house names open
    to general words on first external use; kept metaphors pair `「」` +
    definition (§7.7).
12. **Structure, not symbol-joins.** Repeated `label — explanation` /
    `label: explanation` across cells/cards/bullets = missing structure.
    Split columns (request/result), heading bands, or indent levels instead
    (§4.10, §4.25). One-off label cuts may use `:`.
13. **Numbers/symbols/labels only when referenced later.** Round-number chips
    and `pillar 1/2` numbering mean something only when later slides or
    speech call the number back. Unreferenced one-shot numbers are decor —
    use bare names.
14. **Body and bullets keep subject + verb too** (§2.3 titles rule extended).
    Noun-crunched words (`تقویت …`, `پاسخ به …`, `خودکارسازی …`) hide
    who-does-what — open into `actor + action` (`کاربر جریان را ثبت می‌کند`,
    `AI نقص را می‌یابد و مسئول تأیید می‌کند`). When tightening, cut
    modifiers, keep subject/verb. **Writers never coin new abstract-subject
    passives** (`…خودکار می‌شود`); open into `who becomes able to do what`.
    Already-established source terms stay as-is (§7.7 condition: does the
    word hold outside this doc?).
15. **A human re-reads AI Persian before delivery.** Correct content with
    unnatural Persian still fails as a doc. **Titles (key sentences) are the
    LLM weak point** — §2.6 title read-through checks naturalness besides
    sense, via §2.3/§7.2/§7.14 + `ai-smell-fa.md`.
16. **Exec summaries map 1:1 to body titles/chapters, each line `1 main +
    2–3 same-level details`.** Each line matches one body page (or chapter)
    and reads as a TOC. Details parallel with varied angles
    (current-state/reason/consequence), bolding only the single most
    effective line. **Never print structure names** (`current-state/reason/
    consequence`) **or reading instructions** (`bold = …`) on the page
    (maker-side words). Downstream elements appearing only at the deck end
    still get their final line.
    **No single-child nesting.** `main → 1 child → 1 grandchild` layering is
    meaningless indentation. Indent only with 2+ parallels per level;
    undividable content stays flat or folds into one sentence.
17. **Delete one-line notes under figures when the page reads without them.**
    Settle figure-reading with title + legend; no補足 sentences (§3/§4.28
    floating-element pair, wording form).
18. **4+ parallel bullets with a natural upper grouping bundle into 2
    parents** (parent + 2 children each reads faster than flat 4 numbers;
    §4.23 inner-indent pair). Same-axis enumerations (regions, quarters,
    departments) never bundle (bundling falsifies structure). Test: does the
    parent heading mean something alone?
19. **Bullets vary angles per point. No paraphrase-stacking.** Same-claim
    rewordings ×3 read as thinness-by-count. Vary axes (cost/time/people-org/
    decision/means/volume) and ask per point `does deleting it lose info?` —
    cut lossless points. Never pad for count-alignment.
20. **No taken-for-granted facts as bullets.** Reader-known lines
    (`change needs budget and approval`) are not information — swap for the
    unnoticed concrete mechanism. Test: one new info per line for a domain
    reader.
21. **No slang/exaggeration/undefined abbreviations in body.**
    Reader-stopping words get swapped (§7.2/§7.9 applied).
22. **No Latin-caps decor kickers in Persian decks.** Topics use Persian
    top-corner tag chips (§2.12); Latin stays for brand/product/proper nouns
    and sources. Cover English (program names etc.) excepted.
23. **AI automatic moves pair as `AI proposes, human decides`.**
    (§7.10 regulated-audience pair, same shape.)

## 8. Mechanical check before output (mandatory)

```bash
python3 scripts/check_deck_fa.py deck.pptx        # FAIL 0 to ship
python3 scripts/check_layout_fa.py deck.pptx      # geometry audit
# render audit when soffice/poppler exist:
soffice --headless --convert-to pdf --outdir render deck.pptx
pdftoppm -png render/deck.pdf render/slide
```

`OLD_COLORS` (banned legacy HEX, no `#`) at the top of `check_deck_fa.py`
activates the banned-color check when a brand lists legacy colors (default
empty = no check). All other checks are brand-common.

Checklist (fix every FAIL before delivery; only brand-skin size FAILs from
§1 may be waived):

- [ ] 16:9 size (`12192000×6858000` EMU)
- [ ] Titles non-empty, single register, **≤2 lines** (>80 full-width-equiv
      FAILs; >40 WARNs two-line expectation; never shrink-to-fit per §2.1)
- [ ] No placeholders (`متن نمونه` / `عنوان نمونه` / `Text N` / `YYYY` / …;
      §2.8) · no mold-clustering (§2.8)
- [ ] Rect default (large `roundRect` WARNs; pills excepted) / no
      `subtitle`/`figttl` remnants
- [ ] No listed legacy colors (`OLD_COLORS`)
- [ ] Table headers read as headers (header body+2pt, bold, same ink;
      px-only HTML rule becomes run-size comparison in PPTX) · thin-gray
      heading color WARNs
- [ ] No Latin-caps decor kickers / no conjunction-led two-content headings
      (checker WARNs)
- [ ] No last-row rule / no bordered filled boxes
- [ ] Title numbers match body ordinals (`۳ مرحله` vs. max `مرحله N`; §2.9)
- [ ] **Title list printed → human reads the story aloud for connection**
- [ ] Page numbers sequential (cover/back/doors excluded; §4.45); `→ P.n`
      refs renumbered after inserts
- [ ] No item-counts in titles (§2.4)
- [ ] **3 AI factors** eye-checked on every page (§7 head)
- [ ] **Fresh-eye reviewed** (blind second agent reads the file → triage
      table → fix adopted → re-run mechanical checks)
- [ ] Client-facing deliverables passed an experienced reviewer (§0)

**Layout audit** (`check_layout_fa.py`): footer/body overlap, title
right-overflow (silent-clip catcher), slide overflow.

Then eye-QA (PPTX: `soffice` PDF → `pdftoppm` PNGs; HTML: Chrome PDF →
`pdftoppm`) for orphan wraps, overflows, overlaps, left/right bottom
alignment, legends, **page balance** (read whitespace, not only characters).
**Before fixing review feedback, reproduce on the latest canon first**
(reviewers often read stale versions).

**Fresh-eye review (mandatory final step).** After mechanical FAIL 0, run
`references/content-review-fa.md`: blind agent → findings (wording/logic/
breakage/craft) → triage table (prompt tail format; adopt/reject/hold one
by one; client decides meaning-changing rewordings/restructures) → fix
adopted → re-run checks. Reasons stay one line for rejects; the table ships
with the delivery report. Blind re-reading never substitutes self re-reading.

**Review-loop count is the user's choice:** one rally = mechanical → eye-QA
→ fresh-eye → adopted-fixes → re-check. **Stop after 1 rally, show product
+ table, ask before rally 2. Never self-run rally 2+ unasked.** Agree scope
up front (§0): **speed** = mechanical + self-eye only (no fresh-eye) /
**standard** = 1 rally / **quality** = N rallies until asked count (stop
early at zero findings). Report per rally in 3 lines
(findings/adoptions/remaining concerns).

**What eyes catch that machines miss** (passing shape, broken content):
title-vs-figure conclusion clashes on the same page; unevidenced evaluators
(`limited`, `sufficient`, `no problem` over bare countermeasure lists);
late-page repeats of early pages.

## 9. Pitch decks (on-stage, 7–10 min)

> Only differences from proposal decks (§0–§8) live here. Common rules still apply.

1. **A pitch is not an explainer. "Rebuild for pitch" means rewording +
   story order, not type-swapping.** Flowing into white explainer types
   yields a "sales explainer" rejection. If the base has 1-message pages,
   big numbers, strong visuals, **keep the design, fix order + wording**.
   Rebuilds discard order, not looks.
2. **Set page count from stage time first.** 7 min → 8–10 body pages
   (~40 s/page). 16 pages (~26 s/page) is long. Issues cap at 2 pages
   (problem → why-it-happens). Company intro is 1 page.
3. **Standard order** (1 page each; doors at most 2 — `cases`, `proposal` —
   the §4.45 ~10-page exception):
   1. Catch (§9.4) 2. company intro (bare-label title; **newest master**
      version, never a stale embedded copy) 3. what the product is
      (definition sentence as title; before why-it-was-made — never motive
      before object) 4. why made (`problem / why-it-happens` 2-row table;
      §4.57 focus-row highlight for later focus) 5. value behind the problem
      6. mechanism the product loops (make→use→history→improve; AI moves
      pair with human judgment §7.23) 7. cases (published → field →
      trial-calc marked 試算/برآورد) 8. proposal (example tasks → product
      parts → demo → join-us: process/deliverables/asks table) 9. appendix:
      peripheral cases leave the body. No brand/product-specific title
      examples in this file.
4. **Catch page 1 = `large logo + product screen LEFT … (RTL: RIGHT) +
   1-line headline + 1-line complement`.** No 3-point summary strip below;
   kill header mini-logos, keep one giant logo (§4.56). Headline is one
   subject+verb sentence stating the product core.
5. **No competitor table in pitches.** Comparisons belong to Q&A appendices
   or sales decks. Avoid vendor-name exposure.
6. **Kill Latin kickers/subtitles/decor** (§3, §7.22). Spoken decks cut
   10–20% body and go single-register. Proposal-ese
   (`as we understand it`, …) goes.
7. **Modest self-talk** (§2.15–§2.16). Confirm `we-made-it` record handling
   with the owner for juried decks. Mixed named/anonymous cases in one deck
   pre-decide naming (same-firm identifiability).
8. **Submission/demo consistency:** trial calcs footnote `برآورد/想定`,
   never mixed into record parts. Write office deadline + stage minutes into
   Define-before-Produce (§0).

---

*Attribution: methodology adapted from `consulting-pptx-skill`
(github.com/carnot-tech/consulting-pptx-skill, MIT). "Created with
consulting-pptx-skill" style credit belongs on the back cover bottom-left /
final sources row only when that project's material is reused verbatim —
never on other pages. This port adds the RTL mirror, Persian wording/digit/
font policy, and `python-pptx` mechanical checks the source does not cover.*
