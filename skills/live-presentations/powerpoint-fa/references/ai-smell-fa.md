# AI-Smell Lexicon (Persian fa) — wording tics to remove

> Adapted from `consulting-pptx-skill` (`references/ai-smell-lexicon.md`,
> MIT) + the "3 AI factors" in `slide-rules-fa.md` §7. Source of the method;
> word lists below are rewritten for Persian decks (the Japanese lists are
> not portable). Referenced from `slide-rules-fa.md` §7.9. **After building,
> run the词汇 scan + 30-second self-check before delivery.**
> (`check_deck_fa.py` WARNs the high-precision subset.)

## 0. Review核 (what actually gets flagged)

- **Verbose — cut first (guide: 400 chars).** Long = instant rewrite ask.
- **Over-helpful = "monkey-obvious" prose.** Rephrasing everything politely
  is itself the tell (not the words — the redundancy).
- **Roundabout, not straight.** Kill hedged, modifier-heavy indirection.
- **Stamp-loop fixed phrases** — same stock response repeated (`ثبت شد` × N).
- **Decor-as-substance** — traffic-light tables, emoji icon headings.

## 1. Words / phrasings (remove)

- Empty intensifiers: دقیقاً / بسیار / فوق‌العاده / بی‌نهایت / عظیم /
  انقلابی / پیشگام / نسل بعدی / «اغراق نیست اگر بگوییم» / «چیزی نیست جز»
- Buzzword piles: راهکار / سینرژی / یکپارچه / سرتاسری / ارزش افزوده /
  بینش / بهینه‌سازی / توانمندسازی / چابک‌سازی (as fashion words —
  plain verbs first: استفاده‌کردن، ساختن، رساندن)
- Self-beautifiers / slogans: همراهی / هم‌قدم / دوشادوش / با تمام توان /
  به‌سوی قله‌های بالاتر / «دنیایی که می‌خواهیم»
- Stock pillows: در این راستا / در خصوص / در رابطه با / با توجه به /
  در عصر حاضر / در دنیای امروز / همان‌طور که مستحضرید / به شرح ذیل
- Over-polite: «لطفاً … بفرمایید» chains / «خواهشمند است» every line /
  «در صورت امکان» padding
- Abstract paddings: از منظر / از حیث / جنبه / فاز / نکته کلیدی / کلیدی /
  محوری (when content-free)
- Evaluators without evidence: محدود / کافی / مناسب / بدون مشکل / حداقلی
  (allowed only with on-page evidence — `slide-rules-fa.md` eye-QA trio)

## 2. Syntax / grammar (break up)

- Mechanical templates: fixed `conclusion→reasons→summary` × N /
  `نکات ۳ مورد است` frames / `اولاً・ثانیاً・در پایان` numbered narration
- Report-dressing: everything heading + bold + numbered list + icon
- Roundabout: long subject–verb distance / «امکان … وجود دارد» (→ می‌توان) /
  «می‌باشد» bloat (→ است, only where register allows)
- Triplet hammering: «…، … و … می‌کند» chains / stacked noun-endings /
  «نه‌تنها A، بلکه B، حتی C»
- Conjunction overload: همچنین / علاوه‌بر / بنابراین / ازاین‌رو / بدین‌ترتیب
  at every sentence head (80% deletable)
- Translationese: passive piles, every-subject-spelled, English word order
- Hedge chains: «می‌توان گفت» / «به نظر می‌رسد» × N
- Over-nominalized passives: «بهبود … صورت گرفت» → «… را بهبود داد»
- Dash joins:補足 via ` — ` chains → period / `:` / parens

## 3. Tone / endings (cut the tail)

- Fixed closings: «لطفاً بررسی بفرمایید» / «پیشاپیش سپاس» /
  «در صورت سؤال در خدمتم» / «نظرتان چیست؟» / «حتماً بررسی کنید»
- Empty nods/flattery: عالی است / کاملاً درست است / نکته ظریفی است /
  (write it right the first time instead)
- Decor glyphs: emoji/symbol heading decor (✉🔒✅📋⚠👉▪) / traffic-light
  3-color tables / bold-everything
- Zero-wobble: zero `!` or mechanical `!` everywhere / zero hedge
  colloquialism in a perfect official register
- Hype/fill: «داده بیاید درست می‌شود» optimism / too-tidy honor-student
  prose with no body temperature

## 4. Information design

- Verbose/MECE-everything: split-all-fill-all / disclaimers longer than body /
 要点 unfocused bullet piles → cut first (400-char guide)
- Human opposite: short · 1 sentence = 1 line · points only / honest
  plain assertion / concrete numbers-dates-options / occasional `!` warmth

## 5. 30-second self-check (before ship/send)

1. Deletable words (intensifiers, conjunctions, pillows, fixed closings) cut?
2. Glyph/emoji/bold decor removed?
3. 1–2 colloquial touches mixed (not zero-wobble perfect officialese)?
4. Read aloud: "my own words"?
5. Long → cut to ~400 chars?

## Operation

- Slides: applied as `slide-rules-fa.md` §7.9. `check_deck_fa.py` WARNs the
  high-precision subset (context-justified hits are eye-waived).
- **Split with slide-rules:** wording tics live here; vessel-specific
 指摘 (layout/tables/title structure) live in slide-rules. §7.9 is the door.
- Every adopted review指摘 appends one line here (same accumulation
  practice as slide-rules).
