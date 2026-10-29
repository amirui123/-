# hermes-pptx-skill-fa

Persian-first Agent Skill for creating, inspecting, editing, and validating PowerPoint presentations with `python-pptx`.

This repository follows the portable Agent Skills folder format and includes a Hermes-compatible `SKILL.md`. The skill adds Persian/RTL workflow guidance on top of the standard PowerPoint operations: logical Unicode order, bidi handling, Persian fonts, mixed Persian/Latin runs, digit policies, RTL tables/charts, design quality, accessibility, and structural/rendered verification.

## Install in Hermes

Install directly from the public GitHub repository:

```bash
hermes skills install https://raw.githubusercontent.com/eledah/hermes-pptx-skill-fa/main/SKILL.md --name powerpoint-fa
```

The skill uses `python-pptx` for PPTX creation and editing. Install the runtime dependency in the project or environment where you work:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/pip install python-pptx==1.0.2
```

After installation, start a new Hermes session, or use `/reset` if you want the newly installed skill to appear immediately.

Verify installation:

```bash
hermes skills list
hermes skills search powerpoint-fa
```

## Use in Hermes

The skill is available as a slash command:

```text
/powerpoint-fa Create a Persian 16:9 presentation explaining our product roadmap.
```

Or ask Hermes naturally:

```text
Use the powerpoint-fa skill to create a bilingual Persian-English pitch deck from this outline.
```

For an existing deck:

```text
/powerpoint-fa Inspect this PPTX for RTL, font, overflow, chart, and accessibility problems, then produce a corrected copy.
```

## Examples

A working showcase deck (Persian, RTL, 16:9) built with this skill, based on the
Persian Wikipedia article on the Asiatic cheetah (یوزپلنگ آسیایی). The source
is `build_showcase.py` and the rendered slides are checked in under
`examples/showcase/`; the finished `.pptx` is also included as a sample.

Title slide (hero):

![یوزپلنگ آسیایی — اسلاید عنوان](examples/showcase/01-title.png)

Every slide doubles as a worked example of one of the skill's capabilities:

| Slide | Use case |
| ----- | -------- |
| ![01](examples/showcase/01-title.png) | **Title / hero** — dark panel + framed image, RTL Persian title, mixed Persian–Latin runs, source hyperlink |
| ![02](examples/showcase/02-conservation-status.png) | **Stat cards** — conservation-status numbers, Persian digits, semantic accent colors |
| ![03](examples/showcase/03-taxonomy-table.png) | **RTL table** — explicit RTL column order, Latin scientific names as LTR runs |
| ![04](examples/showcase/04-features-rtl.png) | **RTL card list** — accent strip behind text on the RTL-leading (right) side, mixed units |
| ![05](examples/showcase/05-population-chart.png) | **Column chart** — chart fonts forced to Vazirmatn, highlighted "hope" bar, annotated takeaway |
| ![06](examples/showcase/06-habitat.png) | **Image + alt text** — framed photo with caption and accessibility description |
| ![07](examples/showcase/07-threats.png) | **Warning list** — color-coded cards (amber/red) with right-side accents |
| ![08](examples/showcase/08-conservation-efforts.png) | **Positive list** — green-accented cards, speaker notes on every slide |
| ![09](examples/showcase/09-sources.png) | **Sources & metadata** — hyperlink, feature badges, core properties |

## Local development

```bash
cd /root/workspace/pptx-skill-fa
.venv/bin/python -c "import pptx; print(pptx.__version__)"
git status
```

The current skill entry point is:

```text
SKILL.md
```

The repository also contains the research and design rationale in:

```text
research-report.md
```

## Renderer and font requirements

`python-pptx` does not render slides. For visual QA on Linux, install LibreOffice and Poppler (`soffice` and `pdftoppm`/`pdftocairo`). A Persian-capable font such as Vazirmatn, Noto Arabic/Persian, Sahel, or an organization-approved alternative must be installed on the rendering and target machines.

LibreOffice rendering is useful for regression checks but is not proof of pixel-identical Microsoft PowerPoint output. The skill requires honest reporting when rendering or font verification is unavailable.

## Scope

Supported workflows include:

- Persian, Farsi, Arabic-script, bilingual, and mixed RTL/LTR decks
- new presentations and template-based decks
- text, images, shapes, tables, charts, notes, hyperlinks, backgrounds, and metadata
- structural read-back and semantic validation
- Persian Unicode, ZWNJ/ZWJ, bidi, font, digit, table, and chart checks
- rendered visual inspection when external renderers are available

The project is currently at the initial skill-authoring stage. Helper scripts, reference files, fixtures, and automated tests will be added incrementally.

## Methodology port (consulting-pptx-skill)

The consulting workflow is ported from
[`carnot-tech/consulting-pptx-skill`](https://github.com/carnot-tech/consulting-pptx-skill)
(MIT) and RTL-mirrored for Persian decks:

```bash
.venv/bin/python scripts/pptx_new_fa.py --list
.venv/bin/python scripts/pptx_new_fa.py --parts cover,map,axis,chart,decision,back --title "نام سند" -o mydeck.pptx
.venv/bin/python scripts/check_deck_fa.py mydeck.pptx     # FAIL 0 to ship
.venv/bin/python scripts/check_layout_fa.py mydeck.pptx   # geometry audit
.venv/bin/python -m pytest -q
```

| Path | Content |
|---|---|
| `references/slide-rules-fa.md` | Rule canon (wins on conflict). Read before building. |
| `references/archetype-catalog-fa.md` | Pattern catalog (idea book, not target). |
| `references/ai-smell-fa.md` | Persian AI-tic lexicon + 30-second self-check. |
| `references/content-review-fa.md` | Blind fresh-eye prompt + adopt/reject/hold table. |
| `scripts/check_deck_fa.py` | Mechanical gate: titles, placeholders, counts, tables, RTL, AI-smell. |
| `scripts/check_layout_fa.py` | Static geometry audit (no Node needed). |
| `scripts/pptx_new_fa.py` | Pattern draft assembler. |
| `scripts/fa_helpers.py` | Theme tokens + RTL primitives. |

## Sources

- Hermes skills documentation: https://hermes-agent.nousresearch.com/docs/user-guide/features/skills
- Hermes work-with-skills guide: https://hermes-agent.nousresearch.com/docs/guides/work-with-skills
- Hermes bundled PowerPoint skill: https://hermes-agent.nousresearch.com/docs/user-guide/skills/bundled/productivity/productivity-powerpoint
- Agent Skills open standard: https://agentskills.io/
- Anthropic Agent Skills overview: https://platform.claude.com/docs/en/agents-and-tools/agent-skills/overview
- Anthropic skill authoring best practices: https://platform.claude.com/docs/en/agents-and-tools/agent-skills/best-practices
- python-pptx documentation: https://python-pptx.readthedocs.io/en/latest/
- python-pptx repository: https://github.com/scanny/python-pptx
- consulting-pptx-skill (methodology source, MIT): https://github.com/carnot-tech/consulting-pptx-skill
- Microsoft PowerPoint accessibility guidance: https://support.microsoft.com/en-us/accessibility/powerpoint/make-your-powerpoint-presentations-accessible-to-people-with-disabilities

## License

MIT

## Status

Initial skill version: `0.1.0`

Latest research report: `research-report.md`

Repository: https://github.com/eledah/hermes-pptx-skill-fa

## Security

Review a skill before installing it. Skills may contain instructions and executable scripts. This repository currently ships only the Markdown skill entry point and project metadata; inspect future scripts before running them.

## Note for contributors

Keep `SKILL.md` concise and procedural. Put detailed API, RTL, OOXML, design, and testing material into directly referenced `references/` files as they are added. Keep references one level deep, use deterministic UTF-8 output, and verify all mutating operations by reading the resulting PPTX back.

This project is independent of the bundled Hermes `powerpoint` skill; users may install either or both. `powerpoint-fa` is intended for Persian-first and RTL-sensitive work.

## Current environment

The development environment used for the initial version contains Python 3.11.15, `python-pptx 1.0.2`, pytest 9.1.1, lxml, Pillow, and XlsxWriter. LibreOffice was not installed during initial setup; `pdftoppm` was available.

This repository is under active development and should be treated as an initial, research-backed skill rather than a finished production framework.
