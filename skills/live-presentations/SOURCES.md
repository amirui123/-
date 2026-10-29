# Provenance / منابع استخراج

All skills below were fetched with `git clone --depth 1 --filter=blob:none --sparse`
followed by `git sparse-checkout set <subpaths> LICENSE`. No installer, hook, or
package manager was executed (`install.sh`, `install.ps1`, `npm install`, `pip install`
were **never** run). Temporary `.git` directories were deleted after copying.

Snapshot date: **2026-10-29**.

| Destination | Upstream repo | Subpath(s) | Commit SHA | Commit date | Declared license |
|---|---|---|---|---|---|
| `officecli-pptx` | iOfficeAI/OfficeCLI | `skills/officecli-pptx` | `89905744ed6dbb87fb6807a0d9fda6ed3da1c18d` | 2026-10-28 | Apache-2.0 (root `LICENSE` + `NOTICE`) |
| `morph-ppt` | iOfficeAI/OfficeCLI | `skills/morph-ppt` | `89905744ed6dbb87fb6807a0d9fda6ed3da1c18d` | 2026-10-28 | Apache-2.0 |
| `morph-ppt-3d` | iOfficeAI/OfficeCLI | `skills/morph-ppt-3d` | `89905744ed6dbb87fb6807a0d9fda6ed3da1c18d` | 2026-10-28 | Apache-2.0 |
| `officecli-pitch-deck` | iOfficeAI/OfficeCLI | `skills/officecli-pitch-deck` | `89905744ed6dbb87fb6807a0d9fda6ed3da1c18d` | 2026-10-28 | Apache-2.0 |
| `genoffice` | genspark-ai/genoffice | `skills/genoffice` | `a538d04295f08c061774fa0c33c64fe47efe3376` | 2026-10-29 | Apache-2.0 |
| `genoffice-executive-deck` | genspark-ai/genoffice | `skills/recommended/executive-deck` | `a538d04295f08c061774fa0c33c64fe47efe3376` | 2026-10-29 | Apache-2.0 |
| `genoffice-course-slides` | genspark-ai/genoffice | `skills/recommended/course-slides` | `a538d04295f08c061774fa0c33c64fe47efe3376` | 2026-10-29 | Apache-2.0 |
| `pptagent` | icip-cas/PPTAgent | `skills/pptagent` | `833cda553b343be0e486a93b0b57cac962cdd566` | 2026-09-20 | MIT |
| `academic-pptx-skill` | Gabberflast/academic-pptx-skill | repo root (`SKILL.md`, `content_guidelines.md`, `slide_patterns.md`, `README.md`) | `9f2b703ffe8d1449851617665ab1ffb3516d54ac` | 2026-07-14 | MIT |
| `slidev` | slidevjs/slidev | `skills/slidev` | `a42aff28dd75f26eb5abbf3d616ea022c3d08924` | 2026-10-23 | MIT |
| `frontend-slides` | zarazhangrui/frontend-slides | root skill files + `scripts/`, `bold-template-pack/` | `9906a34d640d2111f724544cbc50f7f130569ae1` | 2026-06-23 | MIT |
| `huashu-design` | alchaincyf/huashu-design | root files + `references/`, `scripts/`, `demos/`, `assets/` (`.mp3` removed) | `4fc7bfdbd5ec3fa2d43ce7a6433de60976a5e4b1` | 2026-10-23 | MIT |
| `marp-slides` | robonuggets/marp-slides | `SKILL.md`, `README.md`, `examples/` | `7ef93e98bcc728d4c2f0a8ea83b0d1ed8982ceef` | 2026-04-08 | **no LICENSE file in repo** — see `LICENSE-STATUS.md` |
| `powerpoint-mcp` | kzarzycki/powerpoint-mcp | `skills/powerpoint-mcp` | `9f767046955e7469fa1a06f9abc89cc85442d98f` | 2026-09-19 | MIT |
| `review-slide-visual` | kzarzycki/powerpoint-mcp | `skills/review-slide-visual` | `9f767046955e7469fa1a06f9abc89cc85442d98f` | 2026-09-19 | MIT |
| `powerpoint-fa` | eledah/hermes-pptx-skill-fa | `SKILL.md`, `README.md`, `references/`, `scripts/`, `requirements.txt`, `pyproject.toml`, `validate.py` | `35e083001b28b6a5f6ad0143bbc6b4d8c99f4044` | 2026-09-26 | `license: MIT` in SKILL.md frontmatter; **no LICENSE file** — see `LICENSE-STATUS.md` |

## Deviations from the original request (and why)

1. **`frontend-slides/references/` does not exist** in the upstream repo. It ships
   `scripts/` (`deploy.sh`, `export-pdf.sh`, `extract-pptx.py`), `bold-template-pack/`,
   plus `STYLE_PRESETS.md`, `animation-patterns.md`, `html-template.md` and
   `viewport-base.css` — all included instead.
2. **`marp-slides`: no LICENSE file** exists upstream (GitHub reports `license = null`).
   `LICENSE-STATUS.md` records this instead of fabricating a license.
3. **`powerpoint-fa`: no LICENSE file** at repo root; MIT is declared in the SKILL.md
   frontmatter. `LICENSE-STATUS.md` records this.
4. **`academic-pptx-skill`: the SKILL.md frontmatter says `license: Proprietary`**,
   which contradicts the repo's actual `LICENSE` file (MIT). The MIT `LICENSE` file is
   included and governs. Flagged for your legal review.
5. **Heavy sample assets dropped, as instructed**: `academic-pptx-skill`'s 190 KB
   sample PDF, `huashu-design`'s 27 MB of `.mp3` voiceover files, and
   `powerpoint-fa`'s `assets/*.jpg` + `examples/showcase/*.png|.pptx`.
6. `morph-ppt` legitimately carries 111 files (~1.5 MB) of style references under
   `reference/styles/` — included because the skill's style library depends on them.
