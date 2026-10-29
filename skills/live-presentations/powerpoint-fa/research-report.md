# Research Report: Persian-First Hermes PowerPoint Skill

Status: initial research baseline

## Executive summary

The bundled Hermes PowerPoint skill is a strong operational baseline: it already covers creation, read-back, editing, templates, charts, notes, slide reordering/duplication, and render verification. The main opportunity is not merely adding more API names. It is adding a disciplined pipeline that separates:

1. content and narrative planning,
2. Persian/RTL text preparation,
3. theme/layout construction,
4. python-pptx implementation,
5. structural validation, and
6. rendered visual inspection.

The Persian skill should be Persian-first in its defaults, but must explicitly state which behavior comes from python-pptx, which requires OOXML, and which depends on PowerPoint/LibreOffice/font availability.

## Grounded baseline

The official Hermes skill page identifies version 1.1.0 and provides five helper-script families: create, read, edit, template filling, and rendering. Its verification contract is outline read-back plus rendered slide review when LibreOffice and Poppler are available.

The official python-pptx documentation identifies version 1.0.0 documentation, while the isolated environment installed for this repository currently has python-pptx 1.0.2. The package is industrial-grade in intent and supports PPTX round-tripping, slides, placeholders, text, images, tables, auto-shapes, common charts, core properties, and notes.

## What python-pptx gives us

### Presentations and slides

- `Presentation()` and `Presentation(path)`
- `prs.slide_width`, `prs.slide_height`
- `prs.slide_layouts`, `prs.slides.add_slide(layout)`
- slide masters, layouts, placeholders, notes slides
- core properties such as title and subject

### Text

Text has three layers: `shape.text_frame`, `TextFrame.paragraphs`, and `_Paragraph.runs`.

Useful public controls include:

- `TextFrame.clear()`, `.text`, `.paragraphs`, margins, wrapping, vertical anchor, auto-size
- paragraph alignment, level, indentation, line spacing, before/after spacing
- run font name, size, bold, italic, underline, color, hyperlink
- `run.hyperlink.address`

The skill should teach run-level construction rather than assigning large strings blindly, because mixed Persian/Latin text, emphasis, hyperlinks, and formatting preservation require predictable run boundaries.

### Visual objects

- `slide.shapes.add_textbox(...)`
- `add_picture(...)`
- `add_shape(...)` with `MSO_AUTO_SHAPE_TYPE`
- tables via `add_table(...)`
- charts via `add_chart(...)`
- grouped shapes and connectors when supported by the relevant API

### Charts

Documented chart data models include `CategoryChartData`, `XyChartData`, and `BubbleChartData`. Common chart types include column/bar, line, pie, scatter, and bubble; most non-3D types are supported.

Useful controls include:

- `chart.chart_type`, `chart.plots`, `chart.series`
- category/value axes and tick labels
- axis bounds, gridlines, number formats
- legends and positions
- data labels and their positions
- series smoothing and markers where supported

Design rules should prefer bar/column/line charts for comparisons and trends, restrict pie charts to a small number of parts, and always provide a textual takeaway. Chart colors should come from a coherent theme rather than ad-hoc per-series colors.

### Notes and metadata

`slide.notes_slide` creates/accesses a notes slide; `notes_slide.notes_text_frame` supports speaker notes as rich text. Notes should be a first-class output in the skill: include narrative intent, source citations, speaking cues, and accessibility descriptions when useful.

## Important limitations to document

- python-pptx does not expose every PowerPoint feature. Advanced effects, animations/transitions, SmartArt, embedded media workflows, theme authoring, some accessibility metadata, and several chart features may require OOXML edits or a different tool.
- Slide copying across presentations is not a safe public API operation. Same-deck duplication and XML-level relationship cloning are specialized operations and must be tested.
- Slide reorder/removal is not fully represented by a simple public API and may require XML manipulation.
- Chart data replacement is broader than surgical editing; wrappers that offer per-series operations should read, modify, and replace the complete dataset, clearly documenting normalization risks.
- A shape may have no title or text frame. Code must check `has_text_frame` and handle blank layouts where `slide.shapes.title is None`.
- A `.pptx` file is a ZIP package of XML parts and relationships. XML-level changes can silently create broken relationships; every such helper needs round-trip tests.
- Rendering is not available through pure python-pptx. LibreOffice/Poppler are external dependencies and may render differently from Microsoft PowerPoint.

## Persian and RTL requirements

The current python-pptx public text API exposes alignment and typography controls but does not provide a high-level Persian/Arabic RTL authoring layer. Therefore the skill needs a dedicated RTL adapter and explicit caveats.

Recommended pipeline:

1. Normalize input conservatively: preserve meaningful ZWNJ, normalize accidental Arabic/Persian character variants only when the user asks for normalization, and never silently change identifiers, URLs, or code.
2. Apply paragraph-level right alignment for Persian content.
3. Set RTL/bidi-related OOXML properties through a narrowly scoped helper when the public API cannot express them; keep the helper isolated and covered by XML tests.
4. Keep Persian and Latin/numeric spans in separate runs when directionality or styling changes.
5. Use a declared font policy and verify fonts on the rendering host. Suggested defaults should be configurable, not hard-coded: Vazirmatn, Sahel, IRANSansX, B Nazanin, or another user-approved installed font.
6. Avoid relying on client-side font substitution. Record the selected font in deck metadata or a manifest.
7. Test mixed strings such as `نسخه v1.2 — API` and Persian numbers, dates, URLs, and file paths.
8. Render in both LibreOffice and, where available, Microsoft PowerPoint before shipping; RTL layout, shaping, and font metrics can differ.

The skill must not claim that setting `PP_ALIGN.RIGHT` alone guarantees correct bidi order. Alignment is visual positioning; RTL/bidi semantics, run direction, numerals, and shaping are separate concerns.

## Design system recommendations

### Narrative

Require a slide brief before implementation: audience, decision/action, one-sentence message per slide, evidence, and intended reading order. Prefer one idea per slide. Use a story arc appropriate to the deck type: context → problem → insight → evidence → options → recommendation → next steps.

### Layout

Define a small tokenized layout system rather than freehand coordinates everywhere:

- widescreen 16:9 by default;
- safe margins and baseline grid;
- title, subtitle, body, chart, image, quote, comparison, timeline, and closing layouts;
- consistent header/footer/slide number behavior;
- reusable spacing, radius, stroke, and elevation tokens;
- named semantic zones to support reading order and validation.

### Typography

Set font family, size hierarchy, weight, line spacing, paragraph spacing, and maximum text density explicitly. Persian body text generally needs more vertical room than Latin text. Avoid shrinking text to fit; split or simplify content first. Treat long URLs, Latin acronyms, and numbers as special mixed-direction content.

### Color and charts

Use theme tokens with contrast checks. Do not encode meaning only by color. Use direct labels where possible, restrained gridlines, zero baselines when appropriate, units and number formats, and source notes. Make chart titles state the conclusion rather than merely naming the chart.

### Accessibility

Add meaningful alt text/description where OOXML supports it, preserve intentional reading order, use real title/body placeholders when possible, maintain sufficient contrast, and include speaker notes or a text outline for important visuals. Microsoft accessibility guidance recommends readable sans-serif fonts, 18pt or larger body text where feasible, whitespace, contrast, and correct reading order.

## Proposed repository architecture

```text
hermes-pptx-skill-fa/
├── SKILL.md
├── README.md
├── LICENSE
├── pyproject.toml
├── requirements.txt
├── references/
│   ├── rtl-persian.md
│   ├── python-pptx-api.md
│   ├── design-system.md
│   ├── charts-and-data.md
│   ├── templates-and-masters.md
│   ├── ooxml-escape-hatches.md
│   ├── accessibility.md
│   └── troubleshooting.md
├── scripts/
│   ├── pptx_create_fa.py
│   ├── pptx_read_fa.py
│   ├── pptx_edit_fa.py
│   ├── pptx_rtl.py
│   ├── pptx_validate.py
│   ├── pptx_render.py
│   ├── pptx_from_template_fa.py
│   └── pptx_manifest.py
├── templates/
│   ├── minimal-fa.json
│   ├── report-fa.json
│   └── pitch-fa.json
├── examples/
│   ├── rtl-mixed-text.py
│   ├── charts-fa.py
│   └── notes-and-accessibility.py
├── tests/
│   ├── test_rtl_xml.py
│   ├── test_text_roundtrip.py
│   ├── test_charts.py
│   ├── test_template_fill.py
│   ├── test_manifest.py
│   └── fixtures/
└── .github/workflows/ci.yml
```

## Priority roadmap

### P0: reliable foundation

- Write a detailed SKILL.md with explicit workflow gates.
- Port the bundled helper scripts while preserving its good JSON CLI contract.
- Add RTL text helpers, font checks, Unicode/ZWNJ diagnostics, and mixed-direction test fixtures.
- Add structural validator: slide count, missing titles, overflow-risk heuristics, text density, contrast, image existence, chart data consistency, notes, and source fields.
- Add a manifest containing skill version, python-pptx version, fonts, renderer, theme, and source references.

### P1: design quality

- Add tokenized theme/layout primitives.
- Add content-density and slide-message checks.
- Add chart recommendation and formatting helpers.
- Add visual regression fixtures and rendered PNG inspection workflow.
- Add Persian examples with realistic dates, numbers, tables, and mixed Latin text.

### P2: advanced interoperability

- Add narrowly scoped OOXML adapters for bidi, alt text, custom metadata, and selected theme/reading-order operations.
- Add template inventory and placeholder diagnostics.
- Add optional PowerPoint/LibreOffice cross-render comparison.
- Add documentation of unsupported features and safe fallbacks.

## Verification contract

Every generated deck should pass:

1. package open/save round-trip;
2. outline extraction and expected-content assertions;
3. RTL/mixed-text XML assertions;
4. font and resource checks;
5. chart/table data checks;
6. render check when a renderer is available;
7. visual review of every rendered slide for clipping, overflow, overlap, contrast, and incorrect directionality;
8. final report of limitations, renderer, fonts, and unresolved warnings.

## Sources

- Hermes bundled skill: https://hermes-agent.nousresearch.com/docs/user-guide/skills/bundled/productivity/productivity-powerpoint
- python-pptx documentation: https://python-pptx.readthedocs.io/en/latest/
- python-pptx text guide: https://python-pptx.readthedocs.io/en/latest/user/text.html
- python-pptx charts guide: https://python-pptx.readthedocs.io/en/latest/user/charts.html
- python-pptx notes guide: https://python-pptx.readthedocs.io/en/latest/user/notes.html
- Microsoft accessibility guidance: https://support.microsoft.com/en-us/accessibility/powerpoint/make-your-powerpoint-presentations-accessible-to-people-with-disabilities
- python-pptx source repository: https://github.com/scanny/python-pptx

## Expanded subagent findings

The four-agent review adds the following details to the initial baseline.

### Public API coverage

The implementation should document and exercise these API families: `Presentation()`/`Presentation(path)`, `prs.slides`, `prs.slide_layouts`, `prs.slide_masters`, `prs.slide_width`, `prs.slide_height`, `prs.core_properties`, `prs.notes_master`; slide/layout/placeholder access; text-frame, paragraph, run, font, and language metadata; textboxes, shapes, pictures, tables, connectors, groups, freeforms, movies, and OLE objects; chart data models and chart/plot/axis/series formatting; backgrounds, hyperlinks, notes, and core properties.

This capability matrix is important because python-pptx can preserve many unsupported Open XML parts while loading and saving them, but preservation does not mean those parts are safely understandable or editable.

### Confirmed public-API gaps

The skill should explicitly mark these as XML-level, unsupported, or client-dependent rather than implying they are ordinary python-pptx operations: deleting, reordering, or duplicating slides; copying slides between presentations; editing comments; authoring animations, transitions, timings, SmartArt, sections, arbitrary theme/master XML, advanced presentation settings, and changing chart type after creation. Chart updates are generally dataset replacement through `Chart.replace_data()`, not cell-level workbook editing.

Media and OLE support exist but depend heavily on correct package parts and the consuming Office client. Graphic-frame rotation, shadows, and other advanced formatting have limited public support.

### Persian/RTL implementation contract

Persian content must remain in logical Unicode order. Never manually reverse strings or pre-shape them into presentation-form glyphs. Preserve ZWNJ (`U+200C`) and ZWJ (`U+200D`) deliberately, use Persian code points such as `ی` and `ک` according to the chosen normalization policy, and test combining marks and punctuation-heavy mixed-script strings.

For Persian paragraphs, combine `PP_ALIGN.RIGHT` with a tested DrawingML paragraph bidi patch (`a:pPr rtl="1"`). Alignment is not equivalent to direction. Keep Persian, Latin, URLs, numbers, and code-like fragments in separate runs when their font, language, emphasis, or directional behavior differs. Use directional isolates only for genuinely ambiguous fragments; do not scatter invisible marks throughout normal prose.

`MSO_LANGUAGE_ID.FARSI` is useful language/proofing metadata but is not a bidi switch and does not guarantee a font. Persian font declarations may need both `a:cs` and `a:latin` typeface slots. Font names are not embedded automatically, and fallback changes metrics, wrapping, marks, and chart/table geometry.

Tables need an explicit policy for logical versus visual column order. RTL paragraph properties inside a cell do not automatically mirror the table. Charts require independent tests for category order, legend order, tick labels, axis position, data labels, and horizontal-bar direction. Geometry, arrows, timelines, and diagrams do not mirror automatically when text becomes RTL.

### Design and accessibility gates

The skill should require audience, purpose, constraints, delivery context, desired action, and a one-sentence central message before slide generation. Each slide needs one dominant idea and a claim-style title. Use a small semantic palette, test contrast, avoid color-only meaning, and prefer progressive disclosure over dense manuscript slides. Microsoft and WCAG guidance support meaningful titles, alt text, intended reading order, sufficient contrast, readable fonts, whitespace, non-color cues, captions where needed, and simple tables with headers.

### Architecture refinement

Keep `SKILL.md` concise and procedural; put API, RTL, design, schema, template, OOXML, accessibility, and visual-QA detail into references. Add a tested `pptx_validate.py`, a Persian RTL adapter, manifest generation, realistic mixed-script fixtures, unit/integration tests, and CI. Use deterministic UTF-8 output and `pathlib`. Make mutating operations read-back verifiable and make renderer-based checks a release gate when available, without claiming PowerPoint-equivalent output from LibreOffice.

## Additional python-pptx study findings

The focused API study confirms several implementation priorities for version 1.0.2:

- Geometry is left/top-based and should be expressed with `Inches`, `Cm`, and `Pt`; a Persian layout abstraction should mirror semantic zones rather than asking callers to manually reverse every coordinate.
- There is no complete first-class API for paragraph direction, bidi, Kashida, language metadata, or OpenType shaping controls.
- RTL-related OOXML may involve paragraph properties and script/font slots such as `pPr`, `a:rtl`, `a:latin`, and `a:cs`; these should be isolated behind namespace-qualified helpers and treated as version-sensitive.
- `fit_text()` and automatic measurement should not be trusted as the primary Persian layout strategy. Prefer conservative dimensions, explicit line spacing, and content-density checks.
- Charts, tables, and placeholders need separate RTL fixtures because their directionality behavior can differ from ordinary textboxes.
- Explicit textboxes are often preferable to arbitrary placeholders when precise Persian geometry and directionality matter, while placeholders remain valuable for template inheritance and accessibility.

## Research note

Some details—especially exact bidi OOXML placement and PowerPoint-versus-LibreOffice rendering behavior—must be validated against fixture files in the implementation phase rather than inferred from the high-level python-pptx API documentation. The implementation should treat those as test-backed compatibility layers, not assumptions.

## Environment verification

- Repository: `https://github.com/eledah/hermes-pptx-skill-fa`
- Python: 3.11.15
- Installed package: `python-pptx 1.0.2`
- Installed test runner: pytest 9.1.1
- Available renderer component: `/usr/bin/pdftoppm`
- LibreOffice (`soffice`) was not present during initial setup.
