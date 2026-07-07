---
name: pptxgenjs-template-generate
description: Generate editable PowerPoint .pptx decks with PptxGenJS from structured JSON specs. Use when Codex needs to create business presentations, CMB style decks, magazine or Swiss style decks, text-heavy editable slides, slides with images, charts, tables, icons, speaker notes, capacity warnings, or user-friendly Markdown outlines.
---

# PPTXGenJS Template Generate

This skill creates editable PowerPoint files with pptxgenjs. Slides must be built from native PowerPoint text boxes, shapes, lines, images, tables, charts, speaker notes, and icons. Do not rasterize whole slides into flat images unless the user explicitly asks for a non-editable image deck.

## Core Rules

- Keep output editable in PowerPoint, WPS, and Keynote.
- Write all JSON, Markdown, and script outputs as UTF-8.
- Write specs with JSON.stringify(data, null, 2) or an equivalent structured writer. Do not hand-concatenate JSON strings.
- Escape straight double quotes inside JSON strings as backslash quote.
- Do not read generator source files for ordinary deck generation. Read scripts only when debugging or changing the generator.
- Do not ignore generator warnings. Fix the JSON, reduce text, split slides, change layout, or provide missing media, then regenerate.

## Typography

Use one unified typography system across all styles:

- Chinese and mixed Chinese-English text: Microsoft YaHei.
- Pure English text: Times New Roman.
- Deck-level cover, section, and closing titles: 36 pt.
- Content-page top titles: 28 pt.
- Content hierarchy: 16 pt for point titles, 14 pt for normal body text, 12 pt for dense or small-area text.
- Chart text is normalized to Microsoft YaHei and 12 pt after PPTX generation.
- The generator enforces these tiers in scripts/pptxgen/engine.js. Do not add layout-specific font families.

## Styles

Supported styles:

- magazine: editorial layout, paper-like background, large type, image and data pages.
- swiss: grid-based business layout, flat geometry, restrained accent color.
- cmb: China Merchants Bank inspired style with CMB red themes and brand logo support.

Supported CMB themes: classic, pearl, graphite.
Supported Swiss themes: ikb, lemon, green, orange, cmb.
Supported magazine themes: ink, indigo, forest, kraft, dune, cmb.

## Workflow

1. Choose style and theme. If the user does not specify a style, use cmb for CMB or banking requests, swiss for data/product/strategy decks, and magazine for editorial or narrative decks.
2. Before writing JSON, generate a text-capacity guide for the target style:

    node scripts/generate-pptx.js --capacity-guide cmb --out outputs/cmb-capacity-guide.md

3. Write a structured JSON spec that stays within the recommended field ranges.
4. Convert the JSON spec to a user-friendly Markdown outline before final generation:

    node scripts/spec-to-md.js --spec path/to/deck.json --out path/to/deck-outline.md

5. Generate the PPTX:

    node scripts/generate-pptx.js --spec path/to/deck.json --out outputs/deck.pptx

6. If warnings mention text capacity, media slots, ignored fields, title-only items, repeated layouts, or layout incompatibility, update the JSON and regenerate.
7. Validate native/editable structure and layout safety:

    node scripts/validate-pptx-native.js outputs/deck.pptx
    node scripts/validate-pptx-layout.js outputs/deck.pptx

## Install Dependencies

Install Node dependencies once in the skill directory:

    npm install

Required runtime packages include pptxgenjs, jszip, sharp, and lucide.

- sharp is required. If it is missing, install it instead of falling back to SVG insertion.
- SVG icons and SVG logos are rasterized to PNG for better Office and LibreOffice compatibility.
- Built-in CMB logo paths can be referenced from JSON as logos/cmb-logo-lockup.png and logos/cmb-logo-mark.svg.

## Useful Commands

Generate built-in samples:

    node scripts/generate-pptx.js --sample --sample-style cmb --out outputs/sample-cmb.pptx
    node scripts/generate-pptx.js --sample --sample-style swiss --out outputs/sample-swiss.pptx
    node scripts/generate-pptx.js --sample --sample-style magazine --out outputs/sample-magazine.pptx

Write a normalized spec when using automatic layout diversification:

    node scripts/generate-pptx.js --spec path/to/deck.json --out outputs/deck.pptx --diversify-layouts --write-normalized-spec outputs/deck.normalized.json

Generate a Markdown outline from JSON:

    node scripts/spec-to-md.js --spec path/to/deck.json --out outputs/deck-outline.md

## JSON Spec Shape

Minimum spec fields:

- Top level: style, theme, title, subtitle, author, company, slides.
- Slide level: layout, kicker, title, subtitle, body, summary, conclusion.
- Collections: sections, items, columns, steps, nodes, layers, lanes, metrics.
- Media and data: image, images, media, captions, chart, charts, table.
- Notes: speakerNotes, speaker_notes, presenterNotes, presenter_notes.
- Placeholder control: allowEmptyMediaSlots true only when an empty placeholder is intentional.

## Layout Selection

Text-first layouts:

- briefing: CMB text-heavy summary with top summary, middle analysis cards, and bottom takeaway.
- textWeave: CMB asymmetric text-card layout for 1 to 6 points.
- article, sectionList, textGrid, fourCards, agenda, matrix, radial, pyramid, roadmap, timeline, pipeline, swimlane.

Media layouts:

- statement: one image slot plus a large statement. It is not a pure text layout.
- media, mediaGrid, gallery, imageGrid, imageHero, quoteImage, textImage, caseStudy.

Data layouts:

- chart, dashboard, dataSheet, bigNumbers, kpiTower, compare, duoCompare, splitCompare.

Media rule:

- If the user does not provide images, do not choose media-slot layouts, including statement, unless chart data fills the media area.
- If a blank image placeholder is intentional, set allowEmptyMediaSlots true.
- If user image count is known, choose a layout whose slots match the count.

Briefing rule:

- briefing with conclusion, takeaway, footerSummary, or nextStep supports at most 4 middle text blocks.
- briefing without a bottom takeaway supports at most 5 middle text blocks.
- Use textWeave or split slides for more points.

## Text Capacity

Always run --capacity-guide before writing JSON for a deck. Capacity ranges are calibrated for Microsoft YaHei / Times New Roman and the 36, 28, 16, 14, 12 pt tiers.

Generation performs two checks:

- JSON field-level capacity warnings from scripts/pptxgen/text-capacity.js.
- Final text-box capacity warnings from scripts/pptxgen/engine.js.

If either check warns, shorten the relevant JSON field, split content, enlarge the layout, or choose another layout. The generator warns; it does not silently rewrite or truncate user text.

## Speaker Notes

Use speakerNotes on each slide for explicit notes. Accepted aliases include speaker_notes, presenterNotes, and presenter_notes.

Do not use notes for speaker notes because some layouts render notes as visible slide content.

## Files To Read

For ordinary deck generation, read only:

- SKILL.md
- The user source material
- The deck JSON spec
- Needed assets referenced by the spec

Do not read these source files unless debugging or modifying the generator:

- scripts/generate-pptx.js
- scripts/pptxgen/*.js
- scripts/validate-pptx-native.js
- scripts/validate-pptx-layout.js

## Quality Checks

Before delivery:

1. Confirm the generated PPTX exists and has the expected slide count.
2. Run native and layout validators.
3. Check that warnings were handled or intentionally accepted.
4. For typography changes, inspect PPTX XML or generated output and confirm only the approved fonts and size tiers are used.
5. For media pages, confirm paths resolve and images are not distorted.
6. For chart pages, confirm chart data is present and text uses the unified typography.
7. Open visually when possible for dense or client-facing decks.
