---
name: artifact-template-academic-word
description: "Create, write, finish, or reformat Chinese academic papers, theses, and course reports as submission-ready Word documents using the retained Academic Word template. Use for 论文、课程报告、实验报告、毕业论文、学位论文、Word 完稿、套模板、论文排版、目录题注交叉引用、GB/T 7714 参考文献，or when the user explicitly invokes $artifact-template-academic-word. Supports drafting from supplied facts and sources or formatting an existing DOCX; never invents evidence or citations."
---

# Academic Word 论文与课程报告完稿

Produce a final `.docx` that needs no manual formatting adjustment. Treat factual completeness and layout completeness separately: only write claims supported by the user's materials and only deliver after structural and visual QA pass.

## Resolve bundled files

Resolve all paths relative to this skill directory.

- `artifact-template.json` points to the privacy-clean course-report master.
- `assets/style-authority.dotx` is the privacy-clean style authority. It retains the accepted OOXML styles and settings but contains no source manuscript text or embedded objects.
- `assets/clean-report.docx` is the default generic course-report master.
- `assets/clean-thesis.docx` is the HUST thesis master.
- `references/style-contract.json` contains the exact fallback style tokens.
- `references/request-schema.md` defines the normalized request.
- `scripts/academic_word.py` provides `analyze`, `build`, and `validate` commands.

Load the bundled workspace document runtime before running Python. Use the returned Python executable and package paths; do not use system Python or install packages into a user environment.

## Route the request

Choose one mode:

1. `draft`: write from the user's topic, outline, facts, data, sources, and assets, then build the DOCX.
2. `format`: preserve an existing DOCX's content and media while applying the approved style system.

Choose one document type:

1. `course_report` by default. This is school-neutral and suitable for any university, course, laboratory, or general academic report. Use the generic cover and never add HUST branding, declarations, or degree language.
2. `hust_thesis` only when the user explicitly requests the HUST thesis shell or supplies HUST thesis requirements.

Resolve the cover choice before building. Use the generic course-report cover by default; use the HUST cover only with `hust_thesis`; or preserve an external institutional/course cover when the user supplies that template. Never infer HUST mode from this skill's formatting provenance. Record the choice as `requirements.cover_choice` with `generic`, `hust`, or `external`.

For `hust_thesis`, collect every Chinese/English cover field listed in `references/request-schema.md`. The retained front matter is three physical pages: Chinese cover, English title page, then originality/copyright declarations. Do not omit those pages or translate names, majors, degrees, and dates by guessing.

If the user supplies an external template or written formatting rules, run `analyze` on the template and compare it with `references/style-contract.json`. Present all material conflicts together and wait for one resolution. Do not silently choose between conflicting page geometry, fonts, cover structure, heading numbering, citation style, or page-number zones. After confirmation, record the decision in `requirements.conflict_resolution` and keep it fixed for that run.

For a structured external template containing its own cover, front matter, and sections, use the Documents template/base-replace workflow with that file as the base. Preserve the supplied cover rather than regenerating the generic or HUST cover. Use this skill's style contract only for unspecified roles. For a style-only template, overlay its styles after mapping every used source style.

## Collect complete inputs

Normalize inputs to the schema in `references/request-schema.md`. Ask for all missing required information in one batch.

Never leave placeholders in a final file. Never invent authors, course names, student IDs, advisers, experimental results, statistics, quotations, publication metadata, DOIs, URLs, or references. If required evidence is absent, stop and request it.

Default citation behavior is GB/T 7714—2015 sequential numbering. Replace `[@citation-id]` markers deterministically and build the reference list from structured citation data. Accept another citation style only when explicitly requested; preserve an existing manuscript's citation style unless the user asks to change it.

In the retained template, in-text sequential citations are square-bracket numbers in superscript, for example superscript `[1]`; bibliography numbers remain baseline text. Do not render in-text citations at the normal baseline.

## Build

Write the normalized JSON request to a temporary/work directory and run:

```bash
"$PYTHON_BIN" scripts/academic_word.py build \
  --request /absolute/path/request.json \
  --output /absolute/path/final.docx
```

The builder must:

- use real paragraph/character/table styles and real multilevel numbering;
- copy `word/styles.xml`, `word/numbering.xml`, `word/theme/theme1.xml`, `word/fontTable.xml`, and `word/webSettings.xml` directly from `assets/style-authority.dotx`; do not recreate an equivalent style system;
- remove `stylesWithEffects.xml` and its relationship because Microsoft Word can prefer it over the retained DOTX `styles.xml`;
- map heading paragraphs to the DOTX's actual numeric style IDs `10`, `2`, `3`, and `4`; headings must not contain direct `w:numPr` because the original template binds all levels through style-level `numId=23`;
- number headings as `1`, `1.1`, `1.1.1`, and `1.1.1.1` using the untouched DOTX numbering definition;
- use the retained DOTX heading contract: Heading 1 黑体/Arial 16 pt centered, Heading 2 黑体/Arial 14 pt justified, Heading 3 黑体/Arial 12 pt justified, and Heading 4 黑体/黑体 12 pt justified, with the exact spacing recorded in `references/style-contract.json`;
- use exact retained section geometry in DXA rather than rounded millimetres: margins 2552/1588/1588/1588, header 851, footer 964, first-section page code 9, column spacing 425 then 720, and `linesAndChars` document grid with 312-DXA pitch;
- copy the retained `settings.xml` behavior for 420-DXA tabs, Chinese punctuation compression, zh-CN East-Asian language, compatibility, math, and drawing defaults, adding only `w:updateFields=true`;
- whenever retained `settings.xml` contains `footnotePr`/`endnotePr`, also package the template's separator-only `footnotes.xml` and `endnotes.xml` with valid document relationships and content-type overrides; missing companions cause Microsoft Word's “显示修复：脚注1/尾注1” error;
- apply the retained header/footer/page-number styles (`ad`, `af1`, `af2`) in the header/footer story parts; do not simulate them with direct borders, font sizes, or paragraph centering;
- let ordinary body runs inherit the retained `宋体小四` paragraph style; do not repeat equivalent direct font and size formatting on every run;
- divide pagination into cover without a page number, front matter/TOC with uppercase Roman numerals starting at I, and body text with Arabic numerals restarting at 1;
- use TOC, PAGE, SEQ, REF, PAGEREF, and STYLEREF fields where applicable;
- include only Heading 1 and Heading 2 in the TOC (`TOC \\o "1-2"`); Heading 3 and Heading 4 remain numbered in the body but must not create cached `TOC3`/`TOC4` rows;
- keep cached TOC entries faithful to the retained mixed-script sizing: use direct `w:sz=21` only and never write direct `w:szCs`; TOC numbers and Latin text must inherit 12 pt from the template style;
- create figure captions as an inline `aff3` caption paragraph with a visible `STYLEREF 1 \\n` chapter-number field and a visible `SEQ 图` field; never type the chapter/sequence number or place a caption in a text box;
- use inline images, fixed-DXA table geometry, repeating table headers, and three-line academic tables;
- create simple equations as OMML and keep the equation and its number together;
- use section properties instead of blank paragraphs for pagination or orientation;
- set `w:updateFields=true` and preserve only portable fonts: 宋体, 黑体, Times New Roman, Arial, and Cambria Math;
- write to a new output file and never overwrite source material.

After rendering establishes the real page for each heading, materialize verified cached TOC page numbers without replacing the TOC field:

```bash
"$PYTHON_BIN" scripts/academic_word.py cache-toc \
  --input /absolute/path/final.docx \
  --page-map /absolute/path/toc-page-map.json \
  --output /absolute/path/final-with-toc-cache.docx
```

Revalidate and rerender the cached result. Do not deliver a TOC whose visible cached page numbers are blank or stale.

For `format` mode, preserve paragraph text, tables, images, equations, hyperlinks, footnotes, and section intent. Do not rewrite prose merely to make style application easier. Use the source file as `content.source_docx` and pass it again to validation for a text-preservation check.

## Refresh fields and verify

Run the structural gate:

```bash
"$PYTHON_BIN" scripts/academic_word.py validate \
  --input /absolute/path/final.docx \
  --request /absolute/path/request.json \
  --source /absolute/path/source.docx \
  --json-report /absolute/path/validation.json
```

Omit `--source` for draft mode. A nonzero exit is a hard failure. Fix every error before rendering.

The structural gate must also compare `webSettings.xml`, compare `settings.xml` after ignoring only the required `updateFields` addition, inspect style references in headers and footers, and reject any one-DXA page-geometry drift, wrong document grid, duplicate body header/footer parts, direct header/footer simulations, or body-run font overrides.

The structural gate must reject note declarations whose footnote/endnote parts, relationships, separator IDs, or content-type overrides are incomplete. LibreOffice accepting such a package is not sufficient because Microsoft Word repairs it on open.

The structural gate must byte-compare the final styles, numbering, theme, and font-table parts with the retained DOTX. It must reject `stylesWithEffects.xml`, nonnumeric heading style IDs, or direct heading numbering. LibreOffice rendering alone is not sufficient evidence for heading numbering: when Microsoft Word is installed, open and inspect a generated four-level heading fixture in Word before delivery.

When Microsoft Word is available on macOS, use `scripts/refresh_in_word.applescript` on a temporary copy. The document has `w:updateFields=true`, so Word refreshes fields while opening and saving. The script has a 45-second Apple-event timeout. If Word automation is unavailable, keep `w:updateFields=true`, refresh with LibreOffice when possible, and disclose that Word's native field-refresh gate could not run.

Render the latest DOCX with the Documents renderer. Inspect every page, not a sample. Check cover balance, missing glyphs, clipping, overlaps, table wrapping, captions, equation numbers, headers/footers, Roman/Arabic page-number transitions, orphan headings, isolated lines, unexpected blank pages, and stale field results. Fix, revalidate, and rerender until clean.

Use conservative WPS-compatible OOXML: no NoteExpress/EndNote ADDIN fields, OLE/Visio embeddings, floating legacy shapes, undocumented fonts, theme-only colors, or layout tables with hidden content. Word/LibreOffice visual QA plus this safe subset is the compatibility gate when WPS is unavailable; do not claim pixel-identical WPS rendering.

## Deliver

Deliver only the final `.docx` unless the user explicitly asks for QA artifacts. The final file must contain no placeholders, comments, tracked changes, plugin fields, embedded OLE objects, or internal citation tokens.
