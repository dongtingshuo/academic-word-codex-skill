# Academic Word Codex Skill

[中文](README.md)

A Codex skill for Chinese academic papers, theses, and course reports. It turns user-supplied topics, evidence, citations, assets, and existing Word documents into submission-ready `.docx` files with structural validation and page-by-page visual QA.

![Template preview](artifact-template-academic-word/assets/preview.png)

## Capabilities

- `draft`: create a complete Word document from a topic, outline, factual material, and sources.
- `format`: preserve the text, tables, images, equations, and links in an existing manuscript while applying the accepted format.
- Use native Word styles, four-level numbering, TOC fields, captions, cross-references, OMML equations, and section-based page numbering.
- Show only Heading 1 and Heading 2 in the TOC while retaining automatic Heading 3 and Heading 4 numbering in the body.
- Use GB/T 7714—2015 numeric citations by default, with superscript square-bracket citations in the body.
- Use a generic course-report shell by default; HUST-specific front matter is enabled only when `hust_thesis` is explicitly selected.
- Validate styles, numbering, citations, captions, page numbering, repeating table headers, note parts, placeholders, revisions, and non-portable objects.

## Installation

```bash
git clone https://github.com/dongtingshuo/academic-word-codex-skill.git
mkdir -p ~/.codex/skills
rsync -a academic-word-codex-skill/artifact-template-academic-word/ \
  ~/.codex/skills/artifact-template-academic-word/
```

Codex Desktop provides the document runtime. To run the builder independently, use Python 3.11 or newer and install:

```bash
python -m pip install -r requirements.txt
```

## Usage

Explicit invocation:

```text
Use $artifact-template-academic-word to create a course-report Word document from my topic, materials, and references.
```

Natural-language routing also works:

```text
Reformat this existing paper without changing its content. Keep only first- and second-level headings in the TOC and return a submission-ready Word file.
```

A complete drafting request needs at least a title, author, document type, body or outline, factual materials, and citation sources. The skill never invents experimental results, author metadata, DOIs, or references; it requests missing required information together.

## Formatting Contract

- A4 page geometry, margins, headers, footers, and document grid follow the accepted template.
- Heading numbering is `1`, `1.1`, `1.1.1`, and `1.1.1.1`.
- The TOC field range is fixed at `TOC \\o "1-2"`.
- Figure captions use `STYLEREF` and `SEQ` fields, tables use repeating header rows, and in-text citations and bibliography entries share one structured source.
- The cover has no visible page number, front matter uses uppercase Roman numerals, and the body restarts at Arabic page 1.

## Validation

```bash
python tools/validate_public_package.py
python artifact-template-academic-word/scripts/academic_word.py validate \
  --input artifact-template-academic-word/assets/clean-report.docx
python artifact-template-academic-word/scripts/academic_word.py validate \
  --input artifact-template-academic-word/assets/clean-thesis.docx
```

Microsoft Word is the preferred final environment for field refresh and submission checks. LibreOffice is used for page-by-page rendering. WPS is supported through a conservative OpenXML subset, but pixel-identical rendering with Word is not claimed.

## Privacy, Provenance, and License

The public repository does not contain the original 72-page sample manuscript, QQ contact details, NoteExpress data, or Visio/OLE objects. `style-authority.dotx` retains only the styles and settings required by the builder.

Formatting provenance traces to [liuweifly/hust-thesis-word](https://github.com/liuweifly/hust-thesis-word). The upstream repository does not declare a license. This repository's MIT license applies only to original code, configuration, and documentation; it excludes the sanitized template and derived masters. See [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md). The retained 2016 format is not a claim of current official HUST compliance; always follow the latest institutional requirements.
