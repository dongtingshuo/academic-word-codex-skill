# Normalized request

Use one JSON object per build.

```json
{
  "mode": "draft",
  "document_type": "course_report",
  "metadata": {
    "title": "报告标题",
    "author": "作者",
    "student_id": "学号",
    "institution": "学校或单位",
    "course": "课程名称",
    "instructor": "教师",
    "major": "专业",
    "major_en": "English major name",
    "advisor": "导师",
    "advisor_en": "English supervisor name",
    "date": "2026年6月",
    "date_en": "June 2026",
    "title_en": "English thesis title",
    "author_en": "English author name",
    "degree_en": "Master of Engineering",
    "classification": "分类号",
    "secrecy": "公开"
  },
  "requirements": {
    "citation_style": "GB/T 7714-2015 numeric",
    "include_toc": true,
    "cover_choice": "generic",
    "external_template": null,
    "conflict_resolution": null,
    "required_metadata": []
  },
  "content": {
    "markdown": "# 引言\n正文……[@ref1]",
    "source_docx": null,
    "abstract_cn": "",
    "keywords_cn": [],
    "abstract_en": "",
    "keywords_en": []
  },
  "citations": [
    {
      "id": "ref1",
      "type": "journal",
      "authors": ["张三", "李四"],
      "title": "文献题名",
      "container": "期刊名",
      "year": 2025,
      "volume": "10",
      "issue": "2",
      "pages": "1-10",
      "doi": ""
    }
  ],
  "assets": [
    {"id": "fig1", "kind": "figure", "path": "/absolute/figure.png", "caption": "示意图"}
  ],
  "output_name": "课程报告.docx"
}
```

## Required fields

- All modes: `mode`, `document_type`, `metadata.title`, and `metadata.author`.
- `draft`: `content.markdown` or a readable Markdown path.
- `format`: an existing `content.source_docx`.
- `hust_thesis`: additionally require `student_id`, `major`, `major_en`, `advisor`, `advisor_en`, `institution`, `date`, `date_en`, `title_en`, `author_en`, `degree_en`, `classification`, `secrecy`, Chinese/English abstracts, and both keyword lists. These fields drive the Chinese cover, English title page, and declarations and must never be guessed.
- `requirements.cover_choice`: use `generic` by default for any school, `hust` only with `document_type=hust_thesis`, or `external` when an uploaded institutional/course Word template supplies the cover. An external cover is handled through the template/base-replace workflow rather than the generic builder.
- Every `[@id]` in Markdown must have exactly one citation record with the same `id`.
- Every `{{figure:id}}` marker must have exactly one readable figure asset.

When `requirements.required_metadata` lists extra keys, treat them as mandatory. Never replace missing values with `XXX`, `待补`, or guessed content.

## Markdown subset

- `#` through `####`: real numbered heading levels.
- `## 摘要`-style unnumbered front matter is supplied through dedicated abstract fields, not body Markdown.
- `- item` or `* item`: real bullet list.
- `1. item`: real numbered list.
- Markdown pipe tables: fixed-width three-line academic tables.
- `{{figure:id}}`: inline figure plus chapter-aware SEQ caption.
- `$$ equation $$`: simple OMML equation plus chapter-aware equation number.
- `[@citation-id]`: GB/T numeric in-text citation.
- `---PAGEBREAK---`: explicit page break; do not use repeated blank lines.

## Conflict record

If an external template conflicts with the fallback contract, store the confirmed choices under `requirements.conflict_resolution`, for example:

```json
{
  "page_geometry": "external",
  "cover_structure": "external",
  "body_typography": "academic-word",
  "heading_numbering": "external",
  "citation_style": "GB/T 7714-2015 numeric"
}
```

Do not build while detected material conflicts remain unresolved.
