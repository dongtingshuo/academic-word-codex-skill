# Academic Word Codex Skill

[English](README_EN.md)

面向中文论文、学位论文和课程报告的 Codex Skill。它把用户提供的题目、材料、引用和现有 Word 文档整理为可提交的 `.docx`，并通过结构检查与逐页渲染控制排版质量。

![模板预览](artifact-template-academic-word/assets/preview.png)

## 核心能力

- `draft`：根据题目、提纲、事实材料和来源起草并生成完整 Word。
- `format`：保留已有文稿的正文、表格、图片、公式和链接，重新套用统一格式。
- 使用真正的 Word 样式、四级标题编号、字段目录、题注、交叉引用、OMML 公式和分节页码。
- 目录只显示一级和二级标题；三级、四级标题继续在正文中自动编号。
- 默认使用 GB/T 7714—2015 顺序编码，正文引用为上标方括号编号。
- 通用课程报告默认不带校名；只有明确选择 `hust_thesis` 时才使用华科论文前置结构。
- 最终验证会检查样式、编号、引用、题注、页码、表头、脚注/尾注部件、占位符、修订痕迹和不可移植对象。

## 安装

```bash
git clone https://github.com/dongtingshuo/academic-word-codex-skill.git
mkdir -p ~/.codex/skills
rsync -a academic-word-codex-skill/artifact-template-academic-word/ \
  ~/.codex/skills/artifact-template-academic-word/
```

Codex Desktop 提供文档运行时。若单独执行生成脚本，请使用 Python 3.11 或更高版本并安装：

```bash
python -m pip install -r requirements.txt
```

## 使用

可以显式调用：

```text
使用 $artifact-template-academic-word，根据我的题目、资料和参考文献生成课程报告 Word。
```

也可以直接描述任务：

```text
把这份已有论文重新排版，正文内容不要改，目录只保留一二级标题，生成可提交的 Word。
```

完整写作至少需要题目、作者、文档类型、正文或提纲、事实材料和引用来源。Skill 不会编造实验数据、作者信息、DOI 或参考文献；必要信息缺失时会集中询问。

## 排版契约

- A4 页面，版心、页边距、页眉页脚和文档网格沿用已验收模板。
- 一级至四级标题编号为 `1`、`1.1`、`1.1.1`、`1.1.1.1`。
- 目录字段范围固定为 `TOC \\o "1-2"`。
- 图注使用 `STYLEREF` 与 `SEQ` 字段，表格使用可重复表头，正文引用与文后条目由同一结构化数据生成。
- 封面无可见页码；前置部分使用大写罗马数字；正文重新从阿拉伯数字 1 开始。

## 验证

```bash
python tools/validate_public_package.py
python artifact-template-academic-word/scripts/academic_word.py validate \
  --input artifact-template-academic-word/assets/clean-report.docx
python artifact-template-academic-word/scripts/academic_word.py validate \
  --input artifact-template-academic-word/assets/clean-thesis.docx
```

Microsoft Word 是最终字段刷新与提交检查的优先环境；LibreOffice 用于逐页渲染验证。WPS 使用保守 OpenXML 子集，但不承诺与 Word 像素级完全一致。

## 隐私、来源与许可

公开仓库不包含原始 72 页示例论文、QQ、NoteExpress 数据或 Visio/OLE 对象。`style-authority.dotx` 只保留生成所需的样式和设置。

格式来源可追溯至 [liuweifly/hust-thesis-word](https://github.com/liuweifly/hust-thesis-word)。上游仓库未声明许可证；本仓库的 MIT 许可仅覆盖原创代码、配置和文档，不覆盖净化模板及其衍生母版。详见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。该 2016 模板不是对华中科技大学现行官方要求的声明，提交前应以最新院系规定为准。
