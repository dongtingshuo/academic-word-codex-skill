# Contributing / 参与贡献

## 中文

感谢你改进 Academic Word Codex Skill。本项目处理真实学术文档，因此正确性、隐私和可复现性优先于功能数量。

提交更改前请遵守以下要求：

1. 不要提交未脱敏论文、作者身份信息、联系方式、学校账号、审稿意见或私有引用库。
2. 不要编造格式要求、实验数据、DOI、参考文献或兼容性结论。
3. 原创代码、配置与文档采用 MIT License；净化模板资产不在 MIT 授权范围内，不能用新文件暗示重新许可。
4. 修改 Word 生成逻辑时，必须同时验证 `clean-report.docx` 与 `clean-thesis.docx`。
5. 用户可见行为、安装方式和限制发生变化时，同步更新中英文 README 与 `SKILL.md`。
6. 不要把 `__pycache__`、PDF 渲染结果、临时 Word 锁文件或验证报告提交到仓库。

建议从独立分支提交聚焦单一问题的 Pull Request，并包含：

- 问题与用户影响；
- 实现选择及兼容性考虑；
- 实际运行的验证命令与结果；
- 涉及排版的真实渲染截图；
- 涉及模板来源或许可时的权利说明。

本地最低验证：

```bash
python -m pip install --requirement requirements.txt
python tools/validate_public_package.py
python artifact-template-academic-word/scripts/academic_word.py validate \
  --input artifact-template-academic-word/assets/clean-report.docx
python artifact-template-academic-word/scripts/academic_word.py validate \
  --input artifact-template-academic-word/assets/clean-thesis.docx
```

## English

Thank you for improving the Academic Word Codex Skill. Because this project processes real academic documents, correctness, privacy, and reproducibility take precedence over feature count.

Before submitting a change:

1. Never commit an unredacted manuscript, author identity, contact details, institutional account, reviewer comments, or a private reference library.
2. Never invent formatting rules, experimental data, DOIs, references, or compatibility claims.
3. Original code, configuration, and documentation use the MIT License. Sanitized template assets are excluded; do not imply that a new file relicenses them.
4. Changes to Word-generation logic must validate both `clean-report.docx` and `clean-thesis.docx`.
5. Update both READMEs and `SKILL.md` when user-visible behavior, installation, or limitations change.
6. Do not commit `__pycache__`, rendered PDFs, temporary Word lock files, or validation reports.

Use a focused branch and pull request. Include:

- the problem and user impact;
- the implementation choice and compatibility considerations;
- commands actually run and their results;
- genuine rendering screenshots for layout changes;
- a rights notice when provenance or licensing is affected.

Minimum local validation:

```bash
python -m pip install --requirement requirements.txt
python tools/validate_public_package.py
python artifact-template-academic-word/scripts/academic_word.py validate \
  --input artifact-template-academic-word/assets/clean-report.docx
python artifact-template-academic-word/scripts/academic_word.py validate \
  --input artifact-template-academic-word/assets/clean-thesis.docx
```
