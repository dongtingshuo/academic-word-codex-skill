# Security Policy / 安全策略

## Supported version / 支持版本

Security fixes target the latest commit on the default branch. / 安全修复面向默认分支的最新版本。

## Private reporting / 私密报告

Use GitHub private vulnerability reporting for this repository. If that option is unavailable, contact the maintainer through a private channel listed on the GitHub profile. Do not open a public issue containing exploit details, an unredacted document, personal data, institutional credentials, or a private citation library.

请使用本仓库的 GitHub 私密漏洞报告功能；若该功能不可用，请通过维护者 GitHub 主页列出的私密渠道联系。不要在公开 Issue 中包含漏洞利用细节、未脱敏文档、个人信息、学校凭据或私有文献库。

A useful report includes the affected version or commit, platform, minimal reproduction, expected impact, and a proposed mitigation when available. Remove document content that is not required to reproduce the issue.

有效报告应包含受影响版本或提交、平台、最小复现、预期影响和可行的缓解建议。请删除复现问题所不需要的文稿内容。

## Scope / 范围

In scope:

- path traversal, unintended file overwrite, or unsafe archive handling;
- leakage of document content or metadata;
- malformed OOXML that triggers unsafe or unexpected external behavior;
- command execution or injection through user-controlled fields;
- dependency vulnerabilities with a practical impact on this project.

Out of scope:

- pixel-level rendering differences already documented between Word, LibreOffice, and WPS;
- format requirements that changed after the retained template was created;
- unsupported office suites or operating-system versions without a reproducible security impact.
