---
name: file-preview
description: 在 Antigravity 桌面端侧边栏打开并预览本地 PPT（.pptx）、Word（.docx）与 HTML（.html/.htm）文件。当用户要求预览、查看或打开这些文件，或在任务产出中包含这些格式文件时触发。
---

# File Preview 文件预览助手

本插件在 Antigravity 桌面端辅助侧边栏（Aux Pane）提供本地 `.pptx`、`.docx`、`.html` 和 `.htm` 文件的高保真预览能力。

## 如何打开侧边栏预览器

- 用户可通过斜杠命令 `/file-preview` 触发，或直接点击下方交互式徽章打开侧边栏视图：

  [File Preview](sidecar://file-preview/viewer/)

- 当在工作区中直接点击 `.pptx`、`.docx`、`.html`、`.htm` 文件时，Antigravity 也会自动通过 `file_extensions` 关联唤起 File Preview 侧边栏并加载对应文件。

## 使用指引

1. 在生成、修改或检查 `.pptx`、`.docx`、`.html`、`.htm` 文件后，**主动**在回复末尾附上 `[File Preview](sidecar://file-preview/viewer/)` 徽章链接，方便用户一键打开侧边栏查看效果。
2. 如需打开特定文件，可告知用户点击上述徽章后在侧边栏文件列表中选择目标文件，或直接在工作区点击该文件。
3. 若遇到旧版二进制 `.doc` 或 `.ppt` 文件，提示用户它们暂不支持直接浏览器端渲染，建议在 Word / PowerPoint / WPS 中另存为 `.docx` 或 `.pptx` 后再预览。
