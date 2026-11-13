---
name: file-preview
description: 在 Antigravity 桌面端侧边栏打开并预览本地 PPT（.pptx）、Word（.docx）、Excel（.xlsx/.xls/.csv/.tsv，支持多 Sheet 工作表切换、公式/合并单元格展示及搜索）、PDF（.pdf，含超过 20MB 的大体积文件流式预览）、MP4（.mp4，含超过 20MB 的大体积视频流式播放）与 HTML（.html/.htm）文件。当用户要求预览、查看或打开这些文件，或在任务产出中包含这些格式文件时触发。
---

# File Preview 文件预览助手

本插件在 Antigravity 桌面端辅助侧边栏（Aux Pane）提供本地 `.pptx`、`.docx`、`.xlsx`、`.xls`、`.csv`、`.tsv`、`.pdf`、`.mp4`、`.html` 和 `.htm` 文件的高保真预览能力，支持 Excel 多 Sheet 工作表切换、公式与合并单元格检视及跨表搜索，并针对超过 20MB 的大体积 `.pdf` 与 `.mp4` 文件提供基于 HTTP Range 的按需分片流式预览与播放。

## 如何打开侧边栏预览器

- 用户可通过斜杠命令 `/file-preview` 触发，或直接点击下方交互式徽章打开侧边栏视图：

  [File Preview](sidecar://file-preview/viewer/)

- 当在工作区中直接点击 `.pptx`、`.docx`、`.xlsx`、`.xls`、`.csv`、`.tsv`、`.pdf`、`.mp4`、`.html`、`.htm` 文件时，Antigravity 也会自动通过 `file_extensions` 关联唤起 File Preview 侧边栏并加载对应文件。

## 使用指引

1. 在生成、修改或检查 `.pptx`、`.docx`、`.xlsx`、`.xls`、`.csv`、`.tsv`、`.html`、`.htm` 文件，以及 `.pdf` / `.mp4`（尤其是体积超过 20MB / `20,971,520` 字节）文件后，**主动**在回复末尾附上 `[File Preview](sidecar://file-preview/viewer/)` 徽章链接，方便用户一键打开侧边栏查看效果。
2. 如需打开特定文件，可告知用户点击上述徽章后在侧边栏文件列表中选择目标文件（或粘贴绝对路径 / 直接拖拽文件进入预览区），也可以直接在工作区点击该文件。
3. 预览 Excel（`.xlsx` / `.xls` / `.csv` / `.tsv`）时，支持底部多 Sheet 工作表标签切换、顶部公式栏展示原始公式与计算值（含一键复制）、合并单元格还原、缩放调节以及跨工作表全文搜索；`.csv` / `.tsv` 支持 UTF-8 与 GB18030 自动解码以防中文乱码。
4. 预览大体积 PDF / MP4 时，插件通过本地独立端口 HTTP `206 Partial Content` 分片流式按需读取页面或视频分片（PDF 支持单页/滚动、缩略图、缩放与全文搜索，MP4 支持倍速与循环播放）。
5. 若遇到旧版二进制 `.doc` 或 `.ppt` 文件，提示用户它们暂不支持直接浏览器端渲染，建议在 Word / PowerPoint / WPS 中另存为 `.docx` 或 `.pptx` 后再预览。
