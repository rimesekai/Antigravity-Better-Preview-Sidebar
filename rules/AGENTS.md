# File Preview 侧边栏预览规则

当工作区或当前对话任务中涉及以下本地文件时，请在回复中主动向用户提供 File Preview 侧边栏预览入口 `[File Preview](sidecar://file-preview/viewer/)`：

- `.pptx`（PowerPoint 演示文稿）
- `.docx`（Word 文档）
- `.html` / `.htm`（本地网页文件）
- 体积超过 20MB（`20,971,520` 字节，超出 Antigravity 内置 PDF 预览器限制）的 `.pdf` 文件

提示用户点击 `[File Preview](sidecar://file-preview/viewer/)` 即可在右侧边栏直接高保真查看（大体积 PDF 支持基于 HTTP Range 的按需分片流式加载）。
