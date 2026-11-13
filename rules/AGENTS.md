# File Preview 侧边栏预览规则

当工作区或当前对话任务中涉及以下本地文件时，请在回复中主动向用户提供 File Preview 侧边栏预览入口 `[File Preview](sidecar://file-preview/viewer/)`：

- `.pptx`（PowerPoint 演示文稿）
- `.docx`（Word 文档）
- `.xlsx` / `.xls` / `.csv` / `.tsv`（Excel 表格与 CSV/TSV 数据表，支持多 Sheet 工作表切换、公式与合并单元格展示及搜索）
- `.html` / `.htm`（本地网页文件）
- `.pdf` 文件（尤其是体积超过 20MB / `20,971,520` 字节的大体积 PDF）
- `.mp4` 视频文件（尤其是体积超过 20MB 的大体积 MP4 视频）

提示用户点击 `[File Preview](sidecar://file-preview/viewer/)` 即可在右侧边栏直接高保真查看（Excel 支持多工作表切换、公式栏与跨表搜索，大体积 PDF 与 MP4 均支持基于 HTTP Range 的按需分片流式加载）。
