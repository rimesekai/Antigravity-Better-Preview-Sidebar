# 更新日志 / Changelog

## v0.2.0 (2026-11-14)

### 新增功能 / Features
- **Excel 预览支持（`.xlsx` / `.xls` / `.csv` / `.tsv`）**：支持多 Sheet 工作表切换、顶部公式栏（展示原始公式与计算值）、按钮与 `Ctrl/Cmd+C` 一键复制、合并单元格还原与整区表头高亮、缩放及跨表搜索；`.csv` / `.tsv` 自动识别 UTF-8 与 GB18030 编码。
- **Word 全局搜索（`.docx`）**：支持跨文本 run 段落级高亮匹配、计数展示、`↑/↓` 导航与 `Ctrl/Cmd+F` 快捷聚焦。
- **HTML / PDF / MP4 核心增强**：新增 >20MB 大体积 PDF 与最高 4GB 本地 MP4 的 HTTP Range（`206 Partial Content`）按需分片流式加载、HTML 沙箱/交互/源码三模式隔离切换及手动「自动 / 浅色 / 深色」主题切换。

### 安全与稳定性优化 / Security & Fixes
- **依赖安全升级**：将 `xlsx` 升级至 SheetJS `0.20.3`；在 `pdfRenderer` 中设置 `isEvalSupported: false` 缓解 PDF.js CVE-2024-4367（GHSA-wgrm-67xf-hhpq）。
- **大图/大表保护与性能**：裁剪 Excel 稀疏 `!ref` 尾空行，行/列改为分批展开并设置渲染硬上限；搜索范围与可渲染边界对齐，命中未展开区域时自动扩展后高亮；PDF 远距跳页采用瞬移滚动避免渲染风暴。
- **日期显示与同文件原地刷新**：Excel 解析启用 `cellDates: false` + `cellNF: true`，消除日期兜底路径的 `+08:00` 尾巴；同文件刷新原地复用渲染器，保留当前 Sheet、选中单元格、滚动位置与搜索高亮。
- **服务端优化**：`/api/fs/read` 改为异步读取并实施分级上限（解析类文档 `60MB`，文本类 `80MB`）；静态资源内存缓存支持按 `mtimeMs` 与文件大小自动失效；剪贴板复制抽离公共实现并增加沙箱 `execCommand` 降级与失败提示。

## v0.1.0

- 初始版本：提供侧边栏 `.pptx`、`.docx`、`.pdf`、`.html` 基础预览框架、工作区文件浏览器与路径安全校验后端。
