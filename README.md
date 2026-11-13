# File Preview — Antigravity 侧边栏文件预览插件

在 Antigravity 桌面端侧边栏（Aux Pane）直接打开并高保真预览本地 **PPT（`.pptx`）**、**Word（`.docx`）**、**Excel（`.xlsx` / `.xls` / `.csv` / `.tsv`）**、**PDF（`.pdf`，支持 >20MB 大文件流式预览）**、**MP4（`.mp4`，支持 >20MB 大视频流式播放）** 与 **HTML（`.html` / `.htm`）** 文件。

## 主要功能

- **PPT 预览（`.pptx`）**：基于 `@aiden0z/pptx-renderer` 渲染幻灯片，支持单页/滚动布局、缩放自适应、缩略图导航与文本检索。
- **Word 预览（`.docx`）**：基于 `docx-preview` 保留分页、页眉页脚、脚注尾注与表格样式；支持自适应宽度缩放与跨 run 全文搜索（`Ctrl/Cmd+F` 快捷聚焦、上一处/下一处定位高亮）。
- **Excel 表格预览（`.xlsx` / `.xls` / `.csv` / `.tsv`）**：基于 `xlsx`（SheetJS `0.20.3`）解析工作簿，支持多 Sheet 工作表标签切换、顶部单元格检视与公式栏（展示原始公式 `=` 与计算值，支持按钮与 `Ctrl/Cmd+C` 一键复制）、合并单元格（`rowSpan` / `colSpan`）还原与整区表头高亮、稀疏边界自动裁剪与行/列分批加载、缩放以及跨 Sheet 文本搜索；`.csv` / `.tsv` 自动检测 UTF-8（含 BOM）与 GB18030 编码，避免中文表格乱码。
- **PDF 预览（`.pdf`，含 >20MB 大文件流式预览）**：基于 `pdfjs-dist` 与本地独立端口 HTTP Range（`206 Partial Content`）分片流式加载服务，无需一次性将整份大文件读入内存即可按需渲染可视页面；支持单页/滚动布局、可折叠缩略图栏、适应/缩放、1-based 翻页跳页与跨页全文搜索。
- **MP4 视频播放（`.mp4`，含 >20MB 大视频流式播放）**：基于本地独立端口 HTTP Range（`206 Partial Content`）按需分片流式播放（最高支持 4GB 本地 MP4），支持拖动进度条、0.5x–2x 倍速调节与循环播放。
- **HTML 预览（`.html` / `.htm`）**：通过按需启动的本地独立端口静态预览服务（`127.0.0.1:0`）挂载渲染，完整支持同目录及工作区内相对路径资源（CSS / JS / 图片 / 字体），并与 Sidecar 主控台严格跨源隔离；支持安全沙箱、交互模式与源码视图切换。
- **自动刷新与主题切换**：默认开启文件变更自动刷新（2.5s 检查 `mtimeMs` / `size`），自动同步跟随 Antigravity 浅色 / 深色主题（支持工具栏右上角「自动 / 浅色 / 深色」手动切换）。
- **关联打开与工作区扫描**：支持通过点击工作区文件（`file_extensions`）、侧边栏文件列表、路径直达或拖拽文件快速切换预览对象。

## 快速安装（直接部署到 Antigravity）

本仓库已包含可直接运行的预构建产物 `dist/`。克隆仓库后无需额外安装依赖，直接执行本地安装脚本即可部署到全局插件目录 `~/.gemini/config/plugins/file-preview/`：

```powershell
git clone https://github.com/rimesekai/Antigravity-Better-Preview-Sidebar.git
cd Antigravity-Better-Preview-Sidebar
npm.cmd run install:local
```

（若在 macOS / Linux 下，可直接运行 `node scripts/install-local.mjs`。）

安装完成后重启或刷新 Antigravity 插件列表，即可使用 `/file-preview` 或点击 `[File Preview](sidecar://file-preview/viewer/)` 打开侧边栏，也可直接点击工作区中的 `.pptx`、`.docx`、`.xlsx`、`.xls`、`.csv`、`.tsv`、`.pdf`、`.mp4`、`.html`、`.htm` 文件自动触发预览。

## 从源码构建与测试

在 Windows PowerShell 中请使用 `npm.cmd`（避免 `.ps1` 执行策略限制）；若不需要为测试环境安装 `pdfjs-dist` 的可选服务端依赖 `canvas`（`node-pre-gyp` / `tar`），建议加上 `--omit=optional`：

```powershell
npm.cmd install --omit=optional --registry=https://registry.npmmirror.com
npm.cmd test
npm.cmd run build
npm.cmd run install:local
```

构建产物输出至 `dist/`（`dist/index.html`、`dist/app.js`、`dist/styles.css`、`dist/server.mjs`、`dist/pdf.worker.js`），插件运行时无需携带 `node_modules/`。

## 已知格式与安全说明

- **支持格式**：`.pptx`、`.docx`、`.xlsx`、`.xls`、`.csv`、`.tsv`、`.pdf`（支持超过 Antigravity 内置 20MB 限制的大体积文件流式预览）、`.mp4`（支持超过 20MB 的大体积视频 HTTP Range 流式播放）、`.html`、`.htm`。
- **暂不支持**：旧版二进制 `.ppt` 与 `.doc` 格式（浏览器端无法直接解析二进制复合文档，请在 PowerPoint / Word / WPS 中另存为 `.pptx` 或 `.docx` 后再预览）。
- **路径与内存安全边界**：后端路由校验路径仅允许访问当前工作区及用户允许目录，自动拦截路径穿越、跨边界符号链接及敏感日志目录；需整包载入渲染进程内存解析的格式（`.pptx` / `.docx` / `.xlsx` / `.xls` / `.csv` / `.tsv`）限制单文件上限为 `60MB`，Excel 单表默认最多加载前 `2,000` 行 × `100` 列并设安全渲染硬上限，PDF（最高 `1GB`）与 MP4（最高 `4GB`）均走独立本地端口 HTTP Range 分片流式读取。
- **PDF.js 与三方依赖说明**：
  1. `pdfRenderer` 已强制向 `pdfjs-dist` 传入 `isEvalSupported: false`，以禁用字体/路径中的动态脚本求值，缓解 CVE-2024-4367（GHSA-wgrm-67xf-hhpq）。
  2. `@aiden0z/pptx-renderer` 在内部处理 `.pptx` 嵌套 PDF 数据时仍使用其上游默认 `getDocument` 参数，请勿使用本插件打开不可信来源的 `.pptx` 文件。
  3. 开发依赖树中 `pdfjs-dist` 的可选依赖 `canvas` 会间接引入旧版 `tar`，仅用于 Node 侧可选安装期，不会被打包进浏览器端与 Sidecar 运行时产物 `dist/`。

## License

本项目采用 [MIT License](./LICENSE) 开源许可。
