# File Preview — Antigravity 侧边栏文件预览插件

在 Antigravity 桌面端侧边栏（Aux Pane）直接打开并高保真预览本地 **PPT（`.pptx`）**、**Word（`.docx`）**、**PDF（`.pdf`，支持 >20MB 大文件流式预览）** 与 **HTML（`.html` / `.htm`）** 文件。

## 主要功能

- **PPT 预览（`.pptx`）**：基于 `@aiden0z/pptx-renderer` 渲染幻灯片，支持单页/滚动布局、缩放自适应、缩略图导航与文本检索。
- **Word 预览（`.docx`）**：基于 `docx-preview` 保留分页、页眉页脚、脚注尾注与表格样式。
- **PDF 预览（`.pdf`，含 >20MB 大文件流式预览）**：基于 `pdfjs-dist` 与本地独立端口 HTTP Range（`206 Partial Content`）分片流式加载服务，无需一次性将整份大文件读入内存即可按需渲染可视页面；支持“画布”模式（单页/滚动、可折叠缩略图栏、适应/缩放、1-based 翻页跳页、跨页全文搜索）与“原生”流式视图切换。
- **HTML 预览（`.html` / `.htm`）**：通过按需启动的本地独立端口静态预览服务（`127.0.0.1:0`）挂载渲染，完整支持同目录及工作区内相对路径资源（CSS / JS / 图片 / 字体），并与 Sidecar 主控台严格跨源隔离；支持安全沙箱、交互模式与源码视图切换。
- **自动刷新与主题跟随**：默认开启文件变更自动刷新（2.5s 检查 `mtimeMs` / `size`），并自动同步跟随 Antigravity 浅色 / 深色主题配色。
- **关联打开与工作区扫描**：支持通过点击工作区文件（`file_extensions`）、侧边栏文件列表、路径直达或拖拽文件快速切换预览对象。

## 构建与测试

在 Windows PowerShell 中请使用 `npm.cmd`（避免 `.ps1` 执行策略限制），并推荐使用国内镜像安装依赖：

```powershell
npm.cmd install --registry=https://registry.npmmirror.com
npm.cmd test
npm.cmd run build
```

构建产物输出至 `dist/`（`dist/index.html`、`dist/app.js`、`dist/styles.css`、`dist/server.mjs`、`dist/pdf.worker.js`），运行时无需携带 `node_modules/`。

## 安装到本地 Antigravity 插件目录

运行以下命令可将运行所需的精简文件复制到全局插件目录 `~/.gemini/config/plugins/file-preview/`：

```powershell
npm.cmd run install:local
```

安装完成后重启或刷新 Antigravity 插件列表，即可使用 `/file-preview` 或点击 `[File Preview](sidecar://file-preview/viewer/)` 打开侧边栏，也可直接点击工作区中的 `.pptx`、`.docx`、`.pdf`、`.html`、`.htm` 文件自动触发预览。

## 已知格式说明

- **支持格式**：`.pptx`、`.docx`、`.pdf`（支持超过 Antigravity 内置 20MB 限制的大体积文件流式预览）、`.html`、`.htm`。
- **暂不支持**：旧版二进制 `.ppt` 与 `.doc` 格式（浏览器端无法直接解析二进制复合文档，请在 PowerPoint / Word / WPS 中另存为 `.pptx` 或 `.docx` 后再预览）。
