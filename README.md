# File Preview — Antigravity 侧边栏文件预览插件

在 Antigravity 桌面端侧边栏（Aux Pane）直接打开并高保真预览本地 **PPT（`.pptx`）**、**Word（`.docx`）** 与 **HTML（`.html` / `.htm`）** 文件。

## 主要功能

- **PPT 预览（`.pptx`）**：基于 `@aiden0z/pptx-renderer` 渲染幻灯片，支持缩放自适应、缩略图导航与文本检索。
- **Word 预览（`.docx`）**：基于 `docx-preview` 保留分页、页眉页脚、脚注尾注与表格样式。
- **HTML 预览（`.html` / `.htm`）**：通过按需启动的本地独立端口静态预览服务（`127.0.0.1:0`）挂载渲染，完整支持同目录及工作区内相对路径资源（CSS / JS / 图片 / 字体），并与 Sidecar 主控台严格跨源隔离；支持安全沙箱模式与交互模式切换。
- **关联打开与工作区扫描**：支持通过点击工作区文件（`file_extensions`）或侧边栏文件列表快速切换预览对象。

## 构建与测试

在 Windows PowerShell 中请使用 `npm.cmd`（避免 `.ps1` 执行策略限制），并推荐使用国内镜像安装依赖：

```powershell
npm.cmd install --registry=https://registry.npmmirror.com
npm.cmd test
npm.cmd run build
```

构建产物输出至 `dist/`（`dist/index.html`、`dist/app.js`、`dist/styles.css`、`dist/server.mjs`），运行时无需携带 `node_modules/`。

## 安装到本地 Antigravity 插件目录

运行以下命令可将运行所需的精简文件复制到全局插件目录 `~/.gemini/config/plugins/file-preview/`：

```powershell
npm.cmd run install:local
```

安装完成后重启或刷新 Antigravity 插件列表，即可使用 `/file-preview` 或点击 `[File Preview](sidecar://file-preview/viewer/)` 打开侧边栏，也可直接点击工作区中的 `.pptx`、`.docx`、`.html`、`.htm` 文件自动触发预览。

## 已知格式说明

- **支持格式**：`.pptx`、`.docx`、`.html`、`.htm`。
- **暂不支持**：旧版二进制 `.ppt` 与 `.doc` 格式（浏览器端无法直接解析二进制复合文档，请在 PowerPoint / Word / WPS 中另存为 `.pptx` 或 `.docx` 后再预览）。
