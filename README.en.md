# File Preview — Antigravity Sidebar File Preview Plugin

[简体中文](./README.md) | **English**

Preview local **PPT (`.pptx`)**, **Word (`.docx`)**, **Excel (`.xlsx` / `.xls` / `.csv` / `.tsv`)**, **PDF (`.pdf`, including >20MB large-file streaming)**, **MP4 (`.mp4`, including >20MB large-video streaming)**, and **HTML (`.html` / `.htm`)** files directly inside the Antigravity desktop Aux Pane with high fidelity.

## Key Features

- **PPT Preview (`.pptx`)**: Powered by `@aiden0z/pptx-renderer` with slide/list layout modes, fit/zoom controls, collapsible thumbnails, and text search.
- **Word Preview (`.docx`)**: Powered by `docx-preview` preserving pagination, headers/footers, footnotes/endnotes, and tables; supports fit-width zoom and full-text search across runs (`Ctrl/Cmd+F` shortcut focus, match counter, and next/previous jump highlighting).
- **Excel Preview (`.xlsx` / `.xls` / `.csv` / `.tsv`)**: Powered by `xlsx` (SheetJS `0.20.3`) with multi-sheet tab switching, top inspection/formula bar (showing raw formula `=` and evaluated value, with button and `Ctrl/Cmd+C` copy), merged cell (`rowSpan` / `colSpan`) rendering with full-span header highlighting, sparse `!ref` bounds trimming, batched row/column loading with hard safety caps, zoom controls, and cross-sheet search; `.csv` / `.tsv` auto-detect UTF-8 (with BOM) and GB18030 encoding.
- **PDF Preview (`.pdf`, >20MB streaming)**: Powered by `pdfjs-dist` over a dedicated local HTTP Range (`206 Partial Content`) server so large PDFs load on demand without buffering the entire file into memory; supports slide/list views, collapsible thumbnails, zoom/fit, page jump, and cross-page text search.
- **MP4 Playback (`.mp4`, >20MB streaming)**: Streams local videos up to 4GB via local HTTP Range (`206 Partial Content`) with scrubbing, `0.5x–2x` playback speed, and loop toggle.
- **HTML Preview (`.html` / `.htm`)**: Served on demand from an isolated local static server (`127.0.0.1:0`) supporting relative assets (CSS / JS / images / fonts) within the file/workspace root; provides Sandbox, Interactive, and Read-only Source views.
- **Auto-Refresh & Theme Sync**: Watches disk file `mtimeMs` / `size` every 2.5s and reloads active renderers in place when supported; follows Antigravity light/dark themes automatically with a manual `Auto / Light / Dark` toggle.
- **Workspace Discovery & Associations**: Automatically opens supported extensions from the workspace via `file_extensions`, plus a workspace file drawer, direct path input, and drag-and-drop preview.

## Quick Install (Prebuilt Bundle Included)

This repository includes the prebuilt runtime assets in `dist/`. After cloning, you can install directly into `~/.gemini/config/plugins/file-preview/` without installing dependencies or building from scratch:

```powershell
git clone https://github.com/rimesekai/Antigravity-Better-Preview-Sidebar.git
cd Antigravity-Better-Preview-Sidebar
npm.cmd run install:local
```

On macOS / Linux, run:

```bash
node scripts/install-local.mjs
```

After installation, reload or restart Antigravity plugins, then trigger `/file-preview`, click `[File Preview](sidecar://file-preview/viewer/)`, or click any `.pptx`, `.docx`, `.xlsx`, `.xls`, `.csv`, `.tsv`, `.pdf`, `.mp4`, `.html`, or `.htm` file in your workspace.

## Build & Test from Source

On Windows PowerShell, use `npm.cmd` (or `--omit=optional` to skip `pdfjs-dist`'s optional `canvas` / `tar` dependency when you only need tests and builds):

```powershell
npm.cmd install --omit=optional
npm.cmd test
npm.cmd run build
npm.cmd run install:local
```

Build outputs are emitted to `dist/` (`dist/index.html`, `dist/app.js`, `dist/styles.css`, `dist/server.mjs`, `dist/pdf.worker.js`), with no `node_modules/` required at runtime.

## Supported Formats & Security Notes

- **Supported**: `.pptx`, `.docx`, `.xlsx`, `.xls`, `.csv`, `.tsv`, `.pdf` (including >20MB streaming beyond Antigravity's built-in 20MB limit), `.mp4` (HTTP Range streaming), `.html`, `.htm`.
- **Not Supported**: Legacy binary `.ppt` and `.doc` files (cannot be parsed natively in browser; re-save as `.pptx` or `.docx` in PowerPoint / Word / WPS first).
- **Path & Memory Boundaries**: Backend file APIs strictly validate paths against workspace/allowed roots and block traversal, out-of-root symlinks, and internal agent log paths. Formats parsed in browser memory (`.pptx` / `.docx` / `.xlsx` / `.xls` / `.csv` / `.tsv`) are capped at `60MB`; Excel sheets default to the first `2,000` rows × `100` columns with hard rendering caps; PDF (`1GB` limit) and MP4 (`4GB` limit) stream on demand over HTTP Range.
- **PDF.js & Third-Party Dependencies**:
  1. `pdfRenderer` passes `isEvalSupported: false` to `pdfjs-dist` to disable `Function`/`eval`-based glyph evaluation, mitigating CVE-2024-4367 (GHSA-wgrm-67xf-hhpq).
  2. `@aiden0z/pptx-renderer` internally calls `pdfjsLib.getDocument` with upstream defaults when handling embedded PDF data inside `.pptx` files—avoid opening untrusted `.pptx` files.
  3. The optional `canvas` dependency of `pdfjs-dist` pulls in legacy `tar` in the dev lockfile only; it is not bundled into `dist/`.

## License

Released under the [MIT License](./LICENSE).
