import { defineConfig, type Plugin } from 'vite';

/**
 * 将 Vite 产物改造为适配 Antigravity Sidecar `plugin://` 协议 iframe 的经典资源引用：
 * - 移除 `<script type="module" ...>` 与默认样式标签
 * - 在 `<head>` 注入 `/styles.css` 与 `/preload.js`
 * - 在 `</body>` 前注入经典 IIFE 脚本 `/app.js`
 */
function sidecarClassicHtmlPlugin(): Plugin {
  return {
    name: 'sidecar-classic-html',
    enforce: 'post',
    transformIndexHtml(html) {
      let result = html
        .replace(/<script\b[^>]*\btype=["']module["'][^>]*>[\s\S]*?<\/script>\s*/gi, '')
        .replace(/<link\b[^>]*\brel=["']stylesheet["'][^>]*>\s*/gi, '');

      const headTags = [
        '  <link rel="stylesheet" href="/styles.css" />',
        '  <script src="/preload.js"></script>',
      ].join('\n');

      if (result.includes('</head>')) {
        result = result.replace('</head>', `${headTags}\n</head>`);
      } else {
        result = `${headTags}\n${result}`;
      }

      const bodyScript = '  <script src="/app.js"></script>';
      if (result.includes('</body>')) {
        result = result.replace('</body>', `${bodyScript}\n</body>`);
      } else {
        result = `${result}\n${bodyScript}`;
      }

      return result;
    },
  };
}

export default defineConfig({
  plugins: [sidecarClassicHtmlPlugin()],
  build: {
    outDir: 'dist',
    emptyOutDir: false,
    cssCodeSplit: false,
    sourcemap: false,
    rollupOptions: {
      input: 'index.html',
      output: {
        format: 'iife',
        inlineDynamicImports: true,
        entryFileNames: 'app.js',
        assetFileNames: (assetInfo) => {
          if (assetInfo.name && assetInfo.name.endsWith('.css')) {
            return 'styles.css';
          }
          return 'assets/[name][extname]';
        },
      },
    },
  },
});
