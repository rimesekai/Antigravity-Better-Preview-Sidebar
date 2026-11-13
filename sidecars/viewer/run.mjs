/**
 * Sidecar 极简自举启动器。
 * 宿主以当前文件所在目录为 cwd 拉起进程并注入 sidecar_sdk。
 */
try {
  const { startServer } = await import('../../dist/server.mjs');
  await startServer();
} catch (error) {
  console.error('[file-preview] Sidecar 启动失败:', error);
  process.exit(1);
}
