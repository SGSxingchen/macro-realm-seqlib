import { defineConfig, loadEnv } from 'vite';
import react from '@vitejs/plugin-react';
import { fileURLToPath } from 'node:url';

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '');
  const apiTarget = env.VITE_API_TARGET || 'http://127.0.0.1:8000';
  return {
    plugins: [react()],
    build: { rollupOptions: { input: {
      archive: fileURLToPath(new URL('./index.html', import.meta.url)),
      tabletop: fileURLToPath(new URL('./tabletop.html', import.meta.url)),
    } } },
    server: {
      proxy: {
        '/api': { target: apiTarget, changeOrigin: true, ws: true }
      }
    }
  };
});
