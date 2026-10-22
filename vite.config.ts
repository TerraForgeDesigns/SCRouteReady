import { printAssetsPlugin } from './scripts/print-assets.mjs';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
export default defineConfig(({ mode }) => ({
  base: mode === 'hub' ? '/routes/' : '/',
  plugins: [react(), printAssetsPlugin(), { name: 'opsready-hub-entry', transformIndexHtml: {
    order: 'pre', handler: html => mode === 'hub' ? html.replace('/src/main.tsx', '/src/hub-main.tsx') : html,
  } }],
}));
