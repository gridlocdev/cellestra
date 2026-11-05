import { defineConfig } from 'vite';
import preact from '@preact/preset-vite';

// `base: './'` keeps every asset URL relative, so one build serves from the NAS
// server, from a sub-path behind a reverse proxy, and from Electron's app:// protocol.
export default defineConfig({
  base: './',
  plugins: [preact()],
  build: { outDir: 'dist', emptyOutDir: true, sourcemap: true },
  server: {
    // `pnpm dev` talks to a locally running `apps/server` for the NAS file store.
    proxy: { '/api': 'http://localhost:8787' },
  },
});
