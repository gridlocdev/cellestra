import { defineConfig, type Plugin } from 'vite';
import preact from '@preact/preset-vite';

// Strict CSP for the built client only — the dev server's HMR needs inline scripts.
const csp: Plugin = {
  name: 'cellestra-csp',
  apply: 'build',
  transformIndexHtml: () => [
    {
      tag: 'meta',
      attrs: {
        'http-equiv': 'Content-Security-Policy',
        content:
          "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; connect-src 'self'; object-src 'none'; base-uri 'none'",
      },
      injectTo: 'head-prepend',
    },
  ],
};

// `base: './'` keeps every asset URL relative, so one build serves from the NAS
// server, from a sub-path behind a reverse proxy, and from Electron's app:// protocol.
export default defineConfig({
  base: './',
  plugins: [preact(), csp],
  build: { outDir: 'dist', emptyOutDir: true, sourcemap: true },
  server: {
    // `pnpm dev` talks to a locally running `apps/server` for the NAS file store.
    proxy: { '/api': 'http://localhost:8787' },
  },
});
