/**
 * Vite config for the project website (GitHub Pages): a multi-page app under
 * site/ that renders the library from src/, so the site always shows the code
 * on main rather than the last npm release.
 *
 *   bun run site:dev     # http://localhost:5174/officeview/
 *   bun run site:build   # site-dist/, deployed by .github/workflows/pages.yml
 *
 * Sample documents are generated into site/public/samples by
 * scripts/build-site-samples.ts before either command runs.
 */
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { fileURLToPath, URL } from 'node:url'
import { ovCss } from './scripts/ov-css-config'

const at = (path: string) => fileURLToPath(new URL(path, import.meta.url))

export default defineConfig({
  root: at('./site'),
  // Served from https://<owner>.github.io/officeview/
  base: process.env.SITE_BASE ?? '/officeview/',
  css: ovCss.css,
  plugins: [react(), ...ovCss.plugins],
  resolve: {
    alias: [
      { find: /^officeview\/worker$/, replacement: at('./src/worker/worker-entry.ts') },
      { find: /^officeview$/, replacement: at('./src/index.ts') },
    ],
  },
  // The viewer worker code-splits its lazy large-file modules.
  worker: { format: 'es' },
  server: { port: 5174 },
  build: {
    outDir: at('./site-dist'),
    emptyOutDir: true,
    chunkSizeWarningLimit: 2000,
    rollupOptions: {
      input: {
        index: at('./site/index.html'),
        demos: at('./site/demos/index.html'),
        playground: at('./site/playground/index.html'),
        docs: at('./site/docs/index.html'),
      },
    },
  },
})
