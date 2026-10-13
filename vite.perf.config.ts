/**
 * Production build of the dev-only S10.2 performance page
 * (src/debug/viewer-perf.html + its viewer worker), so the release bench
 * measures the viewer the way a host ships it: production React, minified
 * bundles and one worker chunk, not the dev server's unbundled development
 * modules (whose React development checks dominated a scroll profile).
 *
 *   bunx vite build --config vite.perf.config.ts
 *
 * Output goes to .tmp/viewer-perf-dist (gitignored). It never touches the
 * library build: vite.config.ts and dist/ are unaffected.
 */
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { fileURLToPath } from 'node:url'

export const PERF_OUT_DIR = fileURLToPath(new URL('./.tmp/viewer-perf-dist', import.meta.url))

export default defineConfig({
  root: fileURLToPath(new URL('.', import.meta.url)),
  base: './',
  mode: 'production',
  plugins: [react()],
  logLevel: 'warn',
  build: {
    outDir: PERF_OUT_DIR,
    emptyOutDir: true,
    sourcemap: false,
    // PERF_MINIFY=0 keeps function names for CPU-profile attribution (diagnosis only).
    minify: process.env.PERF_MINIFY !== '0',
    reportCompressedSize: false,
    chunkSizeWarningLimit: 10_000,
    rollupOptions: { input: fileURLToPath(new URL('./src/debug/viewer-perf.html', import.meta.url)) },
  },
  worker: { format: 'es' },
})
