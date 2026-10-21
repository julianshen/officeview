import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { fileURLToPath, URL } from 'node:url'
import { readFileSync } from 'node:fs'
import { moduleMap } from './scripts/vite-module-map'
import { ovCss } from './scripts/ov-css-config'

export default defineConfig({
  // Tailwind + the theme-root scoping stage (scripts/ov-css-config.ts).
  css: ovCss.css,
  plugins: [react(), ...ovCss.plugins, {
    name: 'drawing-preset-notices',
    generateBundle() {
      for (const name of ['LICENSE', 'NOTICE.md']) {
        this.emitFile({
          type: 'asset',
          fileName: `drawing/${name}`,
          source: readFileSync(fileURLToPath(new URL(`./src/drawing/${name}`, import.meta.url))),
        })
      }
    },
  }, {
    name: 'embedded-font-decoder-notices',
    generateBundle() {
      for (const name of ['LICENSE', 'NOTICE.md']) {
        this.emitFile({ type: 'asset', fileName: `fonts/${name}`, source: readFileSync(fileURLToPath(new URL(`./src/core/fonts/vendor/${name}`, import.meta.url))) })
      }
    },
  }, {
    name: 'unicode-vertical-orientation-notices',
    generateBundle() {
      for (const name of ['LICENSE', 'NOTICE.md']) {
        this.emitFile({ type: 'asset', fileName: `drawing/unicode/${name}`, source: readFileSync(fileURLToPath(new URL(`./src/drawing/unicode/${name}`, import.meta.url))) })
      }
    },
  },
  // dist/.module-map.json for scripts/check-eager.ts (eager import-graph gate)
  ...(process.env.OV_MODULE_MAP ? [moduleMap()] : [])],
  build: {
    lib: {
      entry: {
        index: fileURLToPath(new URL('./src/index.ts', import.meta.url)),
        worker: fileURLToPath(new URL('./src/worker/worker-entry.ts', import.meta.url)),
      },
      formats: ['es'],
      // dist/style.css, exported as "officeview/style.css"
      cssFileName: 'style',
    },
    rollupOptions: {
      // React is a peer dependency: keep it and its subpath entries
      // (react/jsx-runtime, react/jsx-dev-runtime, react-dom/*) out of the
      // bundle. A bare 'react' string does not match 'react/jsx-runtime', which
      // used to inline React's CJS JSX runtime (production AND development
      // builds, picked at run time by process.env.NODE_ENV).
      external: [/^react(?:-dom)?(?:\/.*)?$/],
      // The lazy large-source chunk imports shared code from the index entry;
      // allowing the entry to gain internal exports keeps it one eager file
      // instead of a facade plus a shared chunk.
      preserveEntrySignatures: 'allow-extension',
      output: {
        // The lazy chart engine's entry is src/charts/index.ts; name its chunk
        // charts-*.js so it never reads as (or collides with) an index chunk.
        chunkFileNames: chunk => chunk.facadeModuleId?.replace(/\\/g, '/').endsWith('/src/charts/index.ts') ? 'charts-[hash].js' : '[name]-[hash].js',
      },
    },
  },
})
