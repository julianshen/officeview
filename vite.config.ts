import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { fileURLToPath, URL } from 'node:url'
import { readFileSync } from 'node:fs'

export default defineConfig({
  plugins: [react(), {
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
  }],
  build: {
    lib: {
      entry: {
        index: fileURLToPath(new URL('./src/index.ts', import.meta.url)),
        worker: fileURLToPath(new URL('./src/worker/worker-entry.ts', import.meta.url)),
      },
      formats: ['es'],
    },
    rollupOptions: {
      external: ['react', 'react-dom'],
    },
  },
})
