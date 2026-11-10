import { defineConfig } from 'vite'
import { fileURLToPath, URL } from 'node:url'
import { moduleMap } from './scripts/vite-module-map'

/**
 * Second build pass: `dist/core.js`, the React-free `officeview/core` entry
 * (spec §8.1). Runs after `vite build` and must not wipe its output.
 */
export default defineConfig({
  plugins: [...(process.env.OV_MODULE_MAP ? [moduleMap('.module-map.core.json')] : [])],
  build: {
    emptyOutDir: false,
    lib: {
      entry: { core: fileURLToPath(new URL('./src/core-entry.ts', import.meta.url)) },
      formats: ['es'],
    },
    rollupOptions: {
      // Same peer-dependency rule as vite.config.ts (purity gate checks none is imported).
      external: [/^react(?:-dom)?(?:\/.*)?$/],
      output: { chunkFileNames: 'core-[name]-[hash].js' },
    },
  },
})
