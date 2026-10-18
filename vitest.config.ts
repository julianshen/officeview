import { availableParallelism } from 'node:os'
import { configDefaults, defineConfig } from 'vitest/config'
import react from '@vitejs/plugin-react'
import { ovCss } from './scripts/ov-css-config'

export default defineConfig({
  // vitest has its own config: the same CSS pipeline, spread explicitly.
  css: ovCss.css,
  plugins: [react(), ...ovCss.plugins],
  test: {
    environment: 'jsdom',
    globals: true,
    include: ['tests/**/*.test.{ts,tsx}'],
    // Playwright browser specs live under tests/browser and run via
    // `bun run test:browser` — never through vitest.
    exclude: [...configDefaults.exclude, 'tests/browser/**'],
    // Every jsdom test gets the compiled stylesheet (tokens, sentinel).
    setupFiles: ['tests/setup/styles.ts'],
    // vitest blanks CSS it is not told to process; run the real pipeline
    // for the officeview stylesheet.
    css: { include: [/src\/styles\/officeview\.css/] },
    // Bound CPU-heavy canvas/jsdom workers while preserving test assertions.
    maxWorkers: Math.min(2, availableParallelism()),
  },
})
