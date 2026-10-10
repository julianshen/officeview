import { availableParallelism } from 'node:os'
import { configDefaults, defineConfig } from 'vitest/config'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  test: {
    environment: 'jsdom',
    globals: true,
    include: ['tests/**/*.test.{ts,tsx}'],
    // Playwright browser specs live under tests/browser and run via
    // `bun run test:browser` — never through vitest.
    exclude: [...configDefaults.exclude, 'tests/browser/**'],
    // Bound CPU-heavy canvas/jsdom workers while preserving test assertions.
    maxWorkers: Math.min(2, availableParallelism()),
  },
})
