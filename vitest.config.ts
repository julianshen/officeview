import { availableParallelism } from 'node:os'
import { defineConfig } from 'vitest/config'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  test: {
    environment: 'jsdom',
    globals: true,
    include: ['tests/**/*.test.{ts,tsx}'],
    // Playwright browser specs live under tests/browser and run via
    // `bun run test:browser` — never through vitest.
    exclude: ['tests/browser/**', 'node_modules/**', 'dist/**'],
    // Bound CPU-heavy canvas/jsdom workers while preserving test assertions.
    maxWorkers: Math.min(2, availableParallelism()),
  },
})
