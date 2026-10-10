import { availableParallelism } from 'node:os'
import { defineConfig } from 'vitest/config'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  test: {
    environment: 'jsdom',
    globals: true,
    include: ['tests/**/*.test.{ts,tsx}'],
    // Bound CPU-heavy canvas/jsdom workers while preserving test assertions.
    maxWorkers: Math.min(2, availableParallelism()),
  },
})
