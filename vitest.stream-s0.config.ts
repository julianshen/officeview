import { defineConfig } from 'vitest/config'

/** Stage-0 Task 1 backend: the real `canvas` package (node-canvas). */
export default defineConfig({
  test: {
    environment: 'node',
    include: ['tests/stream-s0/**/*.test.ts'],
    env: { S0_BACKEND: 'node-canvas' },
    maxWorkers: 2,
  },
})
