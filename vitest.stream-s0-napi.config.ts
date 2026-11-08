import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

/** Stage-0 Task 1 backend: `canvas` aliased to the @napi-rs/canvas napi shim. */
export default defineConfig({
  resolve: {
    alias: [
      {
        find: /^canvas$/,
        replacement: fileURLToPath(new URL('./scripts/stream-s0/napi-shim.ts', import.meta.url)),
      },
    ],
  },
  test: {
    environment: 'node',
    include: ['tests/stream-s0/**/*.test.ts', 'tests/stream/**/*.test.ts'],
    env: { S0_BACKEND: 'napi' },
    maxWorkers: 2,
  },
})
