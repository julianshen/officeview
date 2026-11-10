/**
 * Stage-0 Task 1 — bun preload for napi-backed scripts.
 *
 * `bun scripts/stream-s0/*.ts` runs in bun's own module resolution, where the
 * vitest alias cannot help. This preload maps the `canvas` specifier to the
 * shim module (import return type = the napi-backed backend) via a runtime
 * `Bun.plugin` before the script's own top-level imports resolve:
 *
 *   bun --preload ./scripts/stream-s0/napi-preload.ts scripts/stream-s0/<x>.ts
 *
 * `package.json` exposes `"s0:napi": "bun --preload ./scripts/stream-s0/napi-preload.ts"`,
 * so `bun run s0:napi scripts/stream-s0/<x>.ts` runs a probe/script under the
 * napi backend (and `results.resolveBackend()` sees the same module).
 */
import * as shim from './napi-shim'

interface BunPluginBuild {
  module(absSpecifier: string, factory: () => { exports: object; loader: string }): void
}
interface BunRuntime {
  plugin(pluginSpec: { name: string; setup: (build: BunPluginBuild) => void }): void
}

const runtime = (globalThis as unknown as { Bun?: BunRuntime }).Bun

if (!runtime) {
  throw new Error('napi-preload requires the bun runtime (it registers a bun module map for "canvas")')
}

runtime.plugin({
  name: 'stream-s0-napi',
  setup: (build) => {
    // Serves the shim module's exports for every `import('canvas')` in the
    // script (including src/ dynamic imports reached through the corpus).
    build.module('canvas', () => ({ exports: shim, loader: 'object' }))
  },
})
