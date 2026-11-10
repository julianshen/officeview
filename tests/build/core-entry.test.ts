/**
 * Built-artifact checks for the `officeview/core` entry (spec §8.1). Runs
 * against `dist/`; build first: `OV_MODULE_MAP=1 bun run build`.
 */
import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import { describe, expect, it } from 'vitest'

const DIST = join(process.cwd(), 'dist')
const built = existsSync(join(DIST, 'core.js'))

describe.skipIf(!built)('officeview/core build (run `OV_MODULE_MAP=1 bun run build` first)', () => {
  it('builds dist/core.js beside the main entries', () => {
    for (const f of ['index.js', 'worker.js', 'core.js', 'core-entry.d.ts']) expect(existsSync(join(DIST, f)), f).toBe(true)
  })

  it('writes a separate module map for the core pass', () => {
    if (!process.env.OV_MODULE_MAP && !existsSync(join(DIST, '.module-map.json'))) return
    expect(existsSync(join(DIST, '.module-map.core.json'))).toBe(true)
    expect(existsSync(join(DIST, '.module-map.json'))).toBe(true)
  })

  it('core.js exports the stream-source surface only', async () => {
    const mod = await import(/* @vite-ignore */ pathToFileURL(join(DIST, 'core.js')).href) as Record<string, unknown>
    expect(Object.keys(mod).sort()).toEqual(['StreamSourceMisuseError', 'isStreamSource', 'loadStreamRuntime'])
    const runtime = await (mod.loadStreamRuntime as () => Promise<{ STREAM_RUNTIME_VERSION: number }>)()
    expect(runtime.STREAM_RUNTIME_VERSION).toBe(1)
  })
})

it.runIf(!built)('officeview/core build checks skipped: dist/core.js absent', () => {
  expect(built).toBe(false)
})
