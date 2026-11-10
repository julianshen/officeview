/** Stage 2b final fix F1: a unit's canonical must not depend on what the
 * process rendered before. A long-lived server's pool worker renders many
 * documents; the trace tool re-renders the reference in a fresh process and
 * compares SHA-256s (fail-closed), so the two must agree. */
import { describe, expect, it } from 'bun:test'
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { DocumentHost, type MarkSpec } from '../src/render'

const root = resolve(import.meta.dir, '../..')
const PRELOAD = resolve(root, 'scripts/stream-s0/napi-preload.ts')
const ALONE = resolve(import.meta.dir, 'render-alone.ts')
const sha = (b: Uint8Array) => createHash('sha256').update(b).digest('hex')
const bytes = (name: string) => new Uint8Array(readFileSync(resolve(root, 'corpus', name)))

async function freshProcessHash(name: string, unit: number, scale: number, mark: MarkSpec): Promise<string> {
  const p = Bun.spawn([process.execPath, '--preload', PRELOAD, ALONE, name, String(unit), String(scale), JSON.stringify(mark)], { stdout: 'pipe', stderr: 'pipe' })
  const [out, err, code] = await Promise.all([new Response(p.stdout).text(), new Response(p.stderr).text(), p.exited])
  if (code !== 0) throw new Error(`render-alone failed: ${err}`)
  return out
}

async function hashIn(host: DocumentHost, name: string, unit: number, scale: number, mark: MarkSpec): Promise<string> {
  const { docKey } = await host.open(name, bytes(name), 'visible-only')
  const r = host.render(docKey, { unitIndex: unit, scale, projection: 'final-raster', mark })
  if (r.kind !== 'raster') throw new Error('expected raster')
  return sha(r.png)
}

describe('reference determinism across a long-lived process', () => {
  const A = 'lo-calc-pivotcompact.xlsx', B = 'lo-bnc591147.pptx'
  for (const mark of [{ kind: 'fixed' }, { kind: 'viewer', text: 'viewer-7 · 2026-11-10' }] as MarkSpec[]) {
    it(`${A}#0@1 after ${B} equals a fresh process (${mark.kind} mark)`, async () => {
      const alone = await freshProcessHash(A, 0, 1, mark)
      const host = new DocumentHost()
      await hashIn(host, B, 0, 1, mark)
      expect(await hashIn(host, A, 0, 1, mark)).toBe(alone)
      // and the other order: B after A equals B alone
      expect(await hashIn(host, B, 0, 1, mark)).toBe(await freshProcessHash(B, 0, 1, mark))
    }, 60_000)
  }
})
