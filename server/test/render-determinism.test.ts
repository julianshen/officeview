/** Stage 2b final fix F1: a unit's canonical must not depend on what the
 * process rendered before. A long-lived server's pool worker renders many
 * documents; the trace tool re-renders the reference in a fresh process and
 * compares SHA-256s (fail-closed), so the two must agree. */
import { describe, expect, it } from 'bun:test'
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { decodeRasterBody, encodeRasterBody } from '../../src/stream/raster'
import { decodeShowMessage, encodeShowMessage, SHOW_PLAN_LIMITS } from '../../src/stream/show-plan'
import { flyInDeck } from '../../src/testdata/show-decks'
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

/** Piece 4 Task 9 (T5 item 4, the mixed-page decoder pin re-run): layer assets reuse the unit raster body path. */
describe('show-plan layer assets on the unit raster path', () => {
  it('a layer asset round-trips decodeRasterBody byte-identically', async () => {
    const built = async () => {
      const host = new DocumentHost()
      const { docKey } = await host.open('fly', await flyInDeck(), 'full')
      const r = await host.showPlan(docKey, { unitIndex: 0, scale: 1, motion: 'full', mark: { kind: 'fixed' }, strictMarks: false })
      host.drop(docKey)
      if ('refused' in r) throw new Error(`refused ${r.refused}`)
      return r.assets
    }
    const assets = await built()
    const layers = assets.filter((a) => a.tag === 2)
    expect(layers.length).toBeGreaterThan(0)
    for (const a of assets) {
      // The unit raster body: what is framed is exactly what decodes.
      const body = decodeRasterBody(encodeRasterBody(a.w, a.h, a.png))
      expect({ w: body.w, h: body.h, sha: sha(body.png) }).toEqual({ w: a.w, h: a.h, sha: sha(a.png) })
      // The same body inside a kind-6 asset message, decoded by the client codec.
      const m = decodeShowMessage(encodeShowMessage({ tag: a.tag, unit: 0, assetId: a.assetId, w: a.w, h: a.h, png: a.png }), { unit: 0, maxDecoded: SHOW_PLAN_LIMITS.assetDecoded })
      expect(m).toMatchObject({ tag: a.tag, unit: 0, assetId: a.assetId, w: a.w, h: a.h })
      expect(sha((m as { png: Uint8Array }).png)).toBe(sha(a.png))
    }
    // A second host (no shared cache) builds the same bytes.
    expect((await built()).map((a) => `${a.assetId}:${sha(a.png)}`)).toEqual(assets.map((a) => `${a.assetId}:${sha(a.png)}`))
  }, 60_000)
})
