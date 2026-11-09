/** Stage 2a Task 5: per-session Layer-1 encode under a scratch lease (§9.3/§9.4/§14). */
import { afterAll, describe, expect, it } from 'bun:test'
import JSZip from 'jszip'
import { decodeRasterBody } from '../../src/stream/raster'
import { decideScores, scoreCapture } from '../../scripts/stream-s0/mark/detect'
import { deriveKeys } from '../src/keys'
import { basis, deriveBasisKey } from '../src/mark/basis'
import { extendedCodeword } from '../src/mark/bch'
import { messageFor } from '../src/mark/message'
import { DEFAULT_PARAMS } from '../src/mark/params'
import { decodePng } from '../src/mark/png'
import { Pool } from '../src/pool'
import { Client, corpus, disposePool, start } from './harness'

afterAll(() => disposePool())
const TAU = { abs: 7.99, margin: 3.85 }

async function session(url: string) {
  const c = new Client(url), m = await c.hello(); await c.ready()
  return { c, m }
}
async function rasterOf(c: Client, reqId: number, unitIndex = 0) {
  c.unit(reqId, unitIndex)
  const { unit } = await c.unitMessages(reqId)
  if (unit.discriminator !== 2) throw new Error('expected raster')
  return decodeRasterBody(unit.body)
}
const bmp = (png: Uint8Array) => { const d = decodePng(png); return { width: d.w, height: d.h, data: d.rgba } }

describe('forensic serving path', () => {
  it('two sessions get different bytes, each detecting to its own row; forensic-off gets the canonical entry; one cache entry', async () => {
    const h = await start()
    const keys = await deriveKeys(new Uint8Array(32).fill(9))
    const a = await session(h.url), b = await session(h.url)
    const h0 = h.stream.stats().cache.entries
    const ra = await rasterOf(a.c, 10), rb = await rasterOf(b.c, 10)
    expect(Buffer.from(ra.png).equals(Buffer.from(rb.png))).toBe(false)
    expect(h.stream.stats().cache.entries - h0).toBe(1)
    // Reference: the canonical bytes, as a forensic-off session receives them.
    const hOff = await start({ policy: { forensicMark: false } })
    const o = await session(hOff.url), canonical = await rasterOf(o.c, 10)
    const fa = h.stream.forensicOf(a.m.session.sid)!, fb = h.stream.forensicOf(b.m.session.sid)!
    const decoys = Array.from({ length: 62 }, (_, i) => extendedCodeword(messageFor(keys.kMsg, 1, 1000 + i, 0)))
    const codes = [fa.code, fb.code, ...decoys]
    const chips = basis(deriveBasisKey(keys.kBasisRoot, a.m.docRef, '0@1'), ra.w, ra.h, DEFAULT_PARAMS)
    expect(decideScores(scoreCapture(bmp(ra.png), bmp(canonical.png), chips, codes, DEFAULT_PARAMS), TAU)).toMatchObject({ kind: 'attributed', index: 0 })
    expect(decideScores(scoreCapture(bmp(rb.png), bmp(canonical.png), chips, codes, DEFAULT_PARAMS), TAU)).toMatchObject({ kind: 'attributed', index: 1 })
    expect(decideScores(scoreCapture(bmp(canonical.png), bmp(canonical.png), chips, codes, DEFAULT_PARAMS), TAU).kind).not.toBe('attributed')
    expect(h.stream.stats().forensic).toMatchObject({ marked: 2, unmarked: 0 })
    expect(h.stream.stats().forensic.carriers.every((n) => n > 400)).toBe(true)
    a.c.close(); b.c.close(); o.c.close(); h.stop(); hOff.stop()
  }, 30_000)

  it('a forensic-off session streams the canonical cache bytes and runs no encode job', async () => {
    const h = await start({ policy: { forensicMark: false } })
    const before = h.stream.stats().pool.encodes
    const o = await session(h.url)
    const r = await rasterOf(o.c, 10)
    expect(h.stream.stats().pool.encodes).toBe(before)
    expect(h.stream.canonicalOf(o.m.session.sid, 0, 1)).toEqual(r.png)
    o.c.close(); h.stop()
  })

  it('a relaxed session receives command units with no encode job', async () => {
    const h = await start({ policy: { projection: 'relaxed' } })
    const before = h.stream.stats().pool.encodes
    const s = await session(h.url)
    s.c.unit(10, 0)
    const { unit } = await s.c.unitMessages(10)
    expect(unit.discriminator).toBe(1)
    expect(h.stream.stats().pool.encodes).toBe(before)
    s.c.close(); h.stop()
  })

  it('a unit below carrier capacity ships canonical bytes plus diag forensic-unmarked', async () => {
    const zip = await JSZip.loadAsync(corpus('pypptx-test.pptx'))
    const pres = await zip.file('ppt/presentation.xml')!.async('string')
    zip.file('ppt/presentation.xml', pres.replace(/<p:sldSz[^>]*\/>/, '<p:sldSz cx="952500" cy="476250"/>'))
    const tiny = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
    const h = await start({ doc: tiny })
    const s = await session(h.url)
    const r = await rasterOf(s.c, 10)
    expect([r.w, r.h]).toEqual([100, 50])
    const diag = await s.c.next((m) => 'text' in m && m.text.type === 'diag')
    expect((diag as { text: Record<string, unknown> }).text).toMatchObject({ reqId: 10, code: 'forensic-unmarked', unitIndex: 0, reason: 'capacity' })
    expect(h.stream.canonicalOf(s.m.session.sid, 0, 1)).toEqual(r.png)
    expect(h.stream.stats().forensic).toMatchObject({ marked: 0, unmarked: 1 })
    s.c.close(); h.stop()
  })

  it('the scratch lease admits at most 2 encodes per session', async () => {
    const pool = new Pool({ workers: 3, testHooks: true })
    const h = await start({ pool, encodeHangMs: 700 })
    const s = await session(h.url)
    await rasterOf(s.c, 9) // warm: render + first encode
    const base = pool.stats().encodes
    for (const r of [10, 11, 12]) s.c.unit(r, 0)
    await Bun.sleep(350)
    expect(pool.stats().encodes - base).toBe(2)
    expect(h.stream.stats().forensic.leased).toBe(2)
    for (const r of [10, 11, 12]) await s.c.unitMessages(r)
    expect(pool.stats().encodes - base).toBe(3)
    expect(h.stream.stats().forensic.leased).toBe(0)
    s.c.close(); h.stop(); pool.dispose()
  }, 30_000)

  it('closing mid-encode: no frame, lease back to 0, charge refunded only after the worker replies', async () => {
    const pool = new Pool({ workers: 2, testHooks: true })
    const h = await start({ pool, encodeHangMs: 800 })
    const s = await session(h.url)
    s.c.unit(10, 0)
    await Bun.sleep(400)
    expect(h.stream.stats().forensic.leased).toBe(1)
    s.c.close(); await s.c.closed
    await Bun.sleep(50)
    expect(pool.stats().reserved).toBeGreaterThan(0)          // the worker is still encoding
    for (let i = 0; i < 40 && pool.stats().reserved > 0; i++) await Bun.sleep(50)
    expect(pool.stats().reserved).toBe(0)
    expect(h.stream.stats().forensic.leased).toBe(0)
    expect(s.c.msgs.filter((m) => 'binary' in m && m.binary.reqId === 10)).toEqual([])
    h.stop(); pool.dispose()
  }, 30_000)
})
