/** Stage 3b Task 4 — evidence: a printed page carries both marks and no hidden content. */
import { afterAll, describe, expect, it } from 'bun:test'
import { decodeRasterBody } from '../../src/stream/raster'
import { decideScores, scoreCapture } from '../../scripts/stream-s0/mark/detect'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import { decodeImage } from '../src/detect/image'
import { trace } from '../src/detect/trace'
import { deriveKeys } from '../src/keys'
import { basis, deriveBasisKey } from '../src/mark/basis'
import { extendedCodeword } from '../src/mark/bch'
import { messageFor } from '../src/mark/message'
import { DEFAULT_PARAMS } from '../src/mark/params'
import { decodePng } from '../src/mark/png'
import { augmentedFixtures } from './fixtures'
import { Client, corpus, disposePool, start } from './harness'

afterAll(() => disposePool())
const TAU = { abs: 7.99, margin: 3.85 }
const DOC = 'pypptx-test.pptx'

/** A decoded print page: `u32 LE unitIndex` then the raster body. */
function page(body: Uint8Array) {
  const unitIndex = new DataView(body.buffer, body.byteOffset, body.byteLength).getUint32(0, true)
  return { unitIndex, ...decodeRasterBody(body.subarray(4)) }
}
const bmp = (png: Uint8Array) => { const d = decodePng(png); return { width: d.w, height: d.h, data: d.rgba } }

describe('print evidence: the Layer-1 forensic mark', () => {
  it('a printed page detects to its own session at print resolution, and trace() resolves the stored delivery', async () => {
    const keys = await deriveKeys(new Uint8Array(32).fill(9))
    const store = AuditStore.open(':memory:')
    const h = await start({ audit: new AuditTable({ store }) })
    const c = new Client(h.url), m = await c.hello(); await c.ready()
    const sid = m.session.sid
    const job = await c.print(10, 0, 0)
    expect(job.start).toMatchObject({ type: 'print-start', pages: 1, scale: 2 })
    expect(job.end).toMatchObject({ type: 'print-done', pages: 1 })
    expect(job.pages).toHaveLength(1)
    const printed = page(job.pages[0]!)
    expect(printed.unitIndex).toBe(0)

    // Reference: the canonical (unforensic) render at the print scale, the bytes the tracer reproduces.
    const canonical = h.stream.canonicalOf(sid, 0, 2)
    expect(canonical).toBeDefined()
    const f = h.stream.forensicOf(sid)!
    const decoys = Array.from({ length: 62 }, (_, i) => extendedCodeword(messageFor(keys.kMsg, 1, 1000 + i, 0)))
    const chips = basis(deriveBasisKey(keys.kBasisRoot, m.docRef, '0@2'), printed.w, printed.h, DEFAULT_PARAMS)
    const scored = scoreCapture(bmp(printed.png), bmp(canonical!), chips, [f.code, ...decoys], DEFAULT_PARAMS)
    expect(decideScores(scored, TAU)).toMatchObject({ kind: 'attributed', index: 0 })
    // The mark is really in the pixels: the page is not the canonical bytes.
    expect(Buffer.from(printed.png).equals(Buffer.from(canonical!))).toBe(false)
    // And the canonical itself is not attributed (the detector is not fooled by the visible mark alone).
    expect(decideScores(scoreCapture(bmp(canonical!), bmp(canonical!), chips, [f.code, ...decoys], DEFAULT_PARAMS), TAU).kind).not.toBe('attributed')

    // The stored-delivery trace path: the printed PNG resolves to this sid at scale 2.
    const answer = await trace({
      capture: await decodeImage(printed.png), doc: corpus(DOC), unitIndex: 0,
      keys: new Map([[1, await deriveKeys(new Uint8Array(32).fill(9), 1)]]), store,
    })
    expect(answer).toMatchObject({ kind: 'traced', sid, viewerId: 'alice', unitIndex: 0, scale: 2 })

    c.close(); h.stop(); store.close()
  }, 60_000)
})

describe('print evidence: the visible mark', () => {
  it('with forensic off, the printed page is byte-equal to the visible-marked canonical render at scale 2', async () => {
    const h = await start({ policy: { forensicMark: false } })
    const c = new Client(h.url), m = await c.hello(); await c.ready()
    const job = await c.print(10, 0, 0)
    expect(job.start).toMatchObject({ scale: 2 })
    const printed = page(job.pages[0]!)
    const canonical = h.stream.canonicalOf(m.session.sid, 0, 2)
    expect(canonical).toBeDefined()
    expect(Buffer.from(printed.png).equals(Buffer.from(canonical!))).toBe(true)

    // The visible mark is present: the same session with the visible mark off renders different pixels.
    const hOff = await start({ policy: { forensicMark: false, visibleMark: 'off' } })
    const c2 = new Client(hOff.url); await c2.hello(); await c2.ready()
    const bare = page((await c2.print(10, 0, 0)).pages[0]!)
    expect([bare.w, bare.h]).toEqual([printed.w, printed.h])
    expect(Buffer.from(bare.png).equals(Buffer.from(printed.png))).toBe(false)
    c.close(); c2.close(); h.stop(); hOff.stop()
  }, 60_000)
})

describe('print evidence: T2 hidden constructs never reach a printed page', async () => {
  const fixtures = await augmentedFixtures()
  for (const f of fixtures) it(`${f.id}`, async () => {
    const run = async (doc: Uint8Array) => {
      const h = await start({ doc, policy: { contentPolicy: 'visible-only', forensicMark: false } })
      const c = new Client(h.url), m = await c.hello(); await c.ready()
      const count = m.units.count
      const job = await c.print(10, 0, count - 1)
      const pages = job.pages.map(page)
      const canon = pages.map((p) => h.stream.canonicalOf(m.session.sid, p.unitIndex, job.start!.scale))
      const stop = () => { c.close(); h.stop() }
      return { count, job, pages, canon, stop }
    }
    const r = await run(f.withHidden)
    // The job covers exactly the manifest's units, in order, and no more.
    expect(r.job.start).toMatchObject({ type: 'print-start', pages: r.count })
    expect(r.job.end).toMatchObject({ type: 'print-done', pages: r.count })
    expect(r.pages.map((p) => p.unitIndex)).toEqual(Array.from({ length: r.count }, (_, i) => i))
    // Each printed page is the canonical render of that unit at that scale, which the T2 raster tests
    // establish as hidden-free (ops.test.ts).
    r.pages.forEach((p, i) => {
      expect(r.canon[i], `${f.id} unit ${i} canonical present`).toBeDefined()
      expect(Buffer.from(p.png).equals(Buffer.from(r.canon[i]!)), `${f.id} unit ${i} printed == canonical`).toBe(true)
    })
    // The hidden construct is not a page: the unit count matches the document without it.
    const w = await run(f.withoutHidden)
    expect(r.count, `${f.id} unit count equals the visible-only document`).toBe(w.count)
    r.stop(); w.stop()
  }, 60_000)
})
