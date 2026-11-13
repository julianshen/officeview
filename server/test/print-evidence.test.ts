/** Stage 3b Task 4 — evidence: a printed page carries both marks and no hidden content. */
import { afterAll, describe, expect, it } from 'bun:test'
import { createCanvas, loadImage } from '@napi-rs/canvas'
import JSZip from 'jszip'
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
import type { Pool } from '../src/pool'
import { Client, corpus, disposePool, pool, start } from './harness'

afterAll(() => disposePool())
const TAU = { abs: 7.99, margin: 3.85 }
const DOC = 'pypptx-test.pptx'

/** A decoded print page: `u32 LE unitIndex` then the raster body. */
function page(body: Uint8Array) {
  const unitIndex = new DataView(body.buffer, body.byteOffset, body.byteLength).getUint32(0, true)
  return { unitIndex, ...decodeRasterBody(body.subarray(4)) }
}
const bmp = (png: Uint8Array) => { const d = decodePng(png); return { width: d.w, height: d.h, data: d.rgba } }

/** First differing RGBA pixel of two PNGs (decoded), or undefined when size and pixels are equal. */
function firstPixelDiff(a: Uint8Array, b: Uint8Array): string | undefined {
  const x = decodePng(a), y = decodePng(b)
  if (x.w !== y.w || x.h !== y.h) return `size ${x.w}x${x.h} vs ${y.w}x${y.h}`
  for (let i = 0; i < x.rgba.length; i++) {
    if (x.rgba[i] !== y.rgba[i]) {
      const px = i >> 2
      return `pixel (${px % x.w},${Math.floor(px / x.w)}) channel ${i & 3}: ${x.rgba[i]} vs ${y.rgba[i]}`
    }
  }
  return undefined
}

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
    expect(firstPixelDiff(bare.png, printed.png)).toBeDefined()
    c.close(); c2.close(); h.stop(); hOff.stop()
  }, 60_000)
})

/** Prints every unit of `doc` (visible-only, forensic off, default visible mark) through `p`. */
async function printAll(doc: Uint8Array, p: Pool = pool()) {
  const h = await start({ doc, pool: p, policy: { contentPolicy: 'visible-only', forensicMark: false } })
  const c = new Client(h.url), m = await c.hello(); await c.ready()
  const count = m.units.count
  const job = await c.print(10, 0, count - 1)
  const pages = job.pages.map(page)
  const canon = pages.map((q) => h.stream.canonicalOf(m.session.sid, q.unitIndex, job.start!.scale))
  return { count, job, pages, canon, scale: job.start!.scale as number, stop: () => { c.close(); h.stop() } }
}
type Printed = Awaited<ReturnType<typeof printAll>>

/** Per page: same scale, same size, equal decoded pixels. Returns the first mismatch, or undefined. */
function comparePrints(a: Printed, b: Printed): string | undefined {
  if (a.scale !== b.scale) return `scale ${a.scale} vs ${b.scale}`
  if (a.count !== b.count) return `count ${a.count} vs ${b.count}`
  for (let i = 0; i < a.pages.length; i++) {
    const d = firstPixelDiff(a.pages[i]!.png, b.pages[i]!.png)
    if (d) return `unit ${i}: ${d}`
  }
  return undefined
}

/** `doc` with every occurrence of `from` swapped for the same-length `to` in its XML parts. */
async function swapText(doc: Uint8Array, from: string, to: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(doc)
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir || !/\.xml$/.test(name)) continue
    const x = await entry.async('string')
    if (x.includes(from)) zip.file(name, x.split(from).join(to))
  }
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
}

describe('print evidence: T2 hidden constructs never reach a printed page', async () => {
  const fixtures = await augmentedFixtures()
  for (const f of fixtures) it(`${f.id}`, async () => {
    const r = await printAll(f.withHidden), w = await printAll(f.withoutHidden)
    try {
      // The job covers exactly the manifest's units, in order, and no more.
      expect(r.job.start).toMatchObject({ type: 'print-start', pages: r.count })
      expect(r.job.end).toMatchObject({ type: 'print-done', pages: r.count })
      expect(r.pages.map((p) => p.unitIndex)).toEqual(Array.from({ length: r.count }, (_, i) => i))
      // Each printed page is the canonical render of that unit at that scale.
      r.pages.forEach((p, i) => {
        expect(r.canon[i], `${f.id} unit ${i} canonical present`).toBeDefined()
        expect(Buffer.from(p.png).equals(Buffer.from(r.canon[i]!)), `${f.id} unit ${i} printed == canonical`).toBe(true)
      })
      // The decisive check: against the hidden-free control, the same scale, size and decoded pixels, per page.
      expect(comparePrints(r, w), `${f.id}: with-hidden print differs from the hidden-free print`).toBeUndefined()
      // Non-vacuous: the control's visible text is really drawn (change it and the pixels change).
      const other = await printAll(await swapText(f.withoutHidden, f.control, f.control.replace(/^VIS/, 'XYZ')))
      try {
        expect(other.count).toBe(w.count)
        expect(comparePrints(w, other), `${f.id}: visible control ${f.control} must show in the print`).toBeDefined()
      } finally { other.stop() }
    } finally { r.stop(); w.stop() }
  }, 60_000)

  it('mutation: a marker painted into the with-hidden scale-2 raster fails the comparison', async () => {
    const f = fixtures[0]!, real = pool()
    const painted: Pool = new Proxy(real, {
      get(t, k) {
        if (k === 'render') return async (sid: string, job: { scale: number }, est: never) => {
          const out = await (t.render as Pool['render'])(sid, job as never, est)
          if (job.scale !== 2 || out.kind !== 'raster') return out
          const cv = createCanvas(out.w, out.h), ctx = cv.getContext('2d')
          ctx.drawImage(await loadImage(Buffer.from(out.png)), 0, 0)
          ctx.fillStyle = 'red'; ctx.font = '24px sans-serif'; ctx.fillText(f.marker, 20, 40)
          return { ...out, png: new Uint8Array(cv.toBuffer('image/png')) }
        }
        const v = Reflect.get(t, k); return typeof v === 'function' ? v.bind(t) : v
      },
    })
    const r = await printAll(f.withHidden, painted), w = await printAll(f.withoutHidden)
    try {
      // Everything the old evidence asserted still passes under this mutation ...
      r.pages.forEach((p, i) => expect(Buffer.from(p.png).equals(Buffer.from(r.canon[i]!))).toBe(true))
      expect(r.count).toBe(w.count)
      // ... and only the control comparison catches it.
      expect(comparePrints(r, w)).toMatch(/^unit 0: pixel \(/)
    } finally { r.stop(); w.stop() }
  }, 60_000)
})
