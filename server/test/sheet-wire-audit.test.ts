/**
 * Stage 4a Task 15 — the 4a wire audit (S4 §3; parent §3: what never reaches the client is checked on transmitted
 * payloads, not screenshots). A tiles session over the whole server path (`hello` with `sheet-tiles@1`, `ready`,
 * `sheet-layout` of every visible sheet, then every tile of each sheet's used extent at (zoomStep 5, dpr 1) and
 * (zoomStep 5, dpr 2)) is captured byte for byte by a listener of its own on the socket: every text frame as received,
 * every binary frame as received, and every binary message reassembled (inflated) by an independent `Reassembler`.
 *
 * - **no sentinel on the wire**: no decoded text (text frames, the manifest, every reassembled payload read as
 *   latin1, the raw frames too) contains a sentinel of `sentinels()` or a fixture's `marker` / `control`; no
 *   transmitted JSON string equals a hidden sheet's name (the Stage-0 fixtures call it `H`, which a substring search
 *   cannot test), and the manifest's sheet names are the visible ones.
 * - **no drawing byte on the wire**: the hidden picture's PNG does not occur in any frame or payload, nor in the
 *   inflated IDAT of any tile PNG. "PNG IDAT included" is read both ways: the needles are the whole file, its IDAT
 *   chunk data, and 64-byte windows of that data.
 *
 * Non-vacuity: every needle is shown present in the source package; the session is in tiles mode; every tile asked
 * for arrives (`refused: []`), and the tile set covers the hidden picture's box at both pairs; a liveness probe finds
 * a tile's own PNG in the captured payloads and the visible sheet name in the decoded manifest.
 *
 * Policy: visible-only (the ruling: the wire audit runs under visible-only; 'full' is T2's only use).
 */
import { afterAll, beforeAll, describe, expect, test } from 'bun:test'
import JSZip from 'jszip'
import { inflateSync } from 'node:zlib'
import { Reassembler } from '../../src/stream/frame'
import { decodeSheetLayout, decodeTileBody, TILE_ZOOM_STEPS } from '../../src/stream/sheet-wire'
import { streamSheetGeometry } from '../../src/xlsx/view/streamGeometry'
import type { AuthDecision } from '../src/hooks'
import { tileLimit } from '../src/sheet/tile'
import { augmentedFixtures } from './fixtures'
import { Client, disposePool, start } from './harness'
import { SENTINELS, sentinelPng, sentinels } from './sheet-fixtures'

afterAll(async () => { await disposePool() })

const TILES = ['sheet-tiles@1']
const POLICY: AuthDecision['policy'] = { contentPolicy: 'visible-only' }
/** (zoomStep, dpr) pairs the audit fetches. */
const PAIRS = [[5, 1], [5, 2]] as const
/** The harness client's decoded caps per kind (encoded, decoded). */
const CAPS: Record<number, [number, number]> = { 1: [256 << 10, 256 << 10], 2: [1 << 20, 8 << 20], 3: [8 << 20, 64 << 20], 4: [8 << 20, 8 << 20], 5: [8 << 20, 64 << 20], 6: [8 << 20, 32 << 20], 7: [8 << 20, 64 << 20], 8: [8 << 20, 64 << 20], 9: [4 << 20, 4 << 20] }

const latin1 = (b: Uint8Array) => Buffer.from(b.buffer, b.byteOffset, b.byteLength).toString('latin1')
const contains = (hay: Uint8Array, needle: Uint8Array) => Buffer.from(hay.buffer, hay.byteOffset, hay.byteLength).indexOf(Buffer.from(needle.buffer, needle.byteOffset, needle.byteLength)) >= 0

/** The concatenated IDAT chunk data of a PNG. */
function idatOf(png: Uint8Array): Uint8Array {
  const v = new DataView(png.buffer, png.byteOffset, png.byteLength), parts: Uint8Array[] = []
  for (let at = 8; at + 12 <= png.length;) {
    const len = v.getUint32(at)
    if (String.fromCharCode(...png.subarray(at + 4, at + 8)) === 'IDAT') parts.push(png.subarray(at + 8, at + 8 + len))
    at += 12 + len
  }
  return new Uint8Array(Buffer.concat(parts))
}

/** Everything the server sent one client: raw frames, and binary messages reassembled. */
interface Capture { texts: string[]; frames: Uint8Array[]; payloads: Array<{ kind: number; reqId: number; payload: Uint8Array }>; drain(): Promise<void> }
function capture(c: Client): Capture {
  const cap: Capture = { texts: [], frames: [], payloads: [], drain: () => inbox }
  const r = new Reassembler({ acceptedVersions: new Set([1]), expects: () => true, decodedCap: (k) => CAPS[k]?.[1] ?? 0, encodedCap: (k) => CAPS[k]?.[0] ?? 0 })
  let inbox = Promise.resolve()
  c.ws.addEventListener('message', (e) => {
    if (typeof e.data === 'string') { cap.texts.push(e.data); return }
    const frame = new Uint8Array((e.data as ArrayBuffer).slice(0))
    cap.frames.push(frame)
    inbox = inbox.then(async () => {
      const res = await r.push(frame)
      if (res && !res.ok) throw new Error(`audit reassembler rejected a frame: ${res.reason}`)
      if (res?.ok) cap.payloads.push({ kind: res.kind, reqId: res.reqId, payload: res.payload })
    })
  })
  return cap
}

interface Session { cap: Capture; manifest: Record<string, any>; tiles: Array<{ sheetId: string; zoomStep: number; dpr: number; cols: number; rows: number; got: number }>; tilePngs: Uint8Array[] }
/** One visible-only tiles session over `doc`: layout of every visible sheet, then every tile of its used extent at both pairs. */
async function audited(doc: Uint8Array): Promise<Session> {
  const h = await start({ doc, policy: POLICY })
  const c = new Client(h.url)
  const cap = capture(c)
  try {
    const m = await c.hello({ features: TILES })
    await c.ready()
    expect(m.sheetMode).toBe('tiles')
    const out: Session = { cap, manifest: m as never, tiles: [], tilePngs: [] }
    let reqId = 10
    for (const { sheetId } of m.sheets!) {
      const lr = await c.sheetLayout(++reqId, sheetId)
      expect(lr.end).toBeUndefined()
      expect(lr.parts.length).toBeGreaterThan(0)
      const layout = decodeSheetLayout(lr.parts.map((p) => p.text))
      const g = streamSheetGeometry(layout), mt = g.metrics
      const right = layout.used ? mt.cols.start(layout.used.c1) + mt.cols.size(layout.used.c1) : 0
      const bottom = layout.used ? mt.rows.start(layout.used.r1) + mt.rows.size(layout.used.r1) : 0
      for (const [zoomStep, dpr] of PAIRS) {
        const zoom = TILE_ZOOM_STEPS[zoomStep]!
        const cols = tileLimit(right, zoom, dpr), rows = tileLimit(bottom, zoom, dpr)
        const coords = Array.from({ length: cols * rows }, (_, i) => ({ col: i % cols, row: Math.floor(i / cols) }))
        let got = 0
        for (let i = 0; i < coords.length; i += 64) {
          const r = await c.tiles(++reqId, { sheetId, viewGen: 0, zoomStep, dpr, palette: 'light', tiles: coords.slice(i, i + 64) }, { ms: 60_000 })
          expect(r.end).toEqual({ reqId, type: 'tiles-done', refused: [] })
          got += r.tiles.length
          out.tilePngs.push(...r.tiles.map((t) => t.png))
        }
        out.tiles.push({ sheetId, zoomStep, dpr, cols, rows, got })
        expect(got).toBe(cols * rows)
      }
    }
    await cap.drain()
    return out
  } finally { c.close(); await h.stop() }
}

/** Every decoded text of a session: text frames, the manifest, every reassembled payload and every raw frame (latin1). */
const decodedTexts = (s: Session): string[] => [...s.cap.texts, ...s.cap.payloads.map((p) => latin1(p.payload)), ...s.cap.frames.map(latin1)]
/** Every string leaf of every JSON text frame and of the manifest. */
function stringLeaves(s: Session): string[] {
  const out: string[] = []
  const walk = (v: unknown): void => {
    if (typeof v === 'string') out.push(v)
    else if (Array.isArray(v)) v.forEach(walk)
    else if (v && typeof v === 'object') for (const [k, x] of Object.entries(v)) { out.push(k); walk(x) }
  }
  for (const t of s.cap.texts) walk(JSON.parse(t))
  for (const p of s.cap.payloads) if (p.kind === 1) walk(JSON.parse(new TextDecoder().decode(p.payload)))
  return out
}
/** Each needle that occurs in any decoded text (an empty result is a pass). */
const leaked = (s: Session, needles: readonly string[]) => { const hay = decodedTexts(s); return needles.filter((n) => hay.some((h) => h.includes(n))) }
/** The text of every part of a package (the needles' presence in the source). */
async function partsText(doc: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(doc)
  const names = Object.keys(zip.files).filter((n) => !zip.files[n]!.dir)
  return (await Promise.all(names.map(async (n) => latin1(await zip.files[n]!.async('uint8array'))))).join('\n')
}

describe('4a wire audit (visible-only)', () => {
  let sent: Session
  const needles = [SENTINELS.formula, SENTINELS.formulaText, SENTINELS.number, SENTINELS.definedName, SENTINELS.comment, SENTINELS.hiddenSheet, SENTINELS.hiddenCell, SENTINELS.visibleText, SENTINELS.drawingName, 'LEN(']
  beforeAll(async () => { sent = await audited(await sentinels()) }, 120_000)

  test('the audit is live: needles are in the source, tiles arrive, the capture sees them', async () => {
    const src = await partsText(await sentinels())
    // OOXML stores the formula without its `=` (`<f>LEN("SENTINEL_F")</f>`): `=LEN` is the display form a leak would take.
    expect(needles.filter((n) => !src.includes(n))).toEqual([SENTINELS.formulaText])
    const zip = await JSZip.loadAsync(await sentinels())
    expect(contains(await zip.file('xl/media/image1.png')!.async('uint8array'), sentinelPng())).toBe(true)
    // The hidden picture (E2:H8: x 256..512, y 20..160 CSS px) lies inside the fetched tile extent at both pairs.
    expect(sent.manifest.sheets).toEqual([{ sheetId: '0', name: 'S' }])
    for (const t of sent.tiles) {
      const s = TILE_ZOOM_STEPS[t.zoomStep]! * t.dpr
      expect([t.zoomStep, t.dpr, t.cols * 512 >= 512 * s, t.rows * 512 >= 160 * s, t.got]).toEqual([t.zoomStep, t.dpr, true, true, t.cols * t.rows])
    }
    expect(sent.tiles.map((t) => [t.dpr, t.cols, t.rows])).toEqual([[1, 2, 1], [2, 3, 1]])
    // Liveness: the searches find what is there.
    const tilePayloads = sent.cap.payloads.filter((p) => p.kind === 5)
    expect(tilePayloads.length).toBe(sent.tilePngs.length)
    expect(contains(tilePayloads[0]!.payload, sent.tilePngs[0]!)).toBe(true)
    expect(decodedTexts(sent).some((t) => t.includes('"name":"S"'))).toBe(true)
    expect(leaked(sent, ['sheet-layout', 'tiles-done'])).toEqual(['sheet-layout', 'tiles-done'])
    // Every kind-5 payload decodes as a 512² tile.
    for (const p of tilePayloads) expect([decodeTileBody(p.payload).w, decodeTileBody(p.payload).h]).toEqual([512, 512])
  })

  test('no sentinel on the wire', () => {
    expect(leaked(sent, [...needles, 'SENTINEL_'])).toEqual([])
    expect(stringLeaves(sent).filter((s) => s.includes('SENTINEL'))).toEqual([])
  })

  test('no drawing byte on the wire', () => {
    const png = sentinelPng(), idat = idatOf(png)
    expect(idat.length).toBeGreaterThan(1000)
    const windows = [0, 256, 512, idat.length - 64].map((at) => idat.subarray(at, at + 64))
    const drawingNeedles = [png, idat, ...windows]
    const hay = [...sent.cap.frames, ...sent.cap.payloads.map((p) => p.payload), ...sent.tilePngs, ...sent.tilePngs.map((t) => new Uint8Array(inflateSync(idatOf(t))))]
    expect(hay.length).toBeGreaterThan(sent.tilePngs.length * 2)
    const hits = drawingNeedles.map((n, i) => hay.filter((h) => contains(h, n)).length > 0 ? i : -1).filter((i) => i >= 0)
    expect(hits).toEqual([])
    // Positive control: the same search finds the needles in a haystack that holds them.
    expect(drawingNeedles.every((n) => contains(new Uint8Array(Buffer.concat([sent.tilePngs[0]!, png])), n))).toBe(true)
  })

  test('no hidden-fixture marker, control or hidden sheet name on the wire', async () => {
    const xlsx = (await augmentedFixtures()).filter((f) => f.kind === 'xlsx')
    expect(xlsx.map((f) => f.id).sort()).toEqual(['xlsx-comment', 'xlsx-hidden-cols', 'xlsx-hidden-rows', 'xlsx-hidden-sheet', 'xlsx-veryhidden-sheet'])
    for (const f of xlsx) {
      const src = await partsText(f.withHidden)
      expect([f.id, src.includes(f.marker), src.includes(f.control)]).toEqual([f.id, true, true])
      const s = await audited(f.withHidden)
      expect([f.id, s.tiles.length, s.tiles.every((t) => t.got > 0)]).toEqual([f.id, PAIRS.length, true])
      expect([f.id, leaked(s, [f.marker, f.control, 'MARKER-', 'VIS-'])]).toEqual([f.id, []])
      // The hidden sheet (`H`) is named nowhere: the manifest lists the visible sheet only, and no transmitted string is it.
      expect([f.id, s.manifest.sheets]).toEqual([f.id, [{ sheetId: '0', name: 'S' }]])
      expect([f.id, stringLeaves(s).filter((x) => x === 'H')]).toEqual([f.id, []])
    }
  }, 300_000)
})
