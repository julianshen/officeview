/** Stage 4a Task 11: the client tile and layout channel — opt-in HELLO, `sheetLayout`, `tiles`, hostile-server drops, resume. */
import { describe, expect, it } from 'vitest'
import { gzipSync } from 'node:zlib'
import { randomBytes } from 'node:crypto'
import { encodeMessage } from '../../src/stream/frame'
import { encodeManifest, type Manifest } from '../../src/stream/manifest'
import { encodeRasterBody } from '../../src/stream/raster'
import { encodeUnit } from '../../src/stream/unit'
import { encodeSheetLayout, encodeTileBody, TILE_PX, type StreamSheetLayout } from '../../src/stream/sheet-wire'
import { StreamSession, type SessionOptions } from '../../src/stream/remote/session'
import { SHEET_REPLY_MS, TILE_RETRY_FALLBACK_MS, TILE_RETRY_MAX_MS, type DecodedTile } from '../../src/stream/remote/sheet-channel'
import { StreamSessionError } from '../../src/stream/remote/errors'
import { openRemoteDocument } from '../../src/stream/remote/open'

const b64url = (o: object) => btoa(JSON.stringify(o)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
const tokenFor = (jti: string) => `h.${b64url({ jti })}.s`
const settle = (ms = 20) => new Promise((r) => setTimeout(r, ms))
const FEATURES = ['sheet-tiles@1']

const unitsManifest = (): Manifest => ({
  v: 1, docRef: 'a'.repeat(24), kind: 'xlsx', name: 'book.xlsx',
  units: { count: 2, list: [{ w: 960, h: 540 }, { w: 960, h: 540 }] },
  layoutFingerprint: 'f'.repeat(64),
  session: { sid: '0123456789ab', token: tokenFor('j1'), expiresSec: 300, resume: 'session' },
  policy: { contentPolicy: 'visible-only', protection: 'final-raster', semantic: 'visible-only', contentPolicyId: 'visible-only@1', viewerMarkId: 'fixed' },
  can: { search: true, copy: false, print: true, textA11y: false, present: false, notes: false, links: false },
  caps: { copy: null, a11y: 'off', print: { maxPages: null, maxSessionPages: null } },
  recorder: { commands: 'ovsd-1', textMode: 'text', scalesBuckets: [1, 2], requiredFaces: [], projection: 'final-raster' },
  limits: { maxUnitPx: 8192, maxAssetBytes: 8 * 1024 * 1024, wsFrameBytes: 1 << 20, tilePx: 512 },
  notices: [],
})
const tilesManifest = (): Manifest => ({
  ...unitsManifest(),
  sheetMode: 'tiles',
  sheets: [{ sheetId: '0', name: 'Data' }, { sheetId: '1', name: 'Summary' }],
  tiles: { px: 512, zoomSteps: [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4], dprs: [1, 2], palettes: ['light', 'dark'], maxBatch: 64 },
})

class FakeSocket extends EventTarget {
  binaryType: BinaryType = 'blob'
  readyState = 0
  sent: Array<Record<string, any>> = []
  closes: Array<{ code?: number; reason?: string }> = []
  send(d: string) { this.sent.push(JSON.parse(d)) }
  close(code?: number, reason?: string) {
    this.closes.push({ code, reason })
    if (this.readyState === 3) return
    this.readyState = 3
    queueMicrotask(() => this.dispatchEvent(Object.assign(new Event('close'), { code: code ?? 1005, reason: reason ?? '' })))
  }
  open() { this.readyState = 1; this.dispatchEvent(new Event('open')) }
  deliver(data: ArrayBuffer | string) { this.dispatchEvent(Object.assign(new Event('message'), { data })) }
  text(m: object) { this.deliver(JSON.stringify(m)) }
  frame(kind: number, reqId: number, payload: Uint8Array) {
    for (const f of encodeMessage({ ver: 1, kind: kind as never, reqId, payload, compress: (b) => new Uint8Array(gzipSync(b)) })) this.deliver(f.slice().buffer)
  }
  of(type: string) { return this.sent.filter((m) => m.type === type) }
  drop(code: number) { this.readyState = 3; this.dispatchEvent(Object.assign(new Event('close'), { code, reason: '' })) }
}

function pngHeader(w: number, h: number) {
  const b = new Uint8Array(33), v = new DataView(b.buffer)
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]); v.setUint32(8, 13); b.set([73, 72, 68, 82], 12)
  v.setUint32(16, w); v.setUint32(20, h); b.set([8, 6, 0, 0, 0], 24)
  return b
}
type TileOver = Partial<{ sheetIndex: number; viewGen: number; zoomStep: number; dpr: 1 | 2; png: Uint8Array }>
/** One XLSX_TILE body for `(col, row)` of the default request (sheet 0, viewGen 0, zoomStep 5, dpr 1). */
const tile = (col: number, row: number, o: TileOver = {}) => encodeTileBody({ sheetIndex: 0, viewGen: 0, zoomStep: 5, dpr: 1, col, row, png: pngHeader(TILE_PX, TILE_PX), w: TILE_PX, h: TILE_PX, ...o })
/** A tile body whose raster claims 256 × 512 (encodeTileBody refuses to make one). */
function smallTile(col: number, row: number) {
  const head = new Uint8Array(16), v = new DataView(head.buffer)
  v.setUint32(0, 0, true); v.setUint8(4, 5); v.setUint8(5, 1); v.setUint32(8, col, true); v.setUint32(12, row, true)
  const raster = encodeRasterBody(256, TILE_PX, pngHeader(256, TILE_PX))
  const body = new Uint8Array(16 + raster.length); body.set(head); body.set(raster, 16)
  return encodeUnit({ unitId: 0, discriminator: 3, body })
}

async function opened(o: Partial<SessionOptions> = {}, m: Manifest = tilesManifest()) {
  const socks: FakeSocket[] = []
  const client = () => { const s = new FakeSocket(); socks.push(s); return s as never }
  const p = StreamSession.open({ kind: 'stream', url: 'wss://x/stream' }, { client, faces: async () => [], features: FEATURES, ...o })
  const sock = socks[0]!
  sock.open(); await settle()
  sock.frame(1, sock.of('hello')[0]!.reqId, encodeManifest(m))
  const s = await p
  return { s, sock, socks, manifest: m }
}
const req = (tiles: Array<{ col: number; row: number }>, o: Record<string, unknown> = {}) => ({ sheetId: '0', viewGen: 0, zoomStep: 5, dpr: 1 as const, palette: 'light' as const, tiles, ...o })
const key = (t: { col: number; row: number }) => `${t.col},${t.row}`

/** Starts a batch and returns its promise, its reqId and the tiles delivered to `onTile`. */
async function batch(s: StreamSession, sock: FakeSocket, tiles: Array<{ col: number; row: number }>, signal?: AbortSignal) {
  const got: DecodedTile[] = []
  const p = s.tiles(req(tiles), (t) => { got.push(t) }, signal)
  void p.catch(() => {})
  await settle()
  return { p, reqId: sock.of('tiles').at(-1)!.reqId as number, got }
}

describe('client sheet session (S4 §6.10–6.11)', () => {
  it('no sheetMode from the server → units', async () => {
    const socks: FakeSocket[] = []
    const client = () => { const s = new FakeSocket(); socks.push(s); return s as never }
    const p = openRemoteDocument({ kind: 'stream', url: 'wss://x/stream' }, { client, streamSheets: 'tiles', faces: async () => [] })
    await settle(); socks[0]!.open(); await settle()
    const hello = socks[0]!.of('hello')[0]!
    expect(hello.features).toEqual(FEATURES)
    socks[0]!.frame(1, hello.reqId, encodeManifest(unitsManifest()))
    const doc = await p
    expect(doc.sheets).toBeUndefined()
    doc.close()
  })

  it('tiles mode: the doc lists the sheets and routes layout and tiles to the session', async () => {
    const socks: FakeSocket[] = []
    const client = () => { const s = new FakeSocket(); socks.push(s); return s as never }
    const p = openRemoteDocument({ kind: 'stream', url: 'wss://x/stream' }, { client, streamSheets: 'tiles', faces: async () => [] })
    await settle(); socks[0]!.open(); await settle()
    socks[0]!.frame(1, socks[0]!.of('hello')[0]!.reqId, encodeManifest(tilesManifest()))
    const doc = await p
    expect(doc.sheets?.list).toEqual([{ sheetId: '0', name: 'Data' }, { sheetId: '1', name: 'Summary' }])
    const got: DecodedTile[] = []
    const t = doc.sheets!.tiles(req([{ col: 0, row: 0 }]), (x) => { got.push(x) })
    await settle()
    const r = socks[0]!.of('tiles')[0]!
    socks[0]!.frame(5, r.reqId, tile(0, 0))
    socks[0]!.text({ reqId: r.reqId, type: 'tiles-done', refused: [] })
    expect(await t).toEqual({ refused: [] })
    expect(got).toHaveLength(1)
    doc.close()
  })

  it('tiles deliver and finish', async () => {
    const { s, sock } = await opened()
    const { p, reqId, got } = await batch(s, sock, [{ col: 0, row: 0 }, { col: 1, row: 0 }])
    expect(sock.of('tiles')[0]).toEqual({ reqId, type: 'tiles', token: tokenFor('j1'), sheetId: '0', viewGen: 0, zoomStep: 5, dpr: 1, palette: 'light', tiles: [{ col: 0, row: 0 }, { col: 1, row: 0 }] })
    sock.frame(5, reqId, tile(0, 0))
    sock.frame(5, reqId, tile(1, 0))
    sock.text({ reqId, type: 'tiles-done', refused: [] })
    expect(await p).toEqual({ refused: [] })
    expect(got.map(key)).toEqual(['0,0', '1,0'])
    expect(got[0]).toMatchObject({ sheetIndex: 0, viewGen: 0, zoomStep: 5, dpr: 1, col: 0, row: 0 })
    expect(got[0]!.png).toEqual(pngHeader(TILE_PX, TILE_PX))
  })

  it('hostile tiles are dropped before decode', async () => {
    const states: string[] = []
    const { s, sock } = await opened({ onState: (st) => { states.push(st) } })
    const { p, reqId, got } = await batch(s, sock, [{ col: 0, row: 0 }, { col: 1, row: 0 }])
    sock.frame(5, 999, tile(0, 0)) // an unknown reqId
    sock.frame(5, reqId, tile(5, 5)) // an unrequested coordinate
    sock.frame(5, reqId, tile(0, 0)) // the real one
    sock.frame(5, reqId, tile(0, 0)) // a duplicate
    sock.frame(5, reqId, tile(1, 0, { viewGen: 1 })) // viewGen off by one
    sock.frame(5, reqId, tile(1, 0, { sheetIndex: 1 })) // sheetIndex off by one
    sock.frame(5, reqId, smallTile(1, 0)) // w = 256
    // Over the 8 MiB encoded cap (incompressible).
    const big = new Uint8Array(randomBytes(9 * 1024 * 1024)); big.set(pngHeader(TILE_PX, TILE_PX))
    sock.frame(5, reqId, tile(1, 0, { png: big }))
    await settle(200)
    expect(got.map(key)).toEqual(['0,0'])
    sock.frame(5, reqId, tile(1, 0))
    sock.text({ reqId, type: 'tiles-done', refused: [] })
    expect(await p).toEqual({ refused: [] })
    sock.frame(5, reqId, tile(1, 0)) // a tile after tiles-done
    await settle()
    expect(got.map(key)).toEqual(['0,0', '1,0'])
    expect(states).toEqual(['open'])
    expect(sock.closes).toEqual([])
    // The session still works.
    const next = await batch(s, sock, [{ col: 2, row: 0 }])
    sock.frame(5, next.reqId, tile(2, 0))
    sock.text({ reqId: next.reqId, type: 'tiles-done', refused: [] })
    expect(await next.p).toEqual({ refused: [] })
    expect(next.got.map(key)).toEqual(['2,0'])
  })

  it('abort sends sheet-cancel', async () => {
    const { s, sock } = await opened()
    const ac = new AbortController()
    const { p, reqId, got } = await batch(s, sock, [{ col: 0, row: 0 }, { col: 1, row: 0 }], ac.signal)
    sock.frame(5, reqId, tile(0, 0))
    await settle()
    ac.abort()
    expect(sock.of('sheet-cancel')).toEqual([{ reqId: expect.any(Number), type: 'sheet-cancel', target: reqId }])
    expect(await p).toEqual({ refused: [], cancelled: true })
    // Nothing more is delivered for it.
    sock.frame(5, reqId, tile(1, 0))
    sock.text({ reqId, type: 'tiles-done', refused: [], cancelled: true })
    await settle()
    expect(got.map(key)).toEqual(['0,0'])
  })

  it('resume mid-batch', async () => {
    const { s, sock, socks, manifest } = await opened()
    const six = [0, 1, 2, 3, 4, 5].map((col) => ({ col, row: 0 }))
    const { p, reqId, got } = await batch(s, sock, six)
    for (const col of [0, 1, 2]) sock.frame(5, reqId, tile(col, 0))
    await settle()
    sock.drop(1006)
    // A late tile from the old socket, before and after the resume.
    sock.frame(5, reqId, tile(3, 0))
    await settle(300)
    const next = socks[1]!
    next.open(); await settle()
    const hello = next.of('hello')[0]!
    expect(hello.features).toEqual(FEATURES)
    expect(hello.resume).toEqual({ sid: manifest.session.sid, resumeProof: tokenFor('j1') })
    next.frame(1, hello.reqId, encodeManifest({ ...manifest, session: { ...manifest.session, token: tokenFor('j2') } }))
    await settle()
    sock.frame(5, reqId, tile(4, 0))
    const again = next.of('tiles')
    expect(again).toHaveLength(1)
    expect(again[0]).toMatchObject({ token: tokenFor('j2'), tiles: [{ col: 3, row: 0 }, { col: 4, row: 0 }, { col: 5, row: 0 }] })
    const r2 = again[0]!.reqId as number
    expect(r2).not.toBe(reqId)
    next.frame(5, reqId, tile(5, 0)) // the old reqId on the new socket
    next.frame(5, r2, tile(0, 0)) // already delivered
    for (const col of [3, 4, 5]) next.frame(5, r2, tile(col, 0))
    next.text({ reqId: r2, type: 'tiles-done', refused: [] })
    expect(await p).toEqual({ refused: [] })
    expect(got.map(key)).toEqual(['0,0', '1,0', '2,0', '3,0', '4,0', '5,0'])
  })

  it('tiles-done: refusals are reported; a coordinate neither delivered nor refused is refused; rate-limited has a bounded retryAfterMs', async () => {
    const { s, sock } = await opened()
    const a = await batch(s, sock, [{ col: 0, row: 0 }, { col: 1, row: 0 }, { col: 2, row: 0 }])
    sock.frame(5, a.reqId, tile(0, 0))
    // (9, 9) was never asked for; (1, 0) twice: both ignored. (2, 0) is not mentioned.
    sock.text({ reqId: a.reqId, type: 'tiles-done', refused: [{ col: 1, row: 0, code: 'rate-limited' }, { col: 1, row: 0, code: 'rate-limited' }, { col: 9, row: 9, code: 'rate-limited' }], retryAfterMs: 400 })
    expect(await a.p).toEqual({ refused: [{ col: 1, row: 0, code: 'rate-limited' }, { col: 2, row: 0, code: 'policy-refused' }], retryAfterMs: 400 })
    // No retryAfterMs from the server: the fallback.
    const b = await batch(s, sock, [{ col: 0, row: 1 }])
    sock.text({ reqId: b.reqId, type: 'tiles-done', refused: [{ col: 0, row: 1, code: 'rate-limited' }] })
    expect(await b.p).toEqual({ refused: [{ col: 0, row: 1, code: 'rate-limited' }], retryAfterMs: TILE_RETRY_FALLBACK_MS })
    // An absurd one is clamped.
    const c = await batch(s, sock, [{ col: 0, row: 2 }])
    sock.text({ reqId: c.reqId, type: 'tiles-done', refused: [{ col: 0, row: 2, code: 'rate-limited' }], retryAfterMs: 1e12 })
    expect(await c.p).toEqual({ refused: [{ col: 0, row: 2, code: 'rate-limited' }], retryAfterMs: TILE_RETRY_MAX_MS })
  })

  it('a whole-request refusal refuses every pending tile under its code', async () => {
    const { s, sock } = await opened()
    const a = await batch(s, sock, [{ col: 0, row: 0 }, { col: 1, row: 0 }])
    sock.frame(5, a.reqId, tile(0, 0))
    sock.text({ reqId: a.reqId, type: 'refused', code: 'stale-view' })
    expect(await a.p).toEqual({ refused: [{ col: 1, row: 0, code: 'stale-view' }] })
    // Admission's rate-limited refusal carries no retryAfterMs: the fallback applies.
    const b = await batch(s, sock, [{ col: 0, row: 1 }])
    sock.text({ reqId: b.reqId, type: 'refused', code: 'rate-limited' })
    expect(await b.p).toEqual({ refused: [{ col: 0, row: 1, code: 'rate-limited' }], retryAfterMs: TILE_RETRY_FALLBACK_MS })
    // A stale-token refusal is re-sent once under a new reqId.
    const c = await batch(s, sock, [{ col: 0, row: 2 }])
    sock.text({ reqId: c.reqId, type: 'refused', code: 'session-rotated' })
    await settle()
    const resent = sock.of('tiles').at(-1)!
    expect(resent.reqId).not.toBe(c.reqId)
    expect(resent.tiles).toEqual([{ col: 0, row: 2 }])
    sock.text({ reqId: resent.reqId, type: 'refused', code: 'session-rotated' })
    expect(await c.p).toEqual({ refused: [{ col: 0, row: 2, code: 'session-rotated' }] })
  })

  it('a malformed request is refused without sending', async () => {
    const { s, sock } = await opened()
    const bad: Array<Record<string, unknown>> = [
      { sheetId: '2' }, { viewGen: -1 }, { zoomStep: 11 }, { dpr: 3 }, { palette: 'blue' },
      { tiles: [] }, { tiles: Array.from({ length: 65 }, (_, i) => ({ col: i, row: 0 })) }, { tiles: [{ col: 0, row: 0 }, { col: 0, row: 0 }] },
    ]
    for (const o of bad) {
      const e = await s.tiles({ ...req([{ col: 0, row: 0 }]), ...o } as never, () => {}).then(() => undefined, (x) => x)
      expect([o, (e as StreamSessionError).code]).toEqual([o, 'policy-refused'])
    }
    expect(sock.of('tiles')).toEqual([])
  })

  it('a units session has no tile channel', async () => {
    const { s, sock } = await opened({}, unitsManifest())
    const e = await s.tiles(req([{ col: 0, row: 0 }]), () => {}).then(() => undefined, (x) => x)
    expect((e as StreamSessionError).code).toBe('unsupported-op')
    const l = await s.sheetLayout('0').then(() => undefined, (x) => x)
    expect((l as StreamSessionError).code).toBe('unsupported-op')
    expect([sock.of('tiles'), sock.of('sheet-layout')]).toEqual([[], []])
  })

  it('a session end rejects pending batches and layouts', async () => {
    const { s, sock } = await opened()
    const a = await batch(s, sock, [{ col: 0, row: 0 }])
    const l = s.sheetLayout('0'); void l.catch(() => {})
    await settle()
    s.close()
    expect((await a.p.then(() => undefined, (x) => x) as Error).name).toBe('AbortError')
    expect((await l.then(() => undefined, (x) => x) as Error).name).toBe('AbortError')
  })
})

const layout: StreamSheetLayout = {
  sheetId: '0', viewGen: 0, defaults: { rowPx: 20, colPx: 64 },
  rows: Array.from({ length: 3000 }, (_, i) => ({ start: i * 2, end: i * 2, sizeCss: 30 })), cols: [{ start: 1, end: 1, sizeCss: 90 }],
  collapsedRows: [{ start: 1, end: 1 }], collapsedCols: [],
  frozen: { rows: 1, cols: 0 }, merges: [{ r0: 0, r1: 0, c0: 0, c1: 2 }], used: { r0: 0, r1: 5999, c0: 0, c1: 3 },
  filters: [], sort: null, filterSummary: [], authorHidden: false, counts: { visibleRows: 1_048_575, totalRows: 1_048_576 },
}

describe('sheetLayout (S4 §6.4)', () => {
  it('collects every part and decodes the layout', async () => {
    const { s, sock } = await opened()
    const p = s.sheetLayout('0')
    await settle()
    const r = sock.of('sheet-layout')[0]!
    expect(r).toEqual({ reqId: r.reqId, type: 'sheet-layout', token: tokenFor('j1'), sheetId: '0' })
    const parts = encodeSheetLayout(layout)
    expect(parts.length).toBeGreaterThan(1)
    for (const part of parts) sock.text({ reqId: r.reqId, ...part })
    expect(await p).toEqual(layout)
  })

  it('a refusal rejects with its code; a stale token is re-sent once', async () => {
    const { s, sock } = await opened()
    for (const code of ['layout-too-large', 'rate-limited', 'policy-refused'] as const) {
      const p = s.sheetLayout('0')
      await settle()
      sock.text({ reqId: sock.of('sheet-layout').at(-1)!.reqId, type: 'refused', code })
      expect((await p.then(() => undefined, (x) => x) as StreamSessionError).code).toBe(code)
    }
    const p = s.sheetLayout('1')
    await settle()
    const first = sock.of('sheet-layout').at(-1)!
    sock.text({ reqId: first.reqId, type: 'refused', code: 'session-rotated' })
    await settle()
    const again = sock.of('sheet-layout').at(-1)!
    expect([again.reqId !== first.reqId, again.sheetId]).toEqual([true, '1'])
    for (const part of encodeSheetLayout({ ...layout, sheetId: '1' })) sock.text({ reqId: again.reqId, ...part })
    expect((await p).sheetId).toBe('1')
  })

  it('a layout that breaks the contract, or answers another sheet, fails policy-refused and the session survives', async () => {
    const states: string[] = []
    const { s, sock } = await opened({ onState: (st) => { states.push(st) } })
    const p1 = s.sheetLayout('0')
    await settle()
    const r1 = sock.of('sheet-layout').at(-1)!
    for (const part of encodeSheetLayout({ ...layout, sheetId: '1' })) sock.text({ reqId: r1.reqId, ...part })
    expect((await p1.then(() => undefined, (x) => x) as StreamSessionError).code).toBe('policy-refused')
    const p2 = s.sheetLayout('0')
    await settle()
    const r2 = sock.of('sheet-layout').at(-1)!
    sock.text({ reqId: r2.reqId, type: 'sheet-layout', seq: 0, last: true, head: { sheetId: '0' } })
    expect((await p2.then(() => undefined, (x) => x) as StreamSessionError).code).toBe('policy-refused')
    // Over 64 KiB.
    const p3 = s.sheetLayout('0')
    await settle()
    const r3 = sock.of('sheet-layout').at(-1)!
    // A well-formed one-part answer, but over 64 KiB.
    const fat = { ...(encodeSheetLayout({ ...layout, rows: [] })[0] as Record<string, unknown>), rows: Array.from({ length: 2000 }, (_, i) => ({ start: i, end: i, sizeCss: 30.123456789 })) }
    expect(JSON.stringify(fat).length).toBeGreaterThan(65_536)
    sock.text({ reqId: r3.reqId, ...fat })
    expect((await p3.then(() => undefined, (x) => x) as StreamSessionError).code).toBe('policy-refused')
    expect(states).toEqual(['open'])
  })

  it('a resume re-requests a layout in flight', async () => {
    const { s, sock, socks, manifest } = await opened()
    const p = s.sheetLayout('0')
    await settle()
    const r1 = sock.of('sheet-layout')[0]!
    const parts = encodeSheetLayout(layout)
    sock.text({ reqId: r1.reqId, ...parts[0]! })
    await settle()
    sock.drop(1006)
    await settle(300)
    const next = socks[1]!
    next.open(); await settle()
    next.frame(1, next.of('hello')[0]!.reqId, encodeManifest({ ...manifest, session: { ...manifest.session, token: tokenFor('j2') } }))
    await settle()
    const r2 = next.of('sheet-layout')[0]!
    expect([r2.token, r2.sheetId, r2.reqId !== r1.reqId]).toEqual([tokenFor('j2'), '0', true])
    for (const part of parts) next.text({ reqId: r2.reqId, ...part })
    expect(await p).toEqual(layout)
  })
})

describe('sheet request deadlines (fix round 1)', () => {
  it('a stalled batch expires: its outstanding tiles are refused policy-refused', async () => {
    let t = 0
    const { s, sock } = await opened({ now: () => t })
    const { p, reqId, got } = await batch(s, sock, [{ col: 0, row: 0 }, { col: 1, row: 0 }])
    sock.frame(5, reqId, tile(0, 0))
    await settle()
    t = SHEET_REPLY_MS - 1
    await settle(1100)
    expect(s.stats()).toMatchObject({ sheetPending: 1 })
    t = SHEET_REPLY_MS
    await settle(1100)
    expect(await p).toEqual({ refused: [{ col: 1, row: 0, code: 'policy-refused' }] })
    expect(got.map(key)).toEqual(['0,0'])
    expect(s.stats()).toMatchObject({ sheetPending: 0 })
    // A late answer for it is dropped.
    sock.frame(5, reqId, tile(1, 0))
    sock.text({ reqId, type: 'tiles-done', refused: [] })
    await settle()
    expect(got.map(key)).toEqual(['0,0'])
  }, 15_000)

  it('a stalled layout rejects policy-refused', async () => {
    let t = 0
    const { s, sock } = await opened({ now: () => t })
    const p = s.sheetLayout('0'); void p.catch(() => {})
    await settle()
    sock.text({ reqId: sock.of('sheet-layout')[0]!.reqId, ...encodeSheetLayout(layout)[0]! })
    await settle()
    t = SHEET_REPLY_MS
    await settle(1100)
    expect((await p.then(() => undefined, (x) => x) as StreamSessionError).code).toBe('policy-refused')
    expect(s.stats()).toMatchObject({ sheetPending: 0 })
  }, 15_000)

  it('a slow but valid reply inside the deadline succeeds (each arrival restarts it)', async () => {
    let t = 0
    const { s, sock } = await opened({ now: () => t })
    const { p, reqId, got } = await batch(s, sock, [{ col: 0, row: 0 }, { col: 1, row: 0 }])
    t = SHEET_REPLY_MS - 1
    await settle(1100)
    sock.frame(5, reqId, tile(0, 0))
    await settle()
    // Past the first deadline, but inside the one the tile restarted.
    t = 2 * SHEET_REPLY_MS - 2
    await settle(1100)
    sock.frame(5, reqId, tile(1, 0))
    sock.text({ reqId, type: 'tiles-done', refused: [] })
    expect(await p).toEqual({ refused: [] })
    expect(got.map(key)).toEqual(['0,0', '1,0'])
    // A layout answered after 4.999 s.
    const l = s.sheetLayout('0')
    await settle()
    t += SHEET_REPLY_MS - 1
    await settle(1100)
    for (const part of encodeSheetLayout(layout)) sock.text({ reqId: sock.of('sheet-layout')[0]!.reqId, ...part })
    expect(await l).toEqual(layout)
  }, 20_000)
})
