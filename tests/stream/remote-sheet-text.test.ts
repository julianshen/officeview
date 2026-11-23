/**
 * Stage 4b Task 9: the client session side of the `sheet-view@1` ops — `RemoteSheets.view` (find, copy, a11y, fit,
 * resize), its presence, the resume policy (P13), refusals with `limit`/`requested`, and hostile answers that fail only
 * their own request.
 */
import { describe, expect, it } from 'vitest'
import { gzipSync } from 'node:zlib'
import { encodeMessage } from '../../src/stream/frame'
import { encodeManifest, type Manifest } from '../../src/stream/manifest'
import { StreamSession, type SessionOptions } from '../../src/stream/remote/session'
import { createRemoteDoc, type RemoteOfficeDoc } from '../../src/stream/remote/remote-doc'
import { StreamSessionError } from '../../src/stream/remote/errors'

const b64url = (o: object) => btoa(JSON.stringify(o)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
const tokenFor = (jti: string) => `h.${b64url({ jti })}.s`
const settle = (ms = 20) => new Promise((r) => setTimeout(r, ms))
const FEATURES = ['sheet-tiles@1', 'sheet-view@1']

const baseManifest = (): Manifest => ({
  v: 1, docRef: 'a'.repeat(24), kind: 'xlsx', name: 'book.xlsx',
  units: { count: 2, list: [{ w: 960, h: 540 }, { w: 960, h: 540 }] },
  layoutFingerprint: 'f'.repeat(64),
  session: { sid: '0123456789ab', token: tokenFor('j1'), expiresSec: 300, resume: 'session' },
  policy: { contentPolicy: 'visible-only', protection: 'final-raster', semantic: 'visible-only', contentPolicyId: 'visible-only@1', viewerMarkId: 'fixed' },
  can: { search: true, copy: true, print: true, textA11y: true, present: false, notes: false, links: false },
  caps: { copy: { maxParagraphs: null, maxSessionParagraphs: null }, a11y: 'unit', print: { maxPages: null, maxSessionPages: null } },
  recorder: { commands: 'ovsd-1', textMode: 'text', scalesBuckets: [1, 2], requiredFaces: [], projection: 'final-raster' },
  limits: { maxUnitPx: 8192, maxAssetBytes: 8 * 1024 * 1024, wsFrameBytes: 1 << 20, tilePx: 512 },
  notices: [],
  sheetMode: 'tiles',
  sheets: [{ sheetId: '0', name: 'Data' }, { sheetId: '1', name: 'Summary' }],
  tiles: { px: 512, zoomSteps: [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4], dprs: [1, 2], palettes: ['light', 'dark'], maxBatch: 64 },
})
const viewManifest = (): Manifest => ({ ...baseManifest(), sheetView: 1 })

class FakeSocket extends EventTarget {
  binaryType: BinaryType = 'blob'
  readyState = 0
  sent: Array<Record<string, any>> = []
  send(d: string) { this.sent.push(JSON.parse(d)) }
  close(code?: number, reason?: string) {
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

async function opened(m: Manifest = viewManifest(), o: Partial<SessionOptions> = {}) {
  const socks: FakeSocket[] = []
  const client = () => { const s = new FakeSocket(); socks.push(s); return s as never }
  const p = StreamSession.open({ kind: 'stream', url: 'wss://x/stream' }, { client, faces: async () => [], features: FEATURES, ...o })
  const sock = socks[0]!
  sock.open(); await settle()
  sock.frame(1, sock.of('hello')[0]!.reqId, encodeManifest(m))
  const s = await p
  const doc = createRemoteDoc(s)
  return { s, doc, sock, socks, manifest: m }
}
const view = (doc: RemoteOfficeDoc) => doc.sheets!.view!
const errorOf = (p: Promise<unknown>) => p.then(() => { throw new Error('resolved') }, (e: unknown) => e as StreamSessionError)
const sel = { active: [0, 0] as [number, number], anchor: [0, 0] as [number, number], ranges: [{ rows: [[0, 1]] as Array<[number, number]>, cols: [[0, 1]] as Array<[number, number]> }] }
const PANE = { x: 0, y: 0, width: 800, height: 600 }
async function resumeOn(socks: FakeSocket[], manifest: Manifest) {
  socks[0]!.drop(1006)
  await settle(300)
  const next = socks[1]!
  next.open(); await settle()
  next.frame(1, next.of('hello')[0]!.reqId, encodeManifest({ ...manifest, session: { ...manifest.session, token: tokenFor('j2') } }))
  await settle()
  return next
}

describe('RemoteSheets.view (S4 §6.7, P1)', () => {
  it('view present only with sheetView', async () => {
    const a = await opened(baseManifest())
    expect(a.doc.sheets).toBeDefined()
    expect(a.doc.sheets!.view).toBeUndefined()
    a.doc.close()
    const b = await opened(viewManifest())
    expect(b.doc.sheets!.view).toBeDefined()
    b.doc.close()
  })

  it('the ops travel as their wire requests and settle on their answers', async () => {
    const { doc, sock } = await opened()
    const find = view(doc).find({ query: 'x', matchCase: true })
    await settle()
    const f = sock.of('sheet-find')[0]!
    expect(f).toMatchObject({ query: 'x', matchCase: true, token: tokenFor('j1') })
    sock.text({ reqId: f.reqId, type: 'sheet-find-result', hits: [[0, 1, 2], [1, 3, 4, 1]] })
    sock.text({ reqId: f.reqId, type: 'sheet-find-done', total: 2, truncated: true })
    expect(await find).toEqual({ hits: [{ sheetIndex: 0, row: 1, col: 2 }, { sheetIndex: 1, row: 3, col: 4, hidden: 'filter' }], total: 2, truncated: true })

    const copy = view(doc).copy('0', 4, sel)
    await settle()
    const c = sock.of('sheet-copy')[0]!
    expect(c).toMatchObject({ sheetId: '0', viewGen: 4, selection: { active: [0, 0], anchor: [0, 0], ranges: [{ rows: [[0, 1]], cols: [[0, 1]] }] } })
    sock.text({ reqId: c.reqId, type: 'copy-result', seq: 0, text: 'a\t', last: false })
    sock.text({ reqId: c.reqId, type: 'copy-result', seq: 1, text: 'b', last: true })
    expect(await copy).toBe('a\tb')

    const a11y = view(doc).a11y('1', 2, [[0, 1], [5, 5]], [0, 2])
    await settle()
    const y = sock.of('sheet-a11y')[0]!
    expect(y).toMatchObject({ sheetId: '1', viewGen: 2, rows: [[0, 1], [5, 5]], cols: [0, 2] })
    sock.text({ reqId: y.reqId, type: 'sheet-a11y-result', seq: 0, cells: [[0, 0, 'A']], last: false })
    sock.text({ reqId: y.reqId, type: 'sheet-a11y-result', seq: 1, cells: [[5, 2, 'B']], last: true })
    expect(await a11y).toEqual([{ row: 0, col: 0, text: 'A' }, { row: 5, col: 2, text: 'B' }])

    const fit = view(doc).fit({ sheetId: '0', viewGen: 4, axis: 'col', index: 3, pane: PANE })
    await settle()
    const t = sock.of('sheet-fit')[0]!
    expect(t).toMatchObject({ sheetId: '0', viewGen: 4, axis: 'col', index: 3, pane: PANE })
    expect('regions' in t).toBe(false)
    sock.text({ reqId: t.reqId, type: 'sheet-fit-result', size: 91 })
    expect(await fit).toBe(91)

    const resize = view(doc).resize({ sheetId: '0', viewGen: 4, axis: 'row', index: 2, size: 30 })
    await settle()
    const r = sock.of('sheet-resize')[0]!
    expect(r).toMatchObject({ sheetId: '0', viewGen: 4, axis: 'row', index: 2, size: 30 })
    sock.text({ reqId: r.reqId, type: 'sheet-view', sheetId: '0', viewGen: 5, change: { axis: 'row', index: 2, size: 29 } })
    expect(await resize).toEqual({ sheetId: '0', viewGen: 5, change: { axis: 'row', index: 2, size: 29 } })

    const reset = view(doc).resize({ sheetId: '0', viewGen: 5, reset: true })
    await settle()
    const z = sock.of('sheet-resize')[1]!
    expect(z).toMatchObject({ sheetId: '0', viewGen: 5, reset: true })
    expect('axis' in z).toBe(false)
    sock.text({ reqId: z.reqId, type: 'sheet-view', sheetId: '0', viewGen: 6, change: { reset: true } })
    expect(await reset).toEqual({ sheetId: '0', viewGen: 6, change: { reset: true } })
    doc.close()
  })

  it('a newer find supersedes an older one', async () => {
    const { doc, sock } = await opened()
    const first = errorOf(view(doc).find({ query: 'a' }))
    await settle()
    const second = view(doc).find({ query: 'ab' })
    await settle()
    expect((await first).code).toBe('closed')
    const [f1, f2] = sock.of('sheet-find')
    // The server's late done for the superseded find is ignored.
    sock.text({ reqId: f1!.reqId, type: 'sheet-find-done', total: 0, truncated: true })
    sock.text({ reqId: f2!.reqId, type: 'sheet-find-done', total: 0, truncated: false })
    expect(await second).toEqual({ hits: [], total: 0, truncated: false })
    doc.close()
  })

  it('copy and a11y off reject text-unavailable without sending', async () => {
    const m = viewManifest()
    const { doc, sock } = await opened({ ...m, can: { ...m.can, copy: false, textA11y: false }, caps: { ...m.caps, copy: null, a11y: 'off' } })
    expect((await errorOf(view(doc).copy('0', 0, sel))).code).toBe('text-unavailable')
    expect((await errorOf(view(doc).a11y('0', 0, [[0, 0]], [0, 0]))).code).toBe('text-unavailable')
    expect(sock.of('sheet-copy').length + sock.of('sheet-a11y').length).toBe(0)
    doc.close()
  })

  it('find is re-issued on resume; copy, a11y and resize reject session-rotated and are never re-sent', async () => {
    const { doc, sock, socks, manifest } = await opened()
    const find = view(doc).find({ query: 'x' })
    const fit = view(doc).fit({ sheetId: '0', viewGen: 0, axis: 'col', index: 1, pane: PANE })
    const copy = errorOf(view(doc).copy('0', 0, sel))
    const a11y = errorOf(view(doc).a11y('0', 0, [[0, 0]], [0, 0]))
    const resize = errorOf(view(doc).resize({ sheetId: '0', viewGen: 0, axis: 'col', index: 1, size: 80 }))
    await settle()
    const f1 = sock.of('sheet-find')[0]!
    // A hit before the drop is void after the resume.
    sock.text({ reqId: f1.reqId, type: 'sheet-find-result', hits: [[0, 9, 9]] })
    await settle()
    const next = await resumeOn(socks, manifest)
    expect((await copy).code).toBe('session-rotated')
    expect((await a11y).code).toBe('session-rotated')
    expect((await resize).code).toBe('session-rotated')
    expect(next.of('sheet-copy').length + next.of('sheet-a11y').length + next.of('sheet-resize').length).toBe(0)
    const f2 = next.of('sheet-find')[0]!, t2 = next.of('sheet-fit')[0]!
    expect([f2.token, f2.query, f2.reqId !== f1.reqId]).toEqual([tokenFor('j2'), 'x', true])
    expect(t2).toMatchObject({ token: tokenFor('j2'), sheetId: '0', axis: 'col', index: 1 })
    next.text({ reqId: f2.reqId, type: 'sheet-find-result', hits: [[1, 0, 0]] })
    next.text({ reqId: f2.reqId, type: 'sheet-find-done', total: 1, truncated: false })
    next.text({ reqId: t2.reqId, type: 'sheet-fit-result', size: 70 })
    expect(await find).toEqual({ hits: [{ sheetIndex: 1, row: 0, col: 0 }], total: 1, truncated: false })
    expect(await fit).toBe(70)
    doc.close()
  })

  it('refusals carry limit and requested', async () => {
    const { doc, sock } = await opened()
    const copy = errorOf(view(doc).copy('0', 0, sel))
    await settle()
    sock.text({ reqId: sock.of('sheet-copy')[0]!.reqId, type: 'refused', code: 'copy-too-large', limit: 3, requested: 4 })
    const e = await copy
    expect(e).toBeInstanceOf(StreamSessionError)
    expect([e.code, e.limit, e.requested]).toEqual(['copy-too-large', 3, 4])
    for (const code of ['stale-view', 'hidden-target', 'coverage-unknown', 'view-state-limit', 'rate-limited', 'unsupported-op', 'copy-quota-exhausted']) {
      const p = errorOf(view(doc).resize({ sheetId: '0', viewGen: 0, axis: 'col', index: 1, size: 80 }))
      await settle()
      sock.text({ reqId: sock.of('sheet-resize').at(-1)!.reqId, type: 'refused', code })
      expect((await p).code).toBe(code)
    }
    const odd = errorOf(view(doc).fit({ sheetId: '0', viewGen: 0, axis: 'col', index: 1, pane: PANE }))
    await settle()
    sock.text({ reqId: sock.of('sheet-fit').at(-1)!.reqId, type: 'refused', code: 'made-up' })
    expect((await odd).code).toBe('policy-refused')
    doc.close()
  })

  it('hostile answers drop the request', async () => {
    const { s, doc, sock } = await opened()
    const cases: Array<[() => Promise<unknown>, string, (reqId: number) => object]> = [
      [() => view(doc).find({ query: 'x' }), 'sheet-find', (reqId) => ({ reqId, type: 'sheet-find-result', hits: [[0, -1, 0]] })],
      // A sheet the manifest does not list.
      [() => view(doc).find({ query: 'x' }), 'sheet-find', (reqId) => ({ reqId, type: 'sheet-find-result', hits: [[2, 0, 0]] })],
      [() => view(doc).find({ query: 'x' }), 'sheet-find', (reqId) => ({ reqId, type: 'sheet-find-result', hits: [] })],
      [() => view(doc).find({ query: 'x' }), 'sheet-find', (reqId) => ({ reqId, type: 'sheet-find-done', total: 3, truncated: false })],
      [() => view(doc).a11y('0', 0, [[0, 1]], [0, 1]), 'sheet-a11y', (reqId) => ({ reqId, type: 'sheet-a11y-result', seq: 0, cells: [[2, 0, 'x']], last: true })],
      [() => view(doc).a11y('0', 0, [[0, 1]], [0, 1]), 'sheet-a11y', (reqId) => ({ reqId, type: 'sheet-a11y-result', seq: 0, cells: [[0, 2, 'x']], last: true })],
      [() => view(doc).a11y('0', 0, [[0, 1]], [0, 1]), 'sheet-a11y', (reqId) => ({ reqId, type: 'sheet-a11y-result', seq: 0, cells: [[0, 0, 'x'], [0, 0, 'y']], last: true })],
      [() => view(doc).resize({ sheetId: '0', viewGen: 0, axis: 'col', index: 1, size: 80 }), 'sheet-resize', (reqId) => ({ reqId, type: 'sheet-view', sheetId: '1', viewGen: 1, change: { axis: 'col', index: 1, size: 80 } })],
      [() => view(doc).resize({ sheetId: '0', viewGen: 0, axis: 'col', index: 1, size: 80 }), 'sheet-resize', (reqId) => ({ reqId, type: 'sheet-view', sheetId: '0', viewGen: 1, change: { axis: 'col', index: 2, size: 80 } })],
      [() => view(doc).resize({ sheetId: '0', viewGen: 3, axis: 'col', index: 1, size: 80 }), 'sheet-resize', (reqId) => ({ reqId, type: 'sheet-view', sheetId: '0', viewGen: 2, change: { axis: 'col', index: 1, size: 80 } })],
      [() => view(doc).fit({ sheetId: '0', viewGen: 0, axis: 'col', index: 1, pane: PANE }), 'sheet-fit', (reqId) => ({ reqId, type: 'sheet-fit-result', size: -1 })],
    ]
    for (const [start, type, answer] of cases) {
      const p = errorOf(start())
      await settle()
      sock.text(answer(sock.of(type).at(-1)!.reqId))
      expect((await p).code).toBe('policy-refused')
    }
    expect(s.stats().pending).toBe(0)
    // The session stays open: the next request is answered.
    const fit = view(doc).fit({ sheetId: '0', viewGen: 0, axis: 'row', index: 0, pane: PANE })
    await settle()
    sock.text({ reqId: sock.of('sheet-fit').at(-1)!.reqId, type: 'sheet-fit-result', size: 20 })
    expect(await fit).toBe(20)
    doc.close()
  })
})
