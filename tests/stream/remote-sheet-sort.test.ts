/**
 * Stage 4c Task 9: the client session side of the sort ops — `RemoteSheets.sortOps` (sort, filter, values, detect,
 * order), its presence, `sheet-order` part assembly (P7), the resume policy (P8), cancel, the local refusal behind
 * `view-refused` (P13), and hostile answers that fail only their own request.
 */
import { describe, expect, it } from 'vitest'
import { gzipSync } from 'node:zlib'
import { encodeMessage } from '../../src/stream/frame'
import { encodeManifest, type Manifest } from '../../src/stream/manifest'
import { StreamSession, type SessionOptions } from '../../src/stream/remote/session'
import { createRemoteDoc, type RemoteOfficeDoc } from '../../src/stream/remote/remote-doc'
import { StreamSessionError } from '../../src/stream/remote/errors'
import { VIEW_REFUSAL_MESSAGES } from '../../src/stream/remote/text-wire'
import { encodeSheetOrder, SHEET_MAX_ROWS, type SheetOrder, type WireRange } from '../../src/stream/sheet-wire'
import { PIVOT_SHEET_CHANGE_MESSAGE } from '../../src/xlsx/view/operationGuard'

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
  sheetView: 1,
})
const sortManifest = (): Manifest => ({ ...baseManifest(), sheetSort: 1 })

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

async function opened(m: Manifest = sortManifest(), o: Partial<SessionOptions> = {}) {
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
const ops = (doc: RemoteOfficeDoc) => doc.sheets!.sortOps!
const errorOf = (p: Promise<unknown>) => p.then(() => { throw new Error('resolved') }, (e: unknown) => e as StreamSessionError)
const sel = { active: [1, 1] as [number, number], anchor: [1, 1] as [number, number], ranges: [] }
const RANGE: WireRange = { r0: 0, r1: 39, c0: 0, c1: 5, header: 1, totals: 0 }
async function resumeOn(socks: FakeSocket[], manifest: Manifest) {
  socks[0]!.drop(1006)
  await settle(300)
  const next = socks[1]!
  next.open(); await settle()
  next.frame(1, next.of('hello')[0]!.reqId, encodeManifest({ ...manifest, session: { ...manifest.session, token: tokenFor('j2') } }))
  await settle()
  return next
}

/** A reversed band over `[r0, r1]` and a mask over it with `hidden` display slots cleared. */
function orderOf(viewGen: number, r0: number, r1: number, hidden: number[] = []): SheetOrder {
  const n = r1 - r0 + 1
  const sourceAtOffset = new Uint32Array(n)
  for (let i = 0; i < n; i++) sourceAtOffset[i] = r1 - i
  const words = new Uint32Array(Math.ceil(n / 32))
  for (let i = 0; i < n; i++) if (!hidden.includes(r0 + i)) words[i >>> 5] = (words[i >>> 5]! | (1 << (i & 31))) >>> 0
  return {
    head: { sheetId: '0', viewGen, counts: { visible: n - 1 - hidden.length, total: n - 1 }, sort: { range: { ...RANGE, r0, r1 }, keys: [{ col: 1, order: 'desc' }] }, filterSummary: [] },
    bands: [{ r0, r1, sourceAtOffset }],
    mask: { r0, r1, words },
  }
}
const send = (sock: FakeSocket, reqId: number, parts: object[]) => { for (const p of parts) sock.text({ reqId, ...p }) }
const VALUES = (reqId: number) => ({ reqId, type: 'filter-values-result', list: 'L1', items: [[0, 'apple', 1], [1, 'pear', 0]], blanks: { present: true, selected: true }, total: 2, truncated: false, classification: { kind: 'text', hasNumeric: false } })
const DETECTED = (reqId: number) => ({ reqId, type: 'range-detect-result', range: RANGE, source: 'region', header: 'heuristic', overrideIgnored: false })

describe('RemoteSheets.sortOps (S4 §6.6, P1)', () => {
  it('sortOps present only with sheetSort', async () => {
    const a = await opened(baseManifest())
    expect(a.doc.sheets!.view).toBeDefined()
    expect(a.doc.sheets!.sortOps).toBeUndefined()
    // A 4c op on a session without sheetSort rejects unsupported-op without sending.
    const e = await errorOf(a.s.text({ type: 'sheet-order', sheetId: '0' }, () => {}))
    expect(e.code).toBe('unsupported-op')
    expect(a.sock.of('sheet-order')).toHaveLength(0)
    a.doc.close()
    const b = await opened()
    expect(b.doc.sheets!.sortOps).toBeDefined()
    expect(b.doc.sheets!.sortOps!.typedFilters).toBe(false)
    b.doc.close()
    const c = await opened({ ...sortManifest(), typedFilters: true })
    expect(c.doc.sheets!.sortOps!.typedFilters).toBe(true)
    c.doc.close()
  })

  it('the ops travel as their wire requests and settle on their answers', async () => {
    const { doc, sock } = await opened()
    const sort = ops(doc).sort('0', 0, { range: RANGE, keys: [{ col: 1, order: 'desc' }], caseSensitive: true })
    await settle()
    const s = sock.of('sheet-sort')[0]!
    expect(s).toEqual({ reqId: s.reqId, type: 'sheet-sort', token: tokenFor('j1'), sheetId: '0', viewGen: 0, range: RANGE, keys: [{ col: 1, order: 'desc' }], caseSensitive: true })
    send(sock, s.reqId, encodeSheetOrder(orderOf(1, 0, 39)))
    const order = await sort
    expect([order.head.viewGen, order.bands[0]!.sourceAtOffset[0]]).toEqual([1, 39])

    const filter = ops(doc).filter('0', 1, { range: RANGE, columns: [[1, { kind: 'values', list: 'L1', mode: 'exclude', ids: [[1, 2]], blanks: false }], [2, { kind: 'committed' }], [3, { kind: 'text', op: 'contains', first: 'a', wildcards: false }]] })
    await settle()
    const f = sock.of('sheet-filter')[0]!
    expect(f).toEqual({ reqId: f.reqId, type: 'sheet-filter', token: tokenFor('j1'), sheetId: '0', viewGen: 1, range: RANGE, columns: [[1, { kind: 'values', list: 'L1', mode: 'exclude', ids: [[1, 2]], blanks: false }], [2, { kind: 'committed' }], [3, { kind: 'text', op: 'contains', first: 'a', wildcards: false }]] })
    send(sock, f.reqId, encodeSheetOrder(orderOf(2, 0, 39, [5])))
    expect((await filter).mask!.words[0]! & (1 << 5)).toBe(0)

    const values = ops(doc).values({ sheetId: '0', viewGen: 2, range: RANGE, col: 1, search: { text: 'a', wildcards: true } })
    await settle()
    const v = sock.of('filter-values')[0]!
    expect(v).toEqual({ reqId: v.reqId, type: 'filter-values', token: tokenFor('j1'), sheetId: '0', viewGen: 2, range: RANGE, col: 1, search: { text: 'a', wildcards: true } })
    sock.text({ ...VALUES(v.reqId), cursor: 'L1~1:2' })
    expect(await values).toEqual({ list: 'L1', items: [{ id: 0, text: 'apple', selected: true, error: false }, { id: 1, text: 'pear', selected: false, error: false }], blanks: { present: true, selected: true }, total: 2, truncated: false, cursor: 'L1~1:2', classification: { kind: 'text', hasNumeric: false } })
    const page = ops(doc).values({ sheetId: '0', viewGen: 2, cursor: 'L1~1:2' })
    await settle()
    const v2 = sock.of('filter-values')[1]!
    expect(v2).toEqual({ reqId: v2.reqId, type: 'filter-values', token: tokenFor('j1'), sheetId: '0', viewGen: 2, cursor: 'L1~1:2' })
    sock.text(VALUES(v2.reqId))
    expect((await page).items).toHaveLength(2)

    const detect = ops(doc).detect('0', 2, 'filter', sel, { header: 0 })
    await settle()
    const d = sock.of('range-detect')[0]!
    expect(d).toEqual({ reqId: d.reqId, type: 'range-detect', token: tokenFor('j1'), sheetId: '0', viewGen: 2, kind: 'filter', selection: sel, override: { header: 0 } })
    sock.text(DETECTED(d.reqId))
    expect(await detect).toEqual({ range: RANGE, source: 'region', header: 'heuristic', overrideIgnored: false })

    const current = ops(doc).order('0')
    await settle()
    const o = sock.of('sheet-order')[0]!
    expect(o).toEqual({ reqId: o.reqId, type: 'sheet-order', token: tokenFor('j1'), sheetId: '0' })
    send(sock, o.reqId, encodeSheetOrder(orderOf(2, 0, 39, [5])))
    expect((await current).head.viewGen).toBe(2)
    doc.close()
  })

  it('order parts assemble', async () => {
    const { s, doc, sock } = await opened()
    const full = orderOf(1, 0, SHEET_MAX_ROWS - 1)
    const parts = encodeSheetOrder(full)
    expect(parts).toHaveLength(90)
    // Out of order (two parts swapped): refused, and the session goes on.
    const bad = errorOf(ops(doc).sort('0', 0, { range: RANGE, keys: [{ col: 1, order: 'desc' }], caseSensitive: false }))
    await settle()
    const swapped = [...parts]
    ;[swapped[40], swapped[41]] = [swapped[41]!, swapped[40]!]
    send(sock, sock.of('sheet-sort')[0]!.reqId, swapped)
    expect((await bad).code).toBe('policy-refused')
    expect(s.isEnded()).toBe(false)
    // In order: assembled into the band and the mask.
    const good = ops(doc).sort('0', 0, { range: RANGE, keys: [{ col: 1, order: 'desc' }], caseSensitive: false })
    await settle()
    send(sock, sock.of('sheet-sort')[1]!.reqId, parts)
    const order = await good
    expect(order.bands[0]!.sourceAtOffset.length).toBe(SHEET_MAX_ROWS)
    expect([order.bands[0]!.sourceAtOffset[0], order.bands[0]!.sourceAtOffset[SHEET_MAX_ROWS - 1]]).toEqual([SHEET_MAX_ROWS - 1, 0])
    expect(order.mask!.words).toEqual(full.mask!.words)
    // Past 160 parts: refused on arrival.
    const many = errorOf(ops(doc).order('0'))
    await settle()
    const reqId = sock.of('sheet-order')[0]!.reqId
    sock.text({ reqId, ...parts[0]!, last: false })
    for (let i = 1; i <= 160; i++) sock.text({ reqId, type: 'sheet-order', seq: i, last: false, mask: { r0: 0, r1: 31, from: 0, data: 'AAAAAA==' } })
    expect((await many).code).toBe('policy-refused')
    expect(s.stats().pending).toBe(0)
    doc.close()
  })

  it('resume: sort and filter reject session-rotated, values and detect are re-issued', async () => {
    const { doc, sock, socks, manifest } = await opened()
    const sort = errorOf(ops(doc).sort('0', 0, { range: RANGE, keys: [{ col: 1, order: 'asc' }], caseSensitive: false }))
    const filter = errorOf(ops(doc).filter('0', 0, { range: RANGE, columns: [] }))
    const values = ops(doc).values({ sheetId: '0', viewGen: 0, range: RANGE, col: 1 })
    const detect = ops(doc).detect('0', 0, 'sort', sel)
    const order = ops(doc).order('0')
    await settle()
    const v1 = sock.of('filter-values')[0]!, o1 = sock.of('sheet-order')[0]!
    // A part before the drop is void after the resume.
    send(sock, o1.reqId, encodeSheetOrder(orderOf(1, 0, 39)).slice(0, 1))
    await settle()
    const next = await resumeOn(socks, manifest)
    expect((await sort).code).toBe('session-rotated')
    expect((await filter).code).toBe('session-rotated')
    expect(next.of('sheet-sort').length + next.of('sheet-filter').length).toBe(0)
    const v2 = next.of('filter-values')[0]!, d2 = next.of('range-detect')[0]!, o2 = next.of('sheet-order')[0]!
    expect([v2.token, v2.reqId !== v1.reqId, v2.col]).toEqual([tokenFor('j2'), true, 1])
    expect(d2).toMatchObject({ token: tokenFor('j2'), kind: 'sort' })
    expect([o2.token, o2.reqId !== o1.reqId]).toEqual([tokenFor('j2'), true])
    next.text(VALUES(v2.reqId))
    next.text(DETECTED(d2.reqId))
    send(next, o2.reqId, encodeSheetOrder(orderOf(1, 0, 39)))
    expect((await values).list).toBe('L1')
    expect((await detect).source).toBe('region')
    expect((await order).head.viewGen).toBe(1)
    doc.close()
  })

  it('abort sends sheet-cancel', async () => {
    const { doc, sock } = await opened()
    const abort = new AbortController()
    const sort = errorOf(ops(doc).sort('0', 0, { range: RANGE, keys: [{ col: 1, order: 'asc' }], caseSensitive: false }, abort.signal))
    await settle()
    const s = sock.of('sheet-sort')[0]!
    abort.abort()
    expect((await sort as unknown as Error).name).toBe('AbortError')
    expect(sock.of('sheet-cancel')).toEqual([{ reqId: expect.any(Number), type: 'sheet-cancel', target: s.reqId }])
    const fAbort = new AbortController()
    const filter = errorOf(ops(doc).filter('0', 0, { range: RANGE, columns: [] }, fAbort.signal))
    await settle()
    fAbort.abort()
    await filter
    expect(sock.of('sheet-cancel').map((m) => m.target)).toEqual([s.reqId, sock.of('sheet-filter')[0]!.reqId])
    // A late answer to the cancelled sort is dropped.
    send(sock, s.reqId, encodeSheetOrder(orderOf(1, 0, 39)))
    // A value list's abort sends nothing: only a sort or filter has a job to cancel.
    const vAbort = new AbortController()
    const values = errorOf(ops(doc).values({ sheetId: '0', viewGen: 0, range: RANGE, col: 1 }, vAbort.signal))
    await settle()
    vAbort.abort()
    await values
    expect(sock.of('sheet-cancel')).toHaveLength(2)
    doc.close()
  })

  it('view-refused carries the local refusal', async () => {
    const { doc, sock } = await opened()
    const at = (p: Promise<unknown>, type: string, m: object) => settle().then(() => { sock.text({ reqId: sock.of(type).at(-1)!.reqId, type: 'refused', code: 'view-refused', ...m }); return errorOf(p) })
    const sortOf = () => ops(doc).sort('0', 0, { range: RANGE, keys: [{ col: 1, order: 'asc' }], caseSensitive: false })
    const merge = await at(sortOf(), 'sheet-sort', { reason: 'merge', range: { r0: 2, r1: 3, c0: 1, c1: 1 } })
    expect(merge).toBeInstanceOf(StreamSessionError)
    expect([merge.code, merge.refusal]).toEqual(['view-refused', { reason: 'merge', range: { r0: 2, r1: 3, c0: 1, c1: 1 } }])
    const pivot = await at(sortOf(), 'sheet-sort', { reason: 'pivot', range: { r0: 0, r1: 9, c0: 0, c1: 3 } })
    expect(pivot.refusal).toEqual({ reason: 'pivot', range: { r0: 0, r1: 9, c0: 0, c1: 3 }, message: PIVOT_SHEET_CHANGE_MESSAGE })
    const formula = await at(sortOf(), 'sheet-sort', { reason: 'formula-value-unavailable', count: 3, at: [4, 1] })
    expect(formula.refusal).toEqual({ reason: 'formula-value-unavailable', count: 3, firstAddress: { row: 4, col: 1 } })
    const drawing = await at(ops(doc).filter('0', 0, { range: RANGE, columns: [] }), 'sheet-filter', { reason: 'drawing-impact' })
    expect(drawing.refusal).toEqual({ reason: 'drawing-impact', drawingId: '' })
    const ambiguous = await at(ops(doc).detect('0', 0, 'sort', sel), 'range-detect', { reason: 'ambiguous-selection' })
    expect(ambiguous.refusal).toEqual({ reason: 'ambiguous-selection', message: VIEW_REFUSAL_MESSAGES['ambiguous-selection'] })
    // Anything off the P13 contract is policy-refused.
    for (const m of [{ reason: 'merge' }, { reason: 'incomplete' }, { reason: 'made-up' }, { reason: 'merge', range: { r0: 3, r1: 2, c0: 0, c1: 0 } }, { reason: 'formula-value-unavailable', count: 1 }]) {
      const e = await at(sortOf(), 'sheet-sort', m)
      expect([e.code, e.refusal]).toEqual(['policy-refused', undefined])
    }
    // The 4c codes reach the caller; any other is policy-refused.
    for (const code of ['stale-filter-values', 'filter-condition-refused', 'cancelled', 'stale-view', 'hidden-target', 'coverage-unknown', 'view-state-limit', 'rate-limited']) {
      const p = errorOf(ops(doc).filter('0', 0, { range: RANGE, columns: [] }))
      await settle()
      sock.text({ reqId: sock.of('sheet-filter').at(-1)!.reqId, type: 'refused', code })
      expect((await p).code).toBe(code)
    }
    doc.close()
  })

  it('view-refused too-large (Task 14a) is the local resource-limit refusal with the too-large text', async () => {
    const { doc, sock } = await opened()
    const at = (p: Promise<unknown>, type: string) => settle().then(() => { sock.text({ reqId: sock.of(type).at(-1)!.reqId, type: 'refused', code: 'view-refused', reason: 'too-large' }); return errorOf(p) })
    const expected = { reason: 'resource-limit', message: VIEW_REFUSAL_MESSAGES['too-large'] }
    const sort = await at(ops(doc).sort('0', 0, { range: RANGE, keys: [{ col: 1, order: 'asc' }], caseSensitive: false }), 'sheet-sort')
    expect([sort.code, sort.refusal]).toEqual(['view-refused', expected])
    const values = await at(ops(doc).values({ sheetId: '0', viewGen: 0, range: RANGE, col: 1 }), 'filter-values')
    expect([values.code, values.refusal]).toEqual(['view-refused', expected])
    // A value list's error is shown by its message (FilterDropdown): the local text, not the code.
    expect(values.message).toBe(VIEW_REFUSAL_MESSAGES['too-large'])
    const detect = await at(ops(doc).detect('0', 0, 'sort', sel), 'range-detect')
    expect([detect.code, detect.refusal]).toEqual(['view-refused', expected])
    expect(VIEW_REFUSAL_MESSAGES['too-large']).toBe('This sheet is too large to sort or filter in this viewing mode.')
    doc.close()
  })

  it('hostile answers', async () => {
    const { s, doc, sock } = await opened()
    const notPermutation = encodeSheetOrder(orderOf(1, 0, 39)).map((p, i) => i === 1 ? { ...p, band: { r0: 0, r1: 1, from: 0, data: 'AAAAAAAAAAA=' } } : p)
    const cases: Array<[() => Promise<unknown>, string, (reqId: number) => object[]]> = [
      [() => ops(doc).order('0'), 'sheet-order', () => notPermutation],
      [() => ops(doc).values({ sheetId: '0', viewGen: 0, range: RANGE, col: 1 }), 'filter-values', (reqId) => [{ ...VALUES(reqId), items: [[2, 'x', 0]] }]],
      [() => ops(doc).detect('0', 0, 'sort', sel), 'range-detect', (reqId) => [{ ...DETECTED(reqId), range: { ...RANGE, r1: SHEET_MAX_ROWS } }]],
      // Another sheet's order, and a sort answered below its own viewGen.
      [() => ops(doc).order('0'), 'sheet-order', () => encodeSheetOrder({ ...orderOf(1, 0, 39), head: { ...orderOf(1, 0, 39).head, sheetId: '1' } })],
      [() => ops(doc).sort('0', 3, { range: RANGE, keys: [{ col: 1, order: 'asc' }], caseSensitive: false }), 'sheet-sort', () => encodeSheetOrder(orderOf(2, 0, 39))],
      // An answer of another op's shape.
      [() => ops(doc).detect('0', 0, 'sort', sel), 'range-detect', (reqId) => [VALUES(reqId)]],
      [() => ops(doc).values({ sheetId: '0', viewGen: 0, range: RANGE, col: 1 }), 'filter-values', () => encodeSheetOrder(orderOf(1, 0, 39))],
    ]
    for (const [start, type, answer] of cases) {
      const p = errorOf(start())
      await settle()
      const reqId = sock.of(type).at(-1)!.reqId
      send(sock, reqId, answer(reqId))
      expect((await p).code).toBe('policy-refused')
    }
    expect(s.stats().pending).toBe(0)
    expect(s.isEnded()).toBe(false)
    // The session stays open: the next request is answered.
    const next = ops(doc).detect('0', 0, 'sort', sel)
    await settle()
    sock.text(DETECTED(sock.of('range-detect').at(-1)!.reqId))
    expect((await next).header).toBe('heuristic')
    doc.close()
  })

  it('a sort clear travels as keys: []; null names no range and is not sent (fix round 1)', async () => {
    const { doc, sock } = await opened()
    const clear = ops(doc).sort('0', 2, { range: RANGE, keys: [], caseSensitive: false })
    await settle()
    const c = sock.of('sheet-sort')[0]!
    expect(c).toEqual({ reqId: c.reqId, type: 'sheet-sort', token: tokenFor('j1'), sheetId: '0', viewGen: 2, range: RANGE, keys: [], caseSensitive: false })
    send(sock, c.reqId, encodeSheetOrder({ head: { sheetId: '0', viewGen: 3, counts: { visible: 39, total: 39 }, sort: null, filterSummary: [] }, bands: [], mask: null }))
    expect((await clear).head.sort).toBeNull()
    expect((await errorOf(ops(doc).sort('0', 3, null))).code).toBe('unsupported-op')
    expect(sock.of('sheet-sort')).toHaveLength(1)
    doc.close()
  })

  it('pins the local strings the refusals are rebuilt with', () => {
    expect(VIEW_REFUSAL_MESSAGES.pivot).toBe(PIVOT_SHEET_CHANGE_MESSAGE)
  })
})
