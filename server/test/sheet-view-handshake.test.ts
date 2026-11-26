/** Stage 4b Task 3 — the `sheet-view@1` capability, the gating of the five 4b ops, the `'sheet'` rate class (P1, P12). */
import { afterAll, describe, expect, test } from 'bun:test'
import { buildXlsx } from '../../src/testdata/ooxml-builders'
import { hiddenFixtures } from '../../scripts/stream-s0/hidden-fixtures'
import { SessionRates, SHEET_RATE_PER_MIN } from '../src/rate'
import { sheetModeOf } from '../src/sheet/mode'
import { CLOSE } from '../src/ws'
import { Client, corpus, disposePool, start } from './harness'

afterAll(async () => { await disposePool() })

const TILES = ['sheet-tiles@1']
const VIEW = ['sheet-view@1']
const BOTH = ['sheet-tiles@1', 'sheet-view@1']
const STAGE3_KEYS = ['can', 'caps', 'docRef', 'kind', 'layoutFingerprint', 'limits', 'name', 'notices', 'policy', 'recorder', 'session', 'units', 'v']
const TILES_KEYS = [...STAGE3_KEYS, 'sheetMode', 'sheets', 'tiles'].sort()
const OPS_4B = ['sheet-find', 'sheet-copy', 'sheet-a11y', 'sheet-fit', 'sheet-resize']

const small = () => buildXlsx([
  { name: 'First', rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 0 }, { ref: 'B1', v: 42 }] }] },
  { name: 'Second', rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 1 }] }] },
], ['small sheet', 'second sheet'])

async function open(h: { url: string }, features?: unknown) {
  const c = new Client(h.url)
  const m = await c.hello(features === undefined ? {} : { features })
  await c.ready()
  return { c, m }
}
async function closedBy(url: string, extra: object) {
  const c = new Client(url)
  await c.opened()
  c.send({ reqId: 1, type: 'hello', client: 'officeview/1.0', recVersions: [1], ...extra })
  return (await c.closed).code
}
const answerOf = async (c: Client, reqId: number) => (await c.next((x) => 'text' in x && x.text.reqId === reqId) as { text: Record<string, any> }).text

describe('sheetModeOf (P1)', () => {
  test('tiles for an XLSX with either feature; never for other formats', () => {
    expect(sheetModeOf('xlsx', [])).toBe('units')
    expect(sheetModeOf('xlsx', TILES)).toBe('tiles')
    expect(sheetModeOf('xlsx', VIEW)).toBe('tiles')
    expect(sheetModeOf('xlsx', BOTH)).toBe('tiles')
    expect(sheetModeOf('xlsx', ['x@9'])).toBe('units')
    expect(sheetModeOf('docx', VIEW)).toBe('units')
    expect(sheetModeOf('pptx', BOTH)).toBe('units')
  })
})

describe('sheet-view@1 handshake', () => {
  test('sheet-view@1 adds sheetView to a tiles manifest only', async () => {
    const h = await start({ doc: await small() })
    const both = await open(h, BOTH)
    expect(both.m.sheetMode).toBe('tiles')
    expect(both.m.sheetView).toBe(1)
    expect(Object.keys(both.m).sort()).toEqual([...TILES_KEYS, 'sheetView', 'sheetSort'].sort())
    // 4a: sheet-tiles@1 alone is key for key the 4a manifest.
    const t = await open(h, TILES)
    expect(t.m.sheetMode).toBe('tiles')
    expect('sheetView' in t.m).toBe(false)
    expect(Object.keys(t.m).sort()).toEqual(TILES_KEYS)
    // P1: sheet-view@1 alone selects tiles mode.
    const v = await open(h, VIEW)
    expect(v.m.sheetMode).toBe('tiles')
    expect(v.m.sheets).toEqual([{ sheetId: '0', name: 'First' }, { sheetId: '1', name: 'Second' }])
    expect(v.m.sheetView).toBe(1)
    // Neither: Stage 3's manifest.
    const u = await open(h)
    expect(Object.keys(u.m).sort()).toEqual(STAGE3_KEYS)
    both.c.close(); t.c.close(); v.c.close(); u.c.close(); await h.stop()
  }, 60_000)

  test('a DOCX gets no new keys', async () => {
    const docx = (await hiddenFixtures()).find((f) => f.kind === 'docx')!.withoutHidden
    const h = await start({ doc: docx })
    for (const features of [VIEW, BOTH]) {
      const { c, m } = await open(h, features)
      expect(Object.keys(m).sort()).toEqual(STAGE3_KEYS)
      c.close()
    }
    await h.stop()
    const p = await start({ doc: corpus('pypptx-test.pptx') })
    const { c, m } = await open(p, VIEW)
    expect(Object.keys(m).sort()).toEqual(STAGE3_KEYS)
    c.close(); await p.stop()
  }, 60_000)

  test('A15: a workbook served in units mode never has sheetView', async () => {
    const doc = await buildXlsx([{ name: 'N'.repeat(32), rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }] }], [])
    const h = await start({ doc })
    const { c, m } = await open(h, BOTH)
    expect(Object.keys(m).sort()).toEqual(STAGE3_KEYS)
    expect(h.stream.sheetViewOf(m.session.sid)).toBe(false)
    c.send({ token: m.session.token, reqId: 90, type: 'sheet-find', query: 'x' })
    expect(await answerOf(c, 90)).toMatchObject({ type: 'refused', code: 'unsupported-op' })
    c.close(); await h.stop()
  }, 60_000)
})

describe('the 4b ops need the feature', () => {
  test('the 4b ops need the feature', async () => {
    const h = await start({ doc: await small(), policy: { copy: 'full' } })
    const tiles = await open(h, TILES)
    const units = await open(h)
    const view = await open(h, VIEW)
    expect(h.stream.sheetViewOf(tiles.m.session.sid)).toBe(false)
    expect(h.stream.sheetViewOf(units.m.session.sid)).toBe(false)
    expect(h.stream.sheetViewOf(view.m.session.sid)).toBe(true)
    let reqId = 100
    for (const op of OPS_4B) {
      for (const s of [tiles, units]) {
        s.c.send({ token: s.m.session.token, reqId: ++reqId, type: op })
        expect(await answerOf(s.c, reqId)).toMatchObject({ type: 'refused', code: 'unsupported-op' })
      }
      // With the feature the op is not refused for want of it.
      view.c.send({ token: view.m.session.token, reqId: ++reqId, type: op })
      expect((await answerOf(view.c, reqId)).code).not.toBe('unsupported-op')
    }
    tiles.c.close(); units.c.close(); view.c.close(); await h.stop()
  }, 60_000)

  test('4a tiles gating is unchanged under sheet-view@1', async () => {
    const h = await start({ doc: await small(), policy: { copy: 'full' } })
    const view = await open(h, VIEW)
    let reqId = 200
    for (const type of ['unit', 'search', 'select', 'copy', 'a11y']) {
      view.c.send({ token: view.m.session.token, reqId: ++reqId, type })
      expect(await answerOf(view.c, reqId)).toMatchObject({ type: 'refused', code: 'unsupported-op' })
    }
    view.c.close(); await h.stop()
  }, 60_000)
})

describe("off-mode text ops ride the 'text' class", () => {
  test('61 off-mode sheet-find in a minute: the 61st is rate-limited, never 4008; fit and resize are uncharged', async () => {
    // An injected clock keeps the 'message' class (20/s) out of the way: one request per 500 ms, 62 of them inside 60 s.
    let t = Date.UTC(2026, 10, 23, 12)
    const h = await start({ doc: await small(), now: () => t, tokenTtlSec: 86_400 })
    for (const features of [TILES, undefined] as const) {
      t += 120_000
      const { c, m } = await open(h, features)
      let closed: number | undefined
      void c.closed.then((x) => { closed = x.code })
      for (let i = 1; i <= 60; i++) {
        t += 500
        c.send({ token: m.session.token, reqId: 1000 + i, type: 'sheet-find', query: 'x' })
        expect(await answerOf(c, 1000 + i)).toMatchObject({ type: 'refused', code: 'unsupported-op' })
      }
      t += 500
      c.send({ token: m.session.token, reqId: 1061, type: 'sheet-find', query: 'x' })
      expect(await answerOf(c, 1061)).toMatchObject({ type: 'refused', code: 'rate-limited' })
      // The other text ops share the class; fit and resize are never charged.
      t += 500
      c.send({ token: m.session.token, reqId: 1062, type: 'sheet-copy' })
      expect(await answerOf(c, 1062)).toMatchObject({ type: 'refused', code: 'rate-limited' })
      for (const [i, type] of ['sheet-fit', 'sheet-resize'].entries()) {
        t += 500
        c.send({ token: m.session.token, reqId: 1070 + i, type })
        expect(await answerOf(c, 1070 + i)).toMatchObject({ type: 'refused', code: 'unsupported-op' })
      }
      await Bun.sleep(50)
      expect(closed).toBeUndefined()
      c.close()
    }
    await h.stop()
  }, 60_000)
})

describe('resume keeps the feature', () => {
  test('resume keeps the feature', async () => {
    const h = await start({ doc: await small() })
    const { c, m } = await open(h, VIEW)
    const sid = m.session.sid
    const resume = { sid, resumeProof: m.session.token }
    expect(await closedBy(h.url, { resume })).toBe(CLOSE['auth-failed'])
    expect(await closedBy(h.url, { resume, features: TILES })).toBe(CLOSE['auth-failed'])
    expect(await closedBy(h.url, { resume, features: BOTH })).toBe(CLOSE['auth-failed'])
    const r = new Client(h.url)
    const rm = await r.hello({ resume, features: VIEW })
    expect(rm.session.sid).toBe(sid)
    expect(rm.sheetView).toBe(1)
    r.close(); c.close(); await h.stop()
  }, 60_000)

  test('F21: a resume starts with a clean sheetBusy', async () => {
    const h = await start({ doc: await small() })
    const { c, m } = await open(h, VIEW)
    const sid = m.session.sid
    h.stream.sheetRatesOf(sid)!.sheetBusy = true
    // The socket drops: the flag is cleared with it.
    const ws = c.ws as WebSocket & { terminate?: () => void }
    if (ws.terminate) ws.terminate(); else ws.close(1001)
    await c.closed; await Bun.sleep(50)
    expect(h.stream.table.get(sid)?.status).toBe('detached')
    expect(h.stream.sheetRatesOf(sid)!.sheetBusy).toBe(false)
    // A superseding resume (the old socket still open) starts clean too.
    const a = new Client(h.url)
    const am = await a.hello({ resume: { sid, resumeProof: m.session.token }, features: VIEW })
    h.stream.sheetRatesOf(sid)!.sheetBusy = true
    const b = new Client(h.url)
    await b.hello({ resume: { sid, resumeProof: am.session.token }, features: VIEW })
    expect(h.stream.sheetRatesOf(sid)!.sheetBusy).toBe(false)
    a.close(); b.close(); await h.stop()
  }, 60_000)
})

describe("the 'sheet' rate class", () => {
  test('the sheet class', () => {
    expect(SHEET_RATE_PER_MIN).toBe(30)
    let t = 1_000_000
    const r = new SessionRates()
    for (let i = 0; i < 30; i++) expect(r.check('sheet', t + i)).toBe('ok')
    expect(r.check('sheet', t + 30)).toBe('refuse')
    // Ten refusals never close 4008, and never touch the breach window.
    for (let i = 0; i < 10; i++) expect(r.check('sheet', t + 31 + i)).toBe('refuse')
    for (let i = 0; i < 5; i++) expect(r.check('message', t + 100 + i * 100)).toBe('ok')
    for (let i = 0; i < 3; i++) expect(r.check('unit', t + 200 + i)).toBe('ok')
    // The window slides: a hit frees 60 s later.
    expect(r.check('sheet', t + 60_000)).toBe('ok')
    // Independent of 'text'.
    const s = new SessionRates()
    for (let i = 0; i < 30; i++) s.check('sheet', t)
    expect(s.check('text', t)).toBe('ok')
    t += 1
  })

  test('sheetBusy starts false', () => {
    expect(new SessionRates().sheetBusy).toBe(false)
  })
})
