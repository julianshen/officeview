/** Stage 4a Task 4 — HELLO `features` select tiles mode for XLSX; fixed per sid; op gating; the session clock (R9, R12). */
import { afterAll, describe, expect, test } from 'bun:test'
import { buildXlsx } from '../../src/testdata/ooxml-builders'
import { hiddenFixtures } from '../../scripts/stream-s0/hidden-fixtures'
import { AuditStore } from '../src/audit-store'
import { AuditTable } from '../src/audit'
import type { AuthDecision } from '../src/hooks'
import type { Pool } from '../src/pool'
import { clockEpochMs, sheetClockAt } from '../src/sheet/clock'
import { CLOSE } from '../src/ws'
import { Client, corpus, disposePool, pool, start } from './harness'

afterAll(async () => { await disposePool() })

const TILES = ['sheet-tiles@1']
const STAGE3_KEYS = ['can', 'caps', 'docRef', 'kind', 'layoutFingerprint', 'limits', 'name', 'notices', 'policy', 'recorder', 'session', 'units', 'v']
const TILES_BLOCK = { px: 512, zoomSteps: [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4], dprs: [1, 2], palettes: ['light', 'dark'], maxBatch: 64 }

const small = () => buildXlsx([
  { name: 'First', rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 0 }, { ref: 'B1', v: 42 }] }] },
  { name: 'Second', rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 1 }] }] },
], ['small sheet', 'second sheet'])
const fixture = async (id: string) => (await hiddenFixtures()).find((f) => f.id === id)!
const fixtureOfKind = async (kind: 'docx' | 'pptx') => (await hiddenFixtures()).find((f) => f.kind === kind)!

async function open(h: { url: string }, features?: unknown, ready = true) {
  const c = new Client(h.url)
  const m = await c.hello(features === undefined ? {} : { features })
  if (ready) await c.ready()
  return { c, m }
}
/** A HELLO that the server closes: its close code. */
async function closedBy(url: string, extra: object) {
  const c = new Client(url)
  await c.opened()
  c.send({ reqId: 1, type: 'hello', client: 'officeview/1.0', recVersions: [1], ...extra })
  return (await c.closed).code
}
const refusedCode = async (c: Client, reqId: number) => (await c.next((x) => 'text' in x && x.text.reqId === reqId) as { text: Record<string, any> }).text

describe('HELLO features select the sheet mode', () => {
  test('no features → units mode, Stage 3 manifest', async () => {
    const h = await start({ doc: await small() })
    const { c, m } = await open(h)
    expect(m.kind).toBe('xlsx')
    expect(Object.keys(m).sort()).toEqual(STAGE3_KEYS)
    c.close(); await h.stop()
  }, 60_000)

  test('sheet-tiles@1 on xlsx → tiles manifest', async () => {
    const f = await fixture('xlsx-hidden-sheet')
    const h = await start({ doc: f.withHidden })
    const { c, m } = await open(h, TILES)
    expect(m.sheetMode).toBe('tiles')
    expect(m.sheets).toEqual([{ sheetId: '0', name: 'S' }])
    expect(m.tiles).toEqual(TILES_BLOCK as never)
    expect(Object.keys(m).sort()).toEqual([...STAGE3_KEYS, 'sheetMode', 'sheets', 'tiles'].sort())
    c.close(); await h.stop()
    // Two visible sheets: ids "0".."n-1" in tab order.
    const h2 = await start({ doc: await small() })
    const o2 = await open(h2, TILES)
    expect(o2.m.sheets).toEqual([{ sheetId: '0', name: 'First' }, { sheetId: '1', name: 'Second' }])
    o2.c.close(); await h2.stop()
  }, 60_000)

  test('features on docx and pptx stay units', async () => {
    for (const doc of [(await fixtureOfKind('docx')).withoutHidden, corpus('pypptx-test.pptx')]) {
      const h = await start({ doc })
      const { c, m } = await open(h, TILES)
      expect(Object.keys(m).sort()).toEqual(STAGE3_KEYS)
      c.close(); await h.stop()
    }
  }, 60_000)

  test('unknown well-formed features are ignored', async () => {
    const h = await start({ doc: await small() })
    const a = await open(h, ['x@9', 'sheet-tiles@1'])
    expect(a.m.sheetMode).toBe('tiles')
    const b = await open(h, ['x@9'])
    expect(b.m.sheetMode).toBeUndefined()
    expect(Object.keys(b.m).sort()).toEqual(STAGE3_KEYS)
    a.c.close(); b.c.close(); await h.stop()
  }, 60_000)

  test('malformed features close the HELLO', async () => {
    const h = await start({ doc: await small() })
    const bad: unknown[] = ['sheet-tiles@1', Array.from({ length: 17 }, (_, i) => `f${i}`), ['a'.repeat(33)], [1], ['', 'a'], ['a', 'a']]
    for (const features of bad) expect(await closedBy(h.url, { features })).toBe(CLOSE['unsupported-version'])
    // The bounds themselves are accepted.
    const edge = await open(h, [...Array.from({ length: 15 }, (_, i) => `f${i}`), 'a'.repeat(32)])
    expect(Object.keys(edge.m).sort()).toEqual(STAGE3_KEYS)
    edge.c.close(); await h.stop()
  }, 60_000)

  test('resume must present the same features', async () => {
    const h = await start({ doc: await small() })
    const { c, m } = await open(h, TILES)
    const sid = m.session.sid
    expect(await closedBy(h.url, { resume: { sid, resumeProof: m.session.token } })).toBe(CLOSE['auth-failed'])
    expect(await closedBy(h.url, { resume: { sid, resumeProof: m.session.token }, features: ['x@9', ...TILES] })).toBe(CLOSE['auth-failed'])
    // The failed attempts left the session resumable.
    const r = new Client(h.url)
    const rm = await r.hello({ resume: { sid, resumeProof: m.session.token }, features: TILES })
    expect(rm.session.sid).toBe(sid)
    expect(rm.sheetMode).toBe('tiles')
    expect(rm.sheets).toEqual([{ sheetId: '0', name: 'First' }, { sheetId: '1', name: 'Second' }])
    r.close(); c.close(); await h.stop()
  }, 60_000)
})

describe('op gating (R6)', () => {
  test('op gating', async () => {
    const h = await start({ doc: await small(), policy: { copy: 'full' } })
    const t = await open(h, TILES)
    const range = { anchor: { unit: 0, x: 0, y: 0 }, focus: { unit: 0, x: 10, y: 10 } }
    const token = t.m.session.token
    const ops: Array<Record<string, unknown>> = [
      { type: 'unit', unitIndex: 0, scale: 1 }, { type: 'search', query: 'small' }, { type: 'select', range }, { type: 'copy', range }, { type: 'a11y', unit: 0 },
    ]
    let reqId = 30
    for (const op of ops) {
      t.c.send({ token, reqId: ++reqId, ...op })
      expect(await refusedCode(t.c, reqId)).toMatchObject({ type: 'refused', code: 'unsupported-op' })
    }
    const p = await t.c.print(50, 0, 0)
    expect(p.start?.type).toBe('print-start')
    expect(p.end.type).toBe('print-done')
    const u = await open(h)
    for (const type of ['tiles', 'sheet-layout', 'sheet-cancel']) {
      u.c.send({ token: u.m.session.token, reqId: ++reqId, type })
      expect(await refusedCode(u.c, reqId)).toMatchObject({ type: 'refused', code: 'unsupported-op' })
    }
    t.c.close(); u.c.close(); await h.stop()
  }, 60_000)

  test('tiles open renders no unit until print', async () => {
    let renders = 0
    const base = pool()
    const counting = new Proxy(base, {
      get(target, key) {
        if (key === 'render') return (...a: Parameters<Pool['render']>) => { renders++; return target.render(...a) }
        const v = Reflect.get(target, key, target)
        return typeof v === 'function' ? v.bind(target) : v
      },
    }) as Pool
    const h = await start({ doc: await small(), pool: counting })
    const { c } = await open(h, TILES)
    await Bun.sleep(100)
    expect(renders).toBe(0)
    const p = await c.print(60, 0, 0)
    expect(p.start?.type).toBe('print-start')
    expect(renders).toBe(1)
    c.close(); await h.stop()
  }, 60_000)

  test('units and tiles sessions on one document', async () => {
    const h = await start({ doc: await small() })
    const u = await open(h)
    const t = await open(h, TILES)
    u.c.unit(70, 0, 1)
    const r = await u.c.unitMessages(70)
    expect(r.unit.unitId).toBe(0)
    t.c.unit(71, 0, 1)
    expect(await refusedCode(t.c, 71)).toMatchObject({ type: 'refused', code: 'unsupported-op' })
    u.c.close(); await u.c.closed; await Bun.sleep(50)
    const p = await t.c.print(72, 0, 0)
    expect(p.start?.type).toBe('print-start')
    t.c.close(); await h.stop()
  }, 60_000)

  test('the audit open row records the mode', async () => {
    const store = AuditStore.open(':memory:')
    const h = await start({ doc: await small(), audit: new AuditTable({ store }) })
    const t = await open(h, TILES), u = await open(h)
    expect(store.getRow(t.m.session.sid)!.sheetMode).toBe('tiles')
    expect(store.getRow(u.m.session.sid)!.sheetMode).toBe('units')
    t.c.close(); u.c.close(); await h.stop()
    store.close()
  }, 60_000)
})

describe('session clock (R9)', () => {
  test('sheetClockAt and clockEpochMs', () => {
    const at = Date.UTC(2026, 10, 22, 23, 30)
    expect(sheetClockAt(at, 'Asia/Tokyo')).toEqual({ day: '2026-11-23', timeZone: 'Asia/Tokyo' })
    expect(sheetClockAt(at, 'UTC')).toEqual({ day: '2026-11-22', timeZone: 'UTC' })
    expect(() => sheetClockAt(at, 'Mars/Base')).toThrow(RangeError)
    expect(clockEpochMs({ day: '2026-11-23', timeZone: 'Asia/Tokyo' })).toBe(Date.UTC(2026, 10, 23, 3))
    expect(clockEpochMs({ day: '2026-11-22', timeZone: 'UTC' })).toBe(Date.UTC(2026, 10, 22, 12))
  })

  test('a sid keeps its day across midnight, and its zone on resume', async () => {
    let t = Date.UTC(2026, 10, 22, 23, 30)
    let policy: AuthDecision['policy'] = {}
    const h = await start({ doc: await small(), now: () => t, tokenTtlSec: 86_400, authorize: async () => ({ viewerId: 'alice', policy }) })
    const { c, m } = await open(h, TILES)
    const sid = m.session.sid
    expect(h.stream.sheetClockOf(sid)).toEqual({ day: '2026-11-22', timeZone: 'UTC' })
    t += 45 * 60_000
    const r = new Client(h.url)
    const rm = await r.hello({ resume: { sid, resumeProof: m.session.token }, features: TILES })
    expect(rm.session.sid).toBe(sid)
    expect(h.stream.sheetClockOf(sid)).toEqual({ day: '2026-11-22', timeZone: 'UTC' })
    // A resume whose decision resolves to another zone closes auth-failed.
    policy = { timeZone: 'Asia/Tokyo' }
    expect(await closedBy(h.url, { resume: { sid, resumeProof: rm.session.token }, features: TILES })).toBe(CLOSE['auth-failed'])
    expect(h.stream.sheetClockOf(sid)).toEqual({ day: '2026-11-22', timeZone: 'UTC' })
    r.close(); c.close(); await h.stop()
  }, 60_000)

  test('policy.timeZone overrides the server default; an invalid zone closes policy-refused', async () => {
    const at = Date.UTC(2026, 10, 22, 23, 30)
    const h = await start({ doc: await small(), now: () => at, policy: { timeZone: 'Asia/Tokyo' }, sheetTimeZone: 'Europe/London' })
    const { c, m } = await open(h, TILES)
    expect(h.stream.sheetClockOf(m.session.sid)).toEqual({ day: '2026-11-23', timeZone: 'Asia/Tokyo' })
    c.close(); await h.stop()
    const d = await start({ doc: await small(), now: () => at, sheetTimeZone: 'Europe/London' })
    const o = await open(d, TILES)
    expect(d.stream.sheetClockOf(o.m.session.sid)).toEqual({ day: '2026-11-22', timeZone: 'Europe/London' })
    o.c.close(); await d.stop()
    const b = await start({ doc: await small(), policy: { timeZone: 'bad' } })
    expect(await closedBy(b.url, { features: TILES })).toBe(CLOSE['policy-refused'])
    await b.stop()
  }, 60_000)
})
