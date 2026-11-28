/** Stage 4c Task 3 — the `sheetSort` manifest key, the gating of the five 4c ops, and the worker-kind refusals (P1, P14). */
import { afterAll, describe, expect, test } from 'bun:test'
import { buildXlsx } from '../../src/testdata/ooxml-builders'
import { hiddenFixtures } from '../../scripts/stream-s0/hidden-fixtures'
import type { RateClass, SessionRates } from '../src/rate'
import { Client, disposePool, start, type Harness } from './harness'
import { grid500x100 } from './sheet-fixtures'

afterAll(async () => { await disposePool() })

const TILES = ['sheet-tiles@1']
const VIEW = ['sheet-view@1']
const BOTH = ['sheet-tiles@1', 'sheet-view@1']
const STAGE3_KEYS = ['can', 'caps', 'docRef', 'kind', 'layoutFingerprint', 'limits', 'name', 'notices', 'policy', 'recorder', 'session', 'units', 'v']
const TILES_KEYS = [...STAGE3_KEYS, 'sheetMode', 'sheets', 'tiles'].sort()
const OPS_4C = ['sheet-sort', 'sheet-filter', 'filter-values', 'range-detect', 'sheet-order']
const PIVOT_KINDS = ['pivotOpen', 'pivotApply', 'pivotCommit', 'pivotFormulaStatus', 'pivotRelease', 'controlApply', 'controlCommit', 'controlCancel']

const small = () => buildXlsx([
  { name: 'First', rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 0 }, { ref: 'B1', v: 42 }] }] },
  { name: 'Second', rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 1 }] }] },
], ['small sheet', 'second sheet'])

async function open(h: { url: string }, features?: readonly string[]) {
  const c = new Client(h.url)
  const m = await c.hello(features === undefined ? {} : { features })
  await c.ready()
  return { c, m }
}
const answerOf = async (c: Client, reqId: number) => (await c.next((x) => 'text' in x && x.text.reqId === reqId) as { text: Record<string, any> }).text

/** Wraps the sid's `SessionRates.check` so each call is recorded by class, once the server has taken the READY. */
async function chargesOf(h: Harness, sid: string): Promise<RateClass[]> {
  while (h.stream.table.get(sid)?.status !== 'stream') await Bun.sleep(5)
  const rates = h.stream.sheetRatesOf(sid) as SessionRates
  const seen: RateClass[] = []
  const check = rates.check.bind(rates)
  rates.check = (kind, now) => { seen.push(kind); return check(kind, now) }
  return seen
}

describe('sheetSort on every sheet-view manifest', () => {
  test('sheetSort on every sheet-view manifest', async () => {
    const h = await start({ doc: await small() })
    const v = await open(h, VIEW)
    expect(v.m.sheetView).toBe(1)
    expect(v.m.sheetSort).toBe(1)
    expect('typedFilters' in v.m).toBe(false)
    expect(Object.keys(v.m).sort()).toEqual([...TILES_KEYS, 'sheetView', 'sheetSort'].sort())
    const b = await open(h, BOTH)
    expect(b.m.sheetSort).toBe(1)
    expect('typedFilters' in b.m).toBe(false)
    // 4a: sheet-tiles@1 alone is key for key the 4a manifest.
    const t = await open(h, TILES)
    expect(Object.keys(t.m).sort()).toEqual(TILES_KEYS)
    // Units mode: Stage 3's manifest.
    const u = await open(h)
    expect(Object.keys(u.m).sort()).toEqual(STAGE3_KEYS)
    v.c.close(); b.c.close(); t.c.close(); u.c.close(); await h.stop()

    const f = await start({ doc: await small(), policy: { contentPolicy: 'full' } })
    const fv = await open(f, VIEW)
    expect(fv.m.policy.contentPolicy).toBe('full')
    expect(fv.m.sheetSort).toBe(1)
    expect(fv.m.typedFilters).toBe(true)
    expect(Object.keys(fv.m).sort()).toEqual([...TILES_KEYS, 'sheetView', 'sheetSort', 'typedFilters'].sort())
    const ft = await open(f, TILES)
    expect(Object.keys(ft.m).sort()).toEqual(TILES_KEYS)
    const fu = await open(f)
    expect(Object.keys(fu.m).sort()).toEqual(STAGE3_KEYS)
    fv.c.close(); ft.c.close(); fu.c.close(); await f.stop()
  }, 60_000)

  test('a DOCX gets no new keys', async () => {
    const docx = (await hiddenFixtures()).find((x) => x.kind === 'docx')!.withoutHidden
    for (const policy of [undefined, { contentPolicy: 'full' as const }]) {
      const h = await start({ doc: docx, ...(policy ? { policy } : {}) })
      for (const features of [VIEW, BOTH]) {
        const { c, m } = await open(h, features)
        expect(Object.keys(m).sort()).toEqual(STAGE3_KEYS)
        c.close()
      }
      await h.stop()
    }
  }, 60_000)
})

describe('the 4c ops need sheetView', () => {
  test('the 4c ops need sheetView', async () => {
    // An injected clock keeps the 'message' class (20/s) out of the way: one request per 500 ms, 40 of them inside 60 s.
    let t = Date.UTC(2026, 10, 27, 12)
    const h = await start({ doc: await small(), now: () => t, tokenTtlSec: 86_400 })
    for (const features of [TILES, undefined] as const) {
      t += 120_000
      const { c, m } = await open(h, features)
      expect(h.stream.sheetViewOf(m.session.sid)).toBe(false)
      const charges = await chargesOf(h, m.session.sid)
      let closed: number | undefined
      void c.closed.then((x) => { closed = x.code })
      for (let i = 0; i < 40; i++) {
        t += 500
        const reqId = 2000 + i
        c.send({ token: m.session.token, reqId, type: OPS_4C[i % OPS_4C.length], sheetId: '0', viewGen: 0 })
        expect(await answerOf(c, reqId)).toEqual({ reqId, type: 'refused', code: 'unsupported-op' })
      }
      // Uncharged: each costs only its 'message' check.
      expect(charges).toEqual(Array(40).fill('message'))
      await Bun.sleep(50)
      expect(closed).toBeUndefined()
      c.close()
    }
    await h.stop()
  }, 60_000)
})

describe('4b layout bytes unchanged', () => {
  test('4b layout bytes unchanged', async () => {
    // Pasted from base fb624989: the one `sheet-layout` part of grid500x100 for a `sheet-view@1` session, reqId 7.
    const BASE = '{"reqId":7,"type":"sheet-layout","seq":0,"last":true,"head":{"sheetId":"0","viewGen":0,"defaults":{"rowPx":20,"colPx":64},"used":{"r0":0,"r1":499,"c0":0,"c1":99},"sort":null,"filterSummary":[],"authorHidden":false,"counts":{"visibleRows":500,"totalRows":500}}}'
    const h = await start({ doc: await grid500x100() })
    const { c } = await open(h, VIEW)
    const r = await c.sheetLayout(7, '0')
    expect(r.end).toBeUndefined()
    expect(r.parts.length).toBe(1)
    expect(r.parts[0].bytes).toBe(new TextEncoder().encode(BASE).length)
    expect(JSON.stringify(r.parts[0].text)).toBe(BASE)
    c.close(); await h.stop()
  }, 30_000)
})

describe('pivot and control kinds are refused', () => {
  test('pivot and control kinds are refused', async () => {
    const h = await start({ doc: await small() })
    const expected: Array<[string, string]> = [...PIVOT_KINDS.map((k): [string, string] => [k, 'pivot-refused']), ['prepareColumns', 'server-prepared'], ['read', 'stream-no-cell-data']]
    let reqId = 3000
    for (const features of [TILES, undefined] as const) {
      const { c, m } = await open(h, features)
      const charges = await chargesOf(h, m.session.sid)
      let closed: number | undefined
      void c.closed.then((x) => { closed = x.code })
      for (const [type, code] of expected) {
        const before = charges.length
        // No token: the refusal comes before any token check.
        c.send({ reqId: ++reqId, type })
        expect(await answerOf(c, reqId)).toEqual({ reqId, type: 'refused', code })
        expect(charges.slice(before)).toEqual(['message'])
      }
      await Bun.sleep(50)
      expect(closed).toBeUndefined()
      c.close()
    }
    await h.stop()
  }, 60_000)
})
