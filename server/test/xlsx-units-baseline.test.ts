/** Stage 4a Task 2 — characterization of today's units-mode XLSX session (green before any tile code; not RED-first). */
import { afterAll, describe, expect, test } from 'bun:test'
import { buildXlsx } from '../../src/testdata/ooxml-builders'
import { Client, disposePool, start } from './harness'

afterAll(async () => { await disposePool() })

const A1 = 'north harbour'
const xlsx = () => buildXlsx([
  { name: 'First', cols: '<col min="1" max="2" width="24" customWidth="1"/>', rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 0 }, { ref: 'B1', t: 's', v: 1 }] }] },
  { name: 'Second', rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 2 }] }] },
], [A1, 'east quay', 'second sheet text'])

async function session(o: Parameters<typeof start>[0] = {}) {
  const h = await start({ doc: await xlsx(), ...o }), c = new Client(h.url)
  const m = await c.hello(); await c.ready()
  return { h, c, m, stop: () => { c.close(); return h.stop() } }
}

describe('units-mode xlsx baseline (Stage 3 contract Tasks 4 and 16 must keep)', () => {
  test('units-mode xlsx manifest has exactly the Stage 3 keys', async () => {
    const { m, stop } = await session()
    expect(m.kind).toBe('xlsx')
    expect(Object.keys(m).sort()).toEqual(['can', 'caps', 'docRef', 'kind', 'layoutFingerprint', 'limits', 'name', 'notices', 'policy', 'recorder', 'session', 'units', 'v'])
    expect(m.units.count).toBe(2)
    expect('sheetMode' in m).toBe(false)
    expect('sheets' in m).toBe(false)
    expect('tiles' in m).toBe(false)
    await stop()
  }, 60_000)

  test('units-mode xlsx copy answers sheet text', async () => {
    const { c, m, stop } = await session({ policy: { copy: 'full' } })
    expect(m.can.copy).toBe(true)
    const range = { anchor: { unit: 0, x: 0, y: 0 }, focus: { unit: 0, x: 1e6, y: 1e6 } }
    c.send({ token: m.session.token, reqId: 10, type: 'select', range })
    const sel = await c.next((x) => 'text' in x && x.text.reqId === 10) as { text: Record<string, any> }
    expect(sel.text.type).toBe('select-result')
    expect(sel.text.paragraphs).toBeGreaterThanOrEqual(1)
    c.send({ token: m.session.token, reqId: 11, type: 'copy', range })
    let text = ''
    for (;;) {
      const r = await c.next((x) => 'text' in x && x.text.reqId === 11) as { text: Record<string, any> }
      expect(r.text.type).toBe('copy-result')
      text += r.text.text
      if (r.text.last === true) break
    }
    expect(text).toContain(A1)
    await stop()
  }, 60_000)

  test('units-mode xlsx unit request renders', async () => {
    const { c, stop } = await session()
    c.unit(20, 0, 1)
    const { messages, unit } = await c.unitMessages(20)
    expect(messages.filter((x) => x.kind === 3).length).toBe(1)
    expect(messages.at(-1)!.kind).toBe(3)
    expect(unit).toBeDefined()
    await stop()
  }, 60_000)
})
