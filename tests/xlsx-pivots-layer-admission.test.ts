// @vitest-environment node
/**
 * PR #55 round 3 P2: the receiver's decoded pivot layer retains each pivot's saved
 * layout (lines, resolved item vectors, parent bands). Its pre-build reservation and
 * retained charge include that storage, so a 100,000-line pivot is charged for it
 * and refused when the budget only fits the fixed region allowance.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import type { CacheIdentity, RemotePivotDisplay, WorkerReply } from '../src/xlsx/view/contracts'
import { MAX_REGION_STYLE_MEMO_ENTRIES, PIVOT_FORMAT_MEMO_BYTES, REGION_STYLE_LAYER_BASE_BYTES, REGION_STYLE_MEMO_ENTRY_ESTIMATE_BYTES, REGION_STYLE_REGION_ESTIMATE_BYTES } from '../src/xlsx/view/contracts'
import { RemoteRegionStyleCache } from '../src/worker/viewerRemoteRegionStyles'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { ensureTableStyleEngine } from '../src/xlsx/tableStyleLoader'
import { savedPivotLayout, savedPivotLayoutBytes } from '../src/xlsx/pivots/savedLayout'
import type { PivotTableDefinition } from '../src/xlsx/pivots/model'
import { formatA1 } from '../src/xlsx/view/coordinates'

const bytes = (name: string) => new Uint8Array(readFileSync(resolve(process.cwd(), `tests/fixtures/xlsx-pivots/${name}.xlsx`)))

/** A fixture pivot stretched to `lines` saved row lines, two levels deep in tabular form
 * (every tenth line opens a parent band over the next nine). */
async function longPivot(lines: number): Promise<PivotTableDefinition> {
  const doc = await parseXlsx(await OfficePackage.load(bytes('pivot-formats-override')))
  await ensureTableStyleEngine(doc.sheets)
  const base = doc.sheets.find(sheet => sheet.pivotTables?.length)!.pivotTables![0]!
  const [from] = base.location.ref.split(':')
  const m = /^([A-Z]+)(\d+)$/.exec(from!)!
  const minRow = Number(m[2]) - 1
  const minCol = [...m[1]!].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1
  const colItems = base.colItems!
  const rowItems = Array.from({ length: lines }, (_, n) => n % 10 === 0 ? { type: 'data' as const, r: 0, i: 0, x: [n, 0] } : { type: 'data' as const, r: 1, i: 0, x: [n % 10] })
  const maxRow = minRow + base.location.firstDataRow + lines - 1
  const maxCol = minCol + base.location.firstDataCol + colItems.length - 1
  return { ...base, rowItems, location: { ...base.location, ref: `${formatA1({ row: minRow, col: minCol })}:${formatA1({ row: maxRow, col: maxCol })}` } }
}

async function admit(def: PivotTableDefinition, headroom = Infinity) {
  const owner: CacheIdentity = { docId: `long-${def.rowItems!.length}-${headroom}`, sheetId: 's', valueGen: 1, fontGen: 0, viewGen: 0 }
  const payload: RemotePivotDisplay = { sheetId: 's', stamp: owner, pivotTables: [def], pivotStyles: [], pivotDxfs: [{ id: 0, dxf: { fill: { pattern: 'solid', bg: { rgb: 'FFFFE699' } } } }], diagnostics: [] }
  const reply = { kind: 'complete', commandKind: 'read', requestId: 'r', owner, diagnostics: [], replyBytes: 64,
    result: { kind: 'read', result: { op: 'pivotDisplay', result: payload } } } as unknown as WorkerReply<'read'>
  let charged = 0, asked = 0
  const cache = new RemoteRegionStyleCache({
    identity: () => owner, hasRegionStyles: () => false, hasPivotDisplay: () => true, loadState: () => ({ completeness: 'complete' }),
    grant: () => 4096, releaseGrant: () => {}, request: async () => reply, nextRequestId: () => 'r', release: () => {},
    admitLayer: (_sheet, _request, bytes) => { asked = bytes; if (bytes > headroom) return null; charged = bytes; return () => { charged = 0 } }, changed: () => {},
  }, 1024 * 1024 * 1024, 'pivotDisplay')
  let state = cache.get('s')
  for (let i = 0; i < 2000 && state.state !== 'ready' && state.state !== 'failed'; i++) { await new Promise(tick => setTimeout(tick, 2)); state = cache.get('s') }
  const out = { state: state.state, asked, charged, retained: cache.retainedBytes, layerBytes: state.layer?.retainedBytes, features: state.diagnostics.map(item => item.feature) }
  cache.close()
  return out
}

describe('decoded pivot layer admission includes the saved layout', () => {
  it('charges a 100,000-line pivot for its lines, item vectors and parent bands', async () => {
    const def = await longPivot(100_000)
    const layout = savedPivotLayout(def)
    expect(layout).toBeDefined()
    expect(layout!.parentBands.size).toBeGreaterThan(80_000)
    const layoutBytes = savedPivotLayoutBytes(def)
    // At least the 64 bytes per line regions.ts already assumed, plus vectors and bands.
    expect(layoutBytes).toBeGreaterThan(100_000 * 64 + 90_000 * 8)
    // What 429f26da reserved: the fixed layer allowance and one formatted pivot's memo.
    const fixed = REGION_STYLE_LAYER_BASE_BYTES + REGION_STYLE_REGION_ESTIMATE_BYTES + MAX_REGION_STYLE_MEMO_ENTRIES * REGION_STYLE_MEMO_ENTRY_ESTIMATE_BYTES + PIVOT_FORMAT_MEMO_BYTES
    const ready = await admit(def)
    expect(ready.state).toBe('ready')
    expect(ready.charged).toBe(fixed + layoutBytes)
    expect(ready.retained).toBeGreaterThanOrEqual(ready.charged)
    // The layer's own estimate counts the same layout and never exceeds the charge.
    expect(ready.layerBytes).toBeGreaterThan(layoutBytes)
    expect(ready.layerBytes).toBeLessThanOrEqual(ready.charged)
    // Headroom for the fixed allowance but not the layout: refused before building.
    const refused = await admit(def, fixed + 1024 * 1024)
    expect(refused.state).toBe('failed')
    expect(refused.asked).toBe(fixed + layoutBytes)
    expect(refused.charged).toBe(0)
    expect(refused.features).toContain('pivot-display-reply-limit')
  }, 60_000)
})
