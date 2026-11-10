/**
 * T8 (Stage-0 subset, scope call 4): XLSX tile timing under @napi-rs/canvas.
 *
 *   bun run s0:napi scripts/stream-s0/t8-tiles.ts
 *
 * A 500 × 100 sheet (mixed text / number / date-like values) is painted as
 * viewer tiles — TILE_DEVICE_PX (512) device px square at deviceScale 1 —
 * through the real viewport pipeline: buildSheetGeometry → renderSheetViewport
 * (`src/xlsx/view/renderViewport.ts`), each tile on its own canvas with a
 * shared LayoutCache like the viewer. 200 tiles walk the sheet row-major.
 * Three runs; the gate uses the best run's p95 (another feature's test suite
 * may load the machine — os.loadavg() is recorded at start and end).
 */
import { readFileSync } from 'node:fs'
import { loadavg } from 'node:os'
import { join } from 'node:path'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { preloadTableStyles } from '../../src/xlsx/tableStyleLoader'
import { buildSheetGeometry } from '../../src/xlsx/view/viewport'
import { renderSheetViewport } from '../../src/xlsx/view/renderViewport'
import { LayoutCache } from '../../src/xlsx/view/textLayout'
import { TILE_DEVICE_PX } from '../../src/xlsx/view/tileCache'
import type { XlsxSheet } from '../../src/xlsx/types'
import type { LogicalViewport } from '../../src/xlsx/view/contracts'
import { resolveBackend, writeResult } from './results'

const ROWS = 500
const COLS = 100
const TILES = 200
const TARGET_P95_MS = 120

async function sheet500x100(): Promise<XlsxSheet> {
  await preloadTableStyles()
  const bytes = readFileSync(join(import.meta.dir, '..', '..', 'tests', 'fixtures', 'xlsx-pivots', 'table-medium-2.xlsx'))
  const doc = await parseXlsx(await OfficePackage.load(bytes))
  const sheet = doc.sheets[0]!
  sheet.tables = []
  sheet.rows = Array.from({ length: ROWS }, (_, r) => ({
    index: r,
    cells: Array.from({ length: COLS }, (_, c) => ({
      ref: '', row: r, col: c, styleIndex: 0,
      value: r === 0 ? `Col ${c + 1}` : (c % 3 === 0 ? `Item ${r}-${c}` : (c % 3 === 1 ? r * 1.25 + c : 45000 + r)),
    })),
  })) as XlsxSheet['rows']
  return sheet
}

function percentile(sorted: number[], p: number): number {
  return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * p))]!
}

async function main(): Promise<number> {
  const backend = resolveBackend()
  if (backend !== '@napi-rs/canvas') {
    console.error('[t8] refused: run with `bun run s0:napi scripts/stream-s0/t8-tiles.ts`')
    return 2
  }
  const { createCanvas } = await import('canvas')
  const sheet = await sheet500x100()
  const geometry = buildSheetGeometry(sheet)
  const cols = Math.ceil((COLS * 80) / TILE_DEVICE_PX) // ~80 CSS px per default column
  const loadStart = loadavg()
  const runs: Array<{ run: number; p50: number; p95: number; max: number; kinds: Record<string, number>; records: number }> = []
  for (let run = 1; run <= 3; run++) {
    const cache = new LayoutCache()
    const times: number[] = []
    const kinds: Record<string, number> = {}
    let records = 0
    for (let i = 0; i < TILES; i++) {
      const tx = i % cols
      const ty = Math.floor(i / cols)
      const viewport: LogicalViewport = { box: { x: tx * TILE_DEVICE_PX, y: ty * TILE_DEVICE_PX, width: TILE_DEVICE_PX, height: TILE_DEVICE_PX }, zoom: 1, deviceScale: 1 }
      const t0 = performance.now()
      const canvas = createCanvas(TILE_DEVICE_PX, TILE_DEVICE_PX)
      const ctx = canvas.getContext('2d')
      const result = renderSheetViewport(sheet, ctx as never, geometry, viewport, {}, { cache })
      times.push(performance.now() - t0)
      kinds[result.kind] = (kinds[result.kind] ?? 0) + 1
      records += result.records.length
    }
    times.sort((a, b) => a - b)
    runs.push({ run, p50: percentile(times, 0.5), p95: percentile(times, 0.95), max: times[times.length - 1]!, kinds, records })
  }
  const loadEnd = loadavg()
  const bestOf3P95 = Math.min(...runs.map((r) => r.p95))
  // A tile that painted no cell records would make the timing meaningless.
  const meaningful = runs.every((r) => r.records > 0 && (r.kinds.detail ?? 0) > 0)
  const green = meaningful && bestOf3P95 <= TARGET_P95_MS
  writeResult('t8', {
    status: green ? 'GREEN' : 'RED',
    sheet: { rows: ROWS, cols: COLS }, tilePx: TILE_DEVICE_PX, tiles: TILES, deviceScale: 1,
    pipeline: 'buildSheetGeometry → renderSheetViewport (src/xlsx/view/renderViewport.ts)',
    runs, bestOf3P95, targetP95Ms: TARGET_P95_MS, meaningful, loadavgStart: loadStart, loadavgEnd: loadEnd,
    note: 'Gate = best-of-3 p95; tiles walk the sheet row-major from A1; another worktree\'s tests may share the CPU (loadavg recorded).',
  })
  console.log(`[t8] ${green ? 'GREEN' : 'RED'} best-of-3 p95 ${bestOf3P95.toFixed(2)} ms (runs ${runs.map((r) => r.p95.toFixed(2)).join(', ')}), kinds ${JSON.stringify(runs[0]!.kinds)}, records/run ${runs[0]!.records}, loadavg ${loadStart[0]!.toFixed(1)}→${loadEnd[0]!.toFixed(1)}`)
  return green ? 0 : 3
}

if (import.meta.main) process.exit(await main())
