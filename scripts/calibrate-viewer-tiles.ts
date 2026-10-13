/**
 * S10.2 diagnostic: derives the main-heap cost of a retained tile canvas from
 * the REAL viewer (codex/opencode round 6). It mounts the actual
 * InteractiveSheetDoc through the perf page's calibrateViewer, then scrolls it
 * (the workbook has one value, at A1) so the tile cache retains more tiles that
 * draw gridlines only (no text), and measures (forced GC, median of 3)
 * the heap growth between states that the ledger does not explain:
 *
 *   per-tile object bytes =
 *     (Δheap − Δcompiled code − Δ(ledger outside canvasText) − Δshaping charge) / Δtiles
 *
 * where the shaping charge is canvasText − tiles × TILE_CANVAS_BYTES. The
 * result bounds TILE_CANVAS_BYTES from below when TILE_CANVAS_BYTES is at
 * least the printed figure.
 *
 *   bun scripts/calibrate-viewer-tiles.ts [--port 4192] [--rounds 3]
 */
import { spawnSync } from 'node:child_process'
import { chromium } from 'playwright'
import { startStaticServer } from './harness/static-server'
import { measureCodeSpace } from './harness/codeSpace'
import { TILE_CANVAS_BYTES } from '../src/xlsx/view/mainLedger'
import { PERF_OUT_DIR } from '../vite.perf.config'

const arg = (name: string, fallback: number) => { const i = process.argv.indexOf(`--${name}`); return i > 0 ? Number(process.argv[i + 1]) : fallback }
const port = arg('port', 4192), rounds = arg('rounds', 3)

const build = spawnSync('bunx', ['vite', 'build', '-c', 'vite.perf.config.ts'], { stdio: 'inherit' })
if (build.status !== 0) throw new Error('perf page build failed')
const server = await startStaticServer(PERF_OUT_DIR, port)
const browser = await chromium.launch({ channel: 'chrome', headless: true, args: ['--js-flags=--expose-gc'] })
try {
  const page = await browser.newPage({ viewport: { width: 1600, height: 1000 }, deviceScaleFactor: 2 })
  await page.goto(`${server.url}/src/debug/viewer-perf.html`)
  await page.waitForFunction(() => (window as unknown as { __viewerPerf?: { booted: boolean } }).__viewerPerf?.booted === true)
  const cdp = await page.context().newCDPSession(page)
  type State = { heap: number; code: number; tiles: number; ledgerOther: number; canvasText: number }
  const measure = async (updates: number, width = 1440, height = 900): Promise<State> => {
    const r = await page.evaluate(([w, h, u]) => (window as unknown as { __viewerPerf: { calibrateViewer(w: number, h: number, u: number): Promise<{ ledger: { bytes: number; byPool: Record<string, number> } | null; tiles: number | null }> } }).__viewerPerf.calibrateViewer(w, h, u), [width, height, updates] as const)
    const readings: number[] = []
    let used = 0
    for (let k = 0; k < 3; k++) {
      for (let g = 0; g < 3; g++) { await cdp.send('HeapProfiler.collectGarbage'); await new Promise(x => setTimeout(x, 200)) }
      const h = await cdp.send('Runtime.getHeapUsage') as { usedSize: number; embedderHeapUsedSize: number; backingStorageSize: number }
      used = h.usedSize
      readings.push(h.usedSize + h.embedderHeapUsedSize + h.backingStorageSize)
    }
    readings.sort((a, b) => a - b)
    const code = await measureCodeSpace(cdp, used)
    if (!r.ledger || r.tiles === null || !code) throw new Error('calibration state incomplete')
    return { heap: readings[1], code: code.codeBytes, tiles: r.tiles, ledgerOther: r.ledger.bytes - (r.ledger.byPool.canvasText ?? 0), canvasText: r.ledger.byPool.canvasText ?? 0 }
  }
  const results: number[] = []
  for (let round = 0; round < rounds; round++) {
    await page.evaluate(() => (window as unknown as { __viewerPerf: { calibrateViewer(w: number, h: number, u: number): Promise<unknown> } }).__viewerPerf.calibrateViewer(0, 0, 0))
    const a = await measure(0)
    const b = await measure(-12)
    const dTiles = b.tiles - a.tiles
    const shapingA = a.canvasText - a.tiles * TILE_CANVAS_BYTES, shapingB = b.canvasText - b.tiles * TILE_CANVAS_BYTES
    const objectBytes = (b.heap - a.heap - Math.max(0, b.code - a.code) - (b.ledgerOther - a.ledgerOther) - (shapingB - shapingA)) / dTiles
    console.log(JSON.stringify({ round, tiles: [a.tiles, b.tiles], heapDelta: b.heap - a.heap, codeDelta: b.code - a.code, ledgerOtherDelta: b.ledgerOther - a.ledgerOther, shapingDelta: shapingB - shapingA, perTileObjectBytes: Math.round(objectBytes) }))
    if (dTiles > 0) results.push(objectBytes)
  }
  results.sort((a, b) => a - b)
  const median = results[Math.floor(results.length / 2)]
  console.log(`real-viewer per-tile canvas object bytes: median ${Math.round(median)} of ${results.map(Math.round).join(', ')}; TILE_CANVAS_BYTES = ${TILE_CANVAS_BYTES} (${median <= TILE_CANVAS_BYTES ? 'covers it' : 'BELOW the measurement'})`)
  await page.evaluate(() => (window as unknown as { __viewerPerf: { calibrateViewer(w: number, h: number, u: number): Promise<unknown> } }).__viewerPerf.calibrateViewer(0, 0, 0))
} finally {
  await browser.close()
  await server.close()
}
