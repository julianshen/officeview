/**
 * Stage 4a Task 1: the tile feasibility gate (R1). A measurement spike, no product code.
 *
 *   bun scripts/stream-s4/tile-feasibility.ts [--quick] [--out <json>] [--compare <previous.json>]
 *
 * Legs (run in this order; (d) first so its heap baseline is the cleanest of the process):
 *   (d) million rows: open (project → cached-policy parse → buildSheetGeometry → prepareSheetAssets) wall time and
 *       heap growth, the engine-policy parse beside it, then the render of the tile holding row 1,048,576;
 *   (a) headless paint of tile (0,0) under serverBackend(), as sheetTiles.ts:340–360 paints;
 *   (b) per-tile in-process latency over the used range of grid500x100 at (zoom 1, dpr 1) and (zoom 1, dpr 2):
 *       render, visible mark (`fixed`), PNG encode, markRaster with the per-tile basis, gzip framing;
 *   (c) Layer-1 on tiles: 64 session codes, 30 captures per class of a 2×2-tile pane (900×900 at (40,40)),
 *       registered against the R10 composite reference and decided with tauFor(64); a blank pane; and the
 *       `markLayerRaster` crop of a sheet-wide basis on the same captures.
 * Pinned parse path (F1): every tile paints over `projectPackage(…, 'visible-only')` then
 * `detectAndParse(projected, pkg => parseXlsx(pkg, { valuePolicy: 'cached' }))`.
 *
 * Deterministic inputs: fixture bytes (fixed zip dates), a constant master key and docRef, codes from
 * `messageFor` over fixed tokens. `--quick` samples fewer tiles and captures and never writes the committed JSON.
 */
import { createHash } from 'node:crypto'
import { readFileSync, writeFileSync } from 'node:fs'
import { loadavg } from 'node:os'
import { resolve } from 'node:path'
import { gzipSync } from 'node:zlib'
import { installCanvasBackend } from '../../src/core/canvas-backend'
import { OfficePackage } from '../../src/core/zip'
import { detectAndParse } from '../../src/components/loadOfficeSource'
import { prepareSheetAssets, type SheetAssetsLease } from '../../src/components/xlsx/sheetAssets'
import { parseXlsx } from '../../src/xlsx/parse'
import type { XlsxDocument, XlsxSheet } from '../../src/xlsx/types'
import { buildSheetGeometry, type SheetGeometryResult } from '../../src/xlsx/view/viewport'
import { renderSheetViewport, type ViewportPaintResult } from '../../src/xlsx/view/renderViewport'
import { SHEET_ZOOM_STEPS } from '../../src/xlsx/view/sheetZoom'
import { TILE_DEVICE_PX, tileLogicalBox } from '../../src/xlsx/view/tileCache'
import { LayoutCache } from '../../src/xlsx/view/textLayout'
import { encodeMessage, Kind } from '../../src/stream/frame'
import { encodeRasterBody } from '../../src/stream/raster'
import { encodeUnit } from '../../src/stream/unit'
import { serverBackend } from '../../server/src/backend'
import { admit, projectPackage } from '../../server/src/policy'
import { criticalChunksOnly } from '../../server/src/render'
import { markOptions, paintMark } from '../../server/src/mark/visible'
import { basis, deriveBasisKey } from '../../server/src/mark/basis'
import { extendedCodeword } from '../../server/src/mark/bch'
import { decideScores, prepareCandidates, scoreCapture, type Decision } from '../../server/src/mark/detect'
import { markLayerRaster, markRaster } from '../../server/src/mark/mark-raster'
import { messageFor } from '../../server/src/mark/message'
import { DEFAULT_PARAMS } from '../../server/src/mark/params'
import { decodePng } from '../../server/src/mark/png'
import type { Bitmap, Chip } from '../../server/src/mark/types'
import { tauFor } from '../../server/src/detect/tau'
import { jpeg, scale as downscale } from '../stream-s0/mark/attacks'
import { blankSheet, grid500x100, millionRows, MILLION_ROWS_BAND, plainSeven, savedFormula, cfToday, unfilledShapeOverText, rotatedBox, unresolvedAnchor } from '../../server/test/sheet-fixtures'

const ROOT = resolve(import.meta.dir, '../..')
const argv = process.argv.slice(2)
const QUICK = argv.includes('--quick')
const argOf = (name: string) => { const i = argv.indexOf(name); return i >= 0 ? argv[i + 1] : undefined }
const OUT = argOf('--out') ?? (QUICK ? undefined : resolve(ROOT, 'docs/validation/streaming-s4a-feasibility.json'))
const COMPARE = argOf('--compare')

const T = TILE_DEVICE_PX
const MIB = 1024 * 1024
const sha = (b: Uint8Array) => createHash('sha256').update(b).digest('hex')
const key = (s: string) => new Uint8Array(createHash('sha256').update(s).digest())
const MASTER = key('stream-s4a-task1-feasibility-master')
const K_MSG = key('stream-s4a-task1-feasibility-message')
const DOC_REF = 'doc:stream-s4a-task1-grid500x100'
const MARK = markOptions({ kind: 'fixed' })!
/** F26: Task 7 exports this as `tileBasisLabel` (server/src/sheet/tile.ts); inlined for the spike. */
const tileBasisLabel = (sheetId: string, zoom: number, dpr: number, col: number, row: number) => `sheet:${sheetId}@${String(zoom)}x${dpr}:${col},${row}`

const backend = serverBackend()
installCanvasBackend(backend)

const round = (v: number, d = 2) => Math.round(v * 10 ** d) / 10 ** d
const pct = (sorted: number[], p: number) => sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(p * sorted.length) - 1))]!
function stats(values: number[]) {
  const s = [...values].sort((a, b) => a - b)
  return { n: s.length, p50: round(pct(s, 0.5)), p95: round(pct(s, 0.95)), max: round(s[s.length - 1]!) }
}
const gc = () => { Bun.gc(true); Bun.gc(true) }
const heap = () => { gc(); const m = process.memoryUsage(); return { heapUsed: m.heapUsed, rss: m.rss } }

// ───────────────────────────── the pinned open (F1, R2) ─────────────────────────────
interface Opened { doc: XlsxDocument; sheet: XlsxSheet; geometry: SheetGeometryResult; lease: SheetAssetsLease; cache: LayoutCache; times: Record<string, number> }
async function openSheet(bytes: Uint8Array, valuePolicy: 'cached' | 'engine'): Promise<Opened> {
  const times: Record<string, number> = {}
  let t = performance.now()
  const admitted = admit(bytes)
  if (!admitted.ok || admitted.format !== 'xlsx') throw new Error('fixture is not an admitted xlsx')
  const original = await OfficePackage.load(bytes)
  const projected = await projectPackage(original, 'xlsx', 'visible-only')
  times.project = performance.now() - t; t = performance.now()
  const doc = await detectAndParse(projected.pkg, (pkg) => parseXlsx(pkg, { valuePolicy })) as XlsxDocument
  times.parse = performance.now() - t; t = performance.now()
  const sheet = doc.sheets[0]!
  const geometry = buildSheetGeometry(sheet)
  times.geometry = performance.now() - t; t = performance.now()
  const lease = await prepareSheetAssets(doc, { sheetIndex: 0 })
  times.assets = performance.now() - t
  times.total = times.project + times.parse + times.geometry + times.assets
  return { doc, sheet, geometry, lease, cache: new LayoutCache(), times }
}

// ───────────────────────────── one tile ─────────────────────────────
type Ctx2D = CanvasRenderingContext2D
interface Painted { canvas: { toBuffer(m: 'image/png'): Buffer }; ctx: Ctx2D; result: ViewportPaintResult }
/** Paint one 512² tile exactly as SheetTileRenderer.tile() does (sheetTiles.ts:340–360), minus the client budgets. */
function paintTile(o: Opened, col: number, row: number, zoom: number, dpr: number): Painted {
  const canvas = backend.createCanvas(T, T)
  const ctx = canvas.getContext('2d') as unknown as Ctx2D
  const box = tileLogicalBox(col, row, zoom, dpr)
  const s = zoom * dpr
  ctx.setTransform(s, 0, 0, s, 0, 0)
  const result = renderSheetViewport(o.sheet, ctx, o.geometry, { box, zoom, deviceScale: dpr }, o.lease.assets, { cache: o.cache })
  return { canvas: canvas as unknown as Painted['canvas'], ctx, result }
}
/** Spike approximation of the visible mark: laid out per tile over the tile's CSS box at the tile's scale
 * (the product lays it out in global sheet device space so it lines up across seams; the stamp count and
 * font size per tile are the same order). */
function stampMark(p: Painted, zoom: number, dpr: number) {
  const s = zoom * dpr
  p.ctx.setTransform(s, 0, 0, s, 0, 0)
  paintMark(p.ctx, { widthPx: T / s, heightPx: T / s }, MARK)
}
const encode = (p: Painted) => criticalChunksOnly(new Uint8Array(p.canvas.toBuffer('image/png')))
/** R13 body under an encodeUnit header, gzip level 6 as ws.ts frames, fragmented at 1 MiB. */
function frameTile(png: Uint8Array, zoomStep: number, dpr: number, col: number, row: number): { fragments: number; bytes: number } {
  const head = new Uint8Array(16), v = new DataView(head.buffer)
  v.setUint32(0, 1, true); v.setUint8(4, zoomStep); v.setUint8(5, dpr); v.setUint16(6, 0, true); v.setUint32(8, col, true); v.setUint32(12, row, true)
  const raster = encodeRasterBody(T, T, png)
  const body = new Uint8Array(head.length + raster.length); body.set(head); body.set(raster, head.length)
  const payload = encodeUnit({ unitId: 0, discriminator: 3, body })
  const z = new Uint8Array(gzipSync(payload, { level: 6 }))
  const msg = encodeMessage({ ver: 1, kind: Kind.XLSX_TILE, reqId: 1, payload, compress: () => z, fragmentBytes: 1 << 20 })
  return { fragments: msg.length, bytes: msg.reduce((n, f) => n + f.length, 0) }
}
const bitmapOf = (png: Uint8Array): Bitmap => { const d = decodePng(png); return { width: d.w, height: d.h, data: d.rgba } }

/** Used-range tile grid at a scale (F8: the geometry extent is the whole 1M × 16k domain). */
function usedTiles(o: Opened, lastRow: number, lastCol: number, zoom: number, dpr: number) {
  if (o.geometry.kind !== 'ready') throw new Error('geometry not ready')
  const m = o.geometry.metrics, s = zoom * dpr
  const w = m.cols.start(lastCol) + m.cols.size(lastCol), h = m.rows.start(lastRow) + m.rows.size(lastRow)
  return { cols: Math.ceil((w * s) / T), rows: Math.ceil((h * s) / T), usedCss: { width: w, height: h } }
}

// ───────────────────────────── fixture sanity (F6) ─────────────────────────────
async function fixtureSanity() {
  const out: Record<string, unknown> = {}
  const check = (name: string, ok: boolean, detail: unknown) => { out[name] = { ok, detail }; if (!ok) console.error(`[t1] fixture ${name} FAILED: ${JSON.stringify(detail)}`) }
  const value = (o: Opened, r: number, c: number) => o.sheet.rows.find((x) => x.index === r)?.cells.find((x) => x.col === c)
  {
    const o = await openSheet(await grid500x100(), 'cached')
    check('grid500x100', o.sheet.rows.length === 500 && o.sheet.rows.every((r) => r.cells.length === 100) && value(o, 0, 0)?.value === 'R1C1' && value(o, 499, 98)?.value === 500099 && value(o, 499, 99)?.value === 'R500C100', { rows: o.sheet.rows.length, a1: value(o, 0, 0)?.value, cu500: value(o, 499, 98)?.value, cv500: value(o, 499, 99)?.value })
    o.lease.dispose()
  }
  {
    const o = await openSheet(await millionRows(), 'cached')
    const m = o.geometry.kind === 'ready' ? o.geometry.metrics : undefined
    const bandH = m ? m.rows.size(MILLION_ROWS_BAND.first - 1) : NaN, plainH = m ? m.rows.size(0) : NaN
    check('millionRows', o.sheet.rows.length === 2001 && value(o, 1_048_575, 0)?.value === 'LAST-ROW' && bandH === 40 && plainH === 20, { rows: o.sheet.rows.length, last: value(o, 1_048_575, 0)?.value, bandRowCss: bandH, plainRowCss: plainH })
    o.lease.dispose()
  }
  {
    const o = await openSheet(await blankSheet(), 'cached')
    check('blankSheet', o.sheet.rows.length === 0, { rows: o.sheet.rows.length })
    o.lease.dispose()
  }
  {
    const o = await openSheet(await cfToday(), 'cached')
    const cf = o.sheet.conditionalFormatting as unknown
    const rules = JSON.stringify(cf ?? null)
    check('cfToday', rules.includes('timePeriod') && rules.includes('today') && value(o, 4, 0)?.value === 46349, { hasRule: rules.includes('timePeriod'), a5: value(o, 4, 0)?.value, a5Style: value(o, 4, 0)?.styleIndex })
    o.lease.dispose()
  }
  for (const [name, make, text] of [['unfilledShapeOverText', unfilledShapeOverText, 'COVERED-B2'], ['rotatedBox', rotatedBox, 'ROT-1-1']] as const) {
    const o = await openSheet(await make(), 'cached')
    const drawings = o.sheet.drawings ?? []
    const j = JSON.stringify(drawings, (_k, v) => (typeof v === 'bigint' ? String(v) : v))
    const cellText = name === 'unfilledShapeOverText' ? value(o, 1, 1)?.value : value(o, 0, 0)?.value
    const detail = { drawings: drawings.length, anchors: o.sheet.drawingAnchors?.records.length, unresolved: o.sheet.drawingAnchors?.unresolved, cell: cellText, rot90: j.includes('"rotationDeg":90'), unfilled: j.includes('"fill":{"kind":"none"}'), transformInvalid: j.includes('"transformValid":false') }
    check(name, drawings.length === 1 && o.sheet.drawingAnchors?.unresolved === 0 && cellText === text && !detail.transformInvalid && (name === 'rotatedBox' ? detail.rot90 : detail.unfilled), detail)
    o.lease.dispose()
  }
  {
    const o = await openSheet(await unresolvedAnchor(), 'cached')
    check('unresolvedAnchor', (o.sheet.drawingAnchors?.unresolved ?? 0) > 0, { unresolved: o.sheet.drawingAnchors?.unresolved, records: o.sheet.drawingAnchors?.records.length })
    o.lease.dispose()
  }
  // F1: the pinned cached policy paints the saved 7; the engine recomputes 2.
  {
    const tiles: Record<string, string> = {}, values: Record<string, unknown> = {}
    for (const [name, make, policy] of [['savedFormula.cached', savedFormula, 'cached'], ['savedFormula.engine', savedFormula, 'engine'], ['plainSeven.cached', plainSeven, 'cached']] as const) {
      const o = await openSheet(await make(), policy)
      values[name] = value(o, 0, 0)?.value
      tiles[name] = sha(encode(paintTile(o, 0, 0, 1, 1)))
      o.lease.dispose()
    }
    check('savedFormula', values['savedFormula.cached'] === 7 && tiles['savedFormula.cached'] === tiles['plainSeven.cached'], { values, cachedTileEqualsPlainSeven: tiles['savedFormula.cached'] === tiles['plainSeven.cached'], engineTileEqualsPlainSeven: tiles['savedFormula.engine'] === tiles['plainSeven.cached'] })
  }
  return out
}

// ───────────────────────────── (d) million rows ─────────────────────────────
async function legD() {
  const bytes = await millionRows()
  const result: Record<string, unknown> = { fixture: { bytes: bytes.length, sha256: sha(bytes) } }
  const measureOpen = async (policy: 'cached' | 'engine') => {
    const before = heap()
    const t0 = performance.now()
    const o = await openSheet(bytes, policy)
    const wallMs = performance.now() - t0
    const after = heap()
    return { o, open: { policy, wallMs: round(wallMs), stagesMs: Object.fromEntries(Object.entries(o.times).map(([k, v]) => [k, round(v)])), heapGrowthMiB: round((after.heapUsed - before.heapUsed) / MIB), rssGrowthMiB: round((after.rss - before.rss) / MIB) } }
  }
  const cached = await measureOpen('cached')
  // The last row's tile, located from the metrics (the 30 pt band shifts it).
  const o = cached.o
  if (o.geometry.kind !== 'ready') throw new Error('million rows: geometry not ready')
  const y = o.geometry.metrics.rows.start(1_048_575)
  const tileRow = Math.floor(y / T), t0 = performance.now()
  const p = paintTile(o, 0, tileRow, 1, 1)
  const renderMs = performance.now() - t0
  const painted = p.result.records.some((r) => r.source.row === 1_048_575 && r.source.col === 0)
  const lastTile = { rowY: y, tileRow, renderMs: round(renderMs), kind: p.result.kind, records: p.result.records.length, lastRowRecordPainted: painted }
  result.cached = { ...cached.open, lastRowTile: lastTile }
  o.lease.dispose()
  cached.o = undefined as never
  const engine = await measureOpen('engine')
  result.engine = engine.open
  engine.o.lease.dispose()
  const open = cached.open
  const green = open.wallMs <= 10_000 && open.heapGrowthMiB <= 256 && painted
  result.gate = { openMaxMs: 10_000, heapMaxMiB: 256, verdict: green ? 'GREEN' : 'RED' }
  return result
}

// ───────────────────────────── (a) headless paint ─────────────────────────────
async function legA(grid: Opened, blank: Opened) {
  const run = () => { const o = { ...grid, cache: new LayoutCache() }; const p = paintTile(o, 0, 0, 1, 1); return { png: encode(p), kind: p.result.kind, records: p.result.records.length } }
  const r1 = run(), r2 = run()
  const b = encode(paintTile({ ...blank, cache: new LayoutCache() }, 0, 0, 1, 1))
  const d = decodePng(r1.png)
  const out = {
    viewportPaintKind: r1.kind, records: r1.records, png: { w: d.w, h: d.h, bytes: r1.png.length, sha256: sha(r1.png) }, blankTileSha256: sha(b),
    differsFromBlank: sha(b) !== sha(r1.png), twoRunsByteEqual: sha(r1.png) === sha(r2.png),
    valuePolicy: grid.doc.valuePolicy ?? '(unset: engine)',
  }
  const green = d.w === T && d.h === T && out.differsFromBlank && out.twoRunsByteEqual && r1.kind === 'detail'
  return { ...out, verdict: green ? 'GREEN' : 'RED' }
}

// ───────────────────────────── (b) latency ─────────────────────────────
function codes64(): Int8Array[] {
  const out: Int8Array[] = [], seen = new Set<string>()
  for (let m = 0; m < 64; m++) {
    const msg = messageFor(K_MSG, 1, (m * 104_729 + 17) & 0xffffff, (m * 31 + 7) & 0xff)
    const id = msg.join('')
    if (seen.has(id)) throw new Error('code collision'); seen.add(id)
    out.push(extendedCodeword(msg))
  }
  return out
}

function legB(grid: Opened, code: Int8Array) {
  const scales: Record<string, unknown> = {}
  const zoom = 1, zoomStep = SHEET_ZOOM_STEPS.indexOf(1)
  for (const dpr of [1, 2]) {
    const o = { ...grid, cache: new LayoutCache() }
    const grid2 = usedTiles(o, 499, 99, zoom, dpr)
    const coords: Array<[number, number]> = []
    for (let r = 0; r < grid2.rows; r++) for (let c = 0; c < grid2.cols; c++) coords.push([c, r])
    const list = QUICK ? coords.filter((_, i) => i % Math.ceil(coords.length / 40) === 0) : coords
    const st: Record<string, number[]> = { render: [], mark: [], encode: [], layer1: [], frame: [], total: [] }
    const kinds: Record<string, number> = {}, unmarked: Record<string, number> = {}
    let pngBytes = 0, frameBytes = 0, carriersMin = Infinity
    const load0 = loadavg()[0]!
    for (const [c, r] of list) {
      let t = performance.now(); const t0 = t
      const p = paintTile(o, c, r, zoom, dpr)
      st.render!.push(performance.now() - t); t = performance.now()
      stampMark(p, zoom, dpr)
      st.mark!.push(performance.now() - t); t = performance.now()
      const png = encode(p)
      st.encode!.push(performance.now() - t); t = performance.now()
      const marked = markRaster({ png, w: T, h: T, kBasis: deriveBasisKey(MASTER, DOC_REF, tileBasisLabel('0', zoom, dpr, c, r)), code })
      st.layer1!.push(performance.now() - t); t = performance.now()
      const f = frameTile(marked.png, zoomStep, dpr, c, r)
      st.frame!.push(performance.now() - t)
      st.total!.push(performance.now() - t0)
      kinds[p.result.kind] = (kinds[p.result.kind] ?? 0) + 1
      if ('unmarked' in marked) unmarked[marked.unmarked] = (unmarked[marked.unmarked] ?? 0) + 1
      else carriersMin = Math.min(carriersMin, marked.carriers)
      pngBytes += png.length; frameBytes += f.bytes
    }
    const total = stats(st.total!)
    scales[`zoom1-dpr${dpr}`] = {
      tileGrid: { cols: grid2.cols, rows: grid2.rows, tiles: coords.length, measured: list.length, usedCss: grid2.usedCss },
      stagesMs: Object.fromEntries(Object.entries(st).map(([k, v]) => [k, stats(v)])),
      kinds, unmarked, minCarriers: carriersMin === Infinity ? 0 : carriersMin,
      meanPngBytes: Math.round(pngBytes / list.length), meanFramedBytes: Math.round(frameBytes / list.length),
      loadavg1: { start: round(load0), end: round(loadavg()[0]!) },
      verdict: total.p95 <= 100 ? 'GREEN' : total.p95 <= 120 ? 'AMBER' : 'RED',
    }
  }
  return scales
}

// ───────────────────────────── (c) Layer-1 on tiles ─────────────────────────────
const CAPTURE = { x: 40, y: 40, w: 900, h: 900 }
function compose(tiles: Bitmap[][], cols: number, rows: number): Bitmap {
  const W = cols * T, H = rows * T, data = new Uint8ClampedArray(W * H * 4)
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
    const b = tiles[r]![c]!
    for (let y = 0; y < T; y++) data.set(b.data.subarray(y * T * 4, (y + 1) * T * 4), ((r * T + y) * W + c * T) * 4)
  }
  return { width: W, height: H, data }
}
function cropBitmap(b: Bitmap, x: number, y: number, w: number, h: number): Bitmap {
  const data = new Uint8ClampedArray(w * h * 4)
  for (let row = 0; row < h; row++) data.set(b.data.subarray(((y + row) * b.width + x) * 4, ((y + row) * b.width + x + w) * 4), row * w * 4)
  return { width: w, height: h, data }
}
type Klass = 'lossless' | 'jpeg-q70' | 'downscale-50'
const KLASSES: Klass[] = ['lossless', 'jpeg-q70', 'downscale-50']
async function attack(capture: Bitmap, k: Klass): Promise<Bitmap> {
  if (k === 'lossless') return capture
  if (k === 'jpeg-q70') return jpeg(capture, 70)
  return downscale(capture, 0.5)
}
interface Tally { attributed: number; wrong: number; none: number; collusion: number; flat: number; geometry: number; zTrue: number[]; decisions: string[] }
const tally = (): Tally => ({ attributed: 0, wrong: 0, none: 0, collusion: 0, flat: 0, geometry: 0, zTrue: [], decisions: [] })
function record(t: Tally, d: Decision, truth: number, zTrue: number | undefined) {
  if (d.kind === 'attributed') { if (d.index === truth) t.attributed++; else t.wrong++ }
  else if (d.kind === 'none') t.none++
  else if (d.kind === 'collusion-suspected') { t.collusion++; if (!d.indices.includes(truth)) t.wrong++ }
  else if (d.reason === 'flat') t.flat++
  else t.geometry++
  if (zTrue !== undefined && Number.isFinite(zTrue)) t.zTrue.push(round(zTrue, 3))
  t.decisions.push(d.kind === 'attributed' ? `attributed:${d.index}` : d.kind === 'abstain' ? `abstain:${d.reason}` : d.kind)
}
const summarise = (t: Tally, n: number) => ({ n, attributed: t.attributed, rate: round((100 * t.attributed) / n, 1), wrong: t.wrong, none: t.none, collusion: t.collusion, abstainFlat: t.flat, abstainGeometry: t.geometry, zTrue: t.zTrue.length ? stats(t.zTrue) : null, decisionsSha256: sha(new TextEncoder().encode(t.decisions.join('\n'))) })

async function legC(grid: Opened, blank: Opened, codes: Int8Array[]) {
  const zoom = 1, dpr = 1
  const N = QUICK ? 6 : 30
  const tau = tauFor(codes.length)
  const matrix = prepareCandidates(codes)
  const g = usedTiles(grid, 499, 99, zoom, dpr)
  const canon = new Map<string, Uint8Array>()
  const canonical = (o: Opened, tag: string, c: number, r: number) => {
    const k = `${tag}:${c},${r}`
    let png = canon.get(k)
    if (!png) { const p = paintTile(o, c, r, zoom, dpr); stampMark(p, zoom, dpr); png = encode(p); canon.set(k, png) }
    return png
  }
  // Windows and sessions: deterministic spread over the used grid and the 64 codes.
  const windows = Array.from({ length: N }, (_, i) => ({ c0: (i * 5) % (g.cols - 1), r0: (i * 7) % (g.rows - 1), session: (i * 37 + 3) % codes.length }))
  const tileChips = (c: number, r: number): Chip[] => basis(deriveBasisKey(MASTER, DOC_REF, tileBasisLabel('0', zoom, dpr, c, r)), T, T, DEFAULT_PARAMS)
  // The sheet-wide alternative: one basis over the used tile grid (the full 1M × 16k domain cannot be keyed).
  const sheetW = g.cols * T, sheetH = g.rows * T
  const kSheet = deriveBasisKey(MASTER, DOC_REF, `sheet:0@${zoom}x${dpr}`)
  const sheetChips = basis(kSheet, sheetW, sheetH, DEFAULT_PARAMS)
  let fullDomain: string
  try { basis(kSheet, 16_384 * 64, 1_048_576 * 20, DEFAULT_PARAMS); fullDomain = 'ok' } catch (e) { fullDomain = (e as Error).message }

  const perTile = Object.fromEntries(KLASSES.map((k) => [k, tally()])) as Record<Klass, Tally>
  const crop = Object.fromEntries(KLASSES.map((k) => [k, tally()])) as Record<Klass, Tally>
  const blankT = tally()
  const carriersPerTileCrop: number[] = [], chipsInWindow: number[] = []
  const decideOne = (capture: Bitmap, reference: Bitmap, chips: Chip[], cands: typeof matrix, truth: number) => {
    const s = scoreCapture(capture, reference, chips, cands, DEFAULT_PARAMS)
    return { d: decideScores(s, tau), zTrue: s.scores.length ? s.scores[truth] : undefined }
  }
  for (const w of windows) {
    const code = codes[w.session]!
    const at = (o: Opened, tag: string) => [0, 1].map((dr) => [0, 1].map((dc) => ({ c: w.c0 + dc, r: w.r0 + dr, png: canonical(o, tag, w.c0 + dc, w.r0 + dr) })))
    // Per-tile basis (the default): each tile marked on its own basis; the window's chips are each tile's
    // full chip list offset by its origin, in whole blocks (scoring indexes code[n % 512]).
    const cells = at(grid, 'grid')
    const reference = compose(cells.map((row) => row.map((t) => bitmapOf(t.png))), 2, 2)
    const marked = compose(cells.map((row) => row.map((t) => {
      const m = markRaster({ png: t.png, w: T, h: T, kBasis: deriveBasisKey(MASTER, DOC_REF, tileBasisLabel('0', zoom, dpr, t.c, t.r)), code })
      return bitmapOf(m.png)
    })), 2, 2)
    const chips: Chip[] = []
    for (let dr = 0; dr < 2; dr++) for (let dc = 0; dc < 2; dc++) for (const ch of tileChips(w.c0 + dc, w.r0 + dr)) chips.push({ ...ch, x: ch.x + dc * T, y: ch.y + dr * T })
    const capture = cropBitmap(marked, CAPTURE.x, CAPTURE.y, CAPTURE.w, CAPTURE.h)
    for (const k of KLASSES) { const r = decideOne(await attack(capture, k), reference, chips, matrix, w.session); record(perTile[k], r.d, w.session, r.zTrue) }

    // Sheet-wide basis cropped per tile (markLayerRaster), same window and captures.
    const markedCrop = compose(cells.map((row) => row.map((t) => {
      const m = markLayerRaster({ png: t.png, w: T, h: T, kBasis: kSheet, code, crop: { x: t.c * T, y: t.r * T, pageW: sheetW, pageH: sheetH } })
      carriersPerTileCrop.push('unmarked' in m ? 0 : m.carriers)
      return bitmapOf(m.png)
    })), 2, 2)
    // The window's chips keep their sheet index n (code[n % 512]); the candidate codes are re-indexed to the
    // kept chips (remapped[j] = code[n_j % 512]) so scoring's j % 512 reads the right symbol.
    const ox = w.c0 * T, oy = w.r0 * T
    const kept: Array<{ chip: Chip; n: number }> = []
    sheetChips.forEach((ch, n) => { if (ch.x >= ox && ch.y >= oy && ch.x + ch.cell <= ox + 2 * T && ch.y + ch.cell <= oy + 2 * T) kept.push({ chip: { ...ch, x: ch.x - ox, y: ch.y - oy }, n }) })
    chipsInWindow.push(kept.length)
    const remapped = prepareCandidates(codes.map((cw) => { const out = new Int8Array(512).fill(1); kept.forEach((k, j) => { out[j] = cw[k.n % 512]! }); return out }))
    const captureCrop = cropBitmap(markedCrop, CAPTURE.x, CAPTURE.y, CAPTURE.w, CAPTURE.h)
    for (const k of KLASSES) {
      if (kept.length < 2) { record(crop[k], { kind: 'abstain', reason: 'flat' }, w.session, undefined); continue }
      const r = decideOne(await attack(captureCrop, k), reference, kept.map((k2) => k2.chip), remapped, w.session); record(crop[k], r.d, w.session, r.zTrue)
    }

    // Blank pane: the same window, session and per-tile marking on blankSheet, captured losslessly.
    const bcells = at(blank, 'blank')
    const bref = compose(bcells.map((row) => row.map((t) => bitmapOf(t.png))), 2, 2)
    const bmarked = compose(bcells.map((row) => row.map((t) => bitmapOf(markRaster({ png: t.png, w: T, h: T, kBasis: deriveBasisKey(MASTER, DOC_REF, tileBasisLabel('0', zoom, dpr, t.c, t.r)), code }).png))), 2, 2)
    const br = decideOne(cropBitmap(bmarked, CAPTURE.x, CAPTURE.y, CAPTURE.w, CAPTURE.h), bref, chips, matrix, w.session)
    record(blankT, br.d, w.session, br.zTrue)
  }
  const pt = Object.fromEntries(KLASSES.map((k) => [k, summarise(perTile[k], N)])) as Record<Klass, ReturnType<typeof summarise>>
  const cr = Object.fromEntries(KLASSES.map((k) => [k, summarise(crop[k], N)])) as Record<Klass, ReturnType<typeof summarise>>
  const bl = summarise(blankT, N)
  const wrong = KLASSES.reduce((n, k) => n + pt[k].wrong, 0)
  const verdicts = {
    lossless: pt.lossless.rate >= 90, q70: pt['jpeg-q70'].rate >= 75, downscale: pt['downscale-50'].rate >= 75,
    zeroWrong: wrong === 0,
    // Owner ruling 2026-11-23: a marked blank pane never names a wrong session (was: abstains `flat` N/N; blank tiles stay marked).
    blankNeverWrong: bl.wrong === 0,
  }
  const green = Object.values(verdicts).every(Boolean)
  const basisChoice = cr.lossless.rate >= pt.lossless.rate + 5 ? 'crop' : 'per-tile'
  return {
    tau, candidates: codes.length, capture: CAPTURE, captures: N, windows,
    perTile: pt, cropAlternative: { ...cr, sheetBasis: { w: sheetW, h: sheetH, chips: sheetChips.length, fullDomain }, chipsInWindow: stats(chipsInWindow), carriersPerTile: stats(carriersPerTileCrop) },
    blankPane: bl, verdicts, verdict: green ? 'GREEN' : 'RED', basisChoice,
  }
}

// ───────────────────────────── main ─────────────────────────────
async function main(): Promise<number> {
  const started = new Date().toISOString(), load = loadavg()
  console.log(`[t1] ${QUICK ? 'quick ' : ''}run; bun ${Bun.version}; loadavg ${load.map((v) => v.toFixed(2)).join(' ')}`)
  const d = await legD()
  console.log(`[t1] (d) ${JSON.stringify(d)}`)
  const sanity = await fixtureSanity()
  const gridBytes = await grid500x100(), blankBytes = await blankSheet()
  const grid = await openSheet(gridBytes, 'cached'), blank = await openSheet(blankBytes, 'cached')
  const a = await legA(grid, blank)
  console.log(`[t1] (a) ${JSON.stringify(a)}`)
  const codes = codes64()
  const b = legB(grid, codes[0]!)
  for (const [k, v] of Object.entries(b)) console.log(`[t1] (b) ${k} ${JSON.stringify((v as { stagesMs: unknown }).stagesMs)} ${(v as { verdict: string }).verdict}`)
  const c = await legC(grid, blank, codes)
  console.log(`[t1] (c) perTile ${JSON.stringify(Object.fromEntries(KLASSES.map((k) => [k, c.perTile[k].rate])))} crop ${JSON.stringify(Object.fromEntries(KLASSES.map((k) => [k, c.cropAlternative[k].rate])))} blank ${JSON.stringify(c.blankPane)} ${c.verdict}`)
  const out: Record<string, unknown> = {
    task: 'stream-s4a Task 1 feasibility gate', started, quick: QUICK, bun: Bun.version, platform: `${process.platform}-${process.arch}`, loadavgStart: load.map((v) => round(v)), loadavgEnd: loadavg().map((v) => round(v)),
    parsePath: "projectPackage(original, 'xlsx', 'visible-only') → detectAndParse(projected.pkg, pkg => parseXlsx(pkg, { valuePolicy: 'cached' })) → buildSheetGeometry(sheet) → prepareSheetAssets(doc, { sheetIndex: 0 })",
    basisLabel: 'sheet:<id>@<zoom>x<dpr>:<c>,<r> (inlined; Task 7 exports tileBasisLabel)',
    fixtures: { sanity, grid500x100: sha(gridBytes), blankSheet: sha(blankBytes) },
    a, b, c, d,
  }
  if (COMPARE) {
    const prev = JSON.parse(readFileSync(COMPARE, 'utf8')) as typeof out & { b: Record<string, { stagesMs: { total: { p95: number } } }>; a: { png: { sha256: string } }; c: unknown }
    const strip = (x: unknown) => JSON.stringify(x)
    const p95 = Object.fromEntries(Object.keys(b).map((k) => {
      const now = (b as Record<string, { stagesMs: { total: { p95: number } } }>)[k]!.stagesMs.total.p95, then = prev.b[k]!.stagesMs.total.p95
      return [k, { previous: then, now, deltaPct: round((100 * (now - then)) / then, 1), within10pct: Math.abs(now - then) <= 0.1 * then }]
    }))
    out.reproducibility = { against: COMPARE, a: strip(prev.a) === strip(a), c: strip(prev.c) === strip(c), bP95: p95 }
    console.log(`[t1] reproducibility ${JSON.stringify(out.reproducibility)}`)
  }
  const verdicts = { a: a.verdict, b: Object.fromEntries(Object.entries(b).map(([k, v]) => [k, (v as { verdict: string }).verdict])), c: c.verdict, d: (d.gate as { verdict: string }).verdict, fixtures: Object.values(sanity).every((x) => (x as { ok: boolean }).ok) ? 'OK' : 'FAILED' }
  out.verdicts = verdicts
  console.log(`[t1] verdicts ${JSON.stringify(verdicts)} basis ${c.basisChoice}`)
  if (OUT) { writeFileSync(OUT, `${JSON.stringify(out, null, 2)}\n`); console.log(`[t1] wrote ${OUT}`) }
  grid.lease.dispose(); blank.lease.dispose()
  return 0
}

if (import.meta.main) process.exit(await main())
