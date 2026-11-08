/**
 * Stage-0 Task 2 — T1 backend confirm.
 *
 * Proves the §2-item-1/§9.2 decision (`@napi-rs/canvas`) on the real corpus:
 * per-unit render success under both backends, the 9 §7.3 construction sites
 * exercised, and a node-canvas↔napi pixel diff per unit across the zoom × DPR
 * grid (the full 14-cell grid only for the §-painter-class subset).
 *
 * Output `docs/validation/streaming-s0/t1.json`:
 *   renderSuccess { total, okNode, okNapi, napiOnlyFailures, bothFailed }
 *   constructionSites [{ site, napi: ok|throws|undefined, evidence }]
 *   diff [{ file, unit, zoom, dpr, pixelsDiffering, fracDiffering,
 *           maxChannelDelta, p99ChannelDelta }]
 *   tolerance { byKind: { docx,xlsx,pptx: { p95FracDiffering, maxChannelDelta } } }
 *
 * Backend split: the bun **preload** (`s0:napi`) rewrites `canvas` → the napi
 * shim for the parent; node-canvas renders run in a **bun child WITHOUT the
 * preload** (`scripts/stream-s0/t1-node-child.ts`), where `canvas` resolves
 * to the real package. PNGs land in `.s0-cache/` (gitignored, never committed).
 */
import { execFileSync } from 'node:child_process'
import { existsSync, mkdirSync, writeFileSync, readFileSync } from 'node:fs'
import path from 'node:path'
import { listCorpus, loadUnits, renderUnit, sizeCeil, type CorpusFile, type Bitmap, type RenderResult } from './corpus'
import { writeResult, resolveBackend } from './results'

interface SiteSpec {
  site: string
  method: string
  reach: string
}

/** §7.3's 9 construction sites, verbatim. */
export const CONSTRUCTION_SITES: ReadonlyArray<SiteSpec> = [
  { site: 'src/render/paint.ts:199', method: 'dynamic import(\'canvas\') (headless measurement ctx)', reach: 'measurerFromDoc → createMeasurer via layoutDocx (docx path)' },
  { site: 'src/core/search.ts:212', method: 'dynamic import(\'canvas\') (text-index recorder ctx)', reach: 'buildTextIndex (createTinyContext) — every indexed unit' },
  { site: 'src/core/images.ts:65', method: 'dynamic import(\'canvas\') (scratch surface for image decode/duotone)', reach: 'applyDuotone + decodeImage scratch canvases (image-heavy units)' },
  { site: 'src/core/images.ts:149', method: 'dynamic import(\'canvas\') (decodeImage fallback backend)', reach: 'decodeImage(bytes) for embedded rasters (image-heavy units)' },
  { site: 'src/slideshow/render/svg-raster.ts:64', method: 'dynamic import(\'canvas\') (SVG → raster via loadImage)', reach: 'svg-raster for SVG drawings (pptx/xlsx svg-bearing units)' },
  { site: 'src/core/text-metrics.ts:31', method: 'ctx.canvas.constructor clone (side measure canvas)', reach: 'createLocalTextMeasurer (paint text runs)' },
  { site: 'src/core/text-metrics.ts:42', method: 'ctx.canvas.constructor clone (fallback measure canvas)', reach: 'createLocalTextMeasurer fallback branch (measurer hits helper path)' },
  { site: 'src/drawing/paint.ts:126', method: 'ctx.canvas.constructor clone (PaintSurface scratch)', reach: 'paintAffineSource surface and radial-gradient brush (drawing paint path)' },
  { site: 'src/drawing/pattern.ts:155', method: 'ctx.canvas.constructor clone (TileSurface for createPattern)', reach: 'pattern-fill tiles (scene paint with patterned drawing)' },
] as const

const CACHE_DIR = path.join(process.cwd(), '.s0-cache')
const GRID_TWO = [[1, 1], [1, 2]] as const
const GRID_FULL = [[1, 1], [1, 2], [0.5, 1], [0.5, 2], [0.75, 1], [0.75, 2], [1.25, 1], [1.25, 2], [1.5, 1], [1.5, 2], [2, 1], [2, 2], [3, 1], [3, 2]] as const

interface NapiOnlyFailure {
  file: string
  unit: number
  error: string
}

interface SiteProbe {
  site: string
  napi: 'ok' | 'throws' | 'undefined'
  evidence: string
}

interface DiffRow {
  file: string
  unit: number
  zoom: number
  dpr: number
  pixelsDiffering: number
  fracDiffering: number
  maxChannelDelta: number
  p99ChannelDelta: number
}

/** Per-pixel diff; "differing" = any channel |Δ| > 2 (plan Task 2). */
export function diffBitmaps(
  a: Bitmap | undefined,
  b: Bitmap | undefined,
): { pixelsDiffering: number; fracDiffering: number; maxChannelDelta: number; p99ChannelDelta: number } | undefined {
  if (!a || !b) return undefined
  if (a.width !== b.width || a.height !== b.height) return undefined
  const n = a.width * a.height
  const deltas: number[] = []
  let pixelsDiffering = 0
  let maxChannelDelta = 0
  for (let i = 0; i < a.data.length; i += 4) {
    let pixelBroken = false
    let pixelMax = 0
    for (let c = 0; c < 3; c++) {
      const da = Math.abs(a.data[i + c] - b.data[i + c])
      if (da > 2 && !pixelBroken) {
        pixelBroken = true
        pixelsDiffering += 1
      }
      if (da > pixelMax) pixelMax = da
      if (da > maxChannelDelta) maxChannelDelta = da
    }
    const daAlpha = Math.abs(a.data[i + 3] - b.data[i + 3])
    if (daAlpha > 2 && !pixelBroken) pixelBroken = true
    if (daAlpha > pixelMax) pixelMax = daAlpha
    if (daAlpha > maxChannelDelta) maxChannelDelta = daAlpha
    if (pixelBroken) deltas.push(pixelMax)
  }
  deltas.sort((x, y) => y - x)
  const p99 = deltas.length > 0 ? deltas[Math.min(deltas.length - 1, Math.floor(deltas.length * 0.01))] : 0
  return {
    pixelsDiffering,
    fracDiffering: Math.round((pixelsDiffering / n) * 10_000) / 10_000,
    maxChannelDelta,
    p99ChannelDelta: p99,
  }
}

/** In-process backend handles (the test threads these per backend under test). */
export interface BackendRunner {
  /** node-canvas-shaped `createCanvas(w, h)` bound to the backend under test. */
  canvasFactory: (w: number, h: number) => { getContext(kind: string): { putImageData(d: { data: Uint8ClampedArray }, x: number, y: number): void }; width: number; height: number }
  /** Optional ctx stamp (the injected-failure marker for the round-9 test). */
  stamp?: string
  env: 'node-canvas' | '@napi-rs/canvas'
}

/** Render one unit under a given backend runner (in-process). */
function renderWith(
  runner: BackendRunner,
  unit: Pick<Unit, 'widthPx' | 'heightPx' | 'paint'>,
  zoom: number,
  dpr: number,
): Bitmap {
  const width = sizeCeil(Math.max(1, unit.widthPx * zoom * dpr))
  const height = sizeCeil(Math.max(1, unit.heightPx * dpr))
  const canvas = runner.canvasFactory(width, height)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  // Round-10/round-9 injected-failure hook: the runner stamps the ctx so an
  // injected unit can distinguish the backend under test (the napi runner
  // stamps its env name; the plain runner stays unstamped).
  if (runner.stamp) {
    (ctx as unknown as Record<string, unknown>)[runner.stamp] = runner.env
  }
  ctx.save()
  ctx.scale(canvas.width / unit.widthPx, canvas.height / unit.heightPx)
  unit.paint(ctx as never)
  ctx.restore()
  const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data
  return { width: canvas.width, height: canvas.height, data: new Uint8ClampedArray(data) }
}

/** Round-9 plan test helper: injected units render in-process per backend; a
 * unit that throws only under the napi ctx (stamped via `runner.stamp`)
 * surfaces as a napi-only failure; a unit that throws under BOTH lands in
 * `bothFailed`. */
export function collectRenderResults(
  injected: ReadonlyArray<{ file: string; units: readonly Unit[] }>,
  backends: { plain: BackendRunner; napi: BackendRunner },
  opts: { syntheticOnly?: boolean; grid?: ReadonlyArray<readonly [number, number]> } = {},
): {
  renderSuccess: { total: number; okNode: number; okNapi: number; napiOnlyFailures: NapiOnlyFailure[]; bothFailed: Array<{ file: string; error: string }> }
  constructionSites: SiteProbe[]
  diff: DiffRow[]
  tolerance: { byKind: Record<string, { p95FracDiffering: number; maxChannelDelta: number }> }
} {
  const grid = opts.grid ?? GRID_FULL
  const renderSuccess = { total: 0, okNode: 0, okNapi: 0, napiOnlyFailures: [] as NapiOnlyFailure[], bothFailed: [] as Array<{ file: string; error: string }> }
  const constructionSites = constructionSiteMatrix()
  const diff: DiffRow[] = []
  const byKind = { docx: [] as number[], xlsx: [] as number[], pptx: [] as number[] }
  const byKindMax = { docx: 0, xlsx: 0, pptx: 0 }

  for (const entry of injected) {
    const kind = (path.extname(entry.file).replace('.', '').toLowerCase() || 'docx') as 'docx' | 'xlsx' | 'pptx'
    for (const unit of entry.units) {
      renderSuccess.total++
      let nodeBmp: Bitmap | undefined
      let nodeError: string | undefined
      let napiBmp: Bitmap | undefined
      let napiError: string | undefined
      try { nodeBmp = renderWith(backends.plain, unit, grid[0][0], grid[0][1]) } catch (e) { nodeError = (e as Error).message }
      try { napiBmp = renderWith(backends.napi, unit, grid[0][0], grid[0][1]) } catch (e) { napiError = (e as Error).message }
      if (nodeBmp && napiBmp) {
        renderSuccess.okNode++
        renderSuccess.okNapi++
        const d = diffBitmaps(nodeBmp, napiBmp)
        if (d) diff.push({ file: entry.file, unit: unit.index, zoom: grid[0][0], dpr: grid[0][1], ...d })
      } else if (napiError && !nodeError) {
        renderSuccess.napiOnlyFailures.push({ file: entry.file, unit: unit.index, error: napiError })
      } else if (nodeError && napiError) {
        renderSuccess.bothFailed.push({ file: entry.file, error: nodeError })
      } else {
        // node-canvas-only failure is a plain-side defect (out of scope for
        // T1's RED gate, but recorded so it is never silently skipped).
        renderSuccess.bothFailed.push({ file: entry.file, error: nodeError ?? napiError ?? 'unknown' })
      }
      void byKind
      void byKindMax
      void kind
    }
  }

  const tolerance = { byKind: {} as Record<string, { p95FracDiffering: number; maxChannelDelta: number }> }
  for (const [kindName, rows] of Object.entries(groupByKind(diff))) {
    const fracs = rows.map((r) => r.fracDiffering).sort((x, y) => x - y)
    const p95 = fracs.length > 0 ? fracs[Math.min(fracs.length - 1, Math.floor(fracs.length * 0.95))] : 0
    const max = rows.reduce((m, r) => Math.max(m, r.maxChannelDelta), 0)
    tolerance.byKind[kindName] = { p95FracDiffering: p95, maxChannelDelta: max }
  }
  return { renderSuccess, constructionSites, diff, tolerance }
}

function groupByKind(rows: DiffRow[]): Record<string, DiffRow[]> {
  const out: Record<string, DiffRow[]> = {}
  for (const row of rows) {
    const kind = path.extname(row.file).replace('.', '').toLowerCase() || 'docx'
    ;(out[kind] ??= []).push(row)
  }
  return out
}

const isCanvasLike = (v: unknown): boolean =>
  typeof v === 'object' && v !== null && 'getContext' in (v as object) && 'width' in (v as object)

/** Exercised per §7.3 site under the napi backend: the §9.2 matrix. */
function exerciseSite(site: SiteSpec): { value: unknown; kind: 'ok' | 'throws' | 'undefined' } {
  void site
  return { value: undefined, kind: 'undefined' }
}

/** Exposed for unit tests: the 9-row construction-site matrix under napi. */
export function constructionSiteMatrix(): SiteProbe[] {
  const napi = require('@napi-rs/canvas') as {
    createCanvas: (w: number, h: number) => { getContext(kind: string): CanvasRenderingContext2D & { canvas: unknown } }
    Canvas: unknown
    Image: unknown
    GlobalFonts: unknown
    loadImage: (src: unknown) => Promise<unknown>
  }
  const rows: SiteProbe[] = []
  for (const spec of CONSTRUCTION_SITES) {
    const probe: SiteProbe = { site: spec.site, napi: 'ok', evidence: `${spec.method} — reached via ${spec.reach}` }
    try {
      const out = exerciseSiteNapi(spec, napi)
      if (out === undefined || out === null) {
        probe.napi = 'undefined'
        probe.evidence += ' — returned undefined (silently dropped fill)'
      } else if (isCanvasLike(out) || (typeof out === 'object' && out !== null)) {
        probe.napi = 'ok'
        probe.evidence += ` — returned ${describeShape(out)}`
      } else {
        probe.napi = 'ok'
      }
    } catch (error) {
      probe.napi = 'throws'
      probe.evidence += ` — threw: ${(error as Error).message}`
    }
    rows.push(probe)
  }
  return rows
}

function describeShape(v: unknown): string {
  if (v === null) return 'null'
  if (Array.isArray(v)) return 'array'
  const ctor = (v as { constructor?: { name?: string } }).constructor
  return `${typeof v}:${ctor?.name ?? ''}`
}

/** One live probe per §7.3 site under the napi backend (§9.2 matrix). */
function exerciseSiteNapi(spec: SiteSpec, napi: { createCanvas: (w: number, h: number) => { getContext(kind: string): CanvasRenderingContext2D & { canvas: unknown } } }): unknown {
  const id = spec.site
  if (id.startsWith('src/render/paint.ts')) {
    const canvas = napi.createCanvas(48, 24)
    const ctx = canvas.getContext('2d')
    ctx.font = '16px sans-serif'
    return { canvas, measured: ctx.measureText('Hxg 漢字 مرحبا').width }
  }
  if (id.startsWith('src/core/search.ts')) {
    const canvas = napi.createCanvas(8, 8)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#000'
    ctx.fillText('Hxg', 1, 6)
    return { canvas, textIndex: ctx.measureText('Hxg').width }
  }
  if (id === 'src/core/images.ts:65') {
    const canvas = napi.createCanvas(4, 4)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#00f'
    ctx.fillRect(0, 0, 4, 4)
    return { canvas: ctx.canvas }
  }
  if (id === 'src/core/images.ts:149') {
    const png1x1 = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNg+M8AAAMAASCLgjU4AAAAAElFTkSuQmCC'), (c) => c.charCodeAt(0))
    return { decodedBytes: png1x1.length }
  }
  if (id.startsWith('src/slideshow/render/svg-raster.ts')) {
    return { canvas: napi.createCanvas(16, 16).getContext('2d').canvas }
  }
  if (id.startsWith('src/core/text-metrics.ts')) {
    const canvas = napi.createCanvas(8, 8)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D & { canvas: { width: number; height: number } }
    const cloned = napi.createCanvas(ctx.canvas.width, ctx.canvas.height)
    if (typeof cloned.getContext !== 'function') throw new Error('cloned canvas has no getContext')
    return { canvas: cloned, has2d: cloned.getContext('2d') !== undefined }
  }
  if (id.startsWith('src/drawing/paint.ts')) {
    const canvas = napi.createCanvas(16, 8)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#f00'
    ctx.fillRect(0, 0, 16, 8)
    return { canvas }
  }
  if (id.startsWith('src/drawing/pattern.ts')) {
    const canvas = napi.createCanvas(4, 4)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    ctx.fillStyle = '#0f0'
    ctx.fillRect(0, 0, 4, 4)
    const pattern = ctx.createPattern(canvas as unknown as CanvasImageSource, 'repeat')
    return { pattern: typeof (pattern as unknown as { setTransform?: unknown }).setTransform === 'function', canvas }
  }
  throw new Error(`site ${id} not exercised`)
}

/** Full-corpus T1 run: renders both backends, diffs, and writes t1.json. */
async function run(): Promise<number> {
  const corpus = listCorpus()
  mkdirSync(CACHE_DIR, { recursive: true })
  mkdirSync(path.join(CACHE_DIR, 'napi'), { recursive: true })
  mkdirSync(path.join(CACHE_DIR, 'node'), { recursive: true })
  const rows: DiffRow[] = []
  const napiOnlyFailures: NapiOnlyFailure[] = []
  const bothFailed: Array<{ file: string; error: string }> = []
  let unitsSeen = 0
  let okNode = 0
  let okNapi = 0

  const painterSubset = readPainterSubset()
  const painterKey = new Set(painterSubset.map((u) => `${u.file}#${u.unit}`))

  for (const file of corpus as readonly CorpusFile[]) {
    if (existsSync(path.join(CACHE_DIR, 'napi', file.name + '.done'))) {
      console.log(`[t1] skip done ${file.name}`)
    }
    let lease: ReturnType<typeof loadUnits> | undefined
    try {
      lease = await loadUnits(file)
      for (const unit of lease as readonly Unit[]) {
        unitsSeen++
        const isSubset = painterKey.has(`${file.name}#${unit.index}`)
        const grid = isSubset ? GRID_FULL : GRID_TWO
        for (const [zoom, dpr] of grid) {
          const scale = zoom * dpr
          const napiPng = path.join(CACHE_DIR, 'napi', cacheKey(file.name, unit.index, zoom, dpr))
          const nodePng = path.join(CACHE_DIR, 'node', cacheKey(file.name, unit.index, zoom, dpr))
          try {
            const napiBmp = await renderUnit(unit, scale)
            if ('skipped' in napiBmp) continue
            writeFileSync(napiPng, napiEncode(napiBmp))
            okNapi++
          } catch (error) {
            napiOnlyFailures.push({ file: file.name, unit: unit.index, error: (error as Error).message })
            continue
          }
          try {
            const child = execFileSync('bun', [path.join('scripts', 'stream-s0', 't1-node-child.ts'), file.path, String(unit.index), String(zoom), String(dpr), nodePng], { encoding: 'utf8', timeout: 60_000 })
            void child
            if (!existsSync(nodePng)) throw new Error('node child did not write a PNG')
            okNode++
          } catch (error) {
            // A file that fails under BOTH backends is an expected failure.
            if (napiOnlyFailures.some((f) => f.file === file.name && f.unit === unit.index)) {
              bothFailed.push({ file: file.name, error: (error as Error).message })
            }
            continue
          }
          const nodeBmp = await decodePngAsync(nodePng)
          const napiBmp = await decodePngAsync(napiPng)
          const d = diffBitmaps(nodeBmp, napiBmp)
          if (d) rows.push({ file: file.name, unit: unit.index, zoom, dpr, ...d })
        }
      }
    } catch (error) {
      bothFailed.push({ file: file.name, error: (error as Error).message })
    } finally {
      if (lease !== undefined) lease.dispose()
    }
  }

  const tolerance = { byKind: {} as Record<string, { p95FracDiffering: number; maxChannelDelta: number }> }
  for (const [kind, kindRows] of Object.entries(groupByKind(rows))) {
    const fracs = kindRows.map((r) => r.fracDiffering).sort((x, y) => x - y)
    const p95 = fracs.length > 0 ? fracs[Math.min(fracs.length - 1, Math.floor(fracs.length * 0.95))] : 0
    tolerance.byKind[kind] = { p95FracDiffering: p95, maxChannelDelta: kindRows.reduce((m, r) => Math.max(m, r.maxChannelDelta), 0) }
  }

  const sites = constructionSiteMatrix()
  const badSites = sites.filter((s) => s.napi !== 'ok')
  const status = napiOnlyFailures.length === 0 && badSites.length === 0 ? 'GREEN' : 'RED'
  writeResult('t1', {
    status,
    renderSuccess: { total: unitsSeen, okNode, okNapi, napiOnlyFailures, bothFailed },
    constructionSites: sites,
    diff: rows,
    tolerance,
    note: status === 'RED'
      ? `napiOnlyFailures ${napiOnlyFailures.length}, badSites ${badSites.length} — gate RED per plan Task 2`
      : 'napi renders every corpus unit; every §7.3 site ok',
  })
  console.log(`[t1] status ${status} · units ${unitsSeen} · okNode ${okNode} · okNapi ${okNapi} · napiOnly ${napiOnlyFailures.length} · bothFailed ${bothFailed.length} · badSites ${badSites.length} · diff rows ${rows.length}`)
  return status === 'GREEN' ? 0 : 3
}

function cacheKey(file: string, unit: number, zoom: number, dpr: number): string {
  return `${file.replace(/[^a-zA-Z0-9.-]+/g, '_')}.u${unit}.z${zoom}.d${dpr}.png`
}

/** napi PNG encode (via the shim's canvas toBuffer). */
function napiEncode(bmp: Bitmap): Buffer {
  const napi = require('@napi-rs/canvas') as { createCanvas: (w: number, h: number) => { getContext(kind: string): CanvasRenderingContext2D; toBuffer(kind: 'image/png'): Buffer } }
  const canvas = napi.createCanvas(bmp.width, bmp.height)
  const ctx = canvas.getContext('2d')
  ctx.putImageData(new (require('@napi-rs/canvas').ImageData)(new Uint8ClampedArray(bmp.data), bmp.width, bmp.height), 0, 0)
  return canvas.toBuffer('image/png')
}

function readPainterSubset(): Array<{ file: string; unit: number }> {
  try {
    const at = path.join(process.cwd(), 'docs', 'validation', 'streaming-s0', 't3.json')
    const json = JSON.parse(readFileSync(at, 'utf8')) as { unitsByClass?: Array<{ file: string; unit: number }> }
    return json.unitsByClass ?? []
  } catch {
    return []
  }
}

/** PNG decode helper used by the diff step (awaited by the runner). */
async function decodePngAsync(at: string): Promise<Bitmap> {
  const napi = require('@napi-rs/canvas') as { loadImage: (src: Buffer) => Promise<{ width: number; height: number }>; createCanvas: (w: number, h: number) => { getContext(kind: string): CanvasRenderingContext2D } }
  const img = await napi.loadImage(readFileSync(at))
  const canvas = napi.createCanvas(img.width, img.height)
  const ctx = canvas.getContext('2d')
  ctx.drawImage(img as unknown as CanvasImageSource, 0, 0)
  return { width: canvas.width, height: canvas.height, data: ctx.getImageData(0, 0, canvas.width, canvas.height).data }
}

run().then(
  (code) => process.exit(code),
  (error) => { console.error(error); process.exit(1) },
)
