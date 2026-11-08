/**
 * Stage-0 Task 1 — corpus loading.
 *
 * Every backend-sensitive import here goes through the `canvas` specifier,
 * which `vitest.stream-s0-napi.config.ts` aliases to `napi-shim.ts`
 * (@napi-rs/canvas) while a plain run resolves the real `canvas` package.
 * Probes therefore render with the backend under test without touching `src/`.
 *
 * Font registration is NOT silently defaulted: `loadUnits(file, opts)` threads
 * `PaintOptions` through to `getPaintables`, and with no `registerFont`
 * provided, an explicit headless registrar records `font-load-failed` per
 * embedded face (so Task 4 can pass its own node-side `registerFont` adapter
 * and see which faces loaded).
 *
 * Note for later tasks: never `savePng` from `src/test/pixel-diff.ts` as a
 * Stage-0 tool — its PNG IO path is a native `require('canvas')`
 * (`src/test/pixel-diff.ts:63–64`) which neither the vitest napi alias nor
 * the bun preload can intercept (it imports the real `canvas` package), so a
 * "napi" PNG could silently be node-canvas pixels. Use `renderUnit` above.
 */
import { readFileSync, readdirSync } from 'node:fs'
import path from 'node:path'
import { createCanvas } from 'canvas'
import { loadOfficeFile } from '../../src/components/OfficeFile'
import { getPaintables, type Paintable, type PaintableArray, type PaintOptions } from '../../src/render/paint'

export type CorpusKind = 'docx' | 'xlsx' | 'pptx'

export interface CorpusFile {
  name: string
  /** Absolute path, ready for readFile. */
  path: string
  kind: CorpusKind
}

/** One paintable unit of one corpus file (plan Task 1 `Unit`). */
export interface Unit {
  index: number
  widthPx: number
  heightPx: number
  paint: (ctx: CanvasRenderingContext2D) => void
}

/** Rendered RGBA bitmap (`src/test/pixel-diff.ts` shape). */
export interface Bitmap {
  width: number
  height: number
  data: Uint8ClampedArray
}

/** Render result: a unit over 8192 px/side at the requested scale skips. */
export type RenderResult = Bitmap | { skipped: 'oversize' }

/** Consistent LSB-tolerant ceil: absorbs float error (round 10 finding 7). */
export const sizeCeil = (value: number): number => Math.ceil(value - 1e-9)

const CORPUS_DIR = path.join(process.cwd(), 'corpus')
/** Office lock files (`~$name`) are never documents. */
const KINDS: ReadonlyArray<[extension: string, kind: CorpusKind]> = [
  ['.docx', 'docx'],
  ['.xlsx', 'xlsx'],
  ['.pptx', 'pptx'],
]

function corpusKind(name: string): CorpusKind | undefined {
  const lower = name.toLowerCase()
  for (const [extension, kind] of KINDS) {
    if (lower.endsWith(extension)) return kind
  }
  return undefined
}

/** `corpus/` files by kind; an empty list when the directory is absent. */
export function listCorpus(): CorpusFile[] {
  let entries: string[]
  try {
    entries = readdirSync(CORPUS_DIR)
  } catch {
    // Directory absent: the caller decides whether to skip or fail.
    return []
  }
  return entries
    .filter((name) => !name.startsWith('~$'))
    .map((name) => ({ name, kind: corpusKind(name) }))
    .filter((entry): entry is { name: string; kind: CorpusKind } => entry.kind !== undefined)
    .sort((a, b) => a.name.localeCompare(b.name))
    .map(({ name, kind }) => ({ name, kind, path: path.join(CORPUS_DIR, name) }))
}

/**
 * All paintable units of one corpus file, with the diagnosis plumbing
 * (`PaintOptions`) from `getPaintables` and **no silent registrar fallback**:
 * with `opts.registerFont` unset, an explicit headless registrar records
 * `font-load-failed` for every embedded face instead of falling through to
 * the browser registrar.
 *
 * (round-10 loadUnits doc comment: paint closures stay live for late reads;
 * the file bytes come from `readFileSync` per call; disposal is per-unit set
 * via the returned array's non-enumerable `dispose` — a caller cannot
 * accidentally forget the writer's lease.)
 */
export async function loadUnits(file: CorpusFile, opts?: PaintOptions): Promise<HeadlessUnits> {
  const bytes = readFileSync(file.path)
  const paintOpts: PaintOptions = {
    // Explicit headless registrar: acquireFonts catches this and records a
    // `font-load-failed` diagnostic per embedded face (no silent fallback).
    registerFont: (request) => {
      throw new Error(
        `stream-s0: no registerFont provided — embedded ${request.face?.family ?? 'unknown'} cannot load headlessly`,
      )
    },
    ...(opts as Partial<PaintOptions> | undefined),
  }
  const doc = await loadOfficeFile(bytes)
  const paintables = await getPaintables(doc, paintOpts)
  // Keep the lease alive as long as any paint closure remains reachable;
  // each unit closes over this array so `dispose` releases it.
  const units: HeadlessUnits = paintables.map((paintable: Paintable, index: number) => ({
    index,
    widthPx: paintable.spec.widthPx,
    heightPx: paintable.spec.heightPx,
    paint: (ctx: CanvasRenderingContext2D) => paintable.paint(ctx),
  }))
  // The lease/plan surface is attached via non-enumerable properties so array
  // copies cannot silently lose it (round-10 finding: disposal by array
  // identity is fragile with WeakMaps).
  Object.defineProperty(units, 'dispose', {
    value: () => paintables.dispose(),
    enumerable: false,
    configurable: true,
    writable: true,
  })
  Object.defineProperty(units, 'fontDiagnostics', {
    value: paintables.fontDiagnostics,
    enumerable: false,
    configurable: true,
  })
  if (paintables.chartDiagnostics) {
    Object.defineProperty(units, 'chartDiagnostics', {
      value: paintables.chartDiagnostics,
      enumerable: false,
      configurable: true,
    })
  }
  return units
}

/**
 * Paint one unit at `scale = zoom × dpr` and read back RGBA exactly like
 * `src/test/pixel-diff.ts` but without the silent 8192 clamp: a unit above
 * 8192 px on a side at the requested scale returns `{ skipped: 'oversize' }`
 * instead of a downscaled write (Task 2 counts skips, never hides them).
 */
export async function renderUnit(unit: Unit, scale: number): Promise<RenderResult> {
  const width = sizeCeil(Math.max(1, unit.widthPx * scale))
  const height = sizeCeil(Math.max(1, unit.heightPx * scale))
  if (Math.max(width, height) > 8192) return { skipped: 'oversize' }
  const canvas = createCanvas(width, height)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.save()
  ctx.scale(canvas.width / unit.widthPx, canvas.height / unit.heightPx)
  unit.paint(ctx as never)
  ctx.restore()
  return { width: canvas.width, height: canvas.height, data: ctx.getImageData(0, 0, canvas.width, canvas.height).data }
}

interface HeadlessUnits extends Array<Unit> {
  /** Releases the writer's lease (Task 2+ ownership glue). */
  dispose(): void
  /** Font-load diagnostics from `getPaintables` (`font-load-failed` etc.). */
  fontDiagnostics: unknown[]
  /** Chart diagnostics surfaces when charts are prepared (optional). */
  chartDiagnostics?: unknown
}
