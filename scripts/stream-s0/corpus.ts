/**
 * Stage-0 Task 1 — corpus loading.
 *
 * Every backend-sensitive import here goes through the `canvas` specifier,
 * which `vitest.stream-s0-napi.config.ts` aliases to `napi-shim.ts`
 * (@napi-rs/canvas) while a plain run resolves the real `canvas` package.
 * Probes therefore render with the backend under test without touching `src/`.
 *
 * Font registration is deliberately not part of Task 1 (Task 4 owns font
 * probes): units with embedded faces acquire a backend fallback family under
 * the default `acquireFonts` path, which is enough for renderability.
 */
import { readFileSync, readdirSync, statSync } from 'node:fs'
import path from 'node:path'
import { createCanvas } from 'canvas'
import { loadOfficeFile } from '../../src/components/OfficeFile'
import { getPaintables, type Paintable, type PaintableArray } from '../../src/render/paint'

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
 * All paintable units of one corpus file.
 *
 * Implementation note: each unit is painted through the corpus-file's actual
 * bytes, loaded once per call; if repeated probes re-load the same file they
 * re-pay the parse (deterministic, kept simple — Task 2's probe wires reuse).
 * The file must still parse: an unparseable corpus file **throws** here and
 * callers (Task 2) count it as a render failure of the backend under test.
 */
export async function loadUnits(file: CorpusFile): Promise<Unit[]> {
  const bytes = readFileSync(file.path)
  const doc = await loadOfficeFile(bytes)
  const paintables = await getPaintables(doc)
  // Keep the lease alive as long as any paint closure remains reachable;
  // each unit closes over this array so `disposeUnits` can release it.
  const units: Unit[] = paintables.map((paintable: Paintable, index: number) => ({
    index,
    widthPx: paintable.spec.widthPx,
    heightPx: paintable.spec.heightPx,
    paint: (ctx: CanvasRenderingContext2D) => paintable.paint(ctx),
  }))
  paintablesLeases.set(units, paintables)
  return units
}

/**
 * Lease table: renderUnit-driven callers release a file's leases via
 * `disposeUnits` after rendering; a paint after disposal is undefined
 * behaviour, as documented by `getPaintables`.
 */
const paintablesLeases = new WeakMap<ReadonlyArray<Unit>, PaintableArray>()

/** Release the leases behind one file's units (Task 2+ ownership glue). */
export function disposeUnits(units: ReadonlyArray<Unit>): void {
  const lease = paintablesLeases.get(units)
  if (lease) {
    paintablesLeases.delete(units)
    lease.dispose()
  }
}

/**
 * Paint one unit at `scale = zoom × dpr` and read back RGBA exactly like
 * `src/test/pixel-diff.ts` but without the silent 8192 clamp: a unit above
 * 8192 px on a side at the requested scale returns `{ skipped: 'oversize' }`
 * instead of a downscaled write (Task 2 counts skips, never hides them).
 */
export async function renderUnit(unit: Unit, scale: number): Promise<RenderResult> {
  const width = Math.max(1, Math.ceil(unit.widthPx * scale))
  const height = Math.max(1, Math.ceil(unit.heightPx * scale))
  if (Math.max(width, height) > 8192) return { skipped: 'oversize' }
  const canvas = createCanvas(width, height)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.save()
  ctx.scale(canvas.width / unit.widthPx, canvas.height / unit.heightPx)
  unit.paint(ctx as never)
  ctx.restore()
  return { width: canvas.width, height: canvas.height, data: ctx.getImageData(0, 0, canvas.width, canvas.height).data }
}

/** Unused-but-typed helper guard so tree-shaking cannot claim a wrong export. */
export function corpusFileIsRegular(file: CorpusFile): boolean {
  try {
    return statSync(file.path).isFile()
  } catch {
    return false
  }
}
