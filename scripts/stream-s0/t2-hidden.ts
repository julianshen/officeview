/**
 * Stage-0 Task 5 — T2 hidden-content matrix runner.
 * Run under napi: `bun run s0:napi scripts/stream-s0/t2-hidden.ts`.
 */
import { loadOfficeFile } from '../../src/components/OfficeFile'
import { getPaintables } from '../../src/render/paint'
import { createCanvas } from 'canvas'
import { hiddenFixtures, type Fixture } from './hidden-fixtures'
import { loadUnits, renderUnit } from './corpus'
import { writeResult, resolveBackend } from './results'
import { diffBitmaps } from '../../src/test/pixel-diff'

function walkPaths(obj: unknown, marker: string, prefix = ''): string[] {
  const out: string[] = []
  if (typeof obj === 'string') {
    if (obj.includes(marker)) out.push(prefix || '$')
    return out
  }
  if (Array.isArray(obj)) {
    obj.forEach((v, i) => out.push(...walkPaths(v, marker, `${prefix}[${i}]`)))
    return out
  }
  if (obj && typeof obj === 'object') {
    for (const [k, v] of Object.entries(obj as Record<string, unknown>)) {
      out.push(...walkPaths(v, marker, prefix ? `${prefix}.${k}` : k))
    }
  }
  return out
}

async function renderAll(fileBytes: Uint8Array): Promise<{ count: number; pixels: Uint8ClampedArray[]; dims: { w: number; h: number }[] }> {
  // loadUnits expects a CorpusFile; synthesize one via a temp path is heavy.
  // Instead, parse directly and getPaintables, mirroring loadUnits.
  const doc = await loadOfficeFile(fileBytes)
  const paintables = await getPaintables(doc)
  const units = paintables.map((p: { spec: { widthPx: number; heightPx: number }; paint: (c: unknown) => void }, i: number) => ({ index: i, widthPx: p.spec.widthPx, heightPx: p.spec.heightPx, paint: (ctx: CanvasRenderingContext2D) => p.paint(ctx) }))
  const pixels: Uint8ClampedArray[] = []
  const dims: { w: number; h: number }[] = []
  for (const u of units) {
    const r = await (async () => {
      const w = Math.max(1, Math.ceil(u.widthPx))
      const h = Math.max(1, Math.ceil(u.heightPx))
      const canvas = createCanvas(w, h)
      const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
      ctx.save()
      ctx.scale(canvas.width / u.widthPx, canvas.height / u.heightPx)
      u.paint(ctx as never)
      ctx.restore()
      const img = ctx.getImageData(0, 0, canvas.width, canvas.height)
      return { width: canvas.width, height: canvas.height, data: img.data }
    })()
    pixels.push(r.data)
    dims.push({ w: r.width, h: r.height })
  }
  paintables.dispose()
  return { count: units.length, pixels, dims }
}

async function paintLeak(f: Fixture): Promise<boolean> {
  const a = await renderAll(f.withHidden)
  const b = await renderAll(f.withoutHidden)
  if (a.count !== b.count) return true // hidden unit painted
  for (let i = 0; i < a.count; i++) {
    if (a.dims[i].w !== b.dims[i].w || a.dims[i].h !== b.dims[i].h) return true
    const diff = diffBitmaps(
      { width: a.dims[i].w, height: a.dims[i].h, data: a.pixels[i] as unknown as Uint8ClampedArray },
      { width: b.dims[i].w, height: b.dims[i].h, data: b.pixels[i] as unknown as Uint8ClampedArray },
      8,
    )
    if (diff.changedPixels > 0) return true
  }
  return false
}

async function modelExposure(f: Fixture): Promise<string[]> {
  const doc = await loadOfficeFile(f.withHidden)
  const paths = walkPaths(doc, f.marker)
  // de-dupe + cap
  return [...new Set(paths)].slice(0, 50)
}

async function run(): Promise<number> {
  if (resolveBackend() !== '@napi-rs/canvas') {
    console.error('[t2] REFUSED: run under napi preload')
    return 2
  }
  const fixtures = await hiddenFixtures()
  const perFixture: Array<Record<string, unknown>> = []
  let leaks = 0
  for (const f of fixtures) {
    const leak = await paintLeak(f)
    if (leak) leaks++
    const exposure = await modelExposure(f)
    perFixture.push({ id: f.id, kind: f.kind, paintLeak: leak, modelExposure: exposure, exposureCount: exposure.length })
    console.log(`[t2] ${f.id}: paintLeak=${leak} exposure=${exposure.length}`)
  }
  writeResult('t2', {
    status: leaks === 0 ? 'GREEN' : 'RED',
    perFixture,
    totals: { fixtures: fixtures.length, leaks },
  })
  return 0
}

run().then(
  (code) => process.exit(code),
  (error) => { console.error(error); process.exit(1) },
)
