/**
 * Stage-0 Task 5 fix round 1 — T2 hidden-content matrix runner.
 * Leak = hidden marker appears in captured fillText/strokeText strings.
 * INVALID = control marker not drawn in both variants.
 */
import { loadOfficeFile } from '../../src/components/OfficeFile'
import { getPaintables } from '../../src/render/paint'
import { createCanvas } from 'canvas'
import { hiddenFixtures, type Fixture } from './hidden-fixtures'
import { writeResult, resolveBackend } from './results'

interface RenderCapture { units: number; texts: string[] }

async function renderAndCapture(bytes: Uint8Array): Promise<RenderCapture> {
  const doc = await loadOfficeFile(bytes)
  const paintables: any = await getPaintables(doc)
  const texts: string[] = []
  let units = 0
  for (const p of paintables) {
    units++
    const w = Math.max(1, Math.ceil(p.spec.widthPx))
    const h = Math.max(1, Math.ceil(p.spec.heightPx))
    const canvas = createCanvas(w, h)
    const ctx: any = canvas.getContext('2d')
    const proxy = new Proxy(ctx, {
      get(t, k) {
        const v = t[k]
        if (typeof v === 'function') return (...a: any[]) => {
          if ((k === 'fillText' || k === 'strokeText') && typeof a[0] === 'string') texts.push(a[0])
          return v.apply(t, a)
        }
        return v
      },
      set(t, k, v) { t[k] = v; return true },
    })
    p.paint(proxy)
  }
  paintables.dispose()
  return { units, texts }
}

async function run(): Promise<number> {
  if (resolveBackend() !== '@napi-rs/canvas') {
    console.error('[t2] REFUSED: run under napi preload')
    return 2
  }
  const fixtures = await hiddenFixtures()
  const perFixture: Array<Record<string, unknown>> = []
  let leaks = 0
  let invalid = 0
  for (const f of fixtures) {
    const a = await renderAndCapture(f.withHidden)
    const b = await renderAndCapture(f.withoutHidden)
    const withText = a.texts.join('|')
    const withoutText = b.texts.join('|')
    const controlDrawn = withText.includes(f.control) && withoutText.includes(f.control)
    // Runs may split a string across several fillText calls (e.g. a line
    // break inside a group shape's text box) — match on the concatenation too.
    const hiddenDrawn = withText.includes(f.marker) || a.texts.join('').includes(f.marker)
    let status: 'leak' | 'clean' | 'INVALID'
    if (!controlDrawn) { status = 'INVALID'; invalid++ }
    else if (hiddenDrawn) { status = 'leak'; leaks++ }
    else { status = 'clean' }
    // model exposure walk
    const doc: any = await loadOfficeFile(f.withHidden)
    const paths: string[] = []
    const walk = (obj: unknown, prefix = '') => {
      if (typeof obj === 'string') { if (obj.includes(f.marker)) paths.push(prefix || '$'); return }
      if (Array.isArray(obj)) { obj.forEach((v, i) => walk(v, `${prefix}[${i}]`)); return }
      if (obj && typeof obj === 'object') for (const [k, v] of Object.entries(obj as Record<string, unknown>)) walk(v, prefix ? `${prefix}.${k}` : k)
    }
    walk(doc)
    const modelExposure = [...new Set(paths)].slice(0, 50)
    perFixture.push({ id: f.id, kind: f.kind, status, controlDrawn, hiddenDrawn, unitsWith: a.units, unitsWithout: b.units, modelExposure, exposureCount: modelExposure.length })
    console.log(`[t2] ${f.id}: ${status} controlDrawn=${controlDrawn} hiddenDrawn=${hiddenDrawn} exp=${modelExposure.length}`)
  }
  writeResult('t2', {
    status: (leaks === 0 && invalid === 0) ? 'GREEN' : 'RED',
    perFixture,
    totals: { fixtures: fixtures.length, leaks, invalid, clean: fixtures.length - leaks - invalid },
  })
  return 0
}

if (import.meta.main) run().then(
  (code) => process.exit(code),
  (error) => { console.error(error); process.exit(1) },
)
